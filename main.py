# coding:utf-8
import os
import io
import csv
import json
import time
import logging
import datetime
import requests
import re
from datetime import timezone, timedelta

from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.http import MediaIoBaseDownload
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- 日本時間(JST)の設定 ---
JST = timezone(timedelta(hours=9))
def jst_now(): return datetime.datetime.now(JST)

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- GitHub Secrets等からの設定 ---
GDRIVE_JSON = os.environ.get('GDRIVE_JSON', '{}')
SOURCE_FOLDER_ID = os.environ.get('SOURCE_FOLDER_ID', '')
DESTINATION_FOLDER_ID = os.environ.get('DESTINATION_FOLDER_ID', '')
BASIS_USERNAME = os.environ.get('BASIS_USERNAME', '')
BASIS_PASSWORD = os.environ.get('BASIS_PASSWORD', '')
LARK_WEBHOOK_URL = os.environ.get('LARK_WEBHOOK_URL', '')

TEMP_DIR = './temp'
if not os.path.exists(TEMP_DIR): os.makedirs(TEMP_DIR)

# --- エリア判定用の都道府県リスト ---
KANTO_PREFS = ['東京都', '神奈川県', '埼玉県', '千葉県', '茨城県', '栃木県', '群馬県']
KANSAI_PREFS = ['大阪府', '京都府', '兵庫県', '奈良県', '滋賀県', '和歌山県']

def get_drive_service():
    creds_dict = json.loads(GDRIVE_JSON)
    creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/drive'])
    return build('drive', 'v3', credentials=creds)

def move_drive_file(service, file_id, new_name):
    file = service.files().get(fileId=file_id, fields='parents').execute()
    previous_parents = ",".join(file.get('parents'))
    service.files().update(
        fileId=file_id, 
        addParents=DESTINATION_FOLDER_ID, 
        removeParents=previous_parents, 
        body={'name': new_name}
    ).execute()

def get_region_from_address(address):
    """
    住所から都道府県を抽出し、関東・関西を判定する
    """
    if not address:
        return "不明"
    
    # 住所から都道府県を抽出 (例: 大阪府)
    match = re.search(r'(...??[都道府県])', address)
    if not match:
        return "不明"
    
    prefecture = match.group(1)
    if prefecture in KANTO_PREFS:
        return "関東"
    elif prefecture in KANSAI_PREFS:
        return "関西"
    else:
        return f"その他（{prefecture}）"

def send_combined_lark_report(success_list, failure_list):
    """
    成功と失敗の結果をLarkに送信。地域（エリア）情報も含める。
    """
    if not LARK_WEBHOOK_URL: return
    if not success_list and not failure_list: return

    now_str = jst_now().strftime('%Y-%m-%d %H:%M:%S')
    elements = []

    # --- 成功物件のブロック作成 ---
    for item in success_list:
        elements.append({
            "tag": "div",
            "text": {
                "tag": "lark_md",
                "content": (
                    f"**ステータス:** ✅ SUCCESS\n"
                    f"**詳細:** レジル復旧作業 「{item['name']}」 BLASの登録が完了しました\n"
                    f"**地域:** {item['region']}\n"
                    f"**スイッチ:** {item['switch']}\n"
                    f"**実行日時:** {now_str}"
                )
            }
        })

    # --- 失敗物件がある場合 ---
    if failure_list:
        if success_list: elements.append({"tag": "hr"})
        for name, reason in failure_list:
            elements.append({
                "tag": "div",
                "text": {
                    "tag": "lark_md",
                    "content": (
                        f"**ステータス:** ❌ FAILURE\n"
                        f"**詳細:** {name} の登録に失敗しました\n"
                        f"**理由:** {reason}\n"
                        f"**実行日時:** {now_str}"
                    )
                }
            })

    payload = {
        "msg_type": "interactive",
        "card": {
            "header": {
                "title": {"tag": "plain_text", "content": "🤖 Web自動化処理 SUCCESS" if not failure_list else "⚠️ Web自動化処理 REPORT"},
                "template": "green" if not failure_list else "orange"
            },
            "elements": elements
        }
    }
    
    try:
        response = requests.post(LARK_WEBHOOK_URL, json=payload, timeout=10)
        response.raise_for_status()
    except Exception as e:
        logging.error(f"Lark通知送信エラー: {e}")

def main():
    service = get_drive_service()
    # 処理対象のCSVを検索
    query = f"'{SOURCE_FOLDER_ID}' in parents and name contains 'output' and mimeType = 'text/csv' and trashed = false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get('files', [])
    
    if not files:
        logging.info("処理対象のCSVがありません。")
        return

    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 30)

    success_items = []
    failure_items = []

    try:
        logging.info("BLASにログイン中...")
        driver.get("https://www.basis-service.com/blas70/users/login")
        wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(BASIS_USERNAME)
        driver.find_element(By.NAME, "password").send_keys(BASIS_PASSWORD)
        driver.find_element(By.XPATH, "//input[@type='submit']").click()
        time.sleep(5)

        for f in files:
            path = os.path.join(TEMP_DIR, f['name'])
            display_name = f['name']
            switch_val = "不明"
            region_val = "不明"

            try:
                # 1. Google Driveからダウンロード
                request = service.files().get_media(fileId=f['id'])
                with io.FileIO(path, 'wb') as fh:
                    downloader = MediaIoBaseDownload(fh, request)
                    done = False
                    while not done:
                        status, done = downloader.next_chunk()

                # 2. CSVから物件名、スイッチ(17列目)、地域(7列目の住所から判定)を抽出
                try:
                    with open(path, 'r', encoding='utf-8-sig') as csvf:
                        reader = csv.reader(csvf)
                        next(reader) # ヘッダースキップ
                        row = next(reader, None)
                        if row:
                            # 物件名 + 部屋番号
                            display_name = f"{row[4]} {row[5]}".strip()
                            # 地域判定 (Index 6: 物件住所)
                            region_val = get_region_from_address(row[6] if len(row) > 6 else "")
                            # スイッチ判定 (Index 16: スキルレススイッチの有無)
                            switch_val = row[16] if len(row) > 16 and row[16] else "あり"
                except Exception as csv_err:
                    logging.warning(f"CSV読み取りエラー: {csv_err}")

                logging.info(f"処理開始: {display_name} (地域: {region_val}, スイッチ: {switch_val})")

                # 3. BLAS登録操作
                driver.get("https://www.basis-service.com/blas70/items")
                wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "select2-selection__arrow"))).click()
                search_field = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "select2-search__field")))
                search_field.send_keys("【レジル】停止・復電業務")
                time.sleep(2)
                wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), '【レジル】停止・復電業務')]"))).click()
                
                wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'CSVインポート')]"))).click()
                chk = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='radio' and @value='1']")))
                driver.execute_script("arguments[0].click();", chk)
                driver.find_element(By.XPATH, "//input[@type='file']").send_keys(os.path.abspath(path))
                wait.until(EC.element_to_be_clickable((By.ID, "csv_import_btn"))).click()
                
                try:
                    WebDriverWait(driver, 5).until(EC.alert_is_present())
                    driver.switch_to.alert.accept()
                except: pass
                
                time.sleep(10) # 登録完了待機

                # 成功リストに情報を格納
                success_items.append({
                    "name": display_name, 
                    "switch": switch_val, 
                    "region": region_val
                })

                # 処理済みフォルダへ移動
                timestamp = jst_now().strftime('%H%M%S')
                move_drive_file(service, f['id'], f"processed_{display_name}_{timestamp}.csv")

            except Exception as e:
                logging.error(f"❌ 処理失敗 ({display_name}): {e}")
                failure_items.append((display_name, str(e)))

        # 4. レポート送信
        send_combined_lark_report(success_items, failure_items)

    finally:
        driver.quit()

if __name__ == "__main__":
    main()
