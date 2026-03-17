# coding:utf-8
import os
import io
import csv
import json
import time
import logging
import datetime
import requests
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

def send_combined_lark_report(success_list, failure_list):
    """
    成功と失敗の結果を1つの「統合レポート」として送信する
    """
    if not LARK_WEBHOOK_URL: return
    if not success_list and not failure_list: return

    now_str = jst_now().strftime('%Y-%m-%d %H:%M:%S')
    
    # 成功物件のリスト作成
    success_text = "\n".join([f"✅ **成功:** {name}" for name in success_list]) if success_list else "なし"
    # 失敗物件のリスト作成
    failure_text = "\n".join([f"❌ **失敗:** {name} ({reason})" for name, reason in failure_list]) if failure_list else "なし"

    payload = {
        "msg_type": "interactive",
        "card": {
            "header": {
                "title": {"tag": "plain_text", "content": "🤖 BLAS一括登録 統合レポート"},
                "template": "green" if not failure_list else "orange"
            },
            "elements": [
                {
                    "tag": "div",
                    "text": {"tag": "lark_md", "content": f"**【成功物件】**\n{success_text}"}
                },
                {
                    "tag": "hr"
                },
                {
                    "tag": "div",
                    "text": {"tag": "lark_md", "content": f"**【失敗・要確認】**\n{failure_text}"}
                },
                {
                    "tag": "note",
                    "elements": [{"tag": "plain_text", "content": f"実行日時: {now_str}"}]
                }
            ]
        }
    }
    requests.post(LARK_WEBHOOK_URL, json=payload, timeout=10)

def main():
    service = get_drive_service()
    query = f"'{SOURCE_FOLDER_ID}' in parents and name contains 'output' and mimeType = 'text/csv' and trashed = false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get('files', [])
    
    if not files:
        logging.info("処理対象のCSVがありません。")
        return

    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
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
            display_name = f['name']
            path = os.path.join(TEMP_DIR, f['name'])
            try:
                # DriveからDL
                request = service.files().get_media(fileId=f['id'])
                with io.FileIO(path, 'wb') as fh:
                    MediaIoBaseDownload(fh, request).next_chunk()

                # 物件名の抽出 (失敗しても処理は続行)
                try:
                    with open(path, 'r', encoding='utf-8-sig') as csvf:
                        reader = csv.reader(csvf)
                        next(reader)
                        row = next(reader, None)
                        if row:
                            display_name = f"{row[4]} {row[5]}".strip()
                except: pass

                logging.info(f"処理中: {display_name}")

                # BLAS登録操作
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
                
                time.sleep(10) # 登録完了を待機

                # この時点で「成功」とみなしてリストに追加
                success_items.append(display_name)

                # 後処理：Driveでの移動 (ここで失敗してもsuccess_itemsには残る)
                try:
                    timestamp = jst_now().strftime('%H%M%S')
                    move_drive_file(service, f['id'], f"processed_{display_name}_{timestamp}.csv")
                except Exception as drive_err:
                    logging.warning(f"Drive移動失敗 ({display_name}): {drive_err}")

            except Exception as e:
                logging.error(f"❌ 処理失敗 ({display_name}): {e}")
                failure_items.append((display_name, str(e)))

        # ループ終了後にまとめて通知
        send_combined_lark_report(success_items, failure_items)

    finally:
        driver.quit()

if __name__ == "__main__":
    main()
