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

def jst_now():
    return datetime.datetime.now(JST)

# ログの出力時刻も日本時間にするための設定
def logging_jst_converter(*args):
    return jst_now().timetuple()

logging.Formatter.converter = logging_jst_converter
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- GitHub Secretsから取得する設定 ---
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

def download_csvs(service):
    query = f"'{SOURCE_FOLDER_ID}' in parents and name contains 'output' and mimeType = 'text/csv' and trashed = false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])
    downloaded = []
    for item in items:
        path = os.path.join(TEMP_DIR, item['name'])
        request = service.files().get_media(fileId=item['id'])
        with io.FileIO(path, 'wb') as fh:
            MediaIoBaseDownload(fh, request).next_chunk()
        downloaded.append({'id': item['id'], 'name': item['name'], 'local': path})
    return downloaded

def move_drive_file(service, file_id, new_name):
    file = service.files().get(fileId=file_id, fields='parents').execute()
    previous_parents = ",".join(file.get('parents'))
    service.files().update(
        fileId=file_id, 
        addParents=DESTINATION_FOLDER_ID, 
        removeParents=previous_parents, 
        body={'name': new_name}
    ).execute()

def send_lark_success(display_name):
    """
    指定の画像デザインに基づいた通知送信（日本時間対応）
    """
    if not LARK_WEBHOOK_URL: return
    
    # 日本時間でのフォーマット
    now_str = jst_now().strftime('%Y-%m-%d %H:%M:%S')
    
    payload = {
        "msg_type": "interactive",
        "card": {
            "header": {
                "title": {"tag": "plain_text", "content": "🤖 Web自動化処理 SUCCESS"},
                "template": "green"
            },
            "elements": [{
                "tag": "div",
                "text": {
                    "tag": "lark_md",
                    "content": f"**ステータス:** ✅ SUCCESS\n**詳細:** レジル復旧作業 「{display_name}」 BLASの登録が完了しました\n**実行日時:** {now_str}"
                }
            }]
        }
    }
    try:
        response = requests.post(LARK_WEBHOOK_URL, json=payload, timeout=10)
        response.raise_for_status()
    except Exception as e:
        logging.error(f"Lark通知の送信に失敗しました: {e}")

def main():
    if not GDRIVE_JSON or GDRIVE_JSON == '{}':
        logging.error("Google Driveの認証情報が設定されていません。")
        return
        
    service = get_drive_service()
    files = download_csvs(service)
    
    if not files:
        logging.info("処理対象のCSVがありません。")
        return

    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 30)

    try:
        # ログイン処理
        logging.info("BLASにログイン中...")
        driver.get("https://www.basis-service.com/blas70/users/login")
        wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(BASIS_USERNAME)
        driver.find_element(By.NAME, "password").send_keys(BASIS_PASSWORD)
        driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//input[@type='submit']"))
        time.sleep(5)

        for f in files:
            display_name = "不明"
            try:
                # CSVから物件名(5列目)と部屋番号(6列目)を抽出
                with open(f['local'], 'r', encoding='utf-8-sig') as csvf:
                    reader = csv.reader(csvf)
                    next(reader) # ヘッダーをスキップ
                    row = next(reader, None)
                    if row:
                        prop_name = row[4] # 物件名
                        room_num = row[5]  # 部屋番号
                        display_name = f"{prop_name} {room_num}".strip()

                logging.info(f"処理開始: {display_name}")

                # BLAS操作: メニュー遷移
                wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/ul/li[5]/a"))).click()
                time.sleep(3)
                
                # 業務選択
                wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "select2-selection__arrow"))).click()
                search_field = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "select2-search__field")))
                search_field.send_keys("【レジル】停止・復電業務")
                time.sleep(2)
                wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), '【レジル】停止・復電業務')]"))).click()
                
                # CSVインポート
                wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'CSVインポート')]"))).click()
                chk = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='radio' and @value='1']")))
                driver.execute_script("arguments[0].click();", chk)
                driver.find_element(By.XPATH, "//input[@type='file']").send_keys(os.path.abspath(f['local']))
                wait.until(EC.element_to_be_clickable((By.ID, "csv_import_btn"))).click()
                
                # アラート対応
                try:
                    wait.until(EC.alert_is_present())
                    driver.switch_to.alert.accept()
                except:
                    pass
                
                # 処理完了待ち
                time.sleep(10)

                # 成功時の後処理（Google Drive移動 & Lark通知）
                timestamp = jst_now().strftime('%H%M%S')
                new_file_name = f"processed_{display_name}_{timestamp}.csv"
                move_drive_file(service, f['id'], new_file_name)
                
                send_lark_success(display_name)
                logging.info(f"✅ Success: {display_name} (Saved as: {new_file_name})")

            except Exception as e:
                logging.error(f"❌ Error in {display_name}: {e}")
                driver.save_screenshot(f'error_{display_name}_{jst_now().strftime("%H%M%S")}.png')

    finally:
        driver.quit()

if __name__ == "__main__":
    main()
