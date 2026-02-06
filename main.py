#coding:utf-8
import os
import io
import csv
import json
import glob
import time
import logging
import datetime
import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# --- 設定（GitHub Secretsから取得） ---
BASIS_USERNAME = os.getenv('BASIS_USERNAME')
BASIS_PASSWORD = os.getenv('BASIS_PASSWORD')
LARK_WEBHOOK_URL = os.getenv('LARK_WEBHOOK_URL')
SOURCE_FOLDER_ID = os.getenv('SOURCE_FOLDER_ID')
DESTINATION_FOLDER_ID = os.getenv('DESTINATION_FOLDER_ID')
GDRIVE_JSON = os.getenv('GDRIVE_JSON')

TEMP_DIR = './temp'
if not os.path.exists(TEMP_DIR): os.makedirs(TEMP_DIR)

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def get_drive_service():
    creds_dict = json.loads(GDRIVE_JSON)
    creds = service_account.Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/drive'])
    return build('drive', 'v3', credentials=creds)

def download_csvs(service):
    query = f"'{SOURCE_FOLDER_ID}' in parents and name contains 'output' and mimeType = 'text/csv' and trashed = false"
    items = service.files().list(q=query, fields="files(id, name)").execute().get('files', [])
    downloaded = []
    for item in items:
        path = os.path.abspath(os.path.join(TEMP_DIR, item['name']))
        request = service.files().get_media(fileId=item['id'])
        with io.FileIO(path, 'wb') as fh:
            MediaIoBaseDownload(fh, request).next_chunk()
        downloaded.append({'id': item['id'], 'name': item['name'], 'local': path})
        logging.info(f"Downloaded: {item['name']}")
    return downloaded

def move_drive_file(service, file_id, new_name):
    file = service.files().get(fileId=file_id, fields='parents').execute()
    previous_parents = ",".join(file.get('parents'))
    service.files().update(fileId=file_id, addParents=DESTINATION_FOLDER_ID, removeParents=previous_parents, body={'name': new_name}).execute()

def send_lark(property_name):
    if not LARK_WEBHOOK_URL: return
    payload = {
        "msg_type": "interactive",
        "card": {
            "header": {"title": {"tag": "plain_text", "content": "✅ BLAS登録完了"}, "template": "green"},
            "elements": [{"tag": "div", "text": {"tag": "lark_md", "content": f"物件名: **{property_name}** の登録が完了しました。"}}]
        }
    }
    requests.post(LARK_WEBHOOK_URL, json=payload)

def main():
    if not GDRIVE_JSON:
        logging.error("GDRIVE_JSON is missing!")
        return

    service = get_drive_service()
    files = download_csvs(service)
    if not files:
        logging.info("No files to process.")
        return

    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 30)

    try:
        # 1. ログイン
        logging.info("Logging into BLAS...")
        driver.get("https://www.basis-service.com/blas70/users/login")
        wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(BASIS_USERNAME)
        driver.find_element(By.NAME, "password").send_keys(BASIS_PASSWORD)
        driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//input[@type='submit']"))
        
        # ログイン後の安定待ち
        time.sleep(5)

        for f in files:
            # 物件名取得
            with open(f['local'], 'r', encoding='utf-8-sig') as csvf:
                reader = csv.reader(csvf)
                next(reader) # ヘッダー
                row = next(reader, None)
                prop_name = row[4] if row and len(row) > 4 else "不明"

            logging.info(f"Processing: {prop_name}")

            # 2. BLAS操作（画面遷移とアップロード）
            # サイドバーをクリック
            sidebar_xpath = "/html/body/div[1]/div/div[1]/ul/li[5]/a"
            wait.until(EC.element_to_be_clickable((By.XPATH, sidebar_xpath))).click()
            time.sleep(3)

            # ドロップダウン選択（Select2）
            wait.until(EC.element_to_be_clickable((By.CLASS_NAME, "select2-selection__arrow"))).click()
            search_field = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "select2-search__field")))
            search_field.send_keys("【レジル】停止・復電業務")
            time.sleep(2)
            wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(text(), '【レジル】停止・復電業務')]"))).click()

            # CSVインポートボタン
            wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'CSVインポート')]"))).click()
            
            # ラジオボタン選択（JSで確実に）
            chk = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='radio' and @value='1']")))
            driver.execute_script("arguments[0].click();", chk)

            # ファイル送信
            driver.find_element(By.XPATH, "//input[@type='file']").send_keys(f['local'])
            
            # インポート実行
            wait.until(EC.element_to_be_clickable((By.ID, "csv_import_btn"))).click()

            # アラート承認
            try:
                wait.until(EC.alert_is_present())
                driver.switch_to.alert.accept()
                logging.info("Alert accepted.")
            except:
                logging.warning("No alert appeared.")

            # アップロード完了待ち
            time.sleep(10)
            
            # 3. 完了後の処理
            timestamp = datetime.datetime.now().strftime('%H%M%S')
            move_drive_file(service, f['id'], f"processed_{prop_name}_{timestamp}.csv")
            send_lark(prop_name)
            logging.info(f"Successfully processed: {prop_name}")

    except Exception as e:
        driver.save_screenshot('error_screenshot.png')
        logging.error(f"Error: {e}")
        raise e
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
