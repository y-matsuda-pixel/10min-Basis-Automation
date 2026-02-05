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
    """Google Driveから未処理のCSVをダウンロード"""
    query = f"'{SOURCE_FOLDER_ID}' in parents and name contains 'output' and mimeType = 'text/csv' and trashed = false"
    items = service.files().list(q=query, fields="files(id, name)").execute().get('files', [])
    downloaded = []
    for item in items:
        path = os.path.join(TEMP_DIR, item['name'])
        request = service.files().get_media(fileId=item['id'])
        with io.FileIO(path, 'wb') as fh:
            MediaIoBaseDownload(fh, request).next_chunk()
        downloaded.append({'id': item['id'], 'name': item['name'], 'local': path})
        logging.info(f"Downloaded: {item['name']}")
    return downloaded

def move_drive_file(service, file_id, new_name):
    """処理済みファイルをDrive上で移動"""
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

    # Selenium設定
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
        driver.get("https://www.basis-service.com/blas70/users/login")
        wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(BASIS_USERNAME)
        driver.find_element(By.NAME, "password").send_keys(BASIS_PASSWORD)
        driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//input[@type='submit']"))
        
        for f in files:
            # 物件名取得
            with open(f['local'], 'r', encoding='utf-8-sig') as csvf:
                row = next(csv.reader(csvf)) # Skip header
                row = next(csv.reader(csvf))
                prop_name = row[4] if len(row) > 4 else "不明"

            # 2. BLAS操作（インポート画面への遷移・アップロード）
            # ※ ここにサイドバークリックやCSVインポートのボタン操作を記述
            logging.info(f"Processing: {prop_name}")
            
            # (例) サイドバー -> Select2選択 -> ファイル送信 -> アラート承認
            # ※ 前回のスクリプトのdriver操作部分をここに集約
            
            # 3. 完了後の処理
            move_drive_file(service, f['id'], f"processed_{prop_name}_{datetime.datetime.now().strftime('%H%M%S')}.csv")
            send_lark(prop_name)

    except Exception as e:
        driver.save_screenshot('error.png')
        logging.error(f"Error: {e}")
        raise e
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
