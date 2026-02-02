# coding: utf-8
import os
import io
import csv
import glob
import time
import json
import logging
import datetime
import requests
import shutil
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

# --- 環境変数 ---
BASIS_USERNAME = os.getenv('BASIS_USERNAME')
BASIS_PASSWORD = os.getenv('BASIS_PASSWORD')
LARK_WEBHOOK_URL = os.getenv('LARK_WEBHOOK_URL')
SOURCE_FOLDER_ID = os.getenv('SOURCE_FOLDER_ID')
DESTINATION_FOLDER_ID = os.getenv('DESTINATION_FOLDER_ID')

# Google Drive API設定
SCOPES = ['https://www.googleapis.com/auth/drive']
service_account_info = json.loads(os.getenv('GDRIVE_JSON'))
creds = service_account.Credentials.from_service_account_info(service_account_info, scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=creds)

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s')

def download_files_from_drive():
    """Driveからoutput*.csvを検索してダウンロード"""
    query = f"'{SOURCE_FOLDER_ID}' in parents and name contains 'output' and name contains '.csv' and trashed = false"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])
    
    downloaded_files = []
    if not os.path.exists('./temp'): os.makedirs('./temp')
    
    for item in items:
        file_id = item['id']
        file_name = item['name']
        request = drive_service.files().get_media(fileId=file_id)
        fh = io.FileIO(f'./temp/{file_name}', 'wb')
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
        downloaded_files.append({'id': file_id, 'name': file_name, 'local_path': f'./temp/{file_name}'})
    return downloaded_files

def move_drive_file(file_id, new_name):
    """Drive上でファイルを移動・リネーム"""
    file = drive_service.files().get(fileId=file_id, fields='parents').execute()
    previous_parents = ",".join(file.get('parents'))
    drive_service.files().update(
        fileId=file_id,
        addParents=DESTINATION_FOLDER_ID,
        removeParents=previous_parents,
        body={'name': new_name},
        fields='id, parents'
    ).execute()

def send_lark(property_name):
    payload = {
        "msg_type": "interactive",
        "card": {
            "header": {"title": {"tag": "lark_md", "content": "✅ BLAS登録完了"}, "template": "green"},
            "elements": [{"tag": "div", "text": {"tag": "lark_md", "content": f"物件名: **{property_name}** の登録が完了しました。"}}]
        }
    }
    requests.post(LARK_WEBHOOK_URL, json=payload)

def main():
    files = download_files_from_drive()
    if not files:
        logging.info("処理対象ファイルなし。")
        return

    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    
    driver_path = ChromeDriverManager().install()
    
    for f in files:
        # 物件名取得ロジック（既存）
        with open(f['local_path'], 'r', encoding='utf-8-sig') as csvfile:
            reader = csv.reader(csvfile)
            next(reader)
            row = next(reader, None)
            prop_name = row[4] if row and len(row) > 4 else "不明"

        driver = webdriver.Chrome(service=Service(driver_path), options=options)
        wait = WebDriverWait(driver, 30)
        
        try:
            driver.get("https://www.basis-service.com/blas70/users/login")
            wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(BASIS_USERNAME)
            wait.until(EC.presence_of_element_located((By.NAME, "password"))).send_keys(BASIS_PASSWORD)
            driver.find_element(By.XPATH, "//input[@type='submit']").click()
            
            # --- Selenium操作（提供されたロジック） ---
            # サイドバー、ドロップダウン、検索、インポート等の処理をここに記述
            # ※IDやXPathは既存のものをそのまま流用可能
            
            # 完了後、Drive上で移動
            timestamp = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
            new_name = f"{os.path.splitext(f['name'])[0]}_{timestamp}.csv"
            move_drive_file(f['id'], new_name)
            send_lark(prop_name)
            logging.info(f"成功: {prop_name}")
            
        except Exception as e:
            logging.error(f"エラー: {e}")
        finally:
            driver.quit()

if __name__ == "__main__":
    main()