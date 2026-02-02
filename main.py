import os
import io
import csv
import json
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

# --- 設定の読み込み (GitHub Secretsから取得) ---
BASIS_USERNAME = os.getenv('BASIS_USERNAME')
BASIS_PASSWORD = os.getenv('BASIS_PASSWORD')
LARK_WEBHOOK_URL = os.getenv('LARK_WEBHOOK_URL')
SOURCE_FOLDER_ID = os.getenv('SOURCE_FOLDER_ID')
DESTINATION_FOLDER_ID = os.getenv('DESTINATION_FOLDER_ID')
GDRIVE_JSON_STR = os.getenv('GDRIVE_JSON')

# ログの設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Google Drive 認証
if not GDRIVE_JSON_STR:
    logging.error("GDRIVE_JSON が設定されていません。GitHub Secretsを確認してください。")
    exit(1)

service_account_info = json.loads(GDRIVE_JSON_STR)
creds = service_account.Credentials.from_service_account_info(
    service_account_info, 
    scopes=['https://www.googleapis.com/auth/drive']
)
drive_service = build('drive', 'v3', credentials=creds)

def download_files_from_drive():
    """Google Driveから対象のCSVファイルをダウンロードする"""
    query = f"'{SOURCE_FOLDER_ID}' in parents and name contains 'output' and name contains '.csv' and trashed = false"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])
    downloaded_files = []
    
    if not os.path.exists('./temp'):
        os.makedirs('./temp')
        
    for item in items:
        file_id, file_name = item['id'], item['name']
        request = drive_service.files().get_media(fileId=file_id)
        path = f'./temp/{file_name}'
        with io.FileIO(path, 'wb') as fh:
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
        downloaded_files.append({'id': file_id, 'name': file_name, 'local_path': path})
    return downloaded_files

def move_drive_file(file_id, new_name):
    """処理済みファイルを別フォルダへ移動し、名前を変更する"""
    file = drive_service.files().get(fileId=file_id, fields='parents').execute()
    parents = file.get('parents')
    drive_service.files().update(
        fileId=file_id,
        addParents=DESTINATION_FOLDER_ID,
        removeParents=",".join(parents) if parents else "",
        body={'name': new_name},
        fields='id, parents'
    ).execute()

def send_lark(property_name):
    """Larkへ完了通知を送る"""
    if not LARK_WEBHOOK_URL:
        return
    payload = {
        "msg_type": "interactive",
        "card": {
            "header": {"title": {"tag": "lark_md", "content": "✅ BLAS登録完了"}, "template": "green"},
            "elements": [{"tag": "div", "text": {"tag": "lark_md", "content": f"物件名: **{property_name}** の登録が完了しました。"}}]
        }
    }
    requests.post(LARK_WEBHOOK_URL, json=payload)

def main():
    # 1. Google Driveからファイルを取得
    files = download_files_from_drive()
    if not files:
        logging.info("処理対象のファイルが見つかりませんでした。")
        return

    # 2. Seleniumブラウザの設定
    options = Options()
    options.add_argument("--headless=new")  # GitHub Actionsでは必須
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    
    driver_path = ChromeDriverManager().install()
    driver = webdriver.Chrome(service=Service(driver_path), options=options)
    wait = WebDriverWait(driver, 30)

    try:
        # 3. ログイン処理 (1回だけ実施)
        logging.info("ログインを開始します...")
        driver.get("https://www.basis-service.com/blas70/users/login")
        
        wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(BASIS_USERNAME)
        wait.until(EC.presence_of_element_located((By.NAME, "password"))).send_keys(BASIS_PASSWORD)
        driver.find_element(By.XPATH, "//input[@type='submit']").click()

        # ログイン成功の確認 (必要に応じて調整)
        logging.info("ログイン後の処理を開始します。")

        # 4. ファイルごとの処理
        for f in files:
            prop_name = "不明"
            with open(f['local_path'], 'r', encoding='utf-8-sig') as csvfile:
                reader = csv.reader(csvfile)
                next(reader) # ヘッダーをスキップ
                row = next(reader, None)
                if row and len(row) > 4:
                    prop_name = row[4]

            # Google Drive上での移動処理と通知
            timestamp = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
            new_name = f"{os.path.splitext(f['name'])[0]}_{timestamp}.csv"
            
            move_drive_file(f['id'], new_name)
            send_lark(prop_name)
            logging.info(f"成功: {prop_name}")

    except Exception as e:
        # ❌ 失敗時にスクリーンショットを保存 (重要！)
        driver.save_screenshot('error_screenshot.png')
        logging.error(f"エラーが発生しました: {e}")
        raise e  # GitHub Actionsに失敗を通知

    finally:
        driver.quit()

if __name__ == "__main__":
    main()
