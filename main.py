import os
import io
import csv
import json
import logging
import datetime
import requests
import time
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

# --- 設定の読み込み ---
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
    logging.error("❌ GDRIVE_JSON が設定されていません。")
    exit(1)

service_account_info = json.loads(GDRIVE_JSON_STR)
creds = service_account.Credentials.from_service_account_info(
    service_account_info, 
    scopes=['https://www.googleapis.com/auth/drive']
)
drive_service = build('drive', 'v3', credentials=creds)

def download_files_from_drive():
    """Google Driveから対象のCSVファイルをダウンロードする"""
    logging.info(f"📂 フォルダID: {SOURCE_FOLDER_ID} 内を探索中...")
    
    # ファイル名に 'output' を含み、削除されていないCSVを検索
    query = f"'{SOURCE_FOLDER_ID}' in parents and name contains 'output' and mimeType = 'text/csv' and trashed = false"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])
    
    downloaded_files = []
    if not os.path.exists('./temp'):
        os.makedirs('./temp')
        
    for item in items:
        file_id, file_name = item['id'], item['name']
        logging.info(f"📥 ファイル発見: {file_name} (ID: {file_id})")
        
        request = drive_service.files().get_media(fileId=file_id)
        path = f'./temp/{file_name}'
        with io.FileIO(path, 'wb') as fh:
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
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
    logging.info(f"✅ ファイルを移動しました: {new_name}")

def send_lark(property_name):
    """Lark Webhookへ完了通知を送る"""
    if not LARK_WEBHOOK_URL:
        return
    payload = {
        "msg_type": "interactive",
        "card": {
            "header": {"title": {"tag": "lark_md", "content": "✅ BLAS登録完了"}, "template": "green"},
            "elements": [{"tag": "div", "text": {"tag": "lark_md", "content": f"物件名: **{property_name}**\nの自動登録が正常に完了しました。"}}]
        }
    }
    requests.post(LARK_WEBHOOK_URL, json=payload)

def setup_driver():
    """GitHub Actions環境に最適化されたWebDriverの設定"""
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    # ボット検知回避用
    options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

def main():
    # 1. Google Driveからファイルを取得
    files = download_files_from_drive()
    if not files:
        logging.info("📢 処理対象のCSVファイルが見つかりませんでした。終了します。")
        return

    driver = setup_driver()
    wait = WebDriverWait(driver, 30)

    try:
        # 2. ログイン処理
        logging.info("🌐 BLASログイン画面へアクセス...")
        driver.get("https://www.basis-service.com/blas70/users/login")
        
        # 要素が表示されるまで待機
        wait.until(EC.presence_of_element_located((By.NAME, "username"))).send_keys(BASIS_USERNAME)
        driver.find_element(By.NAME, "password").send_keys(BASIS_PASSWORD)
        
        # 送信ボタンをJavaScriptで確実にクリック
        submit_btn = driver.find_element(By.XPATH, "//input[@type='submit']")
        driver.execute_script("arguments[0].click();", submit_btn)
        
        logging.info("🔑 ログイン試行中...")
        time.sleep(3) # 遷移待ち

        # 3. CSVごとに登録処理（ループ）
        for f in files:
            prop_name = "不明"
            # CSVの読み込み (UTF-8 with BOMに対応)
            with open(f['local_path'], 'r', encoding='utf-8-sig') as csvfile:
                reader = csv.reader(csvfile)
                header = next(reader) # ヘッダー
                row = next(reader, None) # 1行目のデータ
                if row and len(row) > 4:
                    prop_name = row[4] # 物件名を取得

            logging.info(f"🛠 物件: {prop_name} の登録を開始します...")

            # --- ここに具体的な入力処理（driver.find_element...）を記述 ---
            # 例: 登録ページへ移動、フォームに入力、保存など
            # --------------------------------------------------------

            # 4. 後処理: ファイル移動と通知
            timestamp = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
            new_name = f"processed_{prop_name}_{timestamp}.csv"
            
            move_drive_file(f['id'], new_name)
            send_lark(prop_name)
            logging.info(f"✨ 完了通知を送信しました: {prop_name}")

    except Exception as e:
        # ❌ 失敗時にスクリーンショットを保存
        driver.save_screenshot('error_screenshot.png')
        logging.error(f"🚨 エラーが発生しました: {e}")
        # GitHub ActionsのArtifacts用にファイルパスをログに出す
        logging.info("エラー時のスクリーンショットを保存しました。")
        raise e 

    finally:
        driver.quit()
        logging.info("🏁 ブラウザを閉じました。")

if __name__ == "__main__":
    main()
