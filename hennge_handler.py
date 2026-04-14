# coding: utf-8
import os
import re
import json
import time
import logging
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def get_gmail_service(token_json_str):
    creds_info = json.loads(token_json_str)
    creds = Credentials.from_authorized_user_info(creds_info)
    return build('gmail', 'v1', credentials=creds)

def fetch_latest_msg(service, query):
    result = service.users().messages().list(userId='me', q=query, maxResults=1).execute()
    messages = result.get('messages', [])
    if not messages: return ""
    msg = service.users().messages().get(userId='me', id=messages[0]['id']).execute()
    return msg.get('snippet', '')

def run_hennge_download(driver, wait, token_json, target_email, temp_dir):
    service = get_gmail_service(token_json)
    
    # 1. URL取得
    logging.info("HENNGE URL取得中...")
    body = fetch_latest_msg(service, 'from:no-reply@hennge.com "download.transfer.hennge.com"')
    url_match = re.search(r'https://download\.transfer\.hennge\.com/[a-zA-Z0-9]+', body)
    if not url_match: raise Exception("HENNGE URLが見つかりません")
    
    # 2. ログイン操作
    driver.get(url_match.group(0))
    wait.until(EC.presence_of_element_located((By.XPATH, "//input[@placeholder='メールアドレスを入力']"))).send_keys(target_email)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='認証コードを送信']]"))).click()
    
    # 3. 認証コード取得
    logging.info("認証コード待機中...")
    auth_code = None
    for _ in range(10):
        time.sleep(10)
        msg = fetch_latest_msg(service, 'subject:"認証コード"')
        match = re.search(r'\d{6}', msg)
        if match:
            auth_code = match.group(0)
            break
    if not auth_code: raise Exception("認証コード取得失敗")
    
    driver.find_elements(By.XPATH, "//input[@type='tel' or @type='text']")[0].send_keys(auth_code)
    
    # 4. パスワード取得 & ダウンロード
    pw_msg = fetch_latest_msg(service, 'subject:"パスワードをお送りします"')
    pw = re.search(r'[A-Za-z0-9]{10,}', pw_msg).group(0)
    
    wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password']"))).send_keys(pw)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[span[text()='ダウンロード']]"))).click()
    
    logging.info("HENNGEダウンロード完了")
    time.sleep(10) # ファイル確定待ち
