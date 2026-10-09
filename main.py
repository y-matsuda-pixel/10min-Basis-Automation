import sys
print("==========================================", flush=True)
print("=== PERFECT_CODE_VERSION_V18_HYBRID_FINAL ===", flush=True)
print("==========================================", flush=True)

# coding: utf-8

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import shutil
import datetime
from datetime import timezone, timedelta
import logging
import requests
from pathlib import Path
import re
import glob
import base64
import zipfile

# PDF解析用ライブラリ
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer

# Gmail API用ライブラリ
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# --- 日本時間(JST)の設定 ---
JST = timezone(timedelta(hours=9))
def jst_now(): return datetime.datetime.now(JST)

# --- 設定情報 ---
BASIS_USERNAME = os.getenv('BASIS_USERNAME', '')
BASIS_PASSWORD = os.getenv('BASIS_PASSWORD', '')
LARK_WEBHOOK_URL = os.getenv('LARK_WEBHOOK_URL', "https://open.larksuite.com/open-apis/bot/v2/hook/3827abd3-eb1b-41b3-8df3-e920884d2d30")

# Lark API 認証情報
LARK_APP_ID = os.getenv('LARK_APP_ID', "cli_aac716ae45b81e14")
LARK_APP_SECRET = os.getenv('LARK_APP_SECRET', '')
LARK_WIKI_TOKEN = "Ykm6w1c70iDJ0kkmM68jPcf5pEf"
DEFAULT_SHEET_ID = "Yb6Zwo"
EMAIL_SHEET_ID = "KTGtTb"

FIXED_COMPANY_NAME = "ベイシス株式会社 IoT推進部"

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_CSV_PATH = BASE_DIR / "output.csv"
PROCESSED_DIR = BASE_DIR / "processed"
PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
DOWNLOAD_DIR = BASE_DIR / "downloads"
DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)

KANTO_PREFS = ['東京都', '神奈川県', '埼玉県', '千葉県', '茨城県', '栃木県', '群馬県', '東京', '神奈川', '埼玉', '千葉', '茨城', '栃木', '群馬']
KANSAI_PREFS = ['大阪府', '京都府', '兵庫県', '奈良県', '滋賀県', '和歌山県', '大阪', '京都', '兵庫', '奈良', '滋賀', '和歌山']

KANSAI_CITIES = ['大阪市', '京都市', '神戸市', '堺市', '奈良市', '和歌山市', '大津市', '東大阪市', '西宮市', '尼崎市', '豊中市', '吹田市', '枚方市', '高槻市', '茨木市', '八尾市', '寝屋川市', '姫路市', '明石市', '加古川市', '宝塚市', '伊丹市', '川西市', '精華', '木津川', '宇治市', '城陽市', '生駒市']
KANTO_CITIES = ['横浜市', '川崎市', 'さいたま市', '千葉市', '相模原市', '船橋市', '川口市', '新宿区', '世田谷区', '港区', '渋谷区', '中央区', '千代田区', '品川区', '目黒区', '大田区', '杉並区', '練馬区', '八王子市', '町田市', '藤沢市', '横須賀市', '平塚市', '茅ヶ崎市', '大和市', '厚木市', '所沢市', '川越市', '越谷市', '草加市', '市川市', '松戸市', '柏市', '市原市', '宇都宮市', '前橋市', '高崎市', '水戸市', '江戸川区']

TEST_DOWNLOAD_ONLY = False
TEST_CSV_ONLY = False

LOG_DIR = BASE_DIR / "log"
LOG_DIR.mkdir(parents=True, exist_ok=True)
log_filename = f"selenium_log_{datetime.datetime.now().strftime('%Y%m%d')}.log"
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_DIR / log_filename, encoding='utf-8'), 
        logging.StreamHandler(sys.stdout)
    ]
)

SCOPES = ['https://www.googleapis.com/auth/gmail.modify']

COLUMN_ALIASES = {
    'MID': ['物件id', 'mid', 'ｍｉｄ', 'id'],
    '物件名': ['物件名', '件名', '物件', '建物名', '対象物件'],
    '部屋番号': ['部屋番号', '号', '部番号', '番号', '部屋'],
    '住所': ['物件住所', '住所', '所在地', '設置場所'],
    '管理人': ['管理員様在籍日時', '管理人駐在時間', '管理員様', '在籍日時', '管理人', '管理員', '管理'],
    'AL': ['al', 'オートロック', '有/解除番号有/無', '有/解除番号有', 'オートロックの有無'],
    '備考': ['備考', 'メモ', '注意事項']
}

def log_flush(msg, level=logging.INFO):
    logging.log(level, msg)
    sys.stdout.flush()

def get_chrome_options():
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36")
    prefs = {
        "download.default_directory": str(DOWNLOAD_DIR),
        "download.prompt_for_download": False,
        "directory_upgrade": True,
        "safebrowsing.enabled": True,
        "safebrowsing.disable_download_protection": True
    }
    options.add_experimental_option("prefs", prefs)
    return options

def identify_column_key(text):
    txt_clean = text.lower().replace(' ', '').replace(' ', '').replace('\n', '')
    for std_key, aliases in COLUMN_ALIASES.items():
        if any(alias in txt_clean for alias in aliases):
            return std_key
    return 'IGNORE'

def is_address_text_strict(txt):
    txt_clean = re.sub(r'^\s*0\s*', '', txt.strip())
    all_prefs = KANTO_PREFS + KANSAI_PREFS
    if any(p in txt_clean for p in all_prefs):
        return True
    if any(c in txt_clean for c in KANTO_CITIES + KANSAI_CITIES):
        return True
    if re.search(r'[一-龠]{2,3}[都道府県]', txt_clean) or re.search(r'\d+丁目|\d+番地|\d+-\d+', txt_clean):
        if not re.match(r'^\d+$', txt_clean) and not re.match(r'^[A-Za-z0-9-]+$', txt_clean):
            return True
    return False

# --- 住所優先で「関東」または「関西」判定 ---
def get_region_from_info(filename, address):
    if address:
        addr_clean = address.strip()
        if any(pref in addr_clean for pref in KANTO_PREFS) or any(city in addr_clean for city in KANTO_CITIES):
            return "関東"
        elif any(pref in addr_clean for pref in KANSAI_PREFS) or any(city in addr_clean for city in KANSAI_CITIES):
            return "関西"

    filename_lower = filename.lower() if filename else ""
    if "関東" in filename_lower or "東日本" in filename_lower:
        return "関東"
    elif "関西" in filename_lower or "西日本" in filename_lower:
        return "関西"

    return "関東"

def get_lark_tenant_access_token():
    url = "https://open.larksuite.com/open-apis/auth/v3/tenant_access_token/internal"
    headers = {"Content-Type": "application/json; charset=utf-8"}
    body = {"app_id": LARK_APP_ID, "app_secret": LARK_APP_SECRET}
    try:
        res = requests.post(url, headers=headers, json=body, timeout=10)
        res_json = res.json()
        if res_json.get("code") == 0:
            return res_json.get("tenant_access_token")
        log_flush(f"Lark API Token取得失敗: {res_json}", logging.ERROR)
    except Exception as e:
        log_flush(f"Lark API接続エラー: {e}", logging.ERROR)
    return None

def get_spreadsheet_token(tenant_token):
    url = f"https://open.larksuite.com/open-apis/wiki/v2/nodes/get_node?token={LARK_WIKI_TOKEN}"
    headers = {"Authorization": f"Bearer {tenant_token}"}
    try:
        res = requests.get(url, headers=headers, timeout=10)
        res_json = res.json()
        if res_json.get("code") == 0:
            obj_token = res_json.get("data", {}).get("node", {}).get("obj_token")
            if obj_token:
                log_flush(f"🔑 WikiトークンからSpreadsheetトークンを取得成功: {obj_token}")
                return obj_token
        log_flush(f"Wikiトークン変換失敗: {res_json}", logging.WARNING)
    except Exception as e:
        log_flush(f"Wikiトークン変換例外: {e}", logging.ERROR)
    return LARK_WIKI_TOKEN

def write_email_to_lark_sheet(email_info):
    if not email_info: return
    tenant_token = get_lark_tenant_access_token()
    if not tenant_token:
        log_flush("Lark APIトークンが取得できなかったため、メール転記をスキップします。", logging.WARNING)
        return

    spreadsheet_token = get_spreadsheet_token(tenant_token)
    headers = {
        "Authorization": f"Bearer {tenant_token}",
        "Content-Type": "application/json; charset=utf-8"
    }

    sheet_id = EMAIL_SHEET_ID
    try:
        sheets_url = f"https://open.larksuite.com/open-apis/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/query"
        res_sheets = requests.get(sheets_url, headers=headers, timeout=10)
        res_data = res_sheets.json()
        if res_sheets.status_code == 200 and res_data.get("code") == 0:
            sheets_list = res_data.get("data", {}).get("sheets", [])
            for s in sheets_list:
                if "復旧依頼メール" in s.get("title", ""):
                    sheet_id = s.get("sheet_id")
                    break
    except Exception as e:
        log_flush(f"メール用シートタブ検索エラー: {e}", logging.WARNING)

    read_url = f"https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values/{sheet_id}!A1:A500"
    target_row = 2
    try:
        res_read = requests.get(read_url, headers=headers, timeout=10)
        res_read_data = res_read.json()
        if res_read.status_code == 200 and res_read_data.get("code") == 0:
            values = res_read_data.get("data", {}).get("valueRange", {}).get("values", [])
            last_filled_row = 1
            for idx, row_val in enumerate(values):
                row_num = idx + 1
                if row_num >= 2 and row_val:
                    a_val = row_val[0] if len(row_val) > 0 else None
                    a_str = str(a_val).strip() if a_val is not None else ""
                    if a_str and a_str.lower() != 'none':
                        last_filled_row = row_num
            target_row = max(last_filled_row + 1, 2)
    except Exception as e:
        log_flush(f"メール用空行判定エラー: {e}", logging.WARNING)

    write_url = f"https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values"
    row_data = [
        email_info.get('date_str', jst_now().strftime('%Y/%m/%d %H:%M:%S')),
        email_info.get('from', ''),
        email_info.get('subject', ''),
        email_info.get('body', '')
    ]

    body_payload = {
        "valueRange": {
            "range": f"{sheet_id}!A{target_row}:D{target_row}",
            "values": [row_data]
        }
    }

    try:
        requests.put(write_url, headers=headers, json=body_payload, timeout=10)
        log_flush(f"📧 Lark [復旧依頼メール] (行{target_row}) にメール情報を転記完了: {email_info.get('subject', '')}")
    except Exception as e:
        log_flush(f"Larkメールシート書き込み例外: {e}", logging.ERROR)

def write_to_lark_sheet(extracted_data, detected_region):
    tenant_token = get_lark_tenant_access_token()
    if not tenant_token:
        log_flush("Lark APIトークンが取得できなかったため、シート書き込みをスキップします。", logging.WARNING)
        return

    spreadsheet_token = get_spreadsheet_token(tenant_token)
    headers = {
        "Authorization": f"Bearer {tenant_token}",
        "Content-Type": "application/json; charset=utf-8"
    }

    now = jst_now()
    target_sheet_title = f"{now.year}年{now.month}月"
    sheet_id = None

    try:
        sheets_url = f"https://open.larksuite.com/open-apis/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/query"
        res_sheets = requests.get(sheets_url, headers=headers, timeout=10)
        res_data = res_sheets.json()
        if res_sheets.status_code == 200 and res_data.get("code") == 0:
            sheets_list = res_data.get("data", {}).get("sheets", [])
            for s in sheets_list:
                if target_sheet_title in s.get("title", ""):
                    sheet_id = s.get("sheet_id")
                    break
    except Exception as e:
        log_flush(f"シートタブ一覧の取得失敗: {e}", logging.WARNING)

    if not sheet_id:
        try:
            copy_url = f"https://open.larksuite.com/open-apis/sheets/v3/spreadsheets/{spreadsheet_token}/sheets/{DEFAULT_SHEET_ID}/copy"
            copy_body = {"destination_name": target_sheet_title}
            res_copy = requests.post(copy_url, headers=headers, json=copy_body, timeout=10)
            res_copy_json = res_copy.json()
            if res_copy_json.get("code") == 0:
                sheet_id = res_copy_json.get("data", {}).get("sheet", {}).get("sheet_id")
            else:
                sheet_id = DEFAULT_SHEET_ID
        except Exception as e:
            sheet_id = DEFAULT_SHEET_ID

    read_url = f"https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values/{sheet_id}!B1:E500"
    target_row = 4
    last_no = 0

    try:
        res_read = requests.get(read_url, headers=headers, timeout=10)
        res_read_data = res_read.json()
        if res_read.status_code == 200 and res_read_data.get("code") == 0:
            values = res_read_data.get("data", {}).get("valueRange", {}).get("values", [])
            last_filled_row = 3
            for idx, row_val in enumerate(values):
                row_num = idx + 1
                if row_num >= 4 and row_val:
                    b_val = row_val[0] if len(row_val) > 0 else None
                    e_val = row_val[3] if len(row_val) > 3 else None
                    
                    b_str = str(b_val).strip() if b_val is not None else ""
                    e_str = str(e_val).strip() if e_val is not None else ""
                    
                    if e_str and e_str.lower() != 'none':
                        last_filled_row = row_num
                        if b_str and b_str.lower() != 'none':
                            try: last_no = int(float(b_str))
                            except ValueError: pass

            target_row = max(last_filled_row + 1, 4)
    except Exception as e:
        log_flush(f"空行判定エラー: {e}", logging.WARNING)

    if last_no == 0 and target_row > 4:
        last_no = target_row - 4

    next_biz_day = get_next_business_day().strftime('%Y/%m/%d')
    write_url = f"https://open.larksuite.com/open-apis/sheets/v2/spreadsheets/{spreadsheet_token}/values"

    stop_items = [d for d in extracted_data if d.get('停止or復旧') != '復旧']
    recovery_items = [d for d in extracted_data if d.get('停止or復旧') == '復旧']

    # 関東＝ガスプラ課、関西＝西日本
    team = "ガスプラ課" if detected_region == "関東" else "西日本"
    current_no = last_no

    if stop_items:
        current_no += 1
        kind = stop_items[0]['物件種別']
        action = "【停止】"
        # B(No), C(作業者名), D(チーム), E(内容), F(物件区分), G(件名), H(作業日), I(数量), J(金額), K(消費税), L(備考)
        row_bl = [current_no, "", team, action, kind, "", next_biz_day, 1, 35000, 3500, ""]
        body_bl = {
            "valueRange": {
                "range": f"{sheet_id}!B{target_row}:L{target_row}",
                "values": [row_bl]
            }
        }
        try:
            requests.put(write_url, headers=headers, json=body_bl, timeout=10)
            log_flush(f"📝 Larkシート [{target_sheet_title}] (行{target_row}, NO:{current_no}) に停止1行転記完了: {action} ({len(stop_items)}件分)")
            target_row += 1
        except Exception as e:
            log_flush(f"Larkシート停止書き込み例外: {e}", logging.ERROR)

    for d in recovery_items:
        current_no += 1
        action = "【復旧】"
        subject = f"{action}{d['物件名']} {d['部屋番号']}"
        # B(No), C(作業者名), D(チーム), E(内容), F(物件区分), G(件名), H(作業日), I(数量), J(金額), K(消費税), L(備考)
        row_bl = [current_no, "", team, action, d['物件種別'], subject, next_biz_day, 1, 15000, 1500, d.get('備考', '')]
        body_bl = {
            "valueRange": {
                "range": f"{sheet_id}!B{target_row}:L{target_row}",
                "values": [row_bl]
            }
        }
        try:
            requests.put(write_url, headers=headers, json=body_bl, timeout=10)
            log_flush(f"📝 Larkシート [{target_sheet_title}] (行{target_row}, NO:{current_no}) に復旧転記完了: {subject}")
            target_row += 1
        except Exception as e:
            log_flush(f"Larkシート復旧書き込み例外: {e}", logging.ERROR)
