import os
import json
import time
import glob
import csv
from urllib.parse import quote
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- 環境変数 (Secretsから取得) ---
USER_ID = os.environ.get("USER_ID", "your_id")
PASSWORD = os.environ.get("USER_PASS", "your_pass")
# GCP_JSONは文字列として読み込む
json_creds = json.loads(os.environ.get("GCP_JSON", "{}")) 
TARGET_URL = os.environ.get("TARGET_URL", "https://example.com/login") 

# --- 設定 ---
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1H2TiCraNjMNoj3547ZB78nQqrdfbfk2a0rMLSbZBE48")
SHEET_NAME = "test今月_raw"
PARTNER_NAME = "株式会社フルアウト"

# ★ここに指定されたドライブフォルダIDを設定
DRIVE_FOLDER_ID = "1hqyzehrzUGWsdV8VQSkWgE9YWoMyc6bs"

def get_google_service(service_name, version, scopes):
    """Google APIサービスを取得するヘルパー関数"""
    creds = Credentials.from_service_account_info(json_creds, scopes=scopes)
    return build(service_name, version, credentials=creds)

def upload_to_drive(file_path):
    """ファイルをGoogle Driveにアップロードする"""
    print(f"Google Driveへアップロード中: {os.path.basename(file_path)}")
    try:
        service = get_google_service('drive', 'v3', ['https://www.googleapis.com/auth/drive'])
        
        file_metadata = {
            'name': os.path.basename(file_path),
            'parents': [DRIVE_FOLDER_ID]
        }
        
        # 拡張子に応じたMIMEタイプ簡易判定
        mimetype = 'application/octet-stream'
        if file_path.endswith('.csv'):
            mimetype = 'text/csv'
        elif file_path.endswith('.png'):
            mimetype = 'image/png'

        media = MediaFileUpload(file_path, mimetype=mimetype)
        
        file = service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id',
            supportsAllDrives=True
        ).execute()
        print(f"アップロード完了 ID: {file.get('id')}")
        
    except Exception as e:
        print(f"ドライブアップロードエラー: {e}")

def update_google_sheet(csv_path):
    """CSVをスプレッドシートに転記"""
    print(f"スプレッドシートへの転記を開始: {SHEET_NAME}")
    service = get_google_service('sheets', 'v4', ['https://www.googleapis.com/auth/spreadsheets'])

    csv_data = []
    try:
        with open(csv_path, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            csv_data = list(reader)
    except UnicodeDecodeError:
        print("UTF-8での読み込みに失敗しました。Shift_JIS(CP932)で再試行します。")
        try:
            with open(csv_path, 'r', encoding='cp932') as f:
                reader = csv.reader(f)
                csv_data = list(reader)
        except Exception as e:
            print(f"CSV読み込みエラー: {e}")
            return

    if not csv_data:
        print("CSVデータが空のため転記をスキップします。")
        return

    try:
        service.spreadsheets().values().clear(
            spreadsheetId=SPREADSHEET_ID,
            range=SHEET_NAME
        ).execute()
    except Exception as e:
        print(f"シートクリアエラー: {e}")

    body = {'values': csv_data}
    try:
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET_NAME}!A1",
            valueInputOption='USER_ENTERED',
            body=body
        ).execute()
        print("スプレッドシート更新完了")
    except Exception as e:
        print(f"書き込みエラー: {e}")

def highlight(driver, element):
    """
    指定された要素を赤枠と赤背景で強調表示する
    （どこがクリックされようとしているか視覚化するため）
    """
    driver.execute_script(
        "arguments[0].setAttribute('style', 'border: 5px solid red; background-color: rgba(255, 0, 0, 0.5);');", 
        element
    )

def main():
    print("=== Action Log取得処理開始 ===")
    
    download_dir = os.path.join(os.getcwd(), "downloads_action_month")
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

    # フォルダを空にする
    for f in glob.glob(os.path.join(download_dir, "*")):
        os.remove(f)

    options = Options()
    options.add_argument('--headless') 
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--window-size=1920,1080')
    
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 20)

    try:
        # --- 1. ログイン ---
        safe_user = quote(USER_ID, safe='')
        safe_pass = quote(PASSWORD, safe='')
        url_body = TARGET_URL.replace("https://", "").replace("http://", "")
        auth_url = f"https://{safe_user}:{safe_pass}@{url_body}"
        
        print(f"アクセス中: {TARGET_URL}")
        driver.get(auth_url)
        time.sleep(3)
        driver.get(auth_url)
        time.sleep(5) 

        # --- 2. 「絞り込み検索」ボタン ---
        print("検索メニューを開きます...")
        try:
            filter_btn = wait.until(EC.element_to_be_clickable((By.ID, "searchFormOpen")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", filter_btn)
            time.sleep(1)
            
            # ★ハイライト＆キャプチャ
            highlight(driver, filter_btn)
            driver.save_screenshot(os.path.join(download_dir, "01_before_filter_click.png"))
            
            filter_btn.click()
            time.sleep(2)
        except Exception as e:
            print(f"絞り込み検索ボタン操作エラー: {e}")

        # --- 3. 「今月」ボタン ---
        print("「今月」ボタンを選択します...")
        try:
            current_month_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".current_month")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", current_month_btn)
            time.sleep(1)
            
            # ★ハイライト＆キャプチャ
            highlight(driver, current_month_btn)
            driver.save_screenshot(os.path.join(download_dir, "02_before_month_click.png"))
            
            current_month_btn.click()
            time.sleep(3)
        except Exception as e:
            print(f"「今月」ボタン操作エラー: {e}")

        # --- 4. パートナー選択 (入力→待機→Enter) ---
        print(f"パートナー({PARTNER_NAME})を入力します...")
        try:
            # 入力欄の特定
            partner_label = driver.find_element(By.XPATH, "//div[contains(text(), 'パートナー')] | //label[contains(text(), 'パートナー')]")
            partner_target = partner_label.find_element(By.XPATH, "./following::input[contains(@placeholder, '選択')][1]")
            
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", partner_target)
            
            # ★ハイライト＆キャプチャ（クリック前）
            highlight(driver, partner_target)
            driver.save_screenshot(os.path.join(download_dir, "03_before_partner_input.png"))
            
            partner_target.click()
            time.sleep(1)
            
            # 文字入力
            active_elem = driver.switch_to.active_element
            active_elem.send_keys(PARTNER_NAME)
            
            # 3秒待機（候補が出るのを待つ）
            time.sleep(3)
            
            # ★Enterを押す直前の状態をキャプチャ（候補が見えているか確認）
            driver.save_screenshot(os.path.join(download_dir, "04_before_enter.png"))
            
            # Enterで確定
            active_elem.send_keys(Keys.ENTER)
            print("パートナーを選択しました")
            time.sleep(2)

        except Exception as e:
            print(f"パートナー入力エラー: {e}")

        # --- 5. 検索ボタン実行 ---
        print("検索ボタン操作...")
        try:
            search_btns = driver.find_elements(By.XPATH, "//input[@value='検索'] | //button[contains(text(), '検索')]")
            target_search_btn = None
            for btn in search_btns:
                if btn.is_displayed():
                    target_search_btn = btn
                    break
            
            if target_search_btn:
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target_search_btn)
                time.sleep(1)
                
                # ★ハイライト＆キャプチャ
                highlight(driver, target_search_btn)
                driver.save_screenshot(os.path.join(download_dir, "05_before_search_click.png"))
                
                driver.execute_script("arguments[0].click();", target_search_btn)
            else:
                webdriver.ActionChains(driver).send_keys(Keys.ENTER).perform()

        except Exception as e:
            print(f"検索ボタン操作エラー: {e}")
            webdriver.ActionChains(driver).send_keys(Keys.ENTER).perform()
        
        print("検索結果を待機中(15秒)...")
        time.sleep(15)
        
        # ★検索結果画面をキャプチャ（正しく絞り込まれているか確認）
        driver.save_screenshot(os.path.join(download_dir, "06_search_result.png"))

        # --- 6. CSV生成ボタン ---
        print("CSV生成ボタン操作...")
        try:
            csv_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@value='CSV生成' or contains(text(), 'CSV生成')]")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", csv_btn)
            time.sleep(1)
            
            # ★ハイライト
            highlight(driver, csv_btn)
            driver.save_screenshot(os.path.join(download_dir, "07_before_csv_click.png"))
            
            driver.execute_script("arguments[0].click();", csv_btn)
            
        except Exception as e:
            print(f"CSVボタンエラー: {e}")
            return
        
        print("ダウンロード待機中...")
        time.sleep(5)
        csv_file_path = None
        for i in range(30):
            files = glob.glob(os.path.join(download_dir, "*.csv"))
            if files:
                csv_file_path = files[0]
                break
            time.sleep(2)
            
        if not csv_file_path:
            print("【エラー】CSVファイルが見つかりません。")
            # 失敗時も画像をアップロードするためにreturnせず進む
        else:
            print(f"ダウンロード成功: {csv_file_path}")
            # --- 7. スプレッドシートへ転記 ---
            update_google_sheet(csv_file_path)

        # --- 8. 全ファイルをGoogle Driveへアップロード ---
        print("=== Google Driveへの保存処理 ===")
        all_files = glob.glob(os.path.join(download_dir, "*"))
        for file_path in all_files:
            upload_to_drive(file_path)

    except Exception as e:
        print(f"【エラー発生】: {e}")
        import traceback
        traceback.print_exc()
        
        # エラー時も可能な限り画像をアップロード
        try:
            all_files = glob.glob(os.path.join(download_dir, "*"))
            for file_path in all_files:
                upload_to_drive(file_path)
        except:
            pass
        
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
