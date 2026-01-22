import os
import json
import time
import glob
import csv
from urllib.parse import quote
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
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
json_creds = json.loads(os.environ.get("GCP_JSON", "{}")) 
TARGET_URL = os.environ.get("TARGET_URL", "https://example.com/login") 

# --- 設定 ---
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1H2TiCraNjMNoj3547ZB78nQqrdfbfk2a0rMLSbZBE48")
SHEET_NAME = "test今月_raw"

def get_google_service(service_name, version):
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_info(json_creds, scopes=scopes)
    return build(service_name, version, credentials=creds)

def update_google_sheet(csv_path):
    print(f"スプレッドシートへの転記を開始: {SHEET_NAME}")
    service = get_google_service('sheets', 'v4')

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
        print("既存データをクリアしました。")
    except Exception as e:
        print(f"シートクリアエラー: {e}")

    body = {'values': csv_data}
    try:
        result = service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET_NAME}!A1",
            valueInputOption='USER_ENTERED',
            body=body
        ).execute()
        print(f"スプレッドシート更新完了: {result.get('updatedCells')} セル更新")
    except Exception as e:
        print(f"書き込みエラー: {e}")

def main():
    print("=== Action Log取得処理開始(今月分) ===")
    
    download_dir = os.path.join(os.getcwd(), "downloads_action_month")
    if not os.path.exists(download_dir):
        os.makedirs(download_dir)

    for f in glob.glob(os.path.join(download_dir, "*.csv")):
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
        print("「絞り込み検索」ボタン操作...")
        try:
            filter_btn = wait.until(EC.element_to_be_clickable((By.ID, "searchFormOpen")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", filter_btn)
            time.sleep(1)
            filter_btn.click()
            print("「絞り込み検索」ボタンをクリックしました")
            time.sleep(1)
        except Exception as e:
            print(f"絞り込み検索ボタンが見つかりません: {e}")

        # --- 3. 「今月」ボタン ---
        print("「今月」ボタン操作...")
        try:
            current_month_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".current_month")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", current_month_btn)
            time.sleep(1)
            current_month_btn.click()
            print("「今月」ボタンをクリックしました")
            time.sleep(3)
        except Exception as e:
            print(f"「今月」ボタン操作エラー: {e}")

        # --- 4. パートナー選択 (修正版: クリック → Enter) ---
        print("パートナー選択: 入力欄をクリックしてEnterキーを押します...")
        try:
            input_xpath = "//input[@placeholder='選択または検索ができます']"
            partner_input = wait.until(EC.element_to_be_clickable((By.XPATH, input_xpath)))
            
            # 入力欄までスクロールしてクリック
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", partner_input)
            time.sleep(1)
            partner_input.click()
            print("入力欄をクリックしました")
            
            # ドロップダウンが開くのを少し待つ
            time.sleep(1.5)
            
            # フォーカスが当たっているはずなので、Enterキーを送信
            driver.switch_to.active_element.send_keys(Keys.ENTER)
            print("Enterキーを送信しました")

            time.sleep(2)

        except Exception as e:
            print(f"パートナー選択エラー: {e}")

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
                driver.execute_script("arguments[0].click();", target_search_btn)
                print("検索ボタンをクリックしました")
            else:
                print("検索ボタンが見つからないためEnterキーで代用します")
                webdriver.ActionChains(driver).send_keys(Keys.ENTER).perform()

        except Exception as e:
            print(f"検索ボタン操作エラー: {e}")
            webdriver.ActionChains(driver).send_keys(Keys.ENTER).perform()
        
        print("検索結果を待機中(15秒)...")
        time.sleep(15)

        # --- 6. CSV生成ボタン ---
        print("CSV生成ボタン操作...")
        try:
            csv_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@value='CSV生成' or contains(text(), 'CSV生成')]")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", csv_btn)
            time.sleep(1)
            driver.execute_script("arguments[0].click();", csv_btn)
            print("CSV生成ボタンをクリックしました")
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
            return
        
        print(f"ダウンロード成功: {csv_file_path}")
        update_google_sheet(csv_file_path)

    except Exception as e:
        print(f"【エラー発生】: {e}")
        import traceback
        traceback.print_exc()
        
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
