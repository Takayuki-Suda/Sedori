import time
from spreadsheet_utils import open_worksheet  # 外部ファイルから関数をインポート
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait  # 必要なインポート
from selenium.webdriver.support import expected_conditions as EC  # 必要なインポート
from concurrent.futures import ThreadPoolExecutor, as_completed


# スプレッドシートを開く（デフォルトで2番目のシートを開く）
worksheet = open_worksheet()

# Chromeのオプション設定（ヘッドレスモードを有効にする）
options = Options()
options.headless = True  # ヘッドレスモードを有効にする

# Chromeドライバーのセットアップ
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# 一度に全てのデータを取得
all_data = worksheet.get_all_values()

# E列（5列目）とH列（7列目）のURLを取得
urls_e = [row[4] for row in all_data[3:]]  # E列（5列目）のURL
urls_h = [row[7] for row in all_data[3:]]  # H列（7列目）のURL

# 金額を取得する関数
def get_price(url, row, column):
    try:
        # 対象のURLにアクセス
        driver.get(url)
        
        # AmazonとMercariのURLで金額取得用のクラス名が異なるので分岐
        if 'amazon' in url:
            # Amazonの金額を取得（動的待機）
            price_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//span[@class='a-price-whole']"))
            )
        elif 'mercari' in url:
            # Mercariの金額を取得（動的待機）
            price_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//span[@class='number__6b270ca7']"))
            )
        else:
            return f"対応するURLの形式ではないURLです (URL: {url})"

        first_price = int(price_element.text.replace('¥', '').replace(',', ''))
        
        # 指定された列に金額を入力
        worksheet.update_cell(row, column, first_price)
        return None  # 成功時は何も返さない
    
    except Exception as e:
        return f"エラーが発生しました (URL: {url}): {str(e)}"


# 並列処理でE列のURLを処理
with ThreadPoolExecutor(max_workers=10) as executor:  # 10スレッドで並列処理
    future_to_url_e = {executor.submit(get_price, url, row, 6): (url, row) for row, url in enumerate(urls_e, start=4)}
    
    # 結果を順番に取得
    for future in as_completed(future_to_url_e):
        result = future.result()
        if result:  # None以外の結果のみ表示
            print(result)

# 並列処理でH列のURLを処理
with ThreadPoolExecutor(max_workers=1) as executor:  # 1スレッドで並列処理
    future_to_url_h = {executor.submit(get_price, url, row, 9): (url, row) for row, url in enumerate(urls_h, start=4)}
    
    # 結果を順番に取得
    for future in as_completed(future_to_url_h):
        result = future.result()
        if result:  # None以外の結果のみ表示
            print(result)

# 全ての金額入力が完了したメッセージを表示
print("金額の入力が完了しました。")

# ブラウザを閉じる
driver.quit()
