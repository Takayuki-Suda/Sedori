import time
from spreadsheet_utils import open_worksheet  # 外部ファイルから関数をインポート
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from concurrent.futures import ThreadPoolExecutor, as_completed


# スプレッドシートを開く（デフォルトで2番目のシートを開く）
worksheet = open_worksheet()

# 一度に全てのデータを取得
all_data = worksheet.get_all_values()

# H列（10列目）のURLを取得
urls = [row[10] for row in all_data[3:] if row[10].startswith("http")]  # URLが空でないことを確認

# 金額を取得する関数
def get_price(url, row):
    options = Options()
    options.headless = True  # ヘッドレスモードを有効にする
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)  # 各スレッドごとに新しいdriverを作成

    try:
        # 対象のURLにアクセス
        driver.get(url)

        # 最初の金額を取得（動的待機）
        price_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, "//div[contains(@class, 'price--3zUvK') and contains(@class, 'price-with-price-plus-shipping--Bmgz2')]")
            )
        )

        # "円" の部分を除去
        currency_text = price_element.find_element(By.TAG_NAME, "span").text  # 円の部分
        price_text = price_element.text.replace(currency_text, "").strip()  # 数値部分のみ取得

        # 数値に変換
        first_price = int(price_text.replace(',', ''))  # カンマがあれば除去

        # L列（12列目）に金額を入力
        worksheet.update_cell(row, 12, first_price)
        return None  # 成功時は何も返さない

    except Exception as e:
        return f"エラーが発生しました (URL: {url}): {str(e)}"

    finally:
        driver.quit()  # スレッドごとに作成したdriverを閉じる

# 並列処理でURLを処理
with ThreadPoolExecutor(max_workers=10) as executor:  # 10スレッドで並列処理
    future_to_url = {executor.submit(get_price, url, row): (url, row) for row, url in enumerate(urls, start=4)}

    # 結果を順番に取得
    for future in as_completed(future_to_url):
        result = future.result()
        if result:  # None以外の結果のみ表示
            print(result)

# 全ての金額入力が完了したメッセージを表示
print("金額の入力が完了しました。")
