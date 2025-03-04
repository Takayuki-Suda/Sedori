import time
import re
import traceback  # 追加
from spreadsheet_utils import open_worksheet
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from concurrent.futures import ThreadPoolExecutor, as_completed
from bs4 import BeautifulSoup
from selenium.common.exceptions import TimeoutException  # 追加

# スプレッドシートを開く
worksheet = open_worksheet()

# H列（17列目）のURLを取得（最初の2つのみ）
all_data = worksheet.get_all_values()
urls_h = [row[16] for row in all_data[3:] if row[16].startswith("http")]

# スレッドごとにChromeドライバを作成する関数
def create_driver():
    options = Options()
    options.headless = True  # 確認のためFalseにしてデバッグ可能
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920x1080")

    # 不要なリソースの読み込みを抑制
    prefs = {
        "profile.managed_default_content_settings.stylesheets": 2,
        "profile.managed_default_content_settings.images": 2,
        "profile.managed_default_content_settings.plugins": 2,
        "profile.managed_default_content_settings.popups": 2,
        "profile.default_content_setting_values.notifications": 2,
    }
    options.add_experimental_option("prefs", prefs)

    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# eBayの金額を取得する関数
def get_ebay_price(url, row, column):
    driver = create_driver()
    try:
        driver.get(url)
        
        # ページの読み込みを待機
        WebDriverWait(driver, 30).until(lambda d: d.execute_script("return document.readyState") == "complete")
        time.sleep(3)  # JavaScriptの実行を待つ
        
        # HTMLを取得して解析
        html = driver.page_source
        soup = BeautifulSoup(html, "html.parser")
        
        # 価格を探す
        price_text = None
        
        # 様々なセレクタで価格を探す
        selectors = [".s-item__price", ".x-price-primary", ".notranslate", ".display-price", ".item-price", "span[itemprop='price']"]
        for selector in selectors:
            elements = soup.select(selector)
            for element in elements:
                text = element.get_text(strip=True)
                if text and ('円' in text):
                    price_text = text
                    print(f"{selector} から価格を検出: {price_text}")
                    break
            if price_text:
                break
        
        # 価格がまだ見つからない場合、ページ全体から探す
        if not price_text:
            page_text = soup.get_text()
            price_pattern = r'(\d{1,3}(?:,\d{3})*)(?:\s*)円'
            matches = re.findall(price_pattern, page_text)
            if matches:
                price_text = f"{matches[0]}円"
                print(f"正規表現で価格を検出: {price_text}")
        
        # 価格範囲を処理する（例：926 円 ～ 1,048 円）
        if price_text:
            # 「～」や「から」などの範囲表現がある場合
            if '～' in price_text or 'から' in price_text or ' to ' in price_text:
                # 範囲の数値を全て抽出
                all_prices = re.findall(r'(\d{1,3}(?:,\d{3})*)(?:\s*)(?:円|\$|€)', price_text)
                if all_prices:
                    # 全ての価格を数値に変換
                    numeric_prices = [float(price.replace(',', '')) for price in all_prices]
                    # 最小値を使用
                    float_price = min(numeric_prices)
                    print(f"価格範囲から最小値を使用: {float_price}")
                else:
                    print(f"価格範囲から数値を抽出できませんでした: {price_text}")
                    return
            else:
                # 通常の単一価格の場合
                numeric_price = re.sub(r'[^\d,.]', '', price_text)
                float_price = float(numeric_price.replace(',', ''))
            
            # スプレッドシートに価格を入力
            worksheet.update_cell(row, column, float_price)
            print(f"成功: {url} → {float_price}")
            return
        
        # デバッグ情報
        print(f"価格が見つかりませんでした: {url}")
        print("ページ内の主要な要素:")
        for elem in soup.select("span, div.price, *[class*='price']")[:10]:
            print(f"- {elem.name}.{' '.join(elem.get('class', []))}: {elem.get_text(strip=True)}")
        
    except TimeoutException as e:
        print(f"タイムアウトエラー発生（URL: {url}）: {e}")
        traceback.print_exc()
    
    except Exception as e:
        print(f"エラー発生（URL: {url}）: {e}")
        traceback.print_exc()
    
    finally:
        driver.quit()
# 並列処理でH列のURLを処理
with ThreadPoolExecutor(max_workers=5) as executor:  # スレッド数を5に調整
    future_to_url = {executor.submit(get_ebay_price, url, row, 18): (url, row) for row, url in enumerate(urls_h, start=4)}

    for future in as_completed(future_to_url):
        future.result()  # 例外を発生させずにスレッド処理を実行

print("金額の入力が完了しました。")
