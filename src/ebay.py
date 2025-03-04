import requests
import re
import time
import random
import concurrent.futures
from bs4 import BeautifulSoup
from spreadsheet_utils import open_worksheet

# スプレッドシートを開く
worksheet = open_worksheet()

# H列（17列目）のURLを取得
all_data = worksheet.get_all_values()
urls_h = [(row_idx + 4, row[16]) for row_idx, row in enumerate(all_data[3:]) if row[16].startswith("http")]

# ヘッダー情報（ランダム化）
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
    "Mozilla/5.0 (iPhone; CPU iPhone OS 16_1 like Mac OS X) AppleWebKit/537.36 (KHTML, like Gecko) Version/16.1 Mobile/15E148 Safari/537.36"
]

HEADERS = {
    "User-Agent": random.choice(USER_AGENTS),
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Connection": "keep-alive",
}

# eBayの価格を取得する関数（リトライ付き）
def get_ebay_price(row_url, max_retries=3):
    row, url = row_url
    for attempt in range(max_retries):
        try:
            print(f"取得中 ({attempt+1}/{max_retries}): {url}")

            session = requests.Session()
            response = session.get(url, headers=HEADERS, timeout=15)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, "html.parser")

            # 価格を探す
            price_text = None
            selectors = [".s-item__price", ".x-price-primary", ".notranslate", ".display-price",
                         ".item-price", "span[itemprop='price']"]

            for selector in selectors:
                elements = soup.select(selector)
                for element in elements:
                    raw_text = "".join(str(content) for content in element.contents)
                    text = re.sub(r'<!--.*?-->', '', raw_text).strip()
                    matches = re.findall(r'(\d{1,3}(?:,\d{3})*)\s*円', text)

                    if matches:
                        price_text = matches[0]
                        break
                if price_text:
                    break

            if price_text:
                numeric_price = float(price_text.replace(',', ''))
                return row, numeric_price

        except requests.RequestException as e:
            print(f"リクエストエラー（URL: {url}）: {e}")

        # ランダムスリープ（Bot対策）
        time.sleep(random.uniform(2, 5))

    return row, None

# 並列リクエスト
results = []
with concurrent.futures.ThreadPoolExecutor(max_workers=3) as executor:
    results = list(executor.map(get_ebay_price, urls_h))

# スプレッドシートへ一括書き込み
updates = [(row, 18, price) for row, price in results if price is not None]
if updates:
    worksheet.batch_update([{"range": f"R{row}C{col}", "values": [[price]]} for row, col, price in updates])

print("金額の入力が完了しました。")
