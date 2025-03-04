import requests
import re
import urllib.parse
from bs4 import BeautifulSoup
from spreadsheet_utils import open_worksheet

# スプレッドシートを開く
worksheet = open_worksheet()

# H列（17列目）のURLを取得
all_data = worksheet.get_all_values()
urls_h = [row[16] for row in all_data[3:] if row[16].startswith("http")]

# ヘッダー情報
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Connection": "keep-alive",
}

# eBayの価格を取得する関数
def get_ebay_price(url):
    try:
        response = requests.get(url, headers=HEADERS, timeout=10)
        response.raise_for_status()  # ステータスコードが200以外なら例外を発生
        soup = BeautifulSoup(response.text, "html.parser")

        # 価格を探す
        price_text = None
        selectors = [".s-item__price", ".x-price-primary", ".notranslate", ".display-price",
                     ".item-price", "span[itemprop='price']"]

        for selector in selectors:
            elements = soup.select(selector)
            for element in elements:
                # コメントなどを含めた全コンテンツを取得
                raw_text = "".join(str(content) for content in element.contents)
                text = re.sub(r'<!--.*?-->', '', raw_text)  # コメントを削除
                text = text.strip()

                # 数値（価格）を抽出（カンマ付きの数字＋円）
                price_pattern = r'(\d{1,3}(?:,\d{3})*)\s*円'
                matches = re.findall(price_pattern, text)

                if matches:
                    price_text = f"{matches[0]}円"  # 最初の価格を取得
                    break
            if price_text:
                break

        if price_text:
            # 数値のみ抽出して float に変換
            numeric_price = re.sub(r'[^\d,.]', '', price_text)
            float_price = float(numeric_price.replace(',', ''))
            return float_price

    except requests.RequestException as e:
        print(f"リクエストエラー（URL: {url}）: {e}")

    return None


# スプレッドシートへ書き込み
for row, url in enumerate(urls_h, start=4):
    price = get_ebay_price(url)
    if price is not None:
        worksheet.update_cell(row, 18, price)  # 18列目に価格を入力
        print(f"成功: {url} → {price}")

print("金額の入力が完了しました。")
