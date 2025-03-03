import time
import requests
import re
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed
from spreadsheet_utils import open_worksheet  # 外部ファイルから関数をインポート

# スプレッドシートを開く
worksheet = open_worksheet()
all_data = worksheet.get_all_values()

# 各列のURLを取得
urls = {
    "amazon": [(row[7], row_idx, "I") for row_idx, row in enumerate(all_data[3:], start=4) if len(row) > 7 and row[7].startswith("http")],
    "rakuten": [(row[10], row_idx, "L") for row_idx, row in enumerate(all_data[3:], start=4) if len(row) > 10 and row[10].startswith("http")],
    "yahoo": [(row[13], row_idx, "O") for row_idx, row in enumerate(all_data[3:], start=4) if len(row) > 13 and row[13].startswith("http")]
}

# User-Agent（bot対策）
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
}

# 価格取得関数
def get_price(url, site):
    try:
        response = requests.get(url, headers=HEADERS, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")

        if site == "amazon":
            price_element = soup.find("span", class_="a-price-whole")
        elif site == "rakuten":
            price_element = soup.select_one("div.price--3zUvK.price-with-price-plus-shipping--Bmgz2")
        elif site == "yahoo":
            price_element = soup.select_one(".SearchResultItemPrice_SearchResultItemPrice__value__G8pQV")
        else:
            return url, None, f"{site}: サイト不明"

        if not price_element:
            return url, None, f"{site}: 価格が見つかりません"

        # 価格テキストを取得し、数値以外を削除
        price_text = price_element.get_text(strip=True)
        price_text = re.sub(r"[^\d]", "", price_text)  # 数字以外を削除
        if not price_text:
            return url, None, f"{site}: 数値の抽出に失敗"

        return url, int(price_text), None  # (URL, 価格, エラーなし)

    except Exception as e:
        return url, None, f"{site}エラー: {str(e)}"

# スプレッドシート更新関数
def update_spreadsheet(updates):
    if updates:  # 空リストで実行しないようにする
        try:
            worksheet.batch_update(updates)
            time.sleep(1)  # 429エラーを防ぐためにスリープ
        except Exception as e:
            print(f"スプレッドシート更新エラー: {str(e)}")

# 並列処理で価格取得
task_list = []
error_logs = []
price_updates = []

with ThreadPoolExecutor(max_workers=10) as executor:
    future_to_meta = {}
    
    for site, url_list in urls.items():
        for url, row, col in url_list:
            future_to_meta[executor.submit(get_price, url, site)] = (row, col)
    
    for future in as_completed(future_to_meta):
        row, col = future_to_meta[future]
        url, price, error = future.result()
        
        if error:
            error_logs.append(error)
        elif price is not None:
            price_updates.append({"range": f"{col}{row}", "values": [[price]]})

# スプレッドシート更新
update_spreadsheet(price_updates)

# エラーログ表示
if error_logs:
    print("\n=== エラー一覧 ===")
    for log in error_logs:
        print(log)

print("金額の取得が完了しました。")
