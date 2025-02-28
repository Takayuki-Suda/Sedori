import urllib.parse
from spreadsheet_utils import open_worksheet  # 外部ファイルから関数をインポート

# スプレッドシートを開く（デフォルトで2番目のシートを開く）
worksheet = open_worksheet()

# D列の全てのセルを取得（D4から下のセル）
d_values = worksheet.col_values(4)  # D列（4列目）の値を取得

# 処理開始のメッセージ
print("検索中です...")

# D列の値に基づいて、E列にAmazon URLを追加
for i, keyword in enumerate(d_values[3:], start=4):  # D4から開始（インデックス3がD4に対応）
    if keyword:  # キーワードが空でない場合
        amazon_url = f"https://search.rakuten.co.jp/search/mall/{urllib.parse.quote(keyword)}/?s=2"  # 楽天の商品検索URL
        worksheet.update_cell(i, 11, amazon_url)  # K列（11列目）にURLを入力
        
        # 進捗をコンソールに表示
        print(f"{i - 3} / {len(d_values) - 3} 件目のURLを更新中...")

# 完了メッセージ
print("楽天のURLをD列のキーワードに基づいてK列に更新しました。")
