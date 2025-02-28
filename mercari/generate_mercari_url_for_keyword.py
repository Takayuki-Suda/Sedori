import urllib.parse
from spreadsheet_utils import open_worksheet  # 外部ファイルから関数をインポート

# スプレッドシートを開く（デフォルトで2番目のシートを開く）
worksheet = open_worksheet()

# D列の全てのセルを取得（D4から下のセル）
d_values = worksheet.col_values(4)  # D列（4列目）の値を取得

# 処理開始のメッセージ
print("検索中です...")

# D列の値に基づいて、E列にメルカリURLを追加
for i, keyword in enumerate(d_values[3:], start=4):  # D4から開始（インデックス3がD4に対応）
    if keyword:  # キーワードが空でない場合
        mercari_url = f"https://www.mercari.com/jp/search/?keyword={urllib.parse.quote(keyword)}&order=asc&sort=price"
        worksheet.update_cell(i, 5, mercari_url)  # E列にURLを入力
        
        # 進捗をコンソールに表示
        print(f"{i - 3} / {len(d_values) - 3} 件目のURLを更新中...")

# 完了メッセージ
print("メルカリのURLをD列のキーワードに基づいてE列に更新しました。")
