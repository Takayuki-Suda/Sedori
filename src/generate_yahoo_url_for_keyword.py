import urllib.parse
from spreadsheet_utils import open_worksheet  # 外部ファイルから関数をインポート

# スプレッドシートを開く（デフォルトで2番目のシートを開く）
worksheet = open_worksheet()

# D列の全てのセルを取得（D4から下のセル）
d_values = worksheet.col_values(4)  # D列（4列目）の値を取得

# 処理開始のメッセージ
print("検索中です...")

# D列の値に基づいて、H列にYahoo!ショッピングのURLを追加
for i, keyword in enumerate(d_values[3:], start=4):  # D4から開始（インデックス3がD4に対応）
    if keyword:  # キーワードが空でない場合
        # キーワードをURLエンコードして、複数のキーワードをカンマ区切りで対応
        encoded_keywords = urllib.parse.quote_plus(keyword.replace(",", " "))  # ','をスペースに置き換え、URLエンコード
        # URLを生成
        yahoo_url = f"https://shopping.yahoo.co.jp/search?p={encoded_keywords}&tab_ex=commerce&area=13&X=2&sc_i=shopping-pc-web-result-item-sort_mdl-sortitem"
        worksheet.update_cell(i, 14, yahoo_url)  # N列（14列目）にURLを入力
        
        # 進捗をコンソールに表示
        print(f"{i - 3} / {len(d_values) - 3} 件目のURLを更新中...")

# 完了メッセージ
print("Yahoo!ショッピングのURLをD列のキーワードに基づいてN列に更新しました。")
