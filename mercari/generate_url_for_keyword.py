import urllib.parse
from spreadsheet_utils import open_worksheet  # 外部ファイルから関数をインポート

def update_shopping_urls_optimized(column_index=4, start_row=4, platforms=None):
    """
    D列のキーワードに基づいて、複数のECサイト（Amazon, メルカリ, 楽天, Yahoo!ショッピング）のURLを一度に更新する関数
    
    Args:
        column_index (int): キーワードが入っている列番号 (デフォルト: 4 = D列)
        start_row (int): 処理を開始する行番号 (デフォルト: 4)
        platforms (list): 更新するプラットフォームのリスト (デフォルト: 全プラットフォーム)
    """
    # プラットフォームが指定されていない場合は全てのプラットフォームを使用
    if platforms is None:
        platforms = [ "mercari","amazon", "rakuten", "yahoo"]
    
    # プラットフォームと対応する列のマッピング
    platform_columns = { 
        "mercari": 5,  # E列
        "amazon": 8,   # H列
        "rakuten": 11, # K列
        "yahoo": 14    # N列
    }
    
    # スプレッドシートを開く
    worksheet = open_worksheet()
    
    # D列の全てのキーワードを取得
    d_values = worksheet.col_values(column_index)
    keywords = d_values[start_row-1:]
    
    # キーワードが空でないインデックスを特定
    valid_indices = [i for i, k in enumerate(keywords) if k]
    valid_keywords = [keywords[i] for i in valid_indices]
    
    if not valid_keywords:
        print("更新するキーワードがありません。")
        return
    
    print(f"合計 {len(valid_keywords)} 件のキーワードを処理します...")
    
    # 各プラットフォームごとに一括更新
    for platform in platforms:
        if platform not in platform_columns:
            print(f"プラットフォーム {platform} はサポートされていません。")
            continue
        
        url_column_index = platform_columns[platform]
        url_updates = []
        
        print(f"{platform.capitalize()}のURLを生成中...")

        # 一括でURLを生成
        for keyword in valid_keywords:
            if platform == "mercari":
                shopping_url = f"https://www.mercari.com/jp/search/?keyword={urllib.parse.quote(keyword)}&order=asc&sort=price"
            elif platform == "amazon":
                shopping_url = f"https://www.amazon.co.jp/s?k={urllib.parse.quote(keyword)}&s=price-asc-rank"
            elif platform == "rakuten":
                shopping_url = f"https://search.rakuten.co.jp/search/mall/{urllib.parse.quote(keyword)}/?s=2"
            elif platform == "yahoo":
                encoded_keywords = urllib.parse.quote_plus(keyword.replace(",", " "))
                shopping_url = f"https://shopping.yahoo.co.jp/search?p={encoded_keywords}&tab_ex=commerce&area=13&X=2&sc_i=shopping-pc-web-result-item-sort_mdl-sortitem"
            
            url_updates.append([shopping_url])  # Google Sheetsに渡すためリストのリストとして追加
        
        # 一括で更新
        cell_range = f"{chr(64 + url_column_index)}{start_row}:{chr(64 + url_column_index)}{start_row + len(valid_keywords) - 1}"
        
        # Google Sheetsに一括更新
        worksheet.update(cell_range, url_updates)  # 範囲と一括データ更新
        
        print(f"{platform.capitalize()}のURL ({len(url_updates)} 件) を更新しました。")
    
    print("すべてのURLの更新が完了しました！")

# 使用例
# すべてのプラットフォームを一度に更新
update_shopping_urls_optimized()

# 特定のプラットフォームのみ更新
# update_shopping_urls_optimized(platforms=["amazon", "yahoo"])
