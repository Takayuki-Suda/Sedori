import time
from spreadsheet_utils import open_worksheet
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

# Chromeのオプション設定（ヘッドレスモードを強化）
options = Options()
options.headless = True  # ヘッドレスモードを有効にする
options.add_argument("--no-sandbox")  # サンドボックスを無効化
options.add_argument("--disable-dev-shm-usage")  # システムリソース使用の制限を無視
options.add_argument("--disable-gpu")  # GPUを無効化
options.add_argument("--window-size=1920x1080")  # ウィンドウサイズを最大化
options.add_argument("--disable-extensions")  # 拡張機能を無効化
options.add_argument("--disable-plugins")  # プラグインを無効化

# 必要なJavaScriptを有効にしたまま、ページの不要な要素を非表示にする
prefs = {
    "profile.managed_default_content_settings.styles": 2,  # CSSを無効
    "profile.managed_default_content_settings.images": 2,  # 画像を無効
    "profile.managed_default_content_settings.plugins": 2,  # プラグインを無効
    "profile.managed_default_content_settings.popups": 2,  # ポップアップを無効
    "download.default_directory": "/dev/null",  # 自動ダウンロードを無効化（ダウンロードディレクトリを設定）
    "profile.default_content_setting_values.notifications": 2,  # 通知を無効化
    "profile.managed_default_content_settings.media_stream": 2,  # メディアストリームを無効化
    "profile.managed_default_content_settings.geolocation": 2,  # 位置情報を無効化
    "profile.managed_default_content_settings.javascript": 1,  # JavaScriptを有効化（無効化から変更）
}
options.add_experimental_option("prefs", prefs)

# Chromeドライバーのセットアップ
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# E列（4列目）のURLを取得
all_data = worksheet.get_all_values()  # スプレッドシート全体を取得
urls_h = [row[4] for row in all_data[3:]]  # E列（4列目）のURL

# メルカリの金額を取得する関数
def get_mercari_price(url, row, column):
    try:
        # 対象のURLにアクセス
        driver.get(url)

        # 不要な要素を非表示にするJavaScriptを実行
        driver.execute_script("""
            // 必要な要素のみ表示する
            let filterHeading = document.querySelector('.merHeading');
            if (filterHeading) filterHeading.style.display = 'none';

            // サイドバーを非表示にする
            let sidebars = document.querySelectorAll('.sidebar');
            sidebars.forEach(sidebar => sidebar.style.display = 'none');

            // ヘッダーやその他不要な要素を非表示にする
            let header = document.querySelector('header');
            if (header) header.style.display = 'none';

            // 入力要素（inputタグ）を非表示にする
            let inputElements = document.querySelectorAll('input');
            inputElements.forEach(input => input.style.display = 'none');
        """)

        # 価格の要素が表示されるまで待つ
        price_element = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".number__6b270ca7"))
        )

        # 価格情報の要素を取得
        first_price = int(price_element.text.replace('¥', '').replace(',', ''))

        # スプレッドシートに金額を入力
        worksheet.update_cell(row, column, first_price)

        return None  # 成功時は何も返さない

    except Exception as e:
        return f"エラーが発生しました (URL: {url}): {str(e)}"


# 並列処理でH列のURLを処理
with ThreadPoolExecutor(max_workers=40) as executor:  # スレッド数を最大40に設定
    future_to_url_h = {executor.submit(get_mercari_price, url, row, 6): (url, row) for row, url in enumerate(urls_h, start=4)}

    # 結果を順番に取得
    for future in as_completed(future_to_url_h):
        result = future.result()
        if result:  # None以外の結果のみ表示
            print(result)

# 全ての金額入力が完了したメッセージを表示
print("金額の入力が完了しました。")

# ブラウザを閉じる
driver.quit()
