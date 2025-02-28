# spreadsheet_utils.py
import gspread
from google.oauth2.service_account import Credentials

def open_worksheet(spreadsheet_id = '1EGaKaTiWd0S-tbfmRvGHKAAdlUmSstbZSys5GvuoNp8', worksheet_index=1, credentials_file='../amazonurl-c9f4024ab77f.json'):
    """
    スプレッドシートを開く関数
    :param spreadsheet_id: スプレッドシートのID
    :param worksheet_index: 開きたいシートのインデックス（デフォルトは1）
    :param credentials_file: サービスアカウントのJSONファイルのパス（デフォルトは 'amazonurl-c9f4024ab77f.json'）
    :return: 指定したワークシート
    """
    # サービスアカウントの認証
    credentials = Credentials.from_service_account_file(
        credentials_file,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    
    # Google Sheetsに接続
    gc = gspread.authorize(credentials)
    
    # スプレッドシートを開く
    worksheet = gc.open_by_key(spreadsheet_id).get_worksheet(worksheet_index)  # インデックスは0から始まるので、デフォルトで1番目のシート
    return worksheet
