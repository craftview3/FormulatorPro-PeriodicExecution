"""
Google Sheets クライアントモジュール

Google Sheets API を使用したスプレッドシート操作を提供します。
設定変更時はこのファイルのみを更新すれば、他のコードに影響しません。
"""

import gspread
from typing import List, Dict, Optional, Any
from datetime import datetime
from google.auth.credentials import Credentials

from .auth import get_gcp_credentials


class SheetsClient:
    """Google Sheets API クライアント"""
    
    def __init__(self, credentials: Optional[Credentials] = None):
        """
        初期化
        
        Args:
            credentials: 認証情報。Noneの場合は自動取得
        """
        if credentials is None:
            credentials, _ = get_gcp_credentials()
        
        self.gc = gspread.authorize(credentials)
        self.credentials = credentials
    
    def open_spreadsheet(self, spreadsheet_id: str):
        """
        スプレッドシートを開く
        
        Args:
            spreadsheet_id: スプレッドシートのID
            
        Returns:
            gspread.Spreadsheet: スプレッドシートオブジェクト
        """
        return self.gc.open_by_key(spreadsheet_id)
    
    def get_or_create_worksheet(self, spreadsheet_id: str, sheet_title: str):
        """
        ワークシートを取得または作成
        
        Args:
            spreadsheet_id: スプレッドシートのID
            sheet_title: シート名
            
        Returns:
            gspread.Worksheet: ワークシートオブジェクト
        """
        try:
            spreadsheet = self.open_spreadsheet(spreadsheet_id)
            return spreadsheet.worksheet(sheet_title)
        except gspread.exceptions.WorksheetNotFound:
            # シートが存在しない場合は作成
            spreadsheet = self.open_spreadsheet(spreadsheet_id)
            return spreadsheet.add_worksheet(title=sheet_title, rows=2000, cols=30)
    
    def get_first_empty_row(self, worksheet) -> int:
        """
        最初の空行の行番号を取得
        
        Args:
            worksheet: ワークシートオブジェクト
            
        Returns:
            int: 最初の空行の行番号（1から開始）
        """
        values = worksheet.get_all_values()
        return len(values) + 1
    
    def append_records(self, spreadsheet_id: str, sheet_title: str, 
                      records: List[Dict], url: str = "") -> None:
        """
        レコードをスプレッドシートに追記
        
        Args:
            spreadsheet_id: スプレッドシートのID
            sheet_title: シート名
            records: 追記するレコードのリスト
            url: PDF URL（各レコードに追加される）
        """
        if not records:
            print("[INFO] 追記対象なし。")
            return
        
        worksheet = self.get_or_create_worksheet(spreadsheet_id, sheet_title)
        today_str = datetime.now().strftime("%Y/%m/%d")
        
        # レコードを行データに変換
        rows = [self._record_to_row(record, today_str, url) for record in records]
        
        # 追記位置を計算
        start_row = self.get_first_empty_row(worksheet)
        end_row = start_row + len(rows) - 1
        
        # データを追記
        worksheet.update(f"A{start_row}:O{end_row}", rows, value_input_option="USER_ENTERED")
        
        print(f"[DONE] {len(rows)} 行を {sheet_title}!A{start_row}:O{end_row} に追記しました。")
    
    def _record_to_row(self, record: Dict, today_str: str, url: str = "") -> List[Any]:
        """
        レコード辞書をスプレッドシート行データに変換
        
        Args:
            record: レコード辞書
            today_str: 今日の日付文字列
            url: PDF URL
            
        Returns:
            List: 行データ（A列からO列まで）
        """
        seibun = record.get("seibunn") or record.get("seibun") or ""
        final_url = record.get("url", "") or url
        
        return [
            0,                           # A: 変更フラグ
            today_str,                   # B: 今日の日付
            0,                           # C: グループID
            seibun,                      # D: 成分名
            "",                          # E: 規制区分（空欄）
            record.get("ryou1", ""),     # F: 配合上限値（一般）
            record.get("条件", ""),      # G: 使用対象・条件（一般）
            record.get("ryou2", ""),     # H:非粘膜・洗い流す上限値
            record.get("ryou3", ""),     # I:非粘膜・洗い流さない上限値
            record.get("ryou4", ""),     # J:粘膜用上限値
            record.get("tanni", ""),     # K:単位
            record.get("bikou", ""),     # L: 備考
            "",                          # M: 予備
            "",                          # N: 予備
            final_url,                   # O: PDF URL
        ]


# 従来のインターフェースとの互換性を保つための関数
def connect_gspread():
    """
    互換性のためのラッパー関数
    
    Returns:
        SheetsClient: Google Sheetsクライアント
    """
    return SheetsClient()


def open_or_create_worksheet(gc, spreadsheet_id: str, title: str):
    """
    互換性のためのラッパー関数
    
    Args:
        gc: SheetsClientインスタンス
        spreadsheet_id: スプレッドシートID
        title: シート名
        
    Returns:
        gspread.Worksheet: ワークシート
    """
    return gc.get_or_create_worksheet(spreadsheet_id, title)


def append_records_to_sheet(ws_or_client, records: List[Dict], 
                           spreadsheet_id: str = "", sheet_title: str = ""):
    """
    互換性のためのラッパー関数
    
    Args:
        ws_or_client: ワークシートまたはSheetsClient
        records: レコードリスト
        spreadsheet_id: スプレッドシートID（SheetsClient使用時）
        sheet_title: シート名（SheetsClient使用時）
    """
    if hasattr(ws_or_client, 'append_records'):
        # 新しいSheetsClientを使用
        ws_or_client.append_records(spreadsheet_id, sheet_title, records)
    else:
        # 従来のワークシート直接操作
        if not records:
            print("[INFO] 追記対象なし。")
            return
        
        today_str = datetime.now().strftime("%Y/%m/%d")
        client = SheetsClient()
        rows = [client._record_to_row(r, today_str) for r in records]
        
        # 最初の空行を取得
        values = ws_or_client.get_all_values()
        start = len(values) + 1
        end = start + len(rows) - 1
        
        ws_or_client.update(f"A{start}:O{end}", rows, value_input_option="USER_ENTERED")
        print(f"[DONE] {len(rows)} 行を {ws_or_client.title}!A{start}:O{end} に追記しました。")