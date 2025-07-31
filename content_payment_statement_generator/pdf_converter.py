"""
PDF変換モジュール

ExcelファイルのPDF変換を行います。
"""

import os
from pathlib import Path
from typing import Optional
import logging
import xlwings as xw


class PDFConverter:
    """PDF変換クラス"""
    
    def __init__(self):
        """PDF変換クラスを初期化"""
        self.logger = logging.getLogger(__name__)
    
    def convert_excel_to_pdf(self, excel_path: str) -> str:
        """ExcelファイルをPDF形式に変換"""
        try:
            excel_file = Path(excel_path)
            
            if not excel_file.exists():
                raise FileNotFoundError(f"Excelファイルが見つかりません: {excel_path}")
            
            # PDF出力パスを生成
            pdf_path = excel_file.with_suffix('.pdf')
            
            # xlwingsを使用してPDF変換
            app = None
            wb = None
            
            try:
                # Excelアプリケーションを起動（表示しない）
                app = xw.App(visible=False)
                
                # Excelファイルを開く
                wb = app.books.open(str(excel_file))
                
                # アクティブなワークシートを取得
                ws = wb.sheets.active
                
                # PDFとして保存
                ws.to_pdf(str(pdf_path))
                
                self.logger.info(f"PDF変換完了: {pdf_path}")
                return str(pdf_path)
                
            finally:
                # リソースをクリーンアップ
                if wb:
                    wb.close()
                if app:
                    app.quit()
            
        except Exception as e:
            self.logger.error(f"PDF変換エラー: {e}")
            raise
    
    def convert_excel_to_pdf_com(self, excel_path: str) -> str:
        """COM経由でExcelファイルをPDF形式に変換（Windows用）"""
        try:
            import win32com.client
            
            excel_file = Path(excel_path)
            
            if not excel_file.exists():
                raise FileNotFoundError(f"Excelファイルが見つかりません: {excel_path}")
            
            # PDF出力パスを生成
            pdf_path = excel_file.with_suffix('.pdf')
            
            # COMオブジェクトを作成
            excel_app = None
            
            try:
                excel_app = win32com.client.Dispatch("Excel.Application")
                excel_app.Visible = False
                excel_app.DisplayAlerts = False
                
                # Excelファイルを開く
                workbook = excel_app.Workbooks.Open(str(excel_file.absolute()))
                
                # PDFとして保存
                workbook.ExportAsFixedFormat(
                    Type=0,  # xlTypePDF
                    Filename=str(pdf_path.absolute()),
                    Quality=0,  # xlQualityStandard
                    IncludeDocProperties=True,  # IncludeDocPropsではなくIncludeDocProperties
                    IgnorePrintAreas=False,
                    OpenAfterPublish=False
                )
                
                # ワークブックを閉じる
                workbook.Close(SaveChanges=False)
                
                self.logger.info(f"PDF変換完了 (COM): {pdf_path}")
                return str(pdf_path)
                
            finally:
                # Excelアプリケーションを終了
                if excel_app:
                    excel_app.Quit()
                    excel_app = None
            
        except ImportError:
            self.logger.warning("win32com.clientが利用できません。xlwingsを使用します。")
            return self.convert_excel_to_pdf(excel_path)
        except Exception as e:
            self.logger.error(f"PDF変換エラー (COM): {e}")
            raise
    
    def validate_pdf_output(self, pdf_path: str) -> bool:
        """PDF出力の検証"""
        try:
            pdf_file = Path(pdf_path)
            
            # ファイルの存在確認
            if not pdf_file.exists():
                self.logger.error(f"PDFファイルが存在しません: {pdf_path}")
                return False
            
            # ファイルサイズの確認
            file_size = pdf_file.stat().st_size
            if file_size == 0:
                self.logger.error(f"PDFファイルが空です: {pdf_path}")
                return False
            
            # 最小サイズの確認（1KB以下は異常とみなす）
            if file_size < 1024:
                self.logger.warning(f"PDFファイルサイズが小さすぎます: {file_size} bytes")
                return False
            
            self.logger.info(f"PDF検証完了: {pdf_path} ({file_size} bytes)")
            return True
            
        except Exception as e:
            self.logger.error(f"PDF検証エラー: {e}")
            return False
    
    def convert_and_validate(self, excel_path: str) -> Optional[str]:
        """Excel to PDF変換を実行"""
        try:
            # PDF変換を実行
            pdf_path = self.convert_excel_to_pdf_com(excel_path)
            return pdf_path
                
        except Exception as e:
            self.logger.error(f"PDF変換エラー: {e}")
            return None
    
    def cleanup_temp_files(self, directory: str, pattern: str = "~$*.tmp") -> None:
        """一時ファイルをクリーンアップ"""
        try:
            temp_dir = Path(directory)
            
            if not temp_dir.exists():
                return
            
            # 一時ファイルを検索して削除
            temp_files = list(temp_dir.glob(pattern))
            
            for temp_file in temp_files:
                try:
                    temp_file.unlink()
                    self.logger.debug(f"一時ファイルを削除しました: {temp_file}")
                except Exception as e:
                    self.logger.warning(f"一時ファイル削除エラー: {temp_file} - {e}")
            
            if temp_files:
                self.logger.info(f"一時ファイルクリーンアップ完了: {len(temp_files)}件")
                
        except Exception as e:
            self.logger.error(f"一時ファイルクリーンアップエラー: {e}")