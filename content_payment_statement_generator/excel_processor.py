"""
Excelファイル処理モジュール

Excelテンプレートの複製と明細データの書き込みを行います。
"""

import shutil
from pathlib import Path
from datetime import datetime, timedelta
from typing import List, Dict, Optional
import logging
import openpyxl
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from .config_manager import ConfigManager
from .data_models import SalesRecord


class ExcelProcessor:
    """Excelファイル処理クラス"""
    
    def __init__(self, config_manager: ConfigManager):
        """Excel処理クラスを初期化"""
        self.config = config_manager
        self.logger = logging.getLogger(__name__)
    
    def copy_template(self, template_name: str, output_path: str, target_month: str) -> str:
        """テンプレートファイルを指定の出力パスに複製"""
        try:
            # テンプレートファイルのパスを取得
            template_path = self.config.get_template_file_by_name(template_name)
            
            if not template_path or not Path(template_path).exists():
                raise FileNotFoundError(f"テンプレートファイルが見つかりません: {template_name}")
            
            # 出力ファイル名を生成 (YYYYMM_元ファイル名)
            template_file = Path(template_path)
            output_filename = f"{target_month}_{template_file.name}"
            output_file_path = Path(output_path) / output_filename
            
            # ディレクトリを作成
            output_file_path.parent.mkdir(parents=True, exist_ok=True)
            
            # ファイルを複製
            shutil.copy2(template_path, output_file_path)
            
            self.logger.info(f"テンプレートファイルを複製しました: {output_file_path}")
            return str(output_file_path)
            
        except Exception as e:
            self.logger.error(f"テンプレートファイル複製エラー: {e}")
            raise
    
    def write_payment_date(self, workbook_path: str, target_month: str) -> None:
        """S3セルに対象年月の翌月5日の日付を記入"""
        try:
            # Excelファイルを開く
            workbook = openpyxl.load_workbook(workbook_path)
            worksheet = workbook.active
            
            # 対象年月から翌月5日を計算
            year = int(target_month[:4])
            month = int(target_month[4:])
            
            # 翌月を計算
            if month == 12:
                next_year = year + 1
                next_month = 1
            else:
                next_year = year
                next_month = month + 1
            
            # 翌月5日の日付を作成
            payment_date = datetime(next_year, next_month, 5)
            
            # S3セルに日付を設定
            worksheet['S3'] = payment_date
            
            # ファイルを保存
            workbook.save(workbook_path)
            
            self.logger.info(f"支払日を設定しました: S3セル = {payment_date.strftime('%Y/%m/%d')}")
            
        except Exception as e:
            self.logger.error(f"支払日設定エラー: {e}")
            raise
        finally:
            if 'workbook' in locals():
                workbook.close()
    
    def write_statement_details(self, workbook_path: str, sales_records: List[SalesRecord]) -> None:
        """23行目以降に明細データを書き込み"""
        try:
            # Excelファイルを開く
            workbook = openpyxl.load_workbook(workbook_path)
            worksheet = workbook.active
            
            # 開始行（23行目）
            start_row = 23
            
            for i, record in enumerate(sales_records):
                row_num = start_row + i
                
                # 各列にデータを設定
                self._write_record_to_row(worksheet, row_num, record)
            
            # ファイルを保存
            workbook.save(workbook_path)
            
            self.logger.info(f"明細データを書き込みました: {len(sales_records)}件")
            
        except Exception as e:
            self.logger.error(f"明細データ書き込みエラー: {e}")
            raise
        finally:
            if 'workbook' in locals():
                workbook.close()
    
    def _write_record_to_row(self, worksheet: Worksheet, row_num: int, record: SalesRecord) -> None:
        """1つのSalesRecordを指定行に書き込み"""
        try:
            # A列：対象年月の翌月5日
            year = int(record.target_month[:4])
            month = int(record.target_month[4:])
            
            if month == 12:
                next_year = year + 1
                next_month = 1
            else:
                next_year = year
                next_month = month + 1
            
            payment_date = datetime(next_year, next_month, 5)
            
            # セルへの書き込み（マージセルの場合は上位セルに書き込む）
            self._safe_write_cell(worksheet, f'A{row_num}', payment_date)
            self._safe_write_cell(worksheet, f'D{row_num}', record.platform)
            self._safe_write_cell(worksheet, f'G{row_num}', record.content_name)
            self._safe_write_cell(worksheet, f'M{row_num}', record.target_month)
            self._safe_write_cell(worksheet, f'S{row_num}', record.performance)
            self._safe_write_cell(worksheet, f'Y{row_num}', record.information_fee)
            self._safe_write_cell(worksheet, f'AC{row_num}', record.rate)
            
            self.logger.debug(f"行 {row_num} にレコードを書き込みました: {record.content_name}")
            
        except Exception as e:
            self.logger.error(f"レコード書き込みエラー (行 {row_num}): {e}")
            raise
    
    def _safe_write_cell(self, worksheet: Worksheet, cell_address: str, value) -> None:
        """セルに安全に値を書き込み（マージセル対応）"""
        try:
            # 直接的なアプローチ：マージセルエラーをキャッチして処理
            try:
                worksheet[cell_address].value = value
                self.logger.debug(f"セル {cell_address} に正常に書き込みました: {value}")
            except Exception as write_error:
                if "'MergedCell' object attribute 'value' is read-only" in str(write_error):
                    # マージセルの場合の処理
                    self._handle_merged_cell_write(worksheet, cell_address, value)
                else:
                    # 他のエラーの場合
                    self.logger.warning(f"セル {cell_address} への書き込みエラー: {write_error}")
            
        except Exception as e:
            self.logger.warning(f"セル {cell_address} への書き込み処理エラー: {e}")
    
    def _handle_merged_cell_write(self, worksheet: Worksheet, cell_address: str, value) -> None:
        """マージセル書き込み処理"""
        try:
            # マージされた範囲を検索
            from openpyxl.utils import coordinate_to_tuple
            row, col = coordinate_to_tuple(cell_address)
            
            for merged_range in worksheet.merged_cells.ranges:
                if merged_range.min_row <= row <= merged_range.max_row and \
                   merged_range.min_col <= col <= merged_range.max_col:
                    # マージ範囲の左上セルに書き込み
                    top_left_address = f"{merged_range.start_cell.column_letter}{merged_range.start_cell.row}"
                    worksheet[top_left_address].value = value
                    self.logger.debug(f"マージセル {cell_address} の代わりに {top_left_address} に書き込みました: {value}")
                    return
            
            # マージ範囲が見つからない場合は、強制的に書き込みを試行
            self.logger.warning(f"マージセル {cell_address} の範囲が特定できません。値: {value}")
            
        except Exception as e:
            self.logger.warning(f"マージセル処理エラー {cell_address}: {e}")
    
    def calculate_payment_amount(self, performance: float, rate: float) -> float:
        """支払額を計算（実績 × 料率）"""
        try:
            return performance * rate
        except Exception as e:
            self.logger.error(f"支払額計算エラー: {e}")
            return 0.0
    
    def process_excel_file(
        self, 
        template_name: str, 
        sales_records: List[SalesRecord], 
        target_month: str
    ) -> str:
        """Excelファイルの完全処理（テンプレート複製 + データ書き込み）"""
        try:
            # 出力ディレクトリを取得
            year = target_month[:4]
            month = target_month[4:]
            output_dir = self.config.get_output_directory(year, month)
            
            # テンプレートを複製
            excel_path = self.copy_template(template_name, output_dir, target_month)
            
            # 支払日を設定
            self.write_payment_date(excel_path, target_month)
            
            # 明細データを書き込み
            self.write_statement_details(excel_path, sales_records)
            
            self.logger.info(f"Excelファイル処理完了: {excel_path}")
            return excel_path
            
        except Exception as e:
            self.logger.error(f"Excelファイル処理エラー: {e}")
            raise
    
    def validate_excel_structure(self, workbook_path: str) -> bool:
        """Excelファイルの構造を検証"""
        try:
            workbook = openpyxl.load_workbook(workbook_path)
            worksheet = workbook.active
            
            # 必要なセルの存在確認
            required_cells = ['S3']  # 支払日セル
            
            for cell in required_cells:
                if worksheet[cell].value is None:
                    self.logger.warning(f"必要なセル {cell} が空です")
                    return False
            
            self.logger.info("Excelファイル構造の検証が完了しました")
            return True
            
        except Exception as e:
            self.logger.error(f"Excelファイル構造検証エラー: {e}")
            return False
        finally:
            if 'workbook' in locals():
                workbook.close()