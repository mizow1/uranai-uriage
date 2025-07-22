"""
CSV統合処理モジュール
"""

import csv
import logging
from typing import List, Dict, Any, Optional
from pathlib import Path
import pandas as pd
import os
from datetime import date

from .error_handler import retry_on_error, handle_errors, ErrorHandler, RetryableError, FatalError, ErrorType
from .constants import ConsolidationConstants
from .messages import MessageFormatter


class ConsolidationProcessor:
    """CSV統合処理クラス"""
    
    def __init__(self):
        """統合処理を初期化"""
        self.logger = logging.getLogger(__name__)
        self.error_handler = ErrorHandler(self.logger)
        
    @handle_errors("CSV統合処理")
    def consolidate_csv_files(self, directory: Path, output_filename: str) -> bool:
        """
        ディレクトリ内のすべてのCSVファイルを統合
        要件: 月次フォルダ内のすべてのCSVファイルのデータを含む統合CSVファイルを生成
        
        Args:
            directory: 統合対象のディレクトリ
            output_filename: 統合後のファイル名
            
        Returns:
            bool: 統合が成功した場合True
            
        Raises:
            FatalError: ディレクトリが存在しない場合
        """
        if not isinstance(directory, Path):
            directory = Path(directory)
            
        if not directory.exists():
            raise FatalError(f"統合対象ディレクトリが存在しません: {directory}", ErrorType.FILE_SYSTEM)
            
        # CSVファイルを検索
        csv_files = list(directory.glob("*.csv"))
        
        # 要件: 統合対象ファイルから除外するファイルを定義
        output_path = directory / output_filename
        exclude_patterns = ConsolidationConstants.CONSOLIDATION_EXCLUDE_PATTERNS
        
        # 出力ファイル自体と集計ファイルを除外
        csv_files = [f for f in csv_files if f != output_path and 
                    not any(pattern in f.name for pattern in exclude_patterns)]
        
        if not csv_files:
            self.logger.warning(f"統合対象のCSVファイルが見つかりませんでした: {directory}")
            return False
            
        self.logger.info(MessageFormatter.get_consolidation_message(
            "csv_files_found", count=len(csv_files)
        ))
        for csv_file in csv_files:
            self.logger.debug(f"  - {csv_file.name}")
            
        # ファイルを日付順にソート
        csv_files = self._sort_csv_files_by_date(csv_files)
        
        # CSVファイルを統合
        consolidated_data = self._merge_csv_files(csv_files)
        
        if consolidated_data is None or consolidated_data.empty:
            self.logger.error("CSVファイルの統合に失敗しました")
            return False
            
        # 要件: 統合ファイルがすでに存在する場合は更新版で上書き
        return self._save_consolidated_data(consolidated_data, output_path, overwrite=True)
    
    def _sort_csv_files_by_date(self, csv_files: List[Path]) -> List[Path]:
        """
        CSVファイルを日付順にソート
        
        Args:
            csv_files: CSVファイルのリスト
            
        Returns:
            List[Path]: ソートされたCSVファイルのリスト
        """
        def extract_date_from_filename(filepath: Path) -> tuple:
            """ファイル名から日付を抽出してソート用のタプルを返す"""
            try:
                filename = filepath.stem
                # yyyy-mm-dd形式の日付を探す
                import re
                date_match = re.search(r'(\d{4})-(\d{2})-(\d{2})', filename)
                if date_match:
                    return (int(date_match.group(1)), int(date_match.group(2)), int(date_match.group(3)))
                else:
                    # 日付が見つからない場合はファイル名でソート
                    return (0, 0, 0, filename)
            except:
                return (0, 0, 0, filepath.name)
        
        try:
            sorted_files = sorted(csv_files, key=extract_date_from_filename)
            self.logger.info(f"CSVファイルを日付順にソートしました: {len(sorted_files)} 個のファイル")
            return sorted_files
        except Exception as e:
            self.logger.warning(f"ファイルソート中にエラーが発生しました: {e}")
            return csv_files
    
    @handle_errors("CSVファイルマージ")
    def _merge_csv_files(self, csv_files: List[Path]) -> Optional[pd.DataFrame]:
        """
        複数のCSVファイルをマージ
        要件: ヘッダー行を統合ファイルの先頭に一度だけ維持
        
        Args:
            csv_files: マージ対象のCSVファイルのリスト
            
        Returns:
            pd.DataFrame: マージされたデータ、失敗時はNone
        """
        all_data = []
        headers_validated = False
        expected_headers = None
        
        for csv_file in csv_files:
            try:
                # CSVファイルを読み込み
                df = pd.read_csv(csv_file, encoding='utf-8')
                
                # 空のファイルをスキップ
                if df.empty:
                    self.logger.warning(f"空のCSVファイルです: {csv_file}")
                    continue
                
                # 要件: ヘッダーの一貫性をチェック
                if not headers_validated:
                    expected_headers = df.columns.tolist()
                    headers_validated = True
                    self.logger.debug(f"期待ヘッダー: {expected_headers}")
                elif df.columns.tolist() != expected_headers:
                    self.logger.warning(f"ヘッダーが一致しません: {csv_file}")
                    self.logger.warning(f"  期待: {expected_headers}")
                    self.logger.warning(f"  実際: {df.columns.tolist()}")
                    # ヘッダーが不一致でも処理を継続
                
                all_data.append(df)
                self.logger.debug(f"CSVファイルを読み込みました: {csv_file} ({len(df)} 行)")
                
            except pd.errors.EmptyDataError:
                self.logger.warning(f"CSVファイルが空です: {csv_file}")
                continue
            except pd.errors.ParserError as e:
                self.logger.error(f"CSVファイルのパースエラー: {csv_file}, {e}")
                continue
            except Exception as e:
                self.logger.error(f"CSVファイル読み込み中にエラーが発生しました: {csv_file}, {e}")
                continue
                
        if not all_data:
            self.logger.error("読み込めるCSVファイルがありませんでした")
            return None
            
        # 要件: データを統合（ヘッダーは一度だけ保持）
        consolidated_df = pd.concat(all_data, ignore_index=True, sort=False)
        
        self.logger.info(MessageFormatter.get_consolidation_message(
            "consolidation_complete", output_file="consolidated_data", files=len(csv_files)
        ))
        return consolidated_df
    
    def _aggregate_by_item_code(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        item_codeごとにデータを集計
        
        Args:
            df: 統合された生データ
            
        Returns:
            pd.DataFrame: item_codeごとに集計されたデータ
        """
        try:
            if df.empty:
                return df
            
            # 数値列と非数値列を分離
            numeric_columns = df.select_dtypes(include=[float, int]).columns.tolist()
            non_numeric_columns = df.select_dtypes(exclude=[float, int]).columns.tolist()
            
            # item_codeが存在しない場合はそのまま返す
            if 'item_code' not in df.columns:
                self.logger.warning("item_code列が見つかりません。集計をスキップします。")
                return df
            
            # item_codeごとにグループ化
            grouped = df.groupby('item_code')
            
            # 集計方法を定義
            agg_dict = {}
            
            # 数値列は合計
            for col in numeric_columns:
                agg_dict[col] = 'sum'
            
            # 非数値列（item_code以外）は最初の値を取得
            for col in non_numeric_columns:
                if col != 'item_code':
                    agg_dict[col] = 'first'
            
            # 集計実行
            aggregated = grouped.agg(agg_dict).reset_index()
            
            self.logger.info(f"item_codeごとの集計完了: {len(df)} 行 → {len(aggregated)} 行")
            
            return aggregated
            
        except Exception as e:
            self.logger.error(f"item_codeごとの集計中にエラーが発生しました: {e}")
            return df
    
    @retry_on_error(max_retries=3, base_delay=1.0)
    def _save_consolidated_data(self, data: pd.DataFrame, output_path: Path, overwrite: bool = True) -> bool:
        """
        統合データをファイルに保存
        要件: 統合ファイルがすでに存在する場合は更新版で上書き
        
        Args:
            data: 統合されたデータ
            output_path: 保存先ファイルパス
            overwrite: 上書きを許可するか
            
        Returns:
            bool: 保存が成功した場合True
            
        Raises:
            RetryableError: ファイル保存に失敗した場合
        """
        try:
            # 要件: 既存ファイルが存在する場合は更新版で上書き
            if output_path.exists():
                if overwrite:
                    # バックアップを作成
                    import time
                    timestamp = time.strftime("%Y%m%d_%H%M%S")
                    backup_path = output_path.with_suffix(f".backup_{timestamp}{output_path.suffix}")
                    output_path.rename(backup_path)
                    self.logger.info(f"既存ファイルをバックアップしました: {backup_path}")
                else:
                    raise FatalError(f"ファイルが既に存在し、上書きが禁止されています: {output_path}", ErrorType.FILE_SYSTEM)
            
            # ディレクトリの存在確認
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            # データを保存
            data.to_csv(output_path, index=False, encoding='utf-8')
            
            # 保存後の検証
            if not output_path.exists() or output_path.stat().st_size == 0:
                raise RetryableError(f"統合ファイルの保存に失敗しました: {output_path}", ErrorType.FILE_SYSTEM)
            
            self.logger.info(MessageFormatter.get_consolidation_message(
            "consolidation_complete", output_file=output_path.name, files=len(data)
        ))
            return True
            
        except PermissionError as e:
            raise RetryableError(f"ファイル保存の権限がありません: {output_path}", ErrorType.FILE_SYSTEM, e)
        except OSError as e:
            if "No space left" in str(e):
                raise FatalError(f"ディスク容量不足: {output_path}", ErrorType.FILE_SYSTEM, e)
            else:
                raise RetryableError(f"統合データ保存中にOSエラーが発生: {e}", ErrorType.FILE_SYSTEM, e)
        except Exception as e:
            error_type = self.error_handler.classify_error(e)
            raise RetryableError(f"統合データ保存中にエラーが発生: {e}", error_type, e)
    
    def generate_monthly_filename(self, year: int, month: int) -> str:
        """
        月次統合ファイル名を生成
        要件: 「line-menu-yyyy-mm.csv」と名付ける
        
        Args:
            year: 年
            month: 月
            
        Returns:
            str: 生成されたファイル名
        """
        return f"line-menu-{year:04d}-{month:02d}.csv"
    
    def extract_year_month_from_directory(self, directory: Path) -> Optional[tuple]:
        """
        ディレクトリ名から年月を抽出
        
        Args:
            directory: ディレクトリパス
            
        Returns:
            tuple: (年, 月) のタプル、抽出失敗時はNone
        """
        try:
            # ディレクトリ名からyyyymm形式を抽出
            dir_name = directory.name
            if len(dir_name) == 6 and dir_name.isdigit():
                year = int(dir_name[:4])
                month = int(dir_name[4:])
                return (year, month)
            else:
                self.logger.warning(f"ディレクトリ名から年月を抽出できませんでした: {dir_name}")
                return None
                
        except Exception as e:
            self.logger.error(f"年月抽出中にエラーが発生しました: {e}")
            return None
    
    def consolidate_monthly_data(self, directory: Path) -> bool:
        """
        月次データを統合
        
        Args:
            directory: 統合対象のディレクトリ
            
        Returns:
            bool: 統合が成功した場合True
        """
        try:
            # ディレクトリ名から年月を抽出
            year_month = self.extract_year_month_from_directory(directory)
            if not year_month:
                return False
                
            year, month = year_month
            
            # 統合ファイル名を生成
            output_filename = self.generate_monthly_filename(year, month)
            
            # CSVファイルを統合
            return self.consolidate_csv_files(directory, output_filename)
            
        except Exception as e:
            self.logger.error(f"月次データ統合中にエラーが発生しました: {e}")
            return False
    
    def validate_csv_structure(self, csv_file: Path) -> bool:
        """
        CSVファイルの構造を検証
        
        Args:
            csv_file: 検証対象のCSVファイル
            
        Returns:
            bool: 構造が正しい場合True
        """
        try:
            df = pd.read_csv(csv_file, encoding='utf-8')
            
            # 基本的な検証
            if df.empty:
                self.logger.warning(f"空のCSVファイルです: {csv_file}")
                return False
                
            # 必要に応じて列の検証を追加
            # 例: 必須列の存在確認
            # required_columns = ['column1', 'column2']
            # if not all(col in df.columns for col in required_columns):
            #     self.logger.error(f"必須列が不足しています: {csv_file}")
            #     return False
            
            self.logger.debug(f"CSVファイルの構造検証が完了しました: {csv_file}")
            return True
            
        except Exception as e:
            self.logger.error(f"CSV構造検証中にエラーが発生しました: {csv_file}, {e}")
            return False