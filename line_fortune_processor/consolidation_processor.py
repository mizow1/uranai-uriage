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


class ConsolidationProcessor:
    """CSV統合処理クラス"""
    
    def __init__(self):
        """統合処理を初期化"""
        self.logger = logging.getLogger(__name__)
        
    def consolidate_csv_files(self, directory: Path, output_filename: str) -> bool:
        """
        ディレクトリ内のすべてのCSVファイルを統合
        
        Args:
            directory: 統合対象のディレクトリ
            output_filename: 統合後のファイル名
            
        Returns:
            bool: 統合が成功した場合True
        """
        try:
            if not isinstance(directory, Path):
                directory = Path(directory)
                
            if not directory.exists():
                self.logger.error(f"ディレクトリが存在しません: {directory}")
                return False
                
            # CSVファイルを検索
            csv_files = list(directory.glob("*.csv"))
            
            # 統合対象ファイルから出力ファイル自体を除外
            output_path = directory / output_filename
            csv_files = [f for f in csv_files if f != output_path]
            
            if not csv_files:
                self.logger.warning(f"統合対象のCSVファイルが見つかりませんでした: {directory}")
                return False
                
            # ファイルを日付順にソート
            csv_files = self._sort_csv_files_by_date(csv_files)
            
            # CSVファイルを統合
            consolidated_data = self._merge_csv_files(csv_files)
            
            if consolidated_data is None:
                self.logger.error("CSVファイルの統合に失敗しました")
                return False
                
            # 統合データを保存
            if self._save_consolidated_data(consolidated_data, output_path):
                self.logger.info(f"統合ファイルを作成しました: {output_path}")
                return True
            else:
                return False
                
        except Exception as e:
            self.logger.error(f"CSV統合中にエラーが発生しました: {e}")
            return False
    
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
    
    def _merge_csv_files(self, csv_files: List[Path]) -> Optional[pd.DataFrame]:
        """
        複数のCSVファイルをマージ
        
        Args:
            csv_files: マージ対象のCSVファイルのリスト
            
        Returns:
            pd.DataFrame: マージされたデータ、失敗時はNone
        """
        try:
            all_data = []
            
            for csv_file in csv_files:
                try:
                    # CSVファイルを読み込み
                    df = pd.read_csv(csv_file, encoding='utf-8')
                    
                    # データが空でない場合のみ追加
                    if not df.empty:
                        all_data.append(df)
                        self.logger.debug(f"CSVファイルを読み込みました: {csv_file} ({len(df)} 行)")
                    else:
                        self.logger.warning(f"空のCSVファイルです: {csv_file}")
                        
                except Exception as e:
                    self.logger.error(f"CSVファイル読み込み中にエラーが発生しました: {csv_file}, {e}")
                    continue
                    
            if not all_data:
                self.logger.error("読み込めるCSVファイルがありませんでした")
                return None
                
            # データを統合
            consolidated_df = pd.concat(all_data, ignore_index=True)
            
            self.logger.info(f"CSVファイルを統合しました: {len(csv_files)} 個のファイル, 合計 {len(consolidated_df)} 行")
            return consolidated_df
            
        except Exception as e:
            self.logger.error(f"CSV統合中にエラーが発生しました: {e}")
            return None
    
    def _save_consolidated_data(self, data: pd.DataFrame, output_path: Path) -> bool:
        """
        統合データをファイルに保存
        
        Args:
            data: 統合されたデータ
            output_path: 保存先ファイルパス
            
        Returns:
            bool: 保存が成功した場合True
        """
        try:
            # 既存ファイルが存在する場合はバックアップを作成
            if output_path.exists():
                backup_path = output_path.with_suffix(f".backup{output_path.suffix}")
                output_path.replace(backup_path)
                self.logger.info(f"既存ファイルをバックアップしました: {backup_path}")
            
            # データを保存
            data.to_csv(output_path, index=False, encoding='utf-8')
            
            self.logger.info(f"統合データを保存しました: {output_path} ({len(data)} 行)")
            return True
            
        except Exception as e:
            self.logger.error(f"統合データ保存中にエラーが発生しました: {e}")
            return False
    
    def generate_monthly_filename(self, year: int, month: int) -> str:
        """
        月次統合ファイル名を生成
        
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