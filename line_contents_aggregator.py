#!/usr/bin/env python3
"""
LINE Contents Aggregator - コンテンツ集計機能

line-menu-yyyy-mmファイルからコンテンツごとの売上実績と情報提供料を計算し、
line-contents-yyyy-mmファイルとして出力する
"""

import pandas as pd
import argparse
import os
import logging
from pathlib import Path
from datetime import datetime
import sys


class LineContentsAggregator:
    """LINE コンテンツ集計クラス"""
    
    def __init__(self, mapping_file_path: str = None):
        self.processed_files = []
        self.errors = []
        self.reiwa_mapping = {}
        self.logger = logging.getLogger(self.__class__.__name__)
        
        # ログ設定（未設定の場合）
        if not self.logger.handlers:
            logging.basicConfig(
                level=logging.INFO,
                format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
            )
        
        # デフォルトのマッピングファイルパス
        if mapping_file_path is None:
            mapping_file_path = r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\line-reiwa-contents-menu.csv"
        
        self.load_reiwa_mapping(mapping_file_path)
    
    def calculate_metrics(self, ios_cost: float, android_cost: float, web_amount: float) -> tuple[int, int]:
        """
        売上金額から実績と情報提供料を計算
        
        Args:
            ios_cost: iOS売上金額
            android_cost: Android売上金額
            web_amount: Web売上金額
            
        Returns:
            tuple: (実績, 情報提供料)
        """
        # 実績 = ((ios_paid_cost + android_paid_cost) × 1.528575 + web_paid_amount) / 1.1
        mobile_total = (ios_cost + android_cost) * 1.528575  
        performance = (mobile_total + web_amount) / 1.1
        
        # 情報提供料 = 実績 * 0.366674
        info_fee = performance * 0.366674
        
        return round(performance), round(info_fee)
    
    def load_reiwa_mapping(self, mapping_file_path: str) -> None:
        """
        reiwaseimeiコンテンツのマッピング情報を読み込み
        
        Args:
            mapping_file_path: マッピングファイルのパス
        """
        try:
            if os.path.exists(mapping_file_path):
                # CSVファイルからマッピング情報を読み込み
                encodings = ['utf-8', 'shift_jis', 'cp932', 'utf-8-sig']
                df = None
                
                for encoding in encodings:
                    try:
                        df = pd.read_csv(mapping_file_path, header=None, names=['item_code', 'sub_group'], encoding=encoding)
                        break
                    except UnicodeDecodeError:
                        continue
                
                if df is not None:
                    # 辞書形式でマッピングを保存
                    self.reiwa_mapping = dict(zip(df['item_code'], df['sub_group']))
                    # print(f"reiwaseimeiマッピングを読み込みました: {len(self.reiwa_mapping)}件")
                else:
                    # print(f"マッピングファイルの読み込みに失敗しました: {mapping_file_path}")
                    pass
            else:
                # print(f"マッピングファイルが見つかりません: {mapping_file_path}")
                pass
                
        except Exception as e:
            self.errors.append(f"マッピングファイル読み込みエラー: {str(e)}")
            # print(f"マッピングファイル読み込みエラー: {str(e)}")
    
    def extract_content_group(self, item_code: str) -> str:
        """
        item_codeから「_」より前の値を抽出してコンテンツグループを取得
        reiwaseimeiの場合はさらに細分化されたサブグループを返す
        
        Args:
            item_code: C列の値（例：kmo2_999, reiwaseimei_002）
            
        Returns:
            str: コンテンツグループ（例：kmo2, amano, chamen等）
        """
        if pd.isna(item_code) or not isinstance(item_code, str):
            return "unknown"
        
        base_group = item_code.split('_')[0] if '_' in item_code else item_code
        
        # reiwaseimeiの場合は細分化されたサブグループを返す
        if base_group == "reiwaseimei" and item_code in self.reiwa_mapping:
            return self.reiwa_mapping[item_code]
        
        return base_group
    
    def process_line_menu_file(self, file_path: str) -> pd.DataFrame:
        """
        line-menu-yyyy-mmファイルを処理してコンテンツ別集計を行う
        
        Args:
            file_path: line-menu-yyyy-mmファイルのパス
            
        Returns:
            pd.DataFrame: 集計結果のデータフレーム
        """
        try:
            # CSVファイルを読み込み（文字コード自動判定）
            encodings = ['utf-8', 'shift_jis', 'cp932', 'utf-8-sig']
            df = None
            
            for encoding in encodings:
                try:
                    df = pd.read_csv(file_path, encoding=encoding)
                    break
                except UnicodeDecodeError:
                    continue
            
            if df is None:
                raise ValueError(f"ファイルの読み込みに失敗しました: {file_path}")
            
            # 必要な列の確認
            required_columns = ['item_name', 'item_code', 'ios_paid_cost', 'android_paid_cost', 'web_paid_amount']
            for col in required_columns:
                if col not in df.columns:
                    raise ValueError(f"必要な列が見つかりません: {col}")
            
            # コンテンツグループの抽出
            df['content_group'] = df['item_code'].apply(self.extract_content_group)
            
            # 数値列を数値型に変換（エラーは0に変換）
            df['ios_paid_cost'] = pd.to_numeric(df['ios_paid_cost'], errors='coerce').fillna(0)
            df['android_paid_cost'] = pd.to_numeric(df['android_paid_cost'], errors='coerce').fillna(0)
            df['web_paid_amount'] = pd.to_numeric(df['web_paid_amount'], errors='coerce').fillna(0)
            
            # コンテンツグループごとの集計
            aggregated = df.groupby('content_group').agg({
                'item_name': 'first',  # 代表的な商品名を取得
                'ios_paid_cost': 'sum',  # iOS売上を合計
                'android_paid_cost': 'sum',  # Android売上を合計
                'web_paid_amount': 'sum',  # Web売上を合計
                'item_code': 'count'  # 件数（行数）をカウント
            }).reset_index()
            
            # 実績と情報提供料の計算
            results = []
            for _, row in aggregated.iterrows():
                ios_cost = row['ios_paid_cost']
                android_cost = row['android_paid_cost']
                web_amount = row['web_paid_amount']
                performance, info_fee = self.calculate_metrics(ios_cost, android_cost, web_amount)
                results.append({
                    'コンテンツ名': row['content_group'],
                    '実績': performance,
                    '情報提供料': info_fee,
                    '売上件数': row['item_code']  # 件数を追加
                })
            
            return pd.DataFrame(results)
            
        except Exception as e:
            self.errors.append(f"ファイル処理エラー {file_path}: {str(e)}")
            return pd.DataFrame()
    
    def generate_contents_filename(self, year: int, month: int) -> str:
        """
        line-contents-yyyy-mmファイル名を生成
        
        Args:
            year: 年
            month: 月
            
        Returns:
            str: ファイル名
        """
        return f"line-contents-{year:04d}-{month:02d}.csv"
    
    def save_contents_file(self, df: pd.DataFrame, output_path: str) -> bool:
        """
        line-contents-yyyy-mmファイルを保存
        
        Args:
            df: 集計結果のデータフレーム
            output_path: 出力ファイルパス
            
        Returns:
            bool: 保存成功可否
        """
        try:
            # ディレクトリが存在しない場合は作成
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            # CSVファイルとして保存（UTF-8 BOM付き）
            df.to_csv(output_path, index=False, encoding='utf-8-sig')
            
            self.processed_files.append(output_path)
            # print(f"保存完了: {output_path}")
            return True
            
        except Exception as e:
            self.errors.append(f"ファイル保存エラー {output_path}: {str(e)}")
            return False
    
    def process_directory(self, base_path: str, year: int = None, month: int = None) -> bool:
        """
        指定されたディレクトリまたは年月のline-menuファイルを処理
        
        Args:
            base_path: ベースディレクトリパス
            year: 処理対象年（Noneの場合は全て）
            month: 処理対象月（Noneの場合は全て）
            
        Returns:
            bool: 処理成功可否
        """
        try:
            base_path = Path(base_path)
            
            if year and month:
                # 特定の年月を処理
                target_dirs = [base_path / str(year) / f"{year:04d}{month:02d}"]
            else:
                # 全ての年月ディレクトリを検索
                target_dirs = []
                for year_dir in base_path.glob('????'):
                    if year_dir.is_dir() and year_dir.name.isdigit():
                        for month_dir in year_dir.glob('??????'):
                            if month_dir.is_dir() and month_dir.name.isdigit():
                                target_dirs.append(month_dir)
            
            success_count = 0
            
            for dir_path in target_dirs:
                if not dir_path.exists():
                    continue
                
                # ディレクトリ名から年月を抽出
                dir_name = dir_path.name
                if len(dir_name) == 6 and dir_name.isdigit():
                    dir_year = int(dir_name[:4])
                    dir_month = int(dir_name[4:])
                else:
                    continue
                
                # line-menu-yyyy-mmファイルを検索
                menu_pattern = f"line-menu-{dir_year:04d}-{dir_month:02d}.csv"
                menu_file = dir_path / menu_pattern
                
                if menu_file.exists():
                    self.logger.info(f"処理中: {menu_file}")
                    
                    # ファイルを処理
                    df = self.process_line_menu_file(str(menu_file))
                    
                    if not df.empty:
                        # 出力ファイル名を生成
                        contents_filename = self.generate_contents_filename(dir_year, dir_month)
                        output_path = dir_path / contents_filename
                        
                        # ファイルを保存
                        if self.save_contents_file(df, str(output_path)):
                            success_count += 1
                        
                        # 結果を表示
                        self.logger.info(f"コンテンツ数: {len(df)}")
                        self.logger.info(f"総実績: {df['実績'].sum():,}円")
                        self.logger.info(f"総情報提供料: {df['情報提供料'].sum():,}円")
                    
                else:
                    self.logger.warning(f"line-menuファイルが見つかりません: {menu_file}")
            
            self.logger.info(f"処理完了: {success_count}件のファイルを処理しました")
            
            if self.errors:
                self.logger.error(f"エラー: {len(self.errors)}件")
                for error in self.errors:
                    self.logger.error(f"  {error}")
            
            return success_count > 0
            
        except Exception as e:
            self.errors.append(f"ディレクトリ処理エラー: {str(e)}")
            return False
    
    def get_stats(self) -> dict:
        """処理統計を取得"""
        return {
            'processed_files': len(self.processed_files),
            'errors': len(self.errors),
            'files': self.processed_files,
            'error_messages': self.errors
        }


def main():
    """メイン関数"""
    # ログ設定
    logger = logging.getLogger(__name__)
    if not logger.handlers:
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
    
    parser = argparse.ArgumentParser(
        description="LINE Contents Aggregator - コンテンツ売上集計ツール"
    )
    
    parser.add_argument(
        "--base-path", "-b",
        type=str,
        default=r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書",
        help="ベースディレクトリパス"
    )
    
    parser.add_argument(
        "--year", "-y",
        type=int,
        help="処理対象年（指定しない場合は全て）"
    )
    
    parser.add_argument(
        "--month", "-m",
        type=int,
        help="処理対象月（指定しない場合は全て）"
    )
    
    parser.add_argument(
        "--file", "-f",
        type=str,
        help="直接処理するline-menuファイルのパス"
    )
    
    args = parser.parse_args()
    
    try:
        aggregator = LineContentsAggregator()
        
        if args.file:
            # 単一ファイルを処理
            if not os.path.exists(args.file):
                logger.error(f"ファイルが存在しません: {args.file}")
                return 1
            
            logger.info(f"ファイル処理中: {args.file}")
            df = aggregator.process_line_menu_file(args.file)
            
            if not df.empty:
                # 出力ファイル名を生成（同じディレクトリに保存）
                file_path = Path(args.file)
                output_path = file_path.parent / f"line-contents-{datetime.now().strftime('%Y-%m')}.csv"
                
                if aggregator.save_contents_file(df, str(output_path)):
                    logger.info(f"コンテンツ数: {len(df)}")
                    logger.info(f"総実績: {df['実績'].sum():,}円")
                    logger.info(f"総情報提供料: {df['情報提供料'].sum():,}円")
                else:
                    logger.error("ファイル保存に失敗しました")
                    return 1
            else:
                logger.warning("処理可能なデータがありませんでした")
                return 1
        
        else:
            # ディレクトリを処理
            success = aggregator.process_directory(args.base_path, args.year, args.month)
            if not success:
                logger.error("処理に失敗しました")
                return 1
        
        # 統計情報を表示
        stats = aggregator.get_stats()
        logger.info(f"\n=== 処理統計 ===")
        logger.info(f"処理ファイル数: {stats['processed_files']}")
        logger.info(f"エラー数: {stats['errors']}")
        
        return 0
        
    except KeyboardInterrupt:
        logger.info("\nユーザーによって処理が中断されました")
        return 1
        
    except Exception as e:
        logger.error(f"予期しないエラーが発生しました: {e}")
        return 1


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)