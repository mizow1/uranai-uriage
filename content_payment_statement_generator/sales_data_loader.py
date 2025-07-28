"""
売上データローダーモジュール

各種CSVファイルからのデータ読み込みと統合を行います。
"""

import pandas as pd
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import logging
from .config_manager import ConfigManager
from .data_models import SalesRecord


class SalesDataLoader:
    """売上データ読み込みクラス"""
    
    def __init__(self, config_manager: ConfigManager):
        """データローダーを初期化"""
        self.config = config_manager
        self.logger = logging.getLogger(__name__)
    
    def load_monthly_sales(self, target_month: str) -> pd.DataFrame:
        """月別ISP別コンテンツ別売上.csvからデータを読み込み"""
        file_path = self.config.get_monthly_sales_file()
        
        if not Path(file_path).exists():
            raise FileNotFoundError(f"月別売上ファイルが見つかりません: {file_path}")
        
        try:
            df = pd.read_csv(file_path, encoding='utf-8-sig')
            
            # 対象年月でフィルタリング
            target_month_int = int(target_month)  # 文字列を数値に変換
            
            if '年月' in df.columns:
                filtered_df = df[df['年月'] == target_month_int].copy()
            elif 'YYYYMM' in df.columns:
                filtered_df = df[df['YYYYMM'] == target_month_int].copy()
            else:
                self.logger.warning("年月列が見つかりません。全データを返します。")
                filtered_df = df.copy()
            
            self.logger.info(f"月別売上データを読み込みました: {len(filtered_df)}件")
            return filtered_df
            
        except Exception as e:
            self.logger.error(f"月別売上データ読み込みエラー: {e}")
            raise
    
    def load_line_contents(self, year: str, month: str) -> pd.DataFrame:
        """LINE用line-contents-yyyy-mm.csvからデータを読み込み"""
        file_path = self.config.get_line_contents_file(year, month)
        
        if not Path(file_path).exists():
            self.logger.warning(f"LINEコンテンツファイルが見つかりません: {file_path}")
            return pd.DataFrame()
        
        try:
            df = pd.read_csv(file_path, encoding='utf-8-sig')
            self.logger.info(f"LINEコンテンツデータを読み込みました: {len(df)}件")
            return df
            
        except Exception as e:
            self.logger.error(f"LINEコンテンツデータ読み込みエラー: {e}")
            raise
    
    def load_content_mapping(self) -> pd.DataFrame:
        """contents_mapping.csvからデータを読み込み"""
        file_path = self.config.get_contents_mapping_file()
        
        if not Path(file_path).exists():
            raise FileNotFoundError(f"コンテンツマッピングファイルが見つかりません: {file_path}")
        
        try:
            df = pd.read_csv(file_path, encoding='utf-8-sig')
            self.logger.info(f"コンテンツマッピングデータを読み込みました: {len(df)}件")
            return df
            
        except Exception as e:
            self.logger.error(f"コンテンツマッピングデータ読み込みエラー: {e}")
            raise
    
    def load_rate_data(self) -> pd.DataFrame:
        """rate.csvからデータを読み込み"""
        file_path = self.config.get_rate_data_file()
        
        if not Path(file_path).exists():
            raise FileNotFoundError(f"レートデータファイルが見つかりません: {file_path}")
        
        try:
            df = pd.read_csv(file_path, encoding='utf-8-sig')
            self.logger.info(f"レートデータを読み込みました: {len(df)}件")
            return df
            
        except Exception as e:
            self.logger.error(f"レートデータ読み込みエラー: {e}")
            raise
    
    def load_target_month_data(self) -> pd.DataFrame:
        """target_month.csvからデータを読み込み"""
        file_path = Path(self.config.base_paths['current_dir']) / "target_month.csv"
        
        if not file_path.exists():
            raise FileNotFoundError(f"target_month.csvファイルが見つかりません: {file_path}")
        
        try:
            df = pd.read_csv(file_path, encoding='utf-8-sig')
            self.logger.info(f"target_monthデータを読み込みました: {len(df)}件")
            return df
            
        except Exception as e:
            self.logger.error(f"target_monthデータ読み込みエラー: {e}")
            raise
    
    def merge_sales_data(self, monthly_data: pd.DataFrame, line_data: pd.DataFrame) -> pd.DataFrame:
        """月別売上データとLINEデータを統合"""
        if line_data.empty:
            self.logger.info("LINEデータが空のため、月別売上データのみを使用します")
            return monthly_data
        
        try:
            # データ統合ロジックを実装
            # 共通のキー列でマージ（コンテンツ名など）
            merged_df = pd.concat([monthly_data, line_data], ignore_index=True)
            
            # 重複データの処理
            if 'コンテンツ名' in merged_df.columns:
                merged_df = merged_df.drop_duplicates(subset=['コンテンツ名'], keep='last')
            
            self.logger.info(f"データ統合完了: {len(merged_df)}件")
            return merged_df
            
        except Exception as e:
            self.logger.error(f"データ統合エラー: {e}")
            raise
    
    def create_sales_records(self, year: str, month: str) -> List[SalesRecord]:
        """統合されたデータからSalesRecordリストを作成"""
        target_month = f"{year}{month.zfill(2)}"
        
        # 各データソースを読み込み
        content_mapping = self.load_content_mapping()
        rate_data = self.load_rate_data()
        target_month_data = self.load_target_month_data()
        
        sales_records = []
        
        # target_month.csvの各行について処理
        for _, row in target_month_data.iterrows():
            try:
                content_name = str(row.get('コンテンツ', ''))
                platform = str(row.get('プラットフォーム', ''))
                offset_months = row.get('支払年月', '')
                
                # 支払年月が空白の場合はスキップ
                if pd.isna(offset_months) or offset_months == '':
                    continue
                
                offset_months = int(offset_months)
                
                # 対象年月からoffset_months分マイナスした年月を計算
                actual_target_month = self._calculate_offset_month(target_month, offset_months)
                
                # 計算した年月のデータを読み込み
                actual_year = actual_target_month[:4]
                actual_month = actual_target_month[4:]
                
                monthly_data = self.load_monthly_sales(actual_target_month)
                line_data = self.load_line_contents(actual_year, actual_month)
                merged_data = self.merge_sales_data(monthly_data, line_data)
                
                # 該当するコンテンツとプラットフォームのデータを検索
                matching_data = self._find_matching_sales_data(merged_data, content_name, platform)
                
                if matching_data is not None:
                    # マッピングデータからテンプレートファイル名を取得
                    template_file = self._get_template_file_from_mapping(
                        content_name, platform, content_mapping
                    )
                    
                    # レートデータから料率とメールアドレスを取得
                    rate_info = self._get_rate_info(template_file, rate_data)
                    
                    # SalesRecordを作成
                    record = SalesRecord(
                        platform=platform,
                        content_name=content_name,
                        performance=float(matching_data.get('実績', 0)),
                        information_fee=float(matching_data.get('情報提供料', matching_data.get('情報提供料合計', 0))),
                        target_month=actual_target_month,  # 計算された実際の対象年月
                        template_file=template_file,
                        rate=rate_info['rate'],
                        recipient_email=rate_info['email']
                    )
                    
                    sales_records.append(record)
                
            except Exception as e:
                self.logger.warning(f"レコード作成エラー (コンテンツ: {content_name}, プラットフォーム: {platform}): {e}")
                continue
        
        self.logger.info(f"SalesRecord作成完了: {len(sales_records)}件")
        return sales_records
    
    def _get_template_file_from_mapping(
        self, content_name: str, platform: str, mapping_df: pd.DataFrame
    ) -> str:
        """コンテンツマッピングからテンプレートファイル名を取得"""
        try:
            # プラットフォーム別の列名マッピング
            platform_column_map = {
                'au': 'au',
                'line': 'LINE', 
                'satori': 'ameba',  # satoriプラットフォームはameba列を参照
                'ameba': 'ameba',   # amebaプラットフォームはameba列を参照
                'rakuten': '楽天',  # rakutenは楽天列を使用
                '楽天': '楽天',
                'excite': 'excite'
            }
            
            # プラットフォームに対応する列名を取得
            target_column = platform_column_map.get(platform.lower(), platform)
            
            if target_column not in mapping_df.columns:
                self.logger.warning(f"プラットフォーム列が見つかりません: {target_column}")
                return "default_template.xlsx"
            
            # コンテンツ名で検索（部分一致も試行）
            for _, row in mapping_df.iterrows():
                platform_content = str(row.get(target_column, ''))
                
                # 完全一致
                if platform_content == content_name:
                    template_file = str(row.iloc[0])  # A列のテンプレートファイル名
                    return template_file + '.xlsx' if not template_file.endswith(('.xlsx', '.xls')) else template_file
                
                # 部分一致（コンテンツ名が含まれる場合）
                if content_name in platform_content or platform_content in content_name:
                    template_file = str(row.iloc[0])  # A列のテンプレートファイル名
                    return template_file + '.xlsx' if not template_file.endswith(('.xlsx', '.xls')) else template_file
            
            # マッチしない場合はデフォルト
            self.logger.warning(f"テンプレートファイルが見つかりません: {content_name} ({platform})")
            return "default_template.xlsx"
            
        except Exception as e:
            self.logger.error(f"テンプレートファイル取得エラー: {e}")
            return "default_template.xlsx"
    
    def _get_rate_info(self, template_file: str, rate_df: pd.DataFrame) -> Dict[str, any]:
        """レートデータから料率とメールアドレスを取得"""
        try:
            # テンプレートファイル名（拡張子なし）で検索
            template_name = Path(template_file).stem
            
            mask = rate_df['名称'] == template_name
            matched_rows = rate_df[mask]
            
            if not matched_rows.empty:
                row = matched_rows.iloc[0]
                rate_value = row.get('料率（％）', 0.0)
                
                # 料率を小数に変換（例：8% -> 0.08）
                rate = float(rate_value) / 100.0 if rate_value else 0.0
                
                # メールアドレスはデフォルト値（rateファイルにメール列がない）
                email = 'mizoguchi@outward.jp'  # デフォルトメールアドレス
                
                return {
                    'rate': rate,
                    'email': email
                }
            
            # マッチしない場合はデフォルト値
            self.logger.warning(f"レート情報が見つかりません: {template_file}")
            return {'rate': 0.0, 'email': 'mizoguchi@outward.jp'}
            
        except Exception as e:
            self.logger.error(f"レート情報取得エラー: {e}")
            return {'rate': 0.0, 'email': 'mizoguchi@outward.jp'}
    
    def _calculate_offset_month(self, target_month: str, offset_months: int) -> str:
        """対象年月からoffset_months分マイナスした年月を計算"""
        try:
            year = int(target_month[:4])
            month = int(target_month[4:])
            
            # マイナス月数を加算（負の値なのでマイナスになる）
            total_months = (year * 12 + month - 1) + offset_months
            
            if total_months < 0:
                self.logger.warning(f"計算結果が負の値になりました: {target_month} + {offset_months}")
                return target_month  # デフォルトで元の年月を返す
            
            result_year = total_months // 12
            result_month = (total_months % 12) + 1
            
            result = f"{result_year}{result_month:02d}"
            self.logger.debug(f"年月計算: {target_month} + {offset_months}ヶ月 = {result}")
            
            return result
            
        except Exception as e:
            self.logger.error(f"年月計算エラー: {e}")
            return target_month
    
    def _find_matching_sales_data(self, merged_data: pd.DataFrame, content_name: str, platform: str) -> Optional[pd.Series]:
        """統合データから該当するコンテンツとプラットフォームのデータを検索"""
        try:
            if merged_data.empty:
                return None
            
            # インデックスをリセットしてDataFrameを正規化
            merged_data = merged_data.reset_index(drop=True)
            
            # コンテンツ名での検索（完全一致）
            content_mask = pd.Series([False] * len(merged_data), index=merged_data.index)
            
            if 'コンテンツ' in merged_data.columns:
                content_mask |= (merged_data['コンテンツ'].astype(str) == content_name)
            if 'コンテンツ名' in merged_data.columns:
                content_mask |= (merged_data['コンテンツ名'].astype(str) == content_name)
            
            # プラットフォーム名での検索（完全一致）
            platform_mask = pd.Series([False] * len(merged_data), index=merged_data.index)
            
            if 'プラットフォーム' in merged_data.columns:
                platform_mask |= (merged_data['プラットフォーム'].astype(str) == platform)
            if 'ISP' in merged_data.columns:
                platform_mask |= (merged_data['ISP'].astype(str) == platform)
            
            # 両方の条件を満たすデータを検索
            matching_rows = merged_data[content_mask & platform_mask]
            
            if not matching_rows.empty:
                return matching_rows.iloc[0]  # 最初のマッチするデータを返す
            
            # 完全一致しない場合は部分一致を試行
            content_partial_mask = pd.Series([False] * len(merged_data), index=merged_data.index)
            
            if 'コンテンツ' in merged_data.columns:
                content_partial_mask |= merged_data['コンテンツ'].astype(str).str.contains(content_name, na=False)
            if 'コンテンツ名' in merged_data.columns:
                content_partial_mask |= merged_data['コンテンツ名'].astype(str).str.contains(content_name, na=False)
            
            platform_partial_mask = pd.Series([False] * len(merged_data), index=merged_data.index)
            
            if 'プラットフォーム' in merged_data.columns:
                platform_partial_mask |= merged_data['プラットフォーム'].astype(str).str.contains(platform, na=False)
            if 'ISP' in merged_data.columns:
                platform_partial_mask |= merged_data['ISP'].astype(str).str.contains(platform, na=False)
            
            partial_matching = merged_data[content_partial_mask & platform_partial_mask]
            
            if not partial_matching.empty:
                self.logger.debug(f"部分一致でデータが見つかりました: {content_name}, {platform}")
                return partial_matching.iloc[0]
            
            self.logger.debug(f"該当データが見つかりませんでした: {content_name}, {platform}")
            return None
            
        except Exception as e:
            self.logger.error(f"売上データ検索エラー: {e}")
            return None