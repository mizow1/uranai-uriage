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
            
            # 対象年月でフィルタリング（文字列として比較）
            if '年月' in df.columns:
                # 年月列を文字列として扱い、文字列として比較
                df['年月'] = df['年月'].astype(str)
                filtered_df = df[df['年月'] == target_month].copy()
            elif 'YYYYMM' in df.columns:
                # YYYYMM列を文字列として扱い、文字列として比較
                df['YYYYMM'] = df['YYYYMM'].astype(str)
                filtered_df = df[df['YYYYMM'] == target_month].copy()
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
                
                # 支払年月が空白の場合は対象外（コンテンツ自体が存在しない）として処理をスキップ
                if pd.isna(offset_months) or offset_months == '' or str(offset_months).strip() == '':
                    self.logger.debug(f"支払年月が空白のため対象外（コンテンツ自体が存在しない）: {content_name} ({platform})")
                    continue  # この行をスキップして次の行へ
                else:
                    offset_months = int(offset_months)
                    # 対象年月からoffset_months分マイナスした年月を計算
                    actual_target_month = self._calculate_offset_month(target_month, offset_months)
                
                self.logger.debug(f"処理中: {content_name} ({platform}) - 対象年月: {actual_target_month} (オフセット: {offset_months})")
                
                # 計算した年月のデータを読み込み
                actual_year = actual_target_month[:4]
                actual_month = actual_target_month[4:]
                
                monthly_data = self.load_monthly_sales(actual_target_month)
                line_data = self.load_line_contents(actual_year, actual_month)
                merged_data = self.merge_sales_data(monthly_data, line_data)
                
                self.logger.debug(f"データ統合結果: {len(merged_data)}件のデータ")
                if not merged_data.empty:
                    self.logger.debug(f"統合データの列: {list(merged_data.columns)}")
                    # プラットフォーム列の値を確認
                    if 'プラットフォーム' in merged_data.columns:
                        platforms = merged_data['プラットフォーム'].unique()
                        self.logger.debug(f"プラットフォーム一覧: {platforms}")
                    if 'コンテンツ' in merged_data.columns:
                        contents = merged_data['コンテンツ'].unique()
                        self.logger.debug(f"コンテンツ一覧: {contents[:10]}")  # 最初の10件のみ表示
                
                # contents_mapping.csvから実際のコンテンツ名を取得してから検索
                template_files = self._get_template_files_from_mapping(
                    content_name, platform, content_mapping
                )
                
                matching_data = None
                # 各テンプレートファイルに対応するコンテンツ名で検索
                # C列に値がある場合は0円データも許可
                allow_zero = not pd.isna(offset_months) and str(offset_months).strip() != ''
                
                for template_file, output_content_name, search_candidates in template_files:
                    # 複数の検索候補で売上データを検索
                    for search_content in search_candidates:
                        matching_data = self._find_matching_sales_data(merged_data, search_content, platform, allow_zero_performance=allow_zero)
                        if matching_data is not None:
                            self.logger.debug(f"マッチしたコンテンツ名: {search_content}")
                            break
                    if matching_data is not None:
                        break
                
                if matching_data is not None:
                    self.logger.info(f"マッチしたデータが見つかりました: {content_name} ({platform}) - 実績: {matching_data.get('実績', 0)}")
                    
                    # 各テンプレートファイルに対してSalesRecordを作成
                    for template_file, output_content_name, search_candidates in template_files:
                        # output_content_nameが空文字の場合はそのプラットフォームでコンテンツが存在しないためスキップ
                        if not output_content_name or output_content_name.strip() == '':
                            self.logger.info(f"コンテンツ名が空文字のためスキップ: {content_name} ({platform})")
                            continue
                            
                        # レートデータから料率とメールアドレスを取得
                        rate_info = self._get_rate_info(template_file, rate_data)
                        
                        # SalesRecordを作成（出力用コンテンツ名を使用）
                        record = SalesRecord(
                            platform=platform,
                            content_name=output_content_name,  # 出力用コンテンツ名を使用（D列優先）
                            performance=float(matching_data.get('実績', 0)),
                            information_fee=float(matching_data.get('情報提供料', 0)),
                            target_month=actual_target_month,  # 計算された実際の対象年月
                            template_file=template_file,
                            rate=rate_info['rate'],
                            recipient_email=rate_info['email']
                        )
                        
                        # デバッグ情報を追加
                        self.logger.debug(f"SalesRecord作成: {output_content_name} ({platform}) - 実績:{record.performance}, 情報提供料:{record.information_fee}, テンプレート:{template_file}")
                        
                        sales_records.append(record)
                else:
                    # データが見つからない場合、C列（支払年月）に値があるかチェック
                    # C列に値がある場合は売上0円でも記載する
                    self.logger.warning(f"該当データが見つかりませんでした: {content_name} ({platform}) - 対象年月: {actual_target_month}")
                    
                    # C列に値がある（支払年月が指定されている）場合は0円で記載
                    if not pd.isna(offset_months) and str(offset_months).strip() != '':
                        self.logger.info(f"支払年月が指定されているため0円で記載: {content_name} ({platform})")
                        
                        # 各テンプレートファイルに対してSalesRecordを作成（実績0円）
                        for template_file, output_content_name, search_candidates in template_files:
                            # output_content_nameが空文字の場合はそのプラットフォームでコンテンツが存在しないためスキップ
                            if not output_content_name or output_content_name.strip() == '':
                                self.logger.info(f"コンテンツ名が空文字のためスキップ: {content_name} ({platform})")
                                continue
                                
                            # レートデータから料率とメールアドレスを取得
                            rate_info = self._get_rate_info(template_file, rate_data)
                            
                            # SalesRecordを作成（実績・情報提供料ともに0、出力用コンテンツ名を使用）
                            record = SalesRecord(
                                platform=platform,
                                content_name=output_content_name,  # 出力用コンテンツ名を使用（D列優先）
                                performance=0.0,  # 実績0円
                                information_fee=0.0,  # 情報提供料0円
                                target_month=actual_target_month,  # 計算された実際の対象年月
                                template_file=template_file,
                                rate=rate_info['rate'],
                                recipient_email=rate_info['email']
                            )
                            
                            # デバッグ情報を追加
                            self.logger.debug(f"SalesRecord作成(0円): {output_content_name} ({platform}) - 実績:0, 情報提供料:0, テンプレート:{template_file}")
                            
                            sales_records.append(record)
                
            except Exception as e:
                self.logger.warning(f"レコード作成エラー (コンテンツ: {content_name}, プラットフォーム: {platform}): {e}")
                continue
        
        self.logger.info(f"SalesRecord作成完了: {len(sales_records)}件")
        return sales_records
    
    def _get_template_files_from_mapping(
        self, content_name: str, platform: str, mapping_df: pd.DataFrame
    ) -> List[Tuple[str, str]]:
        """コンテンツマッピングからテンプレートファイル名と実際のコンテンツ名のリストを取得"""
        try:
            # プラットフォーム別の列名マッピング
            platform_column_map = {
                'mediba': 'mediba',  # medibaは独立したプラットフォーム
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
                return [("default_template.xlsx", content_name)]
            
            # コンテンツ名で検索
            for _, row in mapping_df.iterrows():
                # A列（テンプレートファイル名）でまず検索
                row_content_name = str(row.iloc[0])  # A列のコンテンツ名
                platform_content = str(row.get(target_column, ''))  # プラットフォーム列の値
                
                self.logger.debug(f"マッピング検索: {content_name} ({platform}) vs A列='{row_content_name}', {target_column}列='{platform_content}'")
                
                # A列のコンテンツ名と一致するかチェック
                if row_content_name == content_name:
                    templates = []
                    
                    # medibaプラットフォームの場合、D列を優先して使用（出力用）、検索時はC列またはD列
                    if target_column == 'mediba':
                        # D列（4列目）のコンテンツ名を取得
                        mediba_d_content = ''
                        if len(row) > 3:
                            mediba_d_content = str(row.iloc[3])  # D列
                        
                        # C列（3列目）のコンテンツ名を取得
                        mediba_c_content = ''
                        if len(row) > 2:
                            mediba_c_content = str(row.iloc[2])  # C列
                        
                        # 出力用にはD列を優先、D列が空ならC列を使用
                        output_content = ''
                        if mediba_d_content and mediba_d_content != 'nan' and mediba_d_content.strip() != '':
                            output_content = mediba_d_content
                        elif mediba_c_content and mediba_c_content != 'nan' and mediba_c_content.strip() != '':
                            output_content = mediba_c_content
                        else:
                            output_content = ''  # medibaで両方の列が空の場合は空文字（コンテンツ存在しない）
                            
                        # 検索用候補リスト（C列またはD列）
                        search_candidates = []
                        if mediba_d_content and mediba_d_content != 'nan' and mediba_d_content.strip() != '':
                            search_candidates.append(mediba_d_content)
                        if mediba_c_content and mediba_c_content != 'nan' and mediba_c_content.strip() != '':
                            search_candidates.append(mediba_c_content)
                        if not search_candidates:
                            search_candidates.append('')  # medibaで両方の列が空の場合は空文字で検索（マッチしない）
                            
                        self.logger.debug(f"mediba マッピング一致: {content_name} -> output_content='{output_content}', search_candidates={search_candidates}")
                        
                        # 出力用コンテンツ名で1つだけテンプレートを作成
                        # A列のテンプレートファイル
                        template_a = str(row.iloc[0])
                        if template_a and template_a != '' and not pd.isna(template_a):
                            template_file_a = template_a + '.xlsx' if not template_a.endswith(('.xlsx', '.xls')) else template_a
                            templates.append((template_file_a, output_content, search_candidates))
                        
                        # B列のテンプレートファイル（存在する場合）
                        if len(row) > 1 and pd.notna(row.iloc[1]) and str(row.iloc[1]) != '':
                            template_b = str(row.iloc[1])
                            template_file_b = template_b + '.xlsx' if not template_b.endswith(('.xlsx', '.xls')) else template_b
                            templates.append((template_file_b, output_content, search_candidates))
                    else:
                        # mediba以外のプラットフォームの場合
                        if platform_content and platform_content != 'nan' and platform_content.strip() != '':
                            actual_content_name = str(platform_content)
                        else:
                            # プラットフォーム列が空欄の場合は、そのプラットフォームでのコンテンツは存在しないものとして空文字を設定
                            if target_column.lower() in ['line', '楽天', 'excite', 'ameba', 'mediba']:
                                actual_content_name = ''  # 空文字で明示的に存在しないことを示す
                            else:
                                actual_content_name = content_name  # その他のプラットフォームはデフォルト値を使用
                        
                        # 検索候補は1つのみ
                        search_candidates = [actual_content_name] if actual_content_name else [content_name]
                        
                        self.logger.debug(f"マッピング一致: {content_name} -> actual_content_name='{actual_content_name}', search_candidates={search_candidates}")
                        
                        # A列のテンプレートファイル
                        template_a = str(row.iloc[0])
                        if template_a and template_a != '' and not pd.isna(template_a):
                            template_file_a = template_a + '.xlsx' if not template_a.endswith(('.xlsx', '.xls')) else template_a
                            templates.append((template_file_a, actual_content_name, search_candidates))
                        
                        # B列のテンプレートファイル（存在する場合）
                        if len(row) > 1 and pd.notna(row.iloc[1]) and str(row.iloc[1]) != '':
                            template_b = str(row.iloc[1])
                            template_file_b = template_b + '.xlsx' if not template_b.endswith(('.xlsx', '.xls')) else template_b
                            templates.append((template_file_b, actual_content_name, search_candidates))
                    
                    if templates:
                        self.logger.info(f"テンプレートファイルが見つかりました: {content_name} ({platform}) -> {[t[0] for t in templates]}")
                        return templates
            
            # マッチしない場合はデフォルト
            self.logger.warning(f"テンプレートファイルが見つかりません: {content_name} ({platform})")
            return [("default_template.xlsx", content_name, [content_name])]
            
        except Exception as e:
            self.logger.error(f"テンプレートファイル取得エラー: {e}")
            return [("default_template.xlsx", content_name, [content_name])]
    
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
    
    def _find_matching_sales_data(self, merged_data: pd.DataFrame, content_name: str, platform: str, allow_zero_performance: bool = False) -> Optional[pd.Series]:
        """統合データから該当するコンテンツとプラットフォームのデータを検索"""
        try:
            if merged_data.empty:
                self.logger.debug(f"統合データが空です: {content_name}, {platform}")
                return None
            
            # インデックスをリセットしてDataFrameを正規化
            merged_data = merged_data.reset_index(drop=True)
            
            # デバッグ用：統合データの内容を表示
            self.logger.debug(f"統合データ検索: コンテンツ='{content_name}', プラットフォーム='{platform}', データ件数={len(merged_data)}")
            if not merged_data.empty:
                self.logger.debug(f"統合データの列: {list(merged_data.columns)}")
                # exciteデータのみを表示
                if 'プラットフォーム' in merged_data.columns:
                    platform_data = merged_data[merged_data['プラットフォーム'].astype(str).str.lower() == platform.lower()]
                    if not platform_data.empty:
                        self.logger.debug(f"{platform}データ {len(platform_data)}件:")
                        for idx, row in platform_data.head(5).iterrows():
                            content_col = row.get('コンテンツ', row.get('コンテンツ名', 'N/A'))
                            self.logger.debug(f"  {idx}: コンテンツ='{content_col}', プラットフォーム='{row.get('プラットフォーム', 'N/A')}', 実績={row.get('実績', 'N/A')}")
            
            # プラットフォーム名の正規化マッピング
            platform_mapping = {
                'ameba': ['ameba', 'satori'],  # amebaはsatoriデータも含む
                'satori': ['ameba', 'satori'], 
                'mediba': ['mediba'],  # medibaは独立したプラットフォーム
                'line': ['line'],
                'rakuten': ['rakuten', '楽天'],
                '楽天': ['rakuten', '楽天'],
                'excite': ['excite']
            }
            
            # 検索対象のプラットフォーム名リストを取得
            search_platforms = platform_mapping.get(platform.lower(), [platform])
            self.logger.debug(f"検索対象プラットフォーム: {search_platforms}")
            
            # コンテンツ名での検索（完全一致、スペース正規化対応）
            content_mask = pd.Series([False] * len(merged_data), index=merged_data.index)
            
            # 検索対象コンテンツ名を正規化（前後の空白文字、改行文字を除去）
            clean_content_name = content_name.strip()
            
            # スペース文字を正規化した検索名を作成
            normalized_content_name = clean_content_name.replace('　', ' ').replace(' ', ' ')
            
            if 'コンテンツ' in merged_data.columns:
                # データ側も正規化（前後の空白文字、改行文字を除去）
                clean_data = merged_data['コンテンツ'].astype(str).str.strip()
                
                # 完全一致
                exact_match = (clean_data == clean_content_name)
                content_mask |= exact_match
                self.logger.debug(f"コンテンツ列完全一致: '{clean_content_name}' -> {exact_match.sum()}件")
                
                # スペース正規化での一致
                normalized_data = clean_data.str.replace('　', ' ').str.replace(' ', ' ')
                normalized_match = (normalized_data == normalized_content_name)
                content_mask |= normalized_match
                self.logger.debug(f"コンテンツ列正規化一致: '{normalized_content_name}' -> {normalized_match.sum()}件")
                
            if 'コンテンツ名' in merged_data.columns:
                # データ側も正規化（前後の空白文字、改行文字を除去）
                clean_data = merged_data['コンテンツ名'].astype(str).str.strip()
                
                # 完全一致
                exact_match = (clean_data == clean_content_name)
                content_mask |= exact_match
                self.logger.debug(f"コンテンツ名列完全一致: '{clean_content_name}' -> {exact_match.sum()}件")
                
                # スペース正規化での一致
                normalized_data = clean_data.str.replace('　', ' ').str.replace(' ', ' ')
                normalized_match = (normalized_data == normalized_content_name)
                content_mask |= normalized_match
                self.logger.debug(f"コンテンツ名列正規化一致: '{normalized_content_name}' -> {normalized_match.sum()}件")
            
            self.logger.debug(f"コンテンツ検索結果: 合計{content_mask.sum()}件マッチ")
            
            # プラットフォーム名での検索（完全一致）
            platform_mask = pd.Series([False] * len(merged_data), index=merged_data.index)
            
            for search_platform in search_platforms:
                if 'プラットフォーム' in merged_data.columns:
                    platform_match = (merged_data['プラットフォーム'].astype(str).str.lower() == search_platform.lower())
                    platform_mask |= platform_match
                    self.logger.debug(f"プラットフォーム列一致 '{search_platform}': {platform_match.sum()}件")
                if 'ISP' in merged_data.columns:
                    isp_match = (merged_data['ISP'].astype(str).str.lower() == search_platform.lower())
                    platform_mask |= isp_match
                    self.logger.debug(f"ISP列一致 '{search_platform}': {isp_match.sum()}件")
            
            self.logger.debug(f"プラットフォーム検索結果: 合計{platform_mask.sum()}件マッチ")
            
            # 両方の条件を満たすデータを検索
            matching_rows = merged_data[content_mask & platform_mask]
            self.logger.debug(f"最終的な結合結果: {len(matching_rows)}件マッチ")
            
            if not matching_rows.empty:
                self.logger.debug(f"完全一致でデータが見つかりました: {content_name}, {platform}")
                # 実績が0より大きいレコードを優先的に返す
                valid_rows = matching_rows[matching_rows['実績'] > 0]
                if not valid_rows.empty:
                    return valid_rows.iloc[0]
                elif allow_zero_performance:
                    # 0円データも許可する場合は最初の行を返す
                    self.logger.debug(f"完全一致で実績0のデータを返します: {content_name}, {platform}")
                    return matching_rows.iloc[0]
                else:
                    self.logger.debug(f"完全一致したが実績が0のデータのみ: {content_name}, {platform}")
                    return None
            
            # 完全一致しない場合は部分一致を試行
            content_partial_mask = pd.Series([False] * len(merged_data), index=merged_data.index)
            
            # スペース除去版も作成
            content_name_no_space = clean_content_name.replace('　', '').replace(' ', '')
            
            if 'コンテンツ' in merged_data.columns:
                clean_data = merged_data['コンテンツ'].astype(str).str.strip()
                content_partial_mask |= clean_data.str.contains(clean_content_name, na=False, case=False)
                content_partial_mask |= clean_data.str.contains(content_name_no_space, na=False, case=False)
                # 全角・半角スペースを除去したデータでも検索
                no_space_data = clean_data.str.replace('　', '').str.replace(' ', '')
                content_partial_mask |= no_space_data.str.contains(content_name_no_space, na=False, case=False)
            if 'コンテンツ名' in merged_data.columns:
                clean_data = merged_data['コンテンツ名'].astype(str).str.strip()
                content_partial_mask |= clean_data.str.contains(clean_content_name, na=False, case=False)
                content_partial_mask |= clean_data.str.contains(content_name_no_space, na=False, case=False)
                # 全角・半角スペースを除去したデータでも検索
                no_space_data = clean_data.str.replace('　', '').str.replace(' ', '')
                content_partial_mask |= no_space_data.str.contains(content_name_no_space, na=False, case=False)
            
            platform_partial_mask = pd.Series([False] * len(merged_data), index=merged_data.index)
            
            for search_platform in search_platforms:
                if 'プラットフォーム' in merged_data.columns:
                    platform_partial_mask |= merged_data['プラットフォーム'].astype(str).str.contains(search_platform, na=False, case=False)
                if 'ISP' in merged_data.columns:
                    platform_partial_mask |= merged_data['ISP'].astype(str).str.contains(search_platform, na=False, case=False)
            
            partial_matching = merged_data[content_partial_mask & platform_partial_mask]
            
            if not partial_matching.empty:
                self.logger.debug(f"部分一致でデータが見つかりました: {content_name}, {platform}")
                # 実績が0より大きいレコードを優先的に返す
                valid_rows = partial_matching[partial_matching['実績'] > 0]
                if not valid_rows.empty:
                    return valid_rows.iloc[0]
                elif allow_zero_performance:
                    # 0円データも許可する場合は最初の行を返す
                    self.logger.debug(f"部分一致で実績0のデータを返します: {content_name}, {platform}")
                    return partial_matching.iloc[0]
                else:
                    self.logger.debug(f"部分一致したが実績が0のデータのみ: {content_name}, {platform}")
                    return None
            
            # プラットフォームのみの検索は行わない
            # コンテンツ名が一致しない場合は検索結果なしとする
            # これにより異なるプラットフォームのコンテンツ名が誤って表示されることを防ぐ
            
            self.logger.debug(f"該当データが見つかりませんでした: {content_name}, {platform}")
            self.logger.debug(f"利用可能な列: {list(merged_data.columns)}")
            
            # 詳細なデバッグ情報を追加
            if platform.lower() == 'excite' and 'シェイプシフター' in content_name:
                self.logger.debug("=== 詳細デバッグ：excite シェイプシフター検索 ===")
                if 'コンテンツ' in merged_data.columns:
                    unique_contents = merged_data['コンテンツ'].unique()
                    self.logger.debug(f"利用可能なコンテンツ名（最初の10件）: {unique_contents[:10]}")
                    # シェイプシフターが含まれるコンテンツを検索
                    shape_contents = [c for c in unique_contents if 'シェイプシフター' in str(c)]
                    self.logger.debug(f"シェイプシフターを含むコンテンツ: {shape_contents}")
                if 'プラットフォーム' in merged_data.columns:
                    unique_platforms = merged_data['プラットフォーム'].unique()
                    self.logger.debug(f"利用可能なプラットフォーム: {unique_platforms}")
                    excite_rows = merged_data[merged_data['プラットフォーム'].astype(str).str.lower() == 'excite']
                    self.logger.debug(f"exciteプラットフォームのデータ: {len(excite_rows)}件")
                    if not excite_rows.empty:
                        excite_contents = excite_rows['コンテンツ'].unique() if 'コンテンツ' in excite_rows.columns else []
                        self.logger.debug(f"exciteのコンテンツ: {excite_contents}")
                self.logger.debug("=== 詳細デバッグ終了 ===")
            
            return None
            
        except Exception as e:
            self.logger.error(f"売上データ検索エラー: {e}")
            return None