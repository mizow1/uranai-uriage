import os
import pandas as pd
import openpyxl
from pathlib import Path
import re
import csv
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# 共通コンポーネントをインポート
from common import (
    CSVHandler,
    ExcelHandler, 
    UnifiedLogger,
    ErrorHandler,
    ConfigManager,
    ProcessingResult,
    ContentDetail,
    ProcessingSummary
)

class SalesAggregator:
    def __init__(self, base_path):
        self.base_path = Path(base_path)
        self.results = []
        # 共通コンポーネントを初期化
        self.logger = UnifiedLogger(__name__, log_file=Path("logs/sales_aggregator.log"))
        self.error_handler = ErrorHandler(self.logger.logger)
        self.csv_handler = CSVHandler(self.logger.logger, self.error_handler)
        self.excel_handler = ExcelHandler(self.logger.logger, self.error_handler)
        self.config = ConfigManager(logger=self.logger.logger)
    
    def _load_encrypted_workbook(self, file_path: Path, passwords: list):
        """暗号化されたワークブックを読み込み"""
        try:
            import msoffcrypto
            import io
            
            for password in passwords:
                try:
                    with open(file_path, 'rb') as file:
                        office_file = msoffcrypto.OfficeFile(file)
                        office_file.load_key(password=password if password else None)
                        
                        decrypted = io.BytesIO()
                        office_file.save(decrypted)
                        decrypted.seek(0)
                        
                        wb = openpyxl.load_workbook(decrypted, data_only=True)
                        self.logger.info(f"パスワード解除成功: {file_path.name} ('{password}')")
                        return wb
                        
                except Exception:
                    continue
            
            return None
            
        except ImportError:
            self.logger.error("msoffcrypto-toolが必要です")
            return None
        
    def find_files_in_yearmonth_folders(self):
        """年月フォルダ内のファイルを検索"""
        files_by_platform = {
            'ameba': [],
            'rakuten': [],
            'au': [],
            'excite': [],
            'line': []
        }
        
        # 年フォルダ（4桁）内の月フォルダ（6桁）を検索
        for year_folder in self.base_path.iterdir():
            if year_folder.is_dir() and re.match(r'\d{4}', year_folder.name):
                for month_folder in year_folder.iterdir():
                    if month_folder.is_dir() and re.match(r'\d{6}', month_folder.name):
                        # 月フォルダ直下のファイルを検索
                        for file in month_folder.iterdir():
                            if file.is_file():
                                filename = file.name.lower()
                                if '【株式会社アウトワード御中】satori実績_' in file.name or 'satori実績_' in filename:
                                    files_by_platform['ameba'].append(file)
                                elif 'rcms' in filename:
                                    files_by_platform['rakuten'].append(file)
                                elif 'salessummary' in filename:
                                    files_by_platform['au'].append(file)
                                elif 'excite' in filename:
                                    files_by_platform['excite'].append(file)
                            # サブフォルダも検索（LINEファイル用）
                            elif file.is_dir():
                                for subfile in file.iterdir():
                                    if subfile.is_file():
                                        subfilename = subfile.name.lower()
                                        if subfilename.startswith('line-contents-') and (subfile.suffix.lower() in ['.xls', '.xlsx', '.csv']):
                                            files_by_platform['line'].append(subfile)
        
        return files_by_platform
    
    def process_ameba_file(self, file_path: Path) -> ProcessingResult:
        """ameba占い（SATORI実績）ファイルを処理"""
        result = ProcessingResult(
            platform="ameba",
            file_name=file_path.name,
            success=False
        )
        
        start_time = datetime.now()
        
        try:
            # 統一Excelハンドラーを使用してファイルを読み込み
            passwords = self.config.get_processing_settings().get('excel_passwords', ['', 'password', '123456', '000000', 'admin', 'user'])
            
            # ExcelHandlerでパスワード保護を処理
            df = self.excel_handler.read_excel_with_password_handling(file_path, passwords=passwords)
            
            if df is None:
                # 直接openpyxlで試行（シート単位での処理が必要なため）
                try:
                    wb = openpyxl.load_workbook(file_path, data_only=True)
                except Exception as e:
                    if "password" in str(e).lower() or "protected" in str(e).lower():
                        wb = self._load_encrypted_workbook(file_path, passwords)
                        if wb is None:
                            result.add_error("パスワード保護解除に失敗")
                            return result
                    else:
                        result.add_error(f"ファイル読み込みエラー: {str(e)}")
                        return result
            else:
                # DataFrameからワークブックを取得（複雑な処理のため直接openpyxl使用）
                wb = openpyxl.load_workbook(file_path, data_only=True)
            
            self.logger.log_file_operation("読み込み", file_path, True)
            
            # 新仕様：各シートから情報提供料を集計し、同一コンテンツは合算する
            content_groups = {}
            
            # 1. 「従量実績」シートでC列の値が一致するもののJ列の合計額を算出
            self._process_sheet_data(wb, '従量実績', content_groups, 2, 9)  # C列, J列
            
            # 2. 「docomo占い」シートでC列の値が一致するもののJ列の合計額を算出
            self._process_sheet_data(wb, 'docomo占い', content_groups, 2, 9)  # C列, J列
            
            # 3. 「月額実績」シートのB列の値が一致するもののE列の合計額を算出
            self._process_sheet_data(wb, '月額実績', content_groups, 1, 4)  # B列, E列
            
            # 合計値を計算
            情報提供料合計 = sum(content_groups.values())
            実績合計 = 情報提供料合計 / 0.3 if 情報提供料合計 > 0 else 0  # 30%を除算した値
            
            # ContentDetailリストを作成
            for content_name, 情報提供料 in content_groups.items():
                実績 = 情報提供料 / 0.3 if 情報提供料 > 0 else 0
                detail = ContentDetail(
                    content_group=str(content_name),
                    performance=round(実績),
                    information_fee=round(情報提供料)
                )
                result.add_detail(detail)
            
            # 合計を計算
            result.calculate_totals()
            result.success = True
            result.metadata = {
                'content_groups_count': len(content_groups),
                '情報提供料合計': round(情報提供料合計),
                '実績合計': round(実績合計)
            }
            
            self.logger.info(f"ameba処理完了: {len(content_groups)}コンテンツグループ")
            
        except Exception as e:
            result.add_error(str(e))
            self.error_handler.handle_file_processing_error(e, file_path)
        
        finally:
            end_time = datetime.now()
            result.processing_time = (end_time - start_time).total_seconds()
        
        return result
    
    def _process_sheet_data(self, wb, sheet_name: str, content_groups: dict, key_col: int, value_col: int):
        """シートデータを処理してコンテンツグループに集計"""
        try:
            if sheet_name not in wb.sheetnames:
                self.logger.warning(f"シート '{sheet_name}' が存在しません")
                return
                
            sheet = wb[sheet_name]
            df = pd.DataFrame(sheet.values)
            df.columns = df.iloc[0]
            df = df.drop(0).reset_index(drop=True)
            
            for _, row in df.iterrows():
                key_value = row.iloc[key_col] 
                amount_value = row.iloc[value_col]
                
                if pd.notna(key_value) and pd.notna(amount_value):
                    amount_numeric = pd.to_numeric(amount_value, errors='coerce')
                    if pd.notna(amount_numeric):
                        if key_value not in content_groups:
                            content_groups[key_value] = 0
                        content_groups[key_value] += amount_numeric
                        
        except Exception as e:
            self.logger.warning(f"{sheet_name}シート処理エラー: {e}")
    
    def process_rakuten_file(self, file_path: Path) -> ProcessingResult:
        """楽天占い（rcms・楽天明細）ファイルを処理"""
        result = ProcessingResult(
            platform="rakuten",
            file_name=file_path.name,
            success=False
        )
        
        start_time = datetime.now()
        
        try:
            # ファイル名にrcmsを含むファイルのみ処理
            if 'rcms' not in file_path.name.lower():
                result.add_error("rcmsファイルではありません")
                self.logger.warning(f"楽天占いファイル処理スキップ: {file_path.name} - rcmsファイルではありません")
                return result
            
            # ファイル拡張子に応じてデータを読み込み
            if file_path.suffix.lower() == '.csv':
                df = self.csv_handler.read_csv_with_encoding_detection(file_path)
            else:
                df = self.excel_handler.read_excel_with_password_handling(file_path)
            
            self.logger.log_file_operation("読み込み", file_path, True)
            
            # RCMSファイルの処理
            # L列の値「hoge_xxx」のhoge部分が一致するもののN列の値で計算
            if len(df.columns) < 14:
                result.add_error(f"列数が不足: 必要14列以上、実際{len(df.columns)}列")
                return result
                
            l_column = df.iloc[:, 11]  # L列
            n_column = df.iloc[:, 13]  # N列
            
            hoge_groups = {}
            for i, value in enumerate(l_column):
                if pd.notna(value) and '_' in str(value):
                    hoge_part = str(value).split('_')[0]
                    if hoge_part not in hoge_groups:
                        hoge_groups[hoge_part] = []
                    hoge_groups[hoge_part].append(n_column.iloc[i])
            
            # 各グループの計算
            for hoge, values in hoge_groups.items():
                group_sum = sum(pd.to_numeric(v, errors='coerce') for v in values if pd.notna(v))
                実績_sum = group_sum / 1.1  # N列の値の合計額を1.1で除算
                情報提供料_sum = 実績_sum * 0.725  # 実績に0.725を乗算
                
                detail = ContentDetail(
                    content_group=hoge,
                    performance=round(実績_sum),
                    information_fee=round(情報提供料_sum)
                )
                result.add_detail(detail)
            
            # 「月額実績」シートの処理も追加（もしあれば）
            if file_path.suffix.lower() in ['.xlsx', '.xls']:
                try:
                    monthly_df = self.excel_handler.read_excel_with_password_handling(
                        file_path, sheet_name='月額実績'
                    )
                    if monthly_df is not None and len(monthly_df.columns) >= 5:
                        monthly_amount = monthly_df.iloc[:, 4].sum()  # E列
                        monthly_実績 = monthly_amount / 1.1
                        monthly_情報提供料 = monthly_実績 * 0.725
                        
                        detail = ContentDetail(
                            content_group='月額実績',
                            performance=round(monthly_実績),
                            information_fee=round(monthly_情報提供料)
                        )
                        result.add_detail(detail)
                except Exception as e:
                    self.logger.warning(f"月額実績シート処理エラー: {str(e)}")
            
            # 合計を計算
            result.calculate_totals()
            result.success = True
            result.metadata = {
                'hoge_groups_count': len(hoge_groups),
                'rcms_processing': True
            }
            
            self.logger.info(f"楽天処理完了: {len(hoge_groups)}グループ")
            
        except Exception as e:
            result.add_error(str(e))
            self.error_handler.handle_file_processing_error(e, file_path)
        
        finally:
            end_time = datetime.now()
            result.processing_time = (end_time - start_time).total_seconds()
        
        return result
    
    def process_au_file(self, file_path: Path) -> ProcessingResult:
        """au占い（SalesSummary）ファイルを処理"""
        result = ProcessingResult(
            platform="au",
            file_name=file_path.name,
            success=False
        )
        
        start_time = datetime.now()
        
        try:
            # ファイル拡張子に応じてデータを読み込み
            if file_path.suffix.lower() == '.csv':
                df = self.csv_handler.read_csv_with_encoding_detection(file_path)
            elif file_path.suffix.lower() in ['.xlsx', '.xls']:
                df = self.excel_handler.read_excel_with_password_handling(file_path)
                if df is None:
                    # 複数エンジンで試行
                    df = self.excel_handler.try_multiple_engines(file_path)
            else:
                result.add_error(f"サポートされていないファイル形式: {file_path.suffix}")
                return result
            
            self.logger.log_file_operation("読み込み", file_path, True)
            
            # プラットフォーム名決定（auファイルは全てmedibaとして扱う）
            platform_name = "mediba"
            
            # 列数チェック
            if len(df.columns) < 11:
                result.add_error(f"列数が不足: 必要11列以上、実際{len(df.columns)}列")
                return result
            
            # B列（インデックス1）でグループ化してG列の合計を計算
            b_column = df.iloc[:, 1]  # B列
            g_column = df.iloc[:, 6]  # G列
            k_column = df.iloc[:, 10]  # K列
            
            # 数値に変換
            g_column = pd.to_numeric(g_column, errors='coerce').fillna(0)
            k_column = pd.to_numeric(k_column, errors='coerce').fillna(0)
            
            b_groups = {}
            for i, b_value in enumerate(b_column):
                if pd.notna(b_value):
                    if b_value not in b_groups:
                        b_groups[b_value] = {'g_values': [], 'k_values': []}
                    b_groups[b_value]['g_values'].append(g_column.iloc[i])
                    b_groups[b_value]['k_values'].append(k_column.iloc[i])
            
            # 各グループの計算
            for b_value, values in b_groups.items():
                g_sum = sum(values['g_values'])
                k_sum = sum(values['k_values'])
                実績_sum = g_sum  # G列の値
                情報提供料_sum = (g_sum * 0.4) - k_sum  # G列の40%からK列を引いた値
                
                detail = ContentDetail(
                    content_group=str(b_value),
                    performance=round(実績_sum),
                    information_fee=round(情報提供料_sum)
                )
                result.add_detail(detail)
            
            # 合計を計算  
            result.calculate_totals()
            result.success = True
            result.platform = platform_name  # プラットフォーム名を上書き
            result.metadata = {
                'b_groups_count': len(b_groups),
                'platform_name': platform_name
            }
            
            self.logger.info(f"au処理完了: {len(b_groups)}グループ")
            
        except Exception as e:
            result.add_error(str(e))
            self.error_handler.handle_file_processing_error(e, file_path)
        
        finally:
            end_time = datetime.now()
            result.processing_time = (end_time - start_time).total_seconds()
        
        return result
    
    def process_excite_file(self, file_path: Path) -> ProcessingResult:
        """excite占いファイルを処理"""
        result = ProcessingResult(
            platform="excite",
            file_name=file_path.name,
            success=False
        )
        
        start_time = datetime.now()
        
        try:
            # ファイル拡張子に応じてデータを読み込み
            if file_path.suffix.lower() == '.csv':
                # CSVファイルの場合、特殊な読み込み処理（説明文をスキップ）
                df = self.csv_handler.read_csv_with_encoding_detection(
                    file_path, skiprows=3, on_bad_lines='skip', engine='python'
                )
            else:
                df = self.excel_handler.read_excel_with_password_handling(file_path)
            
            self.logger.log_file_operation("読み込み", file_path, True)
            
            # データの妥当性チェック
            if df.shape[1] < 5 or df.shape[0] == 0:
                result.add_error("データが不正または空です")
                return result
            
            # 数値列の存在チェック
            numeric_cols = df.select_dtypes(include=['number']).columns
            if len(numeric_cols) == 0:
                result.add_error("数値列が見つかりません")
                return result
            
            # exciteファイル固有の処理ロジック
            # 基本的な合計計算（具体的な仕様が不明なため簡単な処理）
            total_amount = 0
            for col in numeric_cols:
                column_sum = df[col].sum()
                if column_sum > 0:
                    total_amount += column_sum
            
            # 基本的な情報提供料計算（30%とする）
            information_fee = total_amount * 0.3
            
            detail = ContentDetail(
                content_group="excite_total",
                performance=round(total_amount),
                information_fee=round(information_fee)
            )
            result.add_detail(detail)
            
            # 合計を計算
            result.calculate_totals()
            result.success = True
            result.metadata = {
                'numeric_columns': len(numeric_cols),
                'data_rows': len(df)
            }
            
            self.logger.info(f"excite処理完了: {len(df)}行処理")
            
        except Exception as e:
            result.add_error(str(e))
            self.error_handler.handle_file_processing_error(e, file_path)
        
        finally:
            end_time = datetime.now()
            result.processing_time = (end_time - start_time).total_seconds()
        
        return result
    
    def process_line_file(self, file_path: Path) -> ProcessingResult:
        """LINE占いファイルを処理"""
        result = ProcessingResult(
            platform="line",
            file_name=file_path.name,
            success=False
        )
        
        start_time = datetime.now()
        
        try:
            # ファイル拡張子に応じてデータを読み込み
            if file_path.suffix.lower() == '.csv':
                df = self.csv_handler.read_csv_with_encoding_detection(file_path)
            elif file_path.suffix.lower() in ['.xlsx', '.xls']:
                df = self.excel_handler.read_excel_with_password_handling(file_path)
            else:
                result.add_error(f"サポートされていないファイル形式: {file_path.suffix}")
                return result
            
            self.logger.log_file_operation("読み込み", file_path, True)
            
            # 列数チェック（最低3列必要: コンテンツ名, 実績, 情報提供料）
            if len(df.columns) < 3:
                result.add_error(f"列数が不足: 必要3列以上、実際{len(df.columns)}列")
                return result
            
            # コンテンツ名（A列）でグループ化してrevenue（C列）の合計を計算
            service_name_column = df.iloc[:, 0]  # A列（コンテンツ名）
            revenue_column = df.iloc[:, 2]  # C列（情報提供料）
            
            # 数値に変換
            revenue_column = pd.to_numeric(revenue_column, errors='coerce').fillna(0)
            
            service_groups = {}
            for i, service_name in enumerate(service_name_column):
                if pd.notna(service_name):
                    if service_name not in service_groups:
                        service_groups[service_name] = 0
                    service_groups[service_name] += revenue_column.iloc[i]
            
            # 各サービスの計算
            for service_name, revenue_sum in service_groups.items():
                # B列（実績）も取得
                performance_sum = 0
                for i, name in enumerate(service_name_column):
                    if pd.notna(name) and name == service_name:
                        performance_value = df.iloc[i, 1]  # B列（実績）
                        if pd.notna(performance_value):
                            performance_sum += pd.to_numeric(performance_value, errors='coerce')
                
                実績_sum = performance_sum  # B列の値を実績とする
                情報提供料_sum = revenue_sum  # C列の値をそのまま情報提供料とする
                
                detail = ContentDetail(
                    content_group=str(service_name),
                    performance=round(実績_sum),
                    information_fee=round(情報提供料_sum)
                )
                result.add_detail(detail)
            
            # 合計を計算
            result.calculate_totals()
            result.success = True
            result.metadata = {
                'service_groups_count': len(service_groups),
                'total_revenue': sum(service_groups.values())
            }
            
            self.logger.info(f"LINE処理完了: {len(service_groups)}サービスグループ")
            
        except Exception as e:
            result.add_error(str(e))
            self.error_handler.handle_file_processing_error(e, file_path)
        
        finally:
            end_time = datetime.now()
            result.processing_time = (end_time - start_time).total_seconds()
        
        return result
    
    def _extract_year_month_from_path(self, file_path: Path) -> str:
        """ファイルのパスから年月を抽出（YYYYMM形式）"""
        try:
            # パスの親フォルダ（年月フォルダ）名から6桁の年月を取得
            for parent in file_path.parents:
                folder_name = parent.name
                if re.match(r'\d{6}', folder_name):
                    return folder_name
            
            # フォルダ名から取得できない場合、ファイル名から推測
            filename = file_path.name
            # ファイル名に含まれる年月パターンを検索（例：202302）
            match = re.search(r'(\d{4})(\d{2})', filename)
            if match:
                year = match.group(1)
                month = match.group(2)
                return f"{year}{month}"
            
            # 見つからない場合は空文字を返す
            return ""
            
        except Exception as e:
            self.logger.warning(f"年月抽出エラー: {file_path} - {str(e)}")
            return ""
    
    def process_all_files(self):
        """すべてのファイルを一括処理"""
        self.logger.info("全ファイル処理を開始")
        
        # 年月フォルダ内のファイルを検索
        files_by_platform = self.find_files_in_yearmonth_folders()
        
        total_files = sum(len(files) for files in files_by_platform.values())
        self.logger.info(f"処理対象ファイル数: {total_files}")
        
        # プラットフォーム別にファイルを処理
        for platform, files in files_by_platform.items():
            for file_path in files:
                self.logger.info(f"処理中: {platform} - {file_path.name}")
                
                # ファイルパスから年月を取得
                year_month = self._extract_year_month_from_path(file_path)
                
                if platform == 'ameba':
                    result = self.process_ameba_file(file_path)
                elif platform == 'rakuten':
                    result = self.process_rakuten_file(file_path)
                elif platform == 'au':
                    result = self.process_au_file(file_path)
                elif platform == 'excite':
                    result = self.process_excite_file(file_path)
                elif platform == 'line':
                    result = self.process_line_file(file_path)
                else:
                    self.logger.warning(f"未対応プラットフォーム: {platform}")
                    continue
                
                if result.success:
                    self.results.append({
                        'platform': result.platform,
                        'file_name': result.file_name,
                        'content_details': result.details,
                        '情報提供料合計': result.total_information_fee,
                        '実績合計': result.total_performance,
                        '年月': year_month,
                        '処理日時': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    })
                    self.logger.info(f"処理成功: {file_path.name}")
                else:
                    self.logger.error(f"処理失敗: {file_path.name} - {', '.join(result.errors)}")
        
        self.logger.info(f"全ファイル処理完了: {len(self.results)}件成功")
    
    def export_to_csv(self, output_path: str):
        """結果をCSVファイルに出力"""
        if not self.results:
            self.logger.warning("出力するデータがありません")
            return
        
        try:
            with open(output_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.writer(csvfile)
                
                # ヘッダー
                writer.writerow(['プラットフォーム', 'ファイル名', 'コンテンツ', '実績', '情報提供料合計', '年月', '処理日時'])
                
                # データ
                for result in self.results:
                    platform = result['platform']
                    file_name = result['file_name']
                    year_month = result.get('年月', '')
                    processing_time = result.get('処理日時', '')
                    
                    if result['content_details']:
                        for detail in result['content_details']:
                            writer.writerow([
                                platform,
                                file_name,
                                detail.content_group,
                                detail.performance,
                                detail.information_fee,
                                year_month,
                                processing_time
                            ])
                    else:
                        # 詳細がない場合は合計値を出力
                        writer.writerow([
                            platform,
                            file_name,
                            '合計',
                            result['実績合計'],
                            result['情報提供料合計'],
                            year_month,
                            processing_time
                        ])
            
            self.logger.info(f"CSV出力完了: {output_path}")
            
        except Exception as e:
            self.logger.error(f"CSV出力エラー: {str(e)}")
            raise