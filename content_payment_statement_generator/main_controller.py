"""
メインコントローラーモジュール

システム全体の処理フローを統合管理します。
"""

from pathlib import Path
from datetime import datetime
from typing import List, Dict, Optional, Tuple
from collections import defaultdict
import logging

from .config_manager import ConfigManager
from .sales_data_loader import SalesDataLoader
from .excel_processor import ExcelProcessor
from .pdf_converter import PDFConverter
from .email_processor import EmailProcessor
from .data_models import SalesRecord, PaymentStatement
from .logger import SystemLogger
from .exceptions import (
    ContentPaymentStatementError,
    DataValidationError,
    ExcelProcessingError,
    PDFConversionError,
    EmailSendError
)


class MainController:
    """メインコントローラークラス"""
    
    def __init__(self, log_level: str = "INFO"):
        """メインコントローラーを初期化"""
        # ログシステムを初期化
        self.system_logger = SystemLogger(log_level)
        self.logger = self.system_logger.get_logger(__name__)
        
        # 各コンポーネントを初期化
        self.config = ConfigManager()
        self.data_loader = SalesDataLoader(self.config)
        self.excel_processor = ExcelProcessor(self.config)
        self.pdf_converter = PDFConverter()
        self.email_processor = EmailProcessor()
        
        # 処理統計
        self.statistics = {
            'processed_contents': 0,
            'generated_pdfs': 0,
            'sent_emails': 0,
            'errors': 0
        }
    
    def process_payment_statements(self, year: str, month: str, send_emails: bool = True, template_filter: str = None) -> bool:
        """支払い明細書の完全処理を実行"""
        try:
            self.system_logger.log_system_info()
            self.logger.info(f"支払い明細書処理を開始: {year}年{month}月")
            
            # 必要なファイルの存在確認
            if not self._validate_required_files(year, month):
                return False
            
            # 売上データを読み込み
            sales_records = self._load_sales_data(year, month)
            if not sales_records:
                self.logger.error("有効な売上データが見つかりません")
                return False
            
            # コンテンツ別に支払い明細書を作成
            payment_statements = self._group_records_by_content(sales_records)
            
            # テンプレートフィルターが指定されている場合はフィルタリング
            if template_filter:
                filtered_statements = {}
                for key, statement in payment_statements.items():
                    if statement.template_file == template_filter:
                        filtered_statements[key] = statement
                payment_statements = filtered_statements
                self.logger.info(f"テンプレートフィルター適用: {template_filter} - {len(payment_statements)}件")
            
            # 各支払い明細書を処理
            success_count = 0
            total_count = len(payment_statements)
            
            for i, (content_key, statement) in enumerate(payment_statements.items(), 1):
                try:
                    self.system_logger.log_progress(i, total_count, f"処理中: {content_key}")
                    
                    if self._process_single_statement(statement, year, month, send_emails):
                        success_count += 1
                    else:
                        self.statistics['errors'] += 1
                        
                except Exception as e:
                    self.system_logger.log_error_details(e, f"支払い明細書処理: {content_key}")
                    self.statistics['errors'] += 1
            
            # 処理結果をログに記録
            self._log_processing_results(success_count, total_count)
            
            # 一時ファイルをクリーンアップ
            self._cleanup_temp_files(year, month)
            
            success = success_count > 0
            self.system_logger.log_system_end(success)
            
            return success
            
        except Exception as e:
            self.system_logger.log_error_details(e, "メイン処理")
            self.system_logger.log_system_end(False)
            return False
    
    def _validate_required_files(self, year: str, month: str) -> bool:
        """必要なファイルの存在確認"""
        try:
            validation_results = self.config.validate_required_files(year, month)
            
            # 必須ファイルのみチェック（line_contentsはオプション）
            required_files = ['monthly_sales', 'contents_mapping', 'rate_data', 'template_dir']
            missing_files = [
                file_type for file_type in required_files 
                if not validation_results.get(file_type, False)
            ]
            
            if missing_files:
                self.logger.error(f"必要なファイルが見つかりません: {', '.join(missing_files)}")
                return False
            
            # LINEコンテンツファイルは警告のみ
            if not validation_results.get('line_contents', False):
                self.logger.warning("LINEコンテンツファイルが見つかりません（処理は継続します）")
            
            self.logger.info("必要なファイルの存在確認が完了しました")
            return True
            
        except Exception as e:
            self.system_logger.log_error_details(e, "ファイル存在確認")
            return False
    
    def _load_sales_data(self, year: str, month: str) -> List[SalesRecord]:
        """売上データを読み込み"""
        try:
            self.logger.info("売上データの読み込みを開始")
            
            sales_records = self.data_loader.create_sales_records(year, month)
            
            self.system_logger.log_data_summary(
                "売上レコード", 
                len(sales_records),
                f"{year}年{month}月"
            )
            
            return sales_records
            
        except Exception as e:
            self.system_logger.log_error_details(e, "売上データ読み込み")
            raise DataValidationError(f"売上データ読み込みエラー: {e}")
    
    def _group_records_by_content(self, sales_records: List[SalesRecord]) -> Dict[str, PaymentStatement]:
        """売上レコードをコンテンツ別にグループ化し、同一プラットフォーム・同一コンテンツは合計"""
        try:
            # まず、プラットフォーム・コンテンツ・テンプレートで集計
            aggregated_records = defaultdict(lambda: {
                'platform': '',
                'content_name': '',
                'template_file': '',
                'target_month': '',
                'rate': 0.0,
                'recipient_email': '',
                'total_performance': 0.0,
                'total_information_fee': 0.0
            })
            
            # 同一プラットフォーム・同一コンテンツを集計
            for record in sales_records:
                key = f"{record.platform}_{record.content_name}_{record.template_file}"
                
                agg = aggregated_records[key]
                if not agg['platform']:  # 初回の場合は基本情報を設定
                    agg['platform'] = record.platform
                    agg['content_name'] = record.content_name
                    agg['template_file'] = record.template_file
                    agg['target_month'] = record.target_month
                    agg['rate'] = record.rate
                    agg['recipient_email'] = record.recipient_email
                
                # 実績と情報提供料を累積
                agg['total_performance'] += record.performance
                agg['total_information_fee'] += record.information_fee
            
            # 集計されたレコードをSalesRecordに変換してグループ化
            grouped_records = defaultdict(list)
            
            for key, agg in aggregated_records.items():
                # 集計されたデータからSalesRecordを作成
                aggregated_record = SalesRecord(
                    platform=agg['platform'],
                    content_name=agg['content_name'],
                    performance=agg['total_performance'],
                    information_fee=agg['total_information_fee'],
                    target_month=agg['target_month'],
                    template_file=agg['template_file'],
                    rate=agg['rate'],
                    recipient_email=agg['recipient_email']
                )
                
                # テンプレートファイル・メールアドレス別にグループ化
                group_key = f"{agg['template_file']}_{agg['recipient_email']}"
                grouped_records[group_key].append(aggregated_record)
            
            # PaymentStatementオブジェクトを作成
            payment_statements = {}
            
            for key, records in grouped_records.items():
                if not records:
                    continue
                
                # 代表レコードから情報を取得
                sample_record = records[0]
                
                # 合計値を計算
                total_performance = sum(r.performance for r in records)
                total_information_fee = sum(r.information_fee for r in records)
                
                # 翌月5日の支払日を計算
                target_month = sample_record.target_month
                year = int(target_month[:4])
                month = int(target_month[4:])
                
                if month == 12:
                    payment_date = datetime(year + 1, 1, 5)
                else:
                    payment_date = datetime(year, month + 1, 5)
                
                statement = PaymentStatement(
                    content_name=sample_record.content_name,
                    template_file=sample_record.template_file,
                    sales_records=records,
                    total_performance=total_performance,
                    total_information_fee=total_information_fee,
                    payment_date=payment_date,
                    recipient_email=sample_record.recipient_email
                )
                
                payment_statements[key] = statement
            
            self.logger.info(f"支払い明細書をグループ化しました: {len(payment_statements)}件")
            return payment_statements
            
        except Exception as e:
            self.system_logger.log_error_details(e, "レコードグループ化")
            raise
    
    def _process_single_statement(
        self, 
        statement: PaymentStatement, 
        year: str, 
        month: str, 
        send_email: bool
    ) -> bool:
        """単一の支払い明細書を処理"""
        try:
            target_month = f"{year}{month.zfill(2)}"
            
            # Excelファイルを処理
            excel_path = self.excel_processor.process_excel_file(
                statement.template_file,
                statement.sales_records,
                target_month
            )
            
            self.system_logger.log_file_operation("Excel処理", excel_path, True)
            self.statistics['processed_contents'] += 1
            
            # PDFに変換
            pdf_path = self.pdf_converter.convert_and_validate(excel_path)
            
            if not pdf_path:
                raise PDFConversionError(f"PDF変換に失敗しました: {excel_path}")
            
            self.system_logger.log_file_operation("PDF変換", pdf_path, True)
            self.statistics['generated_pdfs'] += 1
            
            # メール送信
            if send_email and self._validate_email_address(statement.recipient_email):
                if self.email_processor.send_payment_notification(
                    statement.recipient_email,
                    pdf_path,
                    target_month,
                    statement.content_name
                ):
                    self.statistics['sent_emails'] += 1
                    self.logger.info(f"メール送信完了: {statement.recipient_email}")
                else:
                    raise EmailSendError(f"メール送信に失敗しました: {statement.recipient_email}")
            
            return True
            
        except Exception as e:
            self.system_logger.log_error_details(e, f"支払い明細書処理: {statement.content_name}")
            return False
    
    def _validate_email_address(self, email: str) -> bool:
        """メールアドレスの検証"""
        if not email or email == 'default@example.com':
            self.logger.warning(f"無効なメールアドレス: {email}")
            return False
        
        return self.email_processor.validate_email_address(email)
    
    def _log_processing_results(self, success_count: int, total_count: int) -> None:
        """処理結果をログに記録"""
        self.logger.info("=" * 50)
        self.logger.info("処理結果サマリー")
        self.logger.info(f"処理対象: {total_count}件")
        self.logger.info(f"成功: {success_count}件")
        self.logger.info(f"失敗: {total_count - success_count}件")
        self.logger.info(f"生成されたコンテンツ: {self.statistics['processed_contents']}件")
        self.logger.info(f"生成されたPDF: {self.statistics['generated_pdfs']}件")
        self.logger.info(f"送信されたメール: {self.statistics['sent_emails']}件")
        self.logger.info(f"エラー数: {self.statistics['errors']}件")
        self.logger.info("=" * 50)
    
    def _cleanup_temp_files(self, year: str, month: str) -> None:
        """一時ファイルをクリーンアップ"""
        try:
            output_dir = self.config.get_output_directory(year, month)
            
            # 一時ファイルのクリーンアップ
            self.pdf_converter.cleanup_temp_files(output_dir)
            
            self.logger.info("一時ファイルのクリーンアップが完了しました")
            
        except Exception as e:
            self.system_logger.log_error_details(e, "一時ファイルクリーンアップ")
    
    def test_system_components(self) -> Dict[str, bool]:
        """システムコンポーネントのテスト"""
        test_results = {}
        
        try:
            # Gmail API接続テスト
            test_results['gmail_api'] = self.email_processor.test_connection()
            
            # ファイルアクセステスト
            test_results['file_access'] = self._test_file_access()
            
            # 設定テスト
            test_results['configuration'] = self._test_configuration()
            
            # 結果をログに記録
            self.logger.info("システムテスト結果:")
            for component, result in test_results.items():
                status = "OK" if result else "NG"
                self.logger.info(f"  {component}: {status}")
            
            return test_results
            
        except Exception as e:
            self.system_logger.log_error_details(e, "システムテスト")
            return test_results
    
    def _test_file_access(self) -> bool:
        """ファイルアクセステスト"""
        try:
            # 基本ディレクトリの存在確認
            for path_name, path_value in self.config.base_paths.items():
                if not Path(path_value).exists():
                    self.logger.warning(f"パスが存在しません: {path_name} = {path_value}")
                    return False
            
            return True
            
        except Exception as e:
            self.logger.error(f"ファイルアクセステストエラー: {e}")
            return False
    
    def _test_configuration(self) -> bool:
        """設定テスト"""
        try:
            # 設定ファイルの存在確認
            required_files = [
                self.config.get_contents_mapping_file(),
                self.config.get_rate_data_file()
            ]
            
            for file_path in required_files:
                if not Path(file_path).exists():
                    self.logger.warning(f"設定ファイルが存在しません: {file_path}")
                    return False
            
            return True
            
        except Exception as e:
            self.logger.error(f"設定テストエラー: {e}")
            return False