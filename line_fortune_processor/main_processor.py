"""
メインコントローラー
"""

import time
from typing import Dict, Any, List, Optional
from datetime import datetime, date
from pathlib import Path

from .config import Config
from .logger import get_logger
from .email_processor import EmailProcessor
from .file_processor import FileProcessor
from .consolidation_processor import ConsolidationProcessor
from .error_handler import retry_on_error, handle_errors, ErrorHandler, RetryableError, FatalError, ErrorType
from .constants import AppConstants, FileConstants
from .messages import MessageFormatter
from .performance_optimizer import ConcurrentProcessor, PerformanceMonitor, performance_monitor


class LineFortuneProcessor:
    """LINE Fortune メール処理メインクラス"""
    
    def __init__(self, config_file: str = None, start_date = None, end_date = None):
        """
        メイン処理を初期化
        
        Args:
            config_file: 設定ファイルのパス
            start_date: メール検索開始日
            end_date: メール検索終了日
        """
        # 設定の読み込み
        self.config = Config(config_file)
        
        # コマンドラインから指定された日付範囲があれば設定を上書き
        if start_date is not None:
            self.config.set('start_date', start_date)
        if end_date is not None:
            self.config.set('end_date', end_date)
        
        # ログの初期化
        self.logger = get_logger(
            self.config.get('log_file', 'line_fortune_processor.log'),
            self.config.get('log_level', 'INFO'),
            self.config.get('use_json_logs', False)
        )
        
        # エラーハンドラー
        self.error_handler = ErrorHandler(self.logger.get_logger())
        
        # パフォーマンス最適化
        self.concurrent_processor = ConcurrentProcessor(max_workers=self.config.get('max_workers', 4))
        self.performance_monitor = PerformanceMonitor()
        
        # 各モジュールの初期化
        self.email_processor = EmailProcessor(self.config.get('email', {}))
        self.file_processor = FileProcessor(
            self.config.get('base_path', ''),
            self.config.get('retry_count', 3),
            self.config.get('retry_delay', 5)
        )
        self.consolidation_processor = ConsolidationProcessor()
        
        # 処理統計
        self.stats = {
            AppConstants.STATS_EMAILS_PROCESSED: 0,
            AppConstants.STATS_EMAILS_SUCCESS: 0,
            AppConstants.STATS_EMAILS_ERROR: 0,
            AppConstants.STATS_FILES_SAVED: 0,
            AppConstants.STATS_CONSOLIDATIONS_CREATED: 0
        }
        
        # 処理した日付を追跡
        self.processed_dates = set()
        
        self.session_id = None
        
    @handle_errors("メイン処理")
    def process(self) -> bool:
        """
        メイン処理を実行
        要件に基づく処理フロー:
        1. メール監視と処理
        2. ファイル名付けと整理 
        3. 統合ファイル生成
        4. エラー処理とログ記録
        
        Returns:
            bool: 処理が成功した場合True
        """
        # セッション開始
        self.session_id = self.logger.start_session()
        
        try:
            # 要件: 設定の検証
            if not self.config.validate():
                validation_errors = self.config.get_validation_errors()
                for error in validation_errors:
                    self.logger.error(f"設定エラー: {error}")
                raise FatalError("設定の検証に失敗しました", ErrorType.UNKNOWN)
            
            # 要件: メールサーバーへの接続
            if not self.email_processor.connect():
                raise FatalError("メールサーバーへの接続に失敗しました", ErrorType.NETWORK)
                
            try:
                # 要件: 条件に一致するメールを取得
                emails = self._fetch_target_emails()
                
                if not emails:
                    self.logger.info("処理対象のメールが見つかりませんでした")
                    return True
                
                self.logger.info(f"処理対象メール: {len(emails)} 件")
                
                # 要件: 各メールを処理し、エラーが発生しても他のメールの処理を継続
                for email in emails:
                    try:
                        self.handle_email(email)
                    except Exception as e:
                        self.logger.error(MessageFormatter.get_email_message(
                            "email_processing_error", subject=email.get('subject', 'Unknown')
                        ), exception=e)
                        self.stats[AppConstants.STATS_EMAILS_ERROR] += 1
                        continue
                    
                # 全メール処理完了後に統合処理を実行
                self._consolidate_all_data()
                
                # 処理結果をログに記録
                self._log_processing_results()
                
                # パフォーマンスサマリーをログ出力（一時的に無効化）
                # self.performance_monitor.log_performance_summary()
                
                # セッション終了
                success = self.stats[AppConstants.STATS_EMAILS_ERROR] == 0
                self.logger.end_session(success)
                
                return success
                
            finally:
                # メールサーバーから切断
                try:
                    self.email_processor.disconnect()
                except Exception as e:
                    self.logger.warning(f"メールサーバー切断中にエラー: {e}")
                
        except Exception as e:
            self.logger.end_session(False)
            raise
    
    @handle_errors("メール処理")
    def handle_email(self, email_info: Dict[str, Any]) -> bool:
        """
        単一のメールを処理
        要件: エラーが発生しても他のメールの処理を継続
        
        Args:
            email_info: メール情報
            
        Returns:
            bool: 処理が成功した場合True
        """
        self.stats[AppConstants.STATS_EMAILS_PROCESSED] += 1
        email_id = email_info.get('id', 'unknown')
        subject = email_info.get('subject', '')
        
        self.logger.info(f"メール処理開始: {subject}", email_id=email_id)
        
        # 要件: 件名から日付を抽出
        target_date = self.email_processor.extract_date_from_subject(subject)
        
        if not target_date:
            self.logger.warning(f"日付の抽出に失敗しました: {subject}", email_id=email_id)
            self.stats[AppConstants.STATS_EMAILS_ERROR] += 1
            return False
            
        # 処理対象日付をログに記録し、追跡リストに追加
        self.logger.info(f"処理対象日付: {target_date} (件名: {subject[:30]}...)", email_id=email_id)
        self.processed_dates.add(target_date)
        
        # 要件: CSV添付ファイルを抽出
        attachments = self.email_processor.extract_attachments(email_info, '.csv')
        
        if not attachments:
            self.logger.warning(f"CSV添付ファイルが見つかりませんでした", email_id=email_id, subject=subject)
            # デバッグ用：メール構造の詳細を一時的にINFOレベルで出力
            self.logger.info("=== 添付ファイルが見つからないメールの詳細調査 ===")
            email_message = email_info.get('message')
            if email_message:
                for i, part in enumerate(email_message.walk()):
                    content_type = part.get_content_type()
                    content_disposition = part.get_content_disposition()
                    filename = part.get_filename()
                    decoded_filename = self.email_processor._get_attachment_filename(part)
                    is_multipart = part.is_multipart()
                    
                    self.logger.info(f"Part {i}: "
                                   f"content_type={content_type}, "
                                   f"disposition={content_disposition}, "
                                   f"filename={filename}, "
                                   f"decoded_filename={decoded_filename}, "
                                   f"multipart={is_multipart}")
                                   
                    # CSVファイルかどうかをチェック
                    if decoded_filename and decoded_filename.lower().endswith('.csv'):
                        self.logger.info(f"  *** CSVファイルを検出: {decoded_filename} ***")
                        
                    # 全ヘッダーを出力
                    if hasattr(part, 'items') and part.items():
                        for header_name, header_value in part.items():
                            self.logger.info(f"  Header {header_name}: {header_value}")
            self.logger.info("=== 詳細調査終了 ===")
            self.stats[AppConstants.STATS_EMAILS_ERROR] += 1
            return False
        
        self.logger.info(f"CSV添付ファイルを {len(attachments)} 件発見", email_id=email_id)
        
        # 要件: 対象ディレクトリを作成
        try:
            target_dir = self.file_processor.create_directory_structure(target_date)
        except Exception as e:
            self.logger.error(f"ディレクトリの作成に失敗しました: {target_date}", email_id=email_id, exception=e)
            self.stats[AppConstants.STATS_EMAILS_ERROR] += 1
            return False
        
        # 要件: 各添付ファイルを処理（並行処理で最適化）
        try:
            if len(attachments) > 1 and self.config.get('enable_parallel_processing', True):
                # 複数ファイルの場合は並行処理
                results = self.concurrent_processor.process_attachments_concurrently(
                    attachments, self.process_attachment, target_date, target_dir
                )
                saved_files = [
                    attachment['filename'] 
                    for i, attachment in enumerate(attachments) 
                    if i < len(results) and results[i]
                ]
                self.stats[AppConstants.STATS_FILES_SAVED] += len(saved_files)
            else:
                # 単一ファイルまたは順次処理
                saved_files = []
                for attachment in attachments:
                    try:
                        if self.process_attachment(attachment, target_date, target_dir):
                            saved_files.append(attachment['filename'])
                            self.stats[AppConstants.STATS_FILES_SAVED] += 1
                    except Exception as e:
                        self.logger.error(f"添付ファイル処理中にエラーが発生、次のファイルに進みます: {attachment.get('filename', 'Unknown')}", email_id=email_id, exception=e)
                        continue
        except Exception as e:
            self.logger.error("添付ファイル並行処理中にエラーが発生しました", email_id=email_id, exception=e)
            saved_files = []
        
        if not saved_files:
            self.logger.error(f"添付ファイルの保存に失敗しました", email_id=email_id, subject=subject)
            self.stats[AppConstants.STATS_EMAILS_ERROR] += 1
            return False
        
        # 月次統合は全メール処理完了後に実行
        
        # メールを既読にマーク
        try:
            if email_id and email_id != 'unknown':
                self.email_processor.mark_as_read(email_id)
        except Exception as e:
            self.logger.warning(f"メール既読マーク中にエラーが発生しました", email_id=email_id, exception=e)
        
        self.logger.info(f"メール処理が完了しました", email_id=email_id, files_saved=len(saved_files))
        self.stats[AppConstants.STATS_EMAILS_SUCCESS] += 1
        return True
    
    def _fetch_target_emails(self) -> List[Dict[str, Any]]:
        """要件に基づいて対象メールを取得"""
        return self.email_processor.fetch_matching_emails(
            self.config.get('sender'),
            self.config.get('recipient'),
            self.config.get('subject_pattern'),
            start_date=self.config.get('start_date'),
            end_date=self.config.get('end_date'),
            days_back=self.config.get('search_days', 7)
        )
    
    def _log_processing_results(self):
        """処理結果をログに記録"""
        self.logger.info(
            MessageFormatter.get_session_message(
                "processing_results",
                processed=self.stats[AppConstants.STATS_EMAILS_PROCESSED],
                success=self.stats[AppConstants.STATS_EMAILS_SUCCESS],
                error=self.stats[AppConstants.STATS_EMAILS_ERROR]
            )
        )
        self.logger.info(
            MessageFormatter.get_session_message(
                "statistics",
                emails=self.stats[AppConstants.STATS_EMAILS_PROCESSED],
                files=self.stats[AppConstants.STATS_FILES_SAVED],
                consolidations=self.stats[AppConstants.STATS_CONSOLIDATIONS_CREATED]
            )
        )
    
    def process_attachment(self, attachment: Dict[str, Any], target_date: date, target_dir: Path) -> bool:
        """
        添付ファイルを処理
        
        Args:
            attachment: 添付ファイル情報
            target_date: 対象日付
            target_dir: 保存先ディレクトリ
            
        Returns:
            bool: 処理が成功した場合True
        """
        try:
            original_filename = attachment['filename']
            file_content = attachment['content']
            
            # lineサブディレクトリを作成
            line_dir = target_dir / "line"
            line_dir.mkdir(exist_ok=True)
            
            # ファイル名をリネーム
            new_filename = self.file_processor.rename_file(original_filename, target_date)
            
            # 既存ファイルのバックアップを作成
            if self.file_processor.file_exists(new_filename, line_dir):
                self.logger.info(f"既存ファイルを上書きします: {new_filename}")
                # バックアップ処理を無効化し、上書き保存
                # self.file_processor.backup_file(new_filename, line_dir)
            
            # ファイルを保存（lineディレクトリに保存）
            if self.file_processor.save_file(file_content, new_filename, line_dir):
                self.logger.log_file_operation("保存", new_filename, True)
                return True
            else:
                self.logger.log_file_operation("保存", new_filename, False)
                return False
                
        except Exception as e:
            self.logger.error(f"添付ファイル処理中にエラーが発生しました: {attachment.get('filename', 'Unknown')}", exception=e)
            return False
    
    def dry_run(self) -> bool:
        """
        ドライラン実行（実際の処理は行わない）
        
        Returns:
            bool: 処理が成功した場合True
        """
        try:
            self.logger.info("ドライラン実行を開始します")
            
            # 設定の検証
            if not self.config.validate():
                self.logger.error("設定の検証に失敗しました")
                return False
            
            # メールサーバーに接続
            if not self.email_processor.connect():
                self.logger.error("メールサーバーへの接続に失敗しました")
                return False
                
            try:
                # 条件に一致するメールを取得（7日間の範囲で検索）
                emails = self.email_processor.fetch_matching_emails(
                    self.config.get('sender'),
                    self.config.get('recipient'),
                    self.config.get('subject_pattern'),
                    start_date=self.config.get('start_date'),
                    end_date=self.config.get('end_date'),
                    days_back=self.config.get('search_days', 7)  # 日付が指定されていない場合のデフォルト
                )
                
                self.logger.info(f"処理対象のメールが {len(emails)} 件見つかりました")
                
                # 各メールの情報を表示
                for i, email in enumerate(emails, 1):
                    subject = email.get('subject', '')
                    sender = email.get('sender', '')
                    target_date = self.email_processor.extract_date_from_subject(subject)
                    
                    self.logger.info(f"メール {i}: {subject} from {sender}")
                    if target_date:
                        self.logger.info(f"  対象日付: {target_date}")
                    
                    # 添付ファイルを確認
                    attachments = self.email_processor.extract_attachments(email, '.csv')
                    self.logger.info(f"  CSV添付ファイル数: {len(attachments)}")
                    
                    for attachment in attachments:
                        filename = attachment['filename']
                        size = len(attachment['content'])
                        self.logger.info(f"    - {filename} ({size} bytes)")
                
                return True
                
            finally:
                # メールサーバーから切断
                self.email_processor.disconnect()
                
        except Exception as e:
            self.logger.error("ドライラン実行中にエラーが発生しました", exception=e)
            return False
    
    def get_stats(self) -> Dict[str, Any]:
        """
        処理統計を取得
        
        Returns:
            Dict: 処理統計
        """
        return self.stats.copy()
    
    def cleanup_old_files(self, days_to_keep: int = 30) -> bool:
        """
        古いファイルを削除
        
        Args:
            days_to_keep: 保持する日数
            
        Returns:
            bool: 処理が成功した場合True
        """
        try:
            base_path = Path(self.config.get('base_path', ''))
            
            if not base_path.exists():
                self.logger.warning(f"ベースパスが存在しません: {base_path}")
                return False
            
            total_deleted = 0
            
            # 年フォルダを探索
            for year_dir in base_path.iterdir():
                if year_dir.is_dir() and year_dir.name.isdigit():
                    # 月フォルダを探索
                    for month_dir in year_dir.iterdir():
                        if month_dir.is_dir():
                            deleted_count = self.file_processor.cleanup_old_files(
                                month_dir, days_to_keep
                            )
                            total_deleted += deleted_count
            
            self.logger.info(f"古いファイルを合計 {total_deleted} 個削除しました")
            return True
            
        except Exception as e:
            self.logger.error("古いファイル削除中にエラーが発生しました", exception=e)
            return False
    
    def _consolidate_all_data(self):
        """
        処理した年月期間のデータのみを統合する
        """
        try:
            base_path = Path(self.config.get('base_path', ''))
            if not base_path.exists():
                return
            
            # 今回処理したメールから対象の年月を特定
            processed_dates = self._get_processed_date_range()
            if not processed_dates:
                self.logger.info("処理対象の年月が特定できませんでした")
                return
                
            directories_processed = set()
            
            # 処理した年月期間のみを対象とする
            for target_date in processed_dates:
                year_month = target_date.strftime('%Y%m')
                year = target_date.strftime('%Y')
                
                # 対応するディレクトリパスを構築
                year_dir = base_path / year
                month_dir = year_dir / year_month
                line_dir = month_dir / "line"
                
                if line_dir.exists() and str(line_dir) not in directories_processed:
                    # 月次統合処理を実行
                    if self.consolidation_processor.consolidate_monthly_data(line_dir):
                        self.stats[AppConstants.STATS_CONSOLIDATIONS_CREATED] += 1
                        self.logger.info(f"月次統合ファイルを作成しました: {line_dir}")
                        directories_processed.add(str(line_dir))
                    else:
                        self.logger.warning(f"月次統合に失敗しました: {line_dir}")
            
            self.logger.info(f"統合処理完了: {len(directories_processed)} ディレクトリ処理 (処理対象年月: {len(processed_dates)})")
            
        except Exception as e:
            self.logger.error("統合処理中にエラーが発生しました", exception=e)
    
    def _get_processed_date_range(self) -> set:
        """
        処理した日付の範囲を取得
        
        Returns:
            set: 処理した日付のセット
        """
        return self.processed_dates