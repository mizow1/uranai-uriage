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
            self.config.config['start_date'] = start_date
        if end_date is not None:
            self.config.config['end_date'] = end_date
        
        # ログの初期化
        self.logger = get_logger(
            self.config.get('log_file', 'line_fortune_processor.log'),
            self.config.get('log_level', 'INFO')
        )
        
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
            'emails_processed': 0,
            'emails_success': 0,
            'emails_error': 0,
            'files_saved': 0,
            'consolidations_created': 0
        }
        
        self.session_id = None
        
    def process(self) -> bool:
        """
        メイン処理を実行
        
        Returns:
            bool: 処理が成功した場合True
        """
        try:
            # セッション開始
            self.session_id = self.logger.log_session_start()
            
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
                
                if not emails:
                    self.logger.info("処理対象のメールが見つかりませんでした")
                    return True
                
                # 各メールを処理
                for email in emails:
                    self.handle_email(email)
                    
                # 処理結果をログに記録
                self.logger.log_email_processing(
                    self.stats['emails_processed'],
                    self.stats['emails_success'],
                    self.stats['emails_error']
                )
                
                # セッション終了
                success = self.stats['emails_error'] == 0
                self.logger.log_session_end(self.session_id, success)
                
                return success
                
            finally:
                # メールサーバーから切断
                self.email_processor.disconnect()
                
        except Exception as e:
            self.logger.error("メイン処理中にエラーが発生しました", exception=e)
            if self.session_id:
                self.logger.log_session_end(self.session_id, False)
            return False
    
    def handle_email(self, email_info: Dict[str, Any]) -> bool:
        """
        単一のメールを処理
        
        Args:
            email_info: メール情報
            
        Returns:
            bool: 処理が成功した場合True
        """
        try:
            self.stats['emails_processed'] += 1
            
            # 件名から日付を抽出
            subject = email_info.get('subject', '')
            target_date = self.email_processor.extract_date_from_subject(subject)
            
            if not target_date:
                self.logger.warning(f"日付の抽出に失敗しました: {subject}")
                self.stats['emails_error'] += 1
                return False
            
            # 添付ファイルを抽出
            attachments = self.email_processor.extract_attachments(email_info, '.csv')
            
            if not attachments:
                self.logger.warning(f"CSV添付ファイルが見つかりませんでした: {subject}")
                self.stats['emails_error'] += 1
                return False
            
            # 対象ディレクトリを作成
            target_dir = self.file_processor.create_directory_structure(target_date)
            
            if not target_dir:
                self.logger.error(f"ディレクトリの作成に失敗しました: {target_date}")
                self.stats['emails_error'] += 1
                return False
            
            # 各添付ファイルを処理
            saved_files = []
            for attachment in attachments:
                if self.process_attachment(attachment, target_date, target_dir):
                    saved_files.append(attachment['filename'])
                    self.stats['files_saved'] += 1
            
            if not saved_files:
                self.logger.error(f"添付ファイルの保存に失敗しました: {subject}")
                self.stats['emails_error'] += 1
                return False
            
            # 月次統合を実行
            if self.consolidation_processor.consolidate_monthly_data(target_dir):
                self.stats['consolidations_created'] += 1
                self.logger.log_consolidation_result(
                    str(target_dir),
                    len(saved_files),
                    self.consolidation_processor.generate_monthly_filename(
                        target_date.year, target_date.month
                    ),
                    True
                )
            else:
                self.logger.warning(f"月次統合に失敗しました: {target_dir}")
            
            # メールを既読にマーク
            email_id = email_info.get('id')
            if email_id:
                self.email_processor.mark_as_read(email_id)
            
            self.logger.info(f"メール処理が完了しました: {subject}")
            self.stats['emails_success'] += 1
            return True
            
        except Exception as e:
            self.logger.error(f"メール処理中にエラーが発生しました: {email_info.get('subject', 'Unknown')}", exception=e)
            self.stats['emails_error'] += 1
            return False
    
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