"""
ログ設定モジュール

システム全体のログ設定を管理します。
"""

import logging
import sys
from pathlib import Path
from datetime import datetime
from typing import Optional


class SystemLogger:
    """システムログ管理クラス"""
    
    def __init__(self, log_level: str = "INFO", log_file: Optional[str] = None):
        """ログ設定を初期化"""
        self.log_level = getattr(logging, log_level.upper(), logging.INFO)
        self.log_file = log_file or self._get_default_log_file()
        self._setup_logging()
    
    def _get_default_log_file(self) -> str:
        """デフォルトのログファイルパスを取得"""
        timestamp = datetime.now().strftime("%Y%m%d")
        return f"content_payment_statement_{timestamp}.log"
    
    def _setup_logging(self) -> None:
        """ログ設定を構築"""
        # ログフォーマットを定義
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # ルートロガーを設定
        root_logger = logging.getLogger()
        root_logger.setLevel(self.log_level)
        
        # 既存のハンドラーをクリア
        root_logger.handlers.clear()
        
        # コンソールハンドラーを追加
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(self.log_level)
        console_handler.setFormatter(formatter)
        root_logger.addHandler(console_handler)
        
        # ファイルハンドラーを追加
        if self.log_file:
            try:
                # ログディレクトリを作成
                log_path = Path(self.log_file)
                log_path.parent.mkdir(parents=True, exist_ok=True)
                
                file_handler = logging.FileHandler(
                    self.log_file, 
                    encoding='utf-8'
                )
                file_handler.setLevel(self.log_level)
                file_handler.setFormatter(formatter)
                root_logger.addHandler(file_handler)
                
            except Exception as e:
                # ファイルハンドラー設定失敗時のフォールバック
                fallback_logger = logging.getLogger(__name__)
                fallback_logger.error(f"ログファイル設定エラー: {e}")
    
    def get_logger(self, name: str) -> logging.Logger:
        """指定された名前のロガーを取得"""
        return logging.getLogger(name)
    
    def log_system_info(self) -> None:
        """システム情報をログに記録"""
        logger = self.get_logger("system")
        logger.info("=" * 50)
        logger.info("コンテンツ関連支払い明細書生成システム 開始")
        logger.info(f"ログレベル: {logging.getLevelName(self.log_level)}")
        logger.info(f"ログファイル: {self.log_file}")
        logger.info("=" * 50)
    
    def log_system_end(self, success: bool = True) -> None:
        """システム終了情報をログに記録"""
        logger = self.get_logger("system")
        logger.info("=" * 50)
        if success:
            logger.info("コンテンツ関連支払い明細書生成システム 正常終了")
        else:
            logger.error("コンテンツ関連支払い明細書生成システム 異常終了")
        logger.info("=" * 50)
    
    def log_progress(self, current: int, total: int, message: str = "") -> None:
        """進捗情報をログに記録"""
        logger = self.get_logger("progress")
        percentage = (current / total * 100) if total > 0 else 0
        progress_message = f"進捗: {current}/{total} ({percentage:.1f}%)"
        if message:
            progress_message += f" - {message}"
        logger.info(progress_message)
    
    def log_error_details(self, error: Exception, context: str = "") -> None:
        """エラーの詳細をログに記録"""
        logger = self.get_logger("error")
        error_message = f"エラー発生"
        if context:
            error_message += f" [{context}]"
        error_message += f": {str(error)}"
        
        logger.error(error_message)
        
        # スタックトレースも記録
        import traceback
        logger.debug(f"スタックトレース:\n{traceback.format_exc()}")
    
    def log_file_operation(self, operation: str, file_path: str, success: bool = True) -> None:
        """ファイル操作をログに記録"""
        logger = self.get_logger("file_operation")
        status = "成功" if success else "失敗"
        logger.info(f"ファイル操作 {status}: {operation} - {file_path}")
    
    def log_data_summary(self, data_type: str, count: int, details: str = "") -> None:
        """データ処理のサマリーをログに記録"""
        logger = self.get_logger("data_summary")
        summary_message = f"{data_type}: {count}件処理"
        if details:
            summary_message += f" ({details})"
        logger.info(summary_message)


def setup_system_logging(log_level: str = "INFO", log_file: Optional[str] = None) -> SystemLogger:
    """システムログを設定して返す"""
    return SystemLogger(log_level, log_file)