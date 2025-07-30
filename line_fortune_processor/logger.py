"""
ログ記録モジュール
"""

import logging
import logging.handlers
import json
import os
from pathlib import Path
from datetime import datetime
from typing import Optional, Dict, Any
import uuid


class StructuredLogger:
    """構造化ログフォーマッター"""
    
    def format(self, record: logging.LogRecord) -> str:
        """ログレコードをJSON形式でフォーマット"""
        log_entry = {
            'timestamp': datetime.fromtimestamp(record.created).isoformat(),
            'level': record.levelname,
            'logger': record.name,
            'function': record.funcName,
            'line': record.lineno,
            'message': record.getMessage()
        }
        
        # 追加のコンテキスト情報を含める
        if hasattr(record, 'session_id'):
            log_entry['session_id'] = record.session_id
        if hasattr(record, 'email_id'):
            log_entry['email_id'] = record.email_id
        if hasattr(record, 'file_path'):
            log_entry['file_path'] = record.file_path
        if hasattr(record, 'operation'):
            log_entry['operation'] = record.operation
        if hasattr(record, 'error_type'):
            log_entry['error_type'] = record.error_type
        
        return json.dumps(log_entry, ensure_ascii=False)


class Logger:
    """拡張されたカスタムログクラス"""
    
    def __init__(self, log_file: str = "line_fortune_processor.log", log_level: str = "INFO", use_json: bool = False):
        """
        ログ記録を初期化
        
        Args:
            log_file: ログファイル名
            log_level: ログレベル
            use_json: JSON形式でログを出力するか
        """
        self.log_file = log_file
        self.log_level = log_level
        self.use_json = use_json
        self.logger = None
        self.session_id = None
        self._setup_logger()
    
    def _setup_logger(self):
        """ログ記録の設定"""
        try:
            # ログディレクトリの作成
            log_dir = Path("logs")
            log_dir.mkdir(exist_ok=True)
            
            # ログファイルの完全パス
            log_path = log_dir / self.log_file
            
            # ログレベルの設定
            numeric_level = getattr(logging, self.log_level.upper(), logging.INFO)
            
            # ルートロガーの設定
            self.logger = logging.getLogger("line_fortune_processor")
            self.logger.setLevel(numeric_level)
            
            # 既存のハンドラーをクリア
            self.logger.handlers.clear()
            
            # ファイルハンドラーの設定（ローテーション付き）
            file_handler = logging.handlers.RotatingFileHandler(
                log_path,
                maxBytes=10*1024*1024,  # 10MB
                backupCount=5,
                encoding='utf-8'
            )
            file_handler.setLevel(numeric_level)
            
            # コンソールハンドラーの設定
            console_handler = logging.StreamHandler()
            console_handler.setLevel(numeric_level)
            
            # フォーマッターの設定
            if self.use_json:
                # JSONフォーマッターをファイル用に使用
                json_formatter = StructuredLogger()
                file_handler.setFormatter(json_formatter)
                
                # コンソール用は読みやすいフォーマット
                console_formatter = logging.Formatter(
                    '%(asctime)s - %(levelname)s - %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S'
                )
                console_handler.setFormatter(console_formatter)
            else:
                # 総用フォーマッター
                formatter = logging.Formatter(
                    '%(asctime)s - %(name)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S'
                )
                file_handler.setFormatter(formatter)
                console_handler.setFormatter(formatter)
            
            # ハンドラーをロガーに追加
            self.logger.addHandler(file_handler)
            self.logger.addHandler(console_handler)
            
            # 他のロガーの設定
            logging.getLogger().setLevel(logging.WARNING)  # 他のライブラリのログレベルを上げる
            
            self.logger.info(f"ログ記録を初期化しました: {log_path}")
            
        except Exception as e:
            # ログ初期化失敗時のフォールバック
            import logging
            fallback_logger = logging.getLogger(__name__)
            fallback_logger.error(f"ログ記録の初期化に失敗しました: {e}")
            # フォールバックとして基本的なログ設定を行う
            logging.basicConfig(
                level=logging.INFO,
                format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
            )
            self.logger = logging.getLogger("line_fortune_processor")
    
    def info(self, message: str, **kwargs):
        """
        情報メッセージをログに記録
        
        Args:
            message: ログメッセージ
            **kwargs: 追加の情報
        """
        if self.logger:
            extra = self._prepare_extra_data(kwargs)
            if self.use_json:
                self.logger.info(message, extra=extra)
            else:
                extra_info = self._format_extra_info(kwargs)
                self.logger.info(f"{message}{extra_info}")
    
    def warning(self, message: str, **kwargs):
        """
        警告メッセージをログに記録
        
        Args:
            message: ログメッセージ
            **kwargs: 追加の情報
        """
        if self.logger:
            extra_info = self._format_extra_info(kwargs)
            self.logger.warning(f"{message}{extra_info}")
    
    def error(self, message: str, exception: Optional[Exception] = None, **kwargs):
        """
        エラーメッセージをログに記録
        
        Args:
            message: ログメッセージ
            exception: 例外オブジェクト
            **kwargs: 追加の情報
        """
        if self.logger:
            extra_info = self._format_extra_info(kwargs)
            if exception:
                self.logger.error(f"{message}{extra_info}", exc_info=exception)
            else:
                self.logger.error(f"{message}{extra_info}")
    
    def debug(self, message: str, **kwargs):
        """
        デバッグメッセージをログに記録
        
        Args:
            message: ログメッセージ
            **kwargs: 追加の情報
        """
        if self.logger:
            extra_info = self._format_extra_info(kwargs)
            self.logger.debug(f"{message}{extra_info}")
    
    def critical(self, message: str, exception: Optional[Exception] = None, **kwargs):
        """
        クリティカルメッセージをログに記録
        
        Args:
            message: ログメッセージ
            exception: 例外オブジェクト
            **kwargs: 追加の情報
        """
        if self.logger:
            extra_info = self._format_extra_info(kwargs)
            if exception:
                self.logger.critical(f"{message}{extra_info}", exc_info=exception)
            else:
                self.logger.critical(f"{message}{extra_info}")
    
    def _format_extra_info(self, kwargs: dict) -> str:
        """
        追加情報をフォーマット
        
        Args:
            kwargs: 追加情報の辞書
            
        Returns:
            str: フォーマットされた追加情報
        """
        if not kwargs:
            return ""
            
        formatted_info = []
        for key, value in kwargs.items():
            # 機密情報をマスク
            if key.lower() in ['password', 'token', 'secret', 'key']:
                value = '*' * len(str(value))
            formatted_info.append(f"{key}={value}")
            
        return f" [{', '.join(formatted_info)}]"
    
    def _prepare_extra_data(self, kwargs: Dict[str, Any]) -> Dict[str, Any]:
        """追加データを準備"""
        extra = kwargs.copy()
        if self.session_id:
            extra['session_id'] = self.session_id
        return extra
    
    def start_session(self, session_id: str = None) -> str:
        """
        セッションを開始
        
        Args:
            session_id: セッションID
            
        Returns:
            str: 生成されたセッションID
        """
        if not session_id:
            session_id = str(uuid.uuid4())[:8]
            
        self.session_id = session_id
        self.info("セッション開始", operation="session_start")
        return session_id
    
    def end_session(self, success: bool = True):
        """
        セッションを終了
        
        Args:
            success: 成功/失敗フラグ
        """
        status = "success" if success else "failure"
        self.info(f"セッション終了: {status}", operation="session_end", success=success)
        self.session_id = None
    
    def log_session_start(self, session_id: str = None):
        """後方互換性のためのメソッド"""
        return self.start_session(session_id)
    
    def log_session_end(self, session_id: str = None, success: bool = True):
        """
        セッション終了をログに記録
        
        Args:
            session_id: セッションID
            success: 成功/失敗フラグ
        """
        status = "成功" if success else "失敗"
        self.info(f"=== セッション終了 ({status}) ===", session_id=session_id)
    
    def log_email_processing(self, email_count: int, success_count: int, error_count: int):
        """
        メール処理結果をログに記録
        
        Args:
            email_count: 処理したメール数
            success_count: 成功したメール数
            error_count: エラーが発生したメール数
        """
        self.info(
            f"メール処理結果: 処理数={email_count}, 成功={success_count}, エラー={error_count}"
        )
    
    def log_file_operation(self, operation: str, filename: str, success: bool = True):
        """
        ファイル操作をログに記録
        
        Args:
            operation: 操作タイプ
            filename: ファイル名
            success: 成功/失敗フラグ
        """
        status = "成功" if success else "失敗"
        self.info(f"ファイル操作 ({operation}): {filename} - {status}")
    
    def log_consolidation_result(self, directory: str, file_count: int, output_file: str, success: bool = True):
        """
        統合処理結果をログに記録
        
        Args:
            directory: 処理ディレクトリ
            file_count: 統合したファイル数
            output_file: 出力ファイル名
            success: 成功/失敗フラグ
        """
        status = "成功" if success else "失敗"
        self.info(
            f"統合処理結果: ディレクトリ={directory}, ファイル数={file_count}, 出力={output_file} - {status}"
        )
    
    def get_logger(self):
        """
        ログオブジェクトを取得
        
        Returns:
            logging.Logger: ログオブジェクト
        """
        return self.logger
    
    def set_level(self, level: str):
        """
        ログレベルを変更
        
        Args:
            level: 新しいログレベル
        """
        if self.logger:
            numeric_level = getattr(logging, level.upper(), logging.INFO)
            self.logger.setLevel(numeric_level)
            for handler in self.logger.handlers:
                handler.setLevel(numeric_level)
            self.info(f"ログレベルを変更しました: {level}")


# シングルトンパターンでログインスタンスを管理
_logger_instance = None


def get_logger(log_file: str = "line_fortune_processor.log", log_level: str = "INFO", use_json: bool = False) -> Logger:
    """
    ログインスタンスを取得（シングルトン）
    
    Args:
        log_file: ログファイル名
        log_level: ログレベル
        use_json: JSONフォーマットを使用するか
        
    Returns:
        Logger: ログインスタンス
    """
    global _logger_instance
    if _logger_instance is None:
        _logger_instance = Logger(log_file, log_level, use_json)
    return _logger_instance