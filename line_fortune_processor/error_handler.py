"""
共通エラーハンドリングモジュール
"""

import logging
import time
from typing import Any, Callable, Optional, Type, Union
from functools import wraps
from enum import Enum


class ErrorType(Enum):
    """エラータイプの分類"""
    NETWORK = "network"
    AUTHENTICATION = "authentication"
    FILE_SYSTEM = "file_system"
    PARSING = "parsing"
    UNKNOWN = "unknown"


class RetryableError(Exception):
    """再試行可能なエラー"""
    def __init__(self, message: str, error_type: ErrorType = ErrorType.UNKNOWN, original_error: Exception = None):
        super().__init__(message)
        self.error_type = error_type
        self.original_error = original_error


class FatalError(Exception):
    """致命的エラー（再試行不可）"""
    def __init__(self, message: str, error_type: ErrorType = ErrorType.UNKNOWN, original_error: Exception = None):
        super().__init__(message)
        self.error_type = error_type
        self.original_error = original_error


class ErrorHandler:
    """統一エラーハンドラー"""
    
    def __init__(self, logger: logging.Logger):
        self.logger = logger
    
    def classify_error(self, error: Exception) -> ErrorType:
        """エラーを分類"""
        error_str = str(error).lower()
        
        if any(keyword in error_str for keyword in ['connection', 'network', 'timeout', 'socket']):
            return ErrorType.NETWORK
        elif any(keyword in error_str for keyword in ['authentication', 'login', 'password', 'credential']):
            return ErrorType.AUTHENTICATION
        elif any(keyword in error_str for keyword in ['permission', 'access', 'file', 'directory']):
            return ErrorType.FILE_SYSTEM
        elif any(keyword in error_str for keyword in ['parse', 'decode', 'format']):
            return ErrorType.PARSING
        else:
            return ErrorType.UNKNOWN
    
    def is_retryable(self, error: Exception, error_type: ErrorType) -> bool:
        """エラーが再試行可能かどうか判定"""
        if isinstance(error, FatalError):
            return False
        if isinstance(error, RetryableError):
            return True
            
        if error_type == ErrorType.NETWORK:
            return True
        elif error_type == ErrorType.FILE_SYSTEM:
            return True
        elif error_type == ErrorType.AUTHENTICATION:
            return False  # 認証エラーは通常再試行不可
        elif error_type == ErrorType.PARSING:
            return False  # パースエラーは再試行しても解決しない
        else:
            return False  # 不明なエラーは安全のため再試行しない
    
    def log_error(self, error: Exception, context: str = "", error_type: ErrorType = None):
        """エラーを標準化されたフォーマットでログに記録"""
        if error_type is None:
            error_type = self.classify_error(error)
        
        self.logger.error(f"[{error_type.value.upper()}] {context}: {str(error)}")
        
        if hasattr(error, 'original_error') and error.original_error:
            self.logger.error(f"原因: {str(error.original_error)}")


def retry_on_error(max_retries: int = 3, base_delay: float = 1.0, max_delay: float = 60.0, backoff_factor: float = 2.0):
    """
    指数バックオフによる再試行デコレータ
    
    Args:
        max_retries: 最大再試行回数
        base_delay: 初期遅延時間（秒）
        max_delay: 最大遅延時間（秒）
        backoff_factor: バックオフ係数
    """
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs) -> Any:
            last_error = None
            delay = base_delay
            
            # selfからloggerを取得（存在する場合）
            logger = None
            if args and hasattr(args[0], 'logger'):
                logger = args[0].logger
            elif args and hasattr(args[0], '_logger'):
                logger = args[0]._logger
            
            if not logger:
                logger = logging.getLogger(__name__)
            
            error_handler = ErrorHandler(logger)
            
            for attempt in range(max_retries + 1):
                try:
                    return func(*args, **kwargs)
                    
                except Exception as e:
                    last_error = e
                    error_type = error_handler.classify_error(e)
                    
                    if attempt == max_retries:
                        error_handler.log_error(e, f"{func.__name__}の実行に失敗（最大再試行回数に到達）", error_type)
                        raise
                    
                    if not error_handler.is_retryable(e, error_type):
                        error_handler.log_error(e, f"{func.__name__}で致命的エラーが発生", error_type)
                        raise
                    
                    logger.warning(f"{func.__name__}の実行に失敗（試行 {attempt + 1}/{max_retries + 1}）: {str(e)}")
                    logger.info(f"{delay:.1f}秒後に再試行します...")
                    
                    time.sleep(delay)
                    delay = min(delay * backoff_factor, max_delay)
            
            raise last_error
            
        return wrapper
    return decorator


def handle_errors(context: str = ""):
    """
    エラー処理デコレータ（再試行なし）
    
    Args:
        context: エラーのコンテキスト情報
    """
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs) -> Any:
            try:
                return func(*args, **kwargs)
            except Exception as e:
                # selfからloggerを取得（存在する場合）
                logger = None
                if args and hasattr(args[0], 'logger'):
                    logger = args[0].logger
                elif args and hasattr(args[0], '_logger'):
                    logger = args[0]._logger
                
                if not logger:
                    logger = logging.getLogger(__name__)
                
                error_handler = ErrorHandler(logger)
                error_type = error_handler.classify_error(e)
                
                error_context = context or f"{func.__name__}の実行中"
                error_handler.log_error(e, error_context, error_type)
                
                raise
                
        return wrapper
    return decorator