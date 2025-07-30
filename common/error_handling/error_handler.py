"""
統一エラーハンドリングシステム
"""
from pathlib import Path
from typing import Dict, Any, Optional
import traceback
from .exceptions import FileProcessingError, DataValidationError


class ErrorHandler:
    """エラーハンドリングの統一クラス"""
    
    def __init__(self, logger=None):
        self.logger = logger
    
    def handle_file_processing_error(self, error: Exception, file_path: Path) -> None:
        """ファイル処理エラーを処理"""
        error_context = {
            'error_type': type(error).__name__,
            'file_path': str(file_path),
            'file_name': file_path.name if file_path else 'Unknown',
            'error_message': str(error)
        }
        
        self.log_error_with_context(error, error_context)
    
    def handle_data_validation_error(self, error: Exception, data_context: str) -> None:
        """データ検証エラーを処理"""
        error_context = {
            'error_type': type(error).__name__,
            'data_context': data_context,
            'error_message': str(error)
        }
        
        self.log_error_with_context(error, error_context)
    
    def log_and_continue(self, error: Exception, context: str) -> None:
        """エラーをログ出力して処理を継続"""
        if self.logger:
            self.logger.error(f"処理継続エラー [{context}]: {str(error)}")
            self.logger.debug(f"エラー詳細: {traceback.format_exc()}")
        else:
            # logger利用不可時のフォールバック
            import logging
            fallback_logger = logging.getLogger(__name__)
            fallback_logger.error(f"エラー [{context}]: {str(error)}")
    
    def log_and_raise(self, error: Exception, context: str) -> None:
        """エラーをログ出力して例外を再発生"""
        if self.logger:
            self.logger.error(f"致命的エラー [{context}]: {str(error)}")
            self.logger.debug(f"エラー詳細: {traceback.format_exc()}")
        else:
            # logger利用不可時のフォールバック
            import logging
            fallback_logger = logging.getLogger(__name__)
            fallback_logger.error(f"致命的エラー [{context}]: {str(error)}")
        
        raise error
    
    def log_error_with_context(self, error: Exception, context: Dict[str, Any]) -> None:
        """コンテキスト情報付きでエラーをログ出力"""
        if self.logger:
            context_str = ", ".join([f"{k}={v}" for k, v in context.items()])
            self.logger.error(f"エラー詳細: {context_str}")
            self.logger.debug(f"スタックトレース: {traceback.format_exc()}")
        else:
            # logger利用不可時のフォールバック
            import logging
            fallback_logger = logging.getLogger(__name__)
            fallback_logger.error(f"エラー: {context}")
            fallback_logger.error(f"詳細: {str(error)}")
    
    def handle_processing_result(self, result: Optional[Any], operation: str, file_path: Optional[Path] = None) -> bool:
        """処理結果を評価してエラーハンドリング"""
        if result is None:
            error_msg = f"{operation}が失敗しました"
            if file_path:
                error_msg += f": {file_path.name}"
            
            if self.logger:
                self.logger.error(error_msg)
            else:
                # logger利用不可時のフォールバック
                import logging
                fallback_logger = logging.getLogger(__name__)
                fallback_logger.error(error_msg)
            return False
        
        return True
    
    def create_error_summary(self, errors: list) -> Dict[str, Any]:
        """エラーリストから統計情報を作成"""
        if not errors:
            return {'total_errors': 0, 'error_types': {}}
        
        error_types = {}
        for error in errors:
            error_type = type(error).__name__
            error_types[error_type] = error_types.get(error_type, 0) + 1
        
        return {
            'total_errors': len(errors),
            'error_types': error_types,
            'first_error': str(errors[0]) if errors else None,
            'last_error': str(errors[-1]) if errors else None
        }