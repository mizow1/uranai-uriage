"""
統一ロギングシステム
"""
import logging
from pathlib import Path
from typing import Dict, Any, Optional
from datetime import datetime


class UnifiedLogger:
    """統一ロギングシステムクラス"""
    
    def __init__(self, name: str = __name__, level: str = "INFO", log_file: Optional[Path] = None):
        self.logger = self.setup_logger(name, level, log_file)
    
    def setup_logger(self, name: str, level: str = "INFO", log_file: Optional[Path] = None) -> logging.Logger:
        """ロガーをセットアップ"""
        logger = logging.getLogger(name)
        logger.setLevel(getattr(logging, level.upper()))
        
        # 既存のハンドラーをクリア
        logger.handlers.clear()
        
        # フォーマッターを設定
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        # コンソールハンドラーを追加
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
        
        # ファイルハンドラーを追加（指定されている場合）
        if log_file:
            log_file.parent.mkdir(parents=True, exist_ok=True)
            file_handler = logging.FileHandler(log_file, encoding='utf-8')
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
        
        return logger
    
    def log_file_operation(self, operation: str, file_path: Path, success: bool) -> None:
        """ファイル操作のログ出力"""
        status = "成功" if success else "失敗"
        message = f"ファイル操作 [{operation}] {status}: {file_path.name}"
        
        if success:
            self.logger.info(message)
        else:
            self.logger.error(message)
    
    def log_processing_progress(self, current: int, total: int, item: str) -> None:
        """処理進捗のログ出力"""
        percentage = (current / total) * 100 if total > 0 else 0
        message = f"処理進捗: {current}/{total} ({percentage:.1f}%) - {item}"
        self.logger.info(message)
    
    def log_error_with_context(self, error: Exception, context: Dict[str, Any]) -> None:
        """コンテキスト情報付きエラーログ"""
        context_str = ", ".join([f"{k}={v}" for k, v in context.items()])
        self.logger.error(f"エラー: {str(error)} | コンテキスト: {context_str}")
    
    def log_processing_summary(self, processed_count: int, success_count: int, error_count: int, 
                             duration_seconds: float) -> None:
        """処理結果サマリーのログ出力"""
        success_rate = (success_count / processed_count) * 100 if processed_count > 0 else 0
        
        self.logger.info("="*50)
        self.logger.info("処理結果サマリー")
        self.logger.info(f"処理対象数: {processed_count}")
        self.logger.info(f"成功数: {success_count}")
        self.logger.info(f"エラー数: {error_count}")
        self.logger.info(f"成功率: {success_rate:.1f}%")
        self.logger.info(f"処理時間: {duration_seconds:.2f}秒")
        self.logger.info("="*50)
    
    def log_configuration_info(self, config: Dict[str, Any]) -> None:
        """設定情報のログ出力"""
        self.logger.info("設定情報:")
        for key, value in config.items():
            # パスワードや秘密情報をマスク
            if any(secret in key.lower() for secret in ['password', 'secret', 'key', 'token']):
                value = '*' * len(str(value)) if value else 'None'
            self.logger.info(f"  {key}: {value}")
    
    def log_data_statistics(self, data_stats: Dict[str, Any]) -> None:
        """データ統計のログ出力"""
        self.logger.info("データ統計:")
        for key, value in data_stats.items():
            self.logger.info(f"  {key}: {value}")
    
    def log_performance_metrics(self, metrics: Dict[str, float]) -> None:
        """パフォーマンス指標のログ出力"""
        self.logger.info("パフォーマンス指標:")
        for metric, value in metrics.items():
            if 'time' in metric.lower() or 'duration' in metric.lower():
                self.logger.info(f"  {metric}: {value:.3f}秒")
            else:
                self.logger.info(f"  {metric}: {value}")
    
    def log_file_list(self, files: list, operation: str) -> None:
        """ファイルリストのログ出力"""
        self.logger.info(f"{operation}対象ファイル ({len(files)}件):")
        for i, file_path in enumerate(files, 1):
            if isinstance(file_path, Path):
                self.logger.info(f"  {i}. {file_path.name}")
            else:
                self.logger.info(f"  {i}. {file_path}")
    
    def log_platform_results(self, platform: str, results: Dict[str, Any]) -> None:
        """プラットフォーム別結果のログ出力"""
        self.logger.info(f"[{platform}] 処理結果:")
        for key, value in results.items():
            if isinstance(value, (int, float)):
                if 'amount' in key.lower() or 'fee' in key.lower() or '料' in key:
                    self.logger.info(f"  {key}: ¥{value:,.0f}")
                else:
                    self.logger.info(f"  {key}: {value}")
            else:
                self.logger.info(f"  {key}: {value}")
    
    # 既存のロガーメソッドのプロキシ
    def info(self, message: str) -> None:
        """情報レベルのログ出力"""
        self.logger.info(message)
    
    def warning(self, message: str) -> None:
        """警告レベルのログ出力"""
        self.logger.warning(message)
    
    def error(self, message: str) -> None:
        """エラーレベルのログ出力"""
        self.logger.error(message)
    
    def debug(self, message: str) -> None:
        """デバッグレベルのログ出力"""
        self.logger.debug(message)