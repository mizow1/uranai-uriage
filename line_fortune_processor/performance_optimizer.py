"""
パフォーマンス最適化モジュール
設計書要件6: パフォーマンスの最適化
"""

import asyncio
import concurrent.futures
import threading
import time
from typing import List, Dict, Any, Callable, Optional
from pathlib import Path
import logging
from functools import wraps

from .constants import AppConstants
from .messages import MessageFormatter


class PerformanceMonitor:
    """パフォーマンス監視クラス"""
    
    def __init__(self):
        self.logger = logging.getLogger(__name__)
        self.metrics = {
            'processing_times': [],
            'memory_usage': [],
            'file_sizes': [],
            'concurrent_tasks': 0,
            'total_operations': 0
        }
        self._lock = threading.Lock()
    
    def record_processing_time(self, operation: str, duration: float):
        """処理時間を記録"""
        with self._lock:
            self.metrics['processing_times'].append({
                'operation': operation,
                'duration': duration,
                'timestamp': time.time()
            })
    
    def record_file_size(self, filename: str, size: int):
        """ファイルサイズを記録"""
        with self._lock:
            self.metrics['file_sizes'].append({
                'filename': filename,
                'size': size,
                'timestamp': time.time()
            })
    
    def increment_concurrent_tasks(self):
        """同時実行タスク数を増加"""
        with self._lock:
            self.metrics['concurrent_tasks'] += 1
    
    def decrement_concurrent_tasks(self):
        """同時実行タスク数を減少"""
        with self._lock:
            self.metrics['concurrent_tasks'] = max(0, self.metrics['concurrent_tasks'] - 1)
    
    def get_average_processing_time(self, operation: Optional[str] = None) -> float:
        """平均処理時間を取得"""
        with self._lock:
            times = self.metrics['processing_times']
            if operation:
                times = [t for t in times if t['operation'] == operation]
            
            if not times:
                return 0.0
            
            return sum(t['duration'] for t in times) / len(times)
    
    def get_total_file_size(self) -> int:
        """総ファイルサイズを取得"""
        with self._lock:
            return sum(f['size'] for f in self.metrics['file_sizes'])
    
    def log_performance_summary(self):
        """パフォーマンスサマリーをログ出力"""
        with self._lock:
            avg_time = self.get_average_processing_time()
            total_size = self.get_total_file_size()
            max_concurrent = max([len(self.metrics['processing_times'])], default=0)
            
            self.logger.info(f"パフォーマンスサマリー:")
            self.logger.info(f"  平均処理時間: {avg_time:.2f}秒")
            self.logger.info(f"  総処理ファイルサイズ: {total_size / 1024 / 1024:.2f}MB")
            self.logger.info(f"  最大同時実行数: {max_concurrent}")
            self.logger.info(f"  総操作数: {len(self.metrics['processing_times'])}")


def performance_monitor(operation_name: str):
    """パフォーマンス監視デコレータ"""
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            monitor = getattr(args[0], 'performance_monitor', None) if args else None
            
            if monitor:
                monitor.increment_concurrent_tasks()
                
            start_time = time.time()
            try:
                result = func(*args, **kwargs)
                return result
            finally:
                duration = time.time() - start_time
                if monitor:
                    monitor.record_processing_time(operation_name, duration)
                    monitor.decrement_concurrent_tasks()
                    
        return wrapper
    return decorator


class ConcurrentProcessor:
    """並行処理管理クラス"""
    
    def __init__(self, max_workers: int = 4):
        """
        並行処理管理を初期化
        
        Args:
            max_workers: 最大ワーカー数
        """
        self.max_workers = max_workers
        self.logger = logging.getLogger(__name__)
        self.performance_monitor = PerformanceMonitor()
    
    def process_attachments_concurrently(
        self, 
        attachments: List[Dict[str, Any]], 
        processor_func: Callable, 
        *args, 
        **kwargs
    ) -> List[bool]:
        """
        複数の添付ファイルを並行処理
        
        Args:
            attachments: 添付ファイルのリスト
            processor_func: 処理関数
            *args, **kwargs: 処理関数に渡す引数
            
        Returns:
            List[bool]: 各ファイルの処理結果
        """
        if not attachments:
            return []
        
        self.logger.info(f"添付ファイルの並行処理を開始: {len(attachments)} 件")
        
        results = []
        
        # 添付ファイル数が少ない場合は並行処理しない
        if len(attachments) <= 2:
            for attachment in attachments:
                result = processor_func(attachment, *args, **kwargs)
                results.append(result)
        else:
            # ThreadPoolExecutorを使用した並行処理
            with concurrent.futures.ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                # 各添付ファイルに対して処理を並行実行
                future_to_attachment = {
                    executor.submit(processor_func, attachment, *args, **kwargs): attachment
                    for attachment in attachments
                }
                
                for future in concurrent.futures.as_completed(future_to_attachment):
                    attachment = future_to_attachment[future]
                    try:
                        result = future.result(timeout=300)  # 5分のタイムアウト
                        results.append(result)
                        
                        # ファイルサイズを記録
                        if result and 'content' in attachment:
                            self.performance_monitor.record_file_size(
                                attachment.get('filename', 'unknown'),
                                len(attachment['content'])
                            )
                            
                    except concurrent.futures.TimeoutError:
                        self.logger.error(f"添付ファイル処理がタイムアウトしました: {attachment.get('filename', 'unknown')}")
                        results.append(False)
                    except Exception as e:
                        self.logger.error(f"添付ファイル処理中にエラーが発生しました: {attachment.get('filename', 'unknown')}, {e}")
                        results.append(False)
        
        success_count = sum(1 for r in results if r)
        self.logger.info(f"添付ファイル並行処理完了: {success_count}/{len(attachments)} 件成功")
        
        return results
    
    def process_emails_concurrently(
        self,
        emails: List[Dict[str, Any]],
        processor_func: Callable,
        *args,
        **kwargs
    ) -> List[bool]:
        """
        複数のメールを並行処理
        
        Args:
            emails: メールのリスト
            processor_func: 処理関数
            *args, **kwargs: 処理関数に渡す引数
            
        Returns:
            List[bool]: 各メールの処理結果
        """
        if not emails:
            return []
        
        self.logger.info(f"メールの並行処理を開始: {len(emails)} 件")
        
        results = []
        
        # メール数が少ない場合は順次処理
        if len(emails) <= 3:
            for email in emails:
                result = processor_func(email, *args, **kwargs)
                results.append(result)
        else:
            # メール処理は順次実行が安全（IMAPの制約により）
            # ただし、添付ファイル処理は並行化可能
            for email in emails:
                result = processor_func(email, *args, **kwargs)
                results.append(result)
        
        success_count = sum(1 for r in results if r)
        self.logger.info(f"メール並行処理完了: {success_count}/{len(emails)} 件成功")
        
        return results


class MemoryOptimizer:
    """メモリ使用量最適化クラス"""
    
    def __init__(self, max_memory_mb: int = 500):
        """
        メモリ最適化を初期化
        
        Args:
            max_memory_mb: 最大メモリ使用量（MB）
        """
        self.max_memory_bytes = max_memory_mb * 1024 * 1024
        self.logger = logging.getLogger(__name__)
        self.current_memory_usage = 0
        self._lock = threading.Lock()
    
    def check_memory_usage(self, additional_bytes: int = 0) -> bool:
        """
        メモリ使用量をチェック
        
        Args:
            additional_bytes: 追加で使用予定のバイト数
            
        Returns:
            bool: メモリ使用量が制限内の場合True
        """
        try:
            import psutil
            process = psutil.Process()
            current_usage = process.memory_info().rss
            
            if current_usage + additional_bytes > self.max_memory_bytes:
                self.logger.warning(f"メモリ使用量が制限に近づいています: {current_usage / 1024 / 1024:.2f}MB")
                return False
            
            return True
            
        except ImportError:
            # psutilが利用できない場合は簡易チェック
            with self._lock:
                if self.current_memory_usage + additional_bytes > self.max_memory_bytes:
                    return False
                return True
    
    def register_memory_usage(self, bytes_used: int):
        """メモリ使用量を登録"""
        with self._lock:
            self.current_memory_usage += bytes_used
    
    def release_memory_usage(self, bytes_released: int):
        """メモリ使用量を解放"""
        with self._lock:
            self.current_memory_usage = max(0, self.current_memory_usage - bytes_released)
    
    def process_large_csv_streaming(self, file_path: Path, chunk_size: int = 1000):
        """
        大きなCSVファイルをストリーミング処理
        
        Args:
            file_path: CSVファイルパス
            chunk_size: チャンクサイズ（行数）
            
        Yields:
            pd.DataFrame: チャンクデータ
        """
        try:
            import pandas as pd
            
            self.logger.info(f"大きなCSVファイルのストリーミング処理開始: {file_path}")
            
            chunk_count = 0
            for chunk in pd.read_csv(file_path, chunksize=chunk_size, encoding='utf-8'):
                chunk_count += 1
                
                # メモリ使用量をチェック
                estimated_size = chunk.memory_usage(deep=True).sum()
                if not self.check_memory_usage(estimated_size):
                    self.logger.warning(f"メモリ制限により処理を一時停止: チャンク {chunk_count}")
                    time.sleep(1)  # メモリ解放を待つ
                
                self.register_memory_usage(estimated_size)
                
                try:
                    yield chunk
                finally:
                    self.release_memory_usage(estimated_size)
            
            self.logger.info(f"ストリーミング処理完了: {chunk_count} チャンク処理")
            
        except Exception as e:
            self.logger.error(f"ストリーミング処理中にエラーが発生しました: {e}")
            raise


# パフォーマンス監視の便利関数
def time_it(func):
    """実行時間を測定するデコレータ"""
    @wraps(func)
    def wrapper(*args, **kwargs):
        start_time = time.time()
        result = func(*args, **kwargs)
        duration = time.time() - start_time
        
        logger = logging.getLogger(__name__)
        logger.debug(f"{func.__name__} 実行時間: {duration:.2f}秒")
        
        return result
    return wrapper