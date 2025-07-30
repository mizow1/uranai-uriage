"""
パフォーマンス最適化ユーティリティ
"""
import functools
import time
import threading
from typing import Any, Callable, Dict, List, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
import pandas as pd


class PerformanceOptimizer:
    """パフォーマンス最適化のユーティリティクラス"""
    
    def __init__(self, logger=None):
        self.logger = logger
        self._cache = {}
        self._cache_lock = threading.Lock()
    
    def measure_performance(self, func: Callable) -> Callable:
        """関数のパフォーマンスを測定するデコレータ"""
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            start_time = time.time()
            try:
                result = func(*args, **kwargs)
                return result
            finally:
                end_time = time.time()
                duration = end_time - start_time
                if self.logger:
                    self.logger.info(f"Performance: {func.__name__} took {duration:.3f}秒")
                else:
                    print(f"Performance: {func.__name__} took {duration:.3f}秒")
        return wrapper
    
    def cache_result(self, cache_key: str = None, ttl_seconds: int = 300):
        """結果をキャッシュするデコレータ"""
        def decorator(func: Callable) -> Callable:
            @functools.wraps(func)
            def wrapper(*args, **kwargs):
                # キャッシュキーを生成
                if cache_key:
                    key = cache_key
                else:
                    key = f"{func.__name__}_{hash(str(args) + str(sorted(kwargs.items())))}"
                
                with self._cache_lock:
                    # キャッシュから取得を試行
                    if key in self._cache:
                        cached_time, cached_result = self._cache[key]
                        if time.time() - cached_time < ttl_seconds:
                            if self.logger:
                                self.logger.debug(f"Cache hit for {func.__name__}")
                            return cached_result
                
                # キャッシュにない場合は実行
                result = func(*args, **kwargs)
                
                with self._cache_lock:
                    self._cache[key] = (time.time(), result)
                
                return result
            return wrapper
        return decorator
    
    def parallel_process_files(self, files: List[Path], process_func: Callable, 
                             max_workers: int = 4) -> List[Any]:
        """ファイルを並列処理する"""
        results = []
        
        if self.logger:
            self.logger.info(f"並列処理開始: {len(files)}ファイル, 最大{max_workers}スレッド")
        
        start_time = time.time()
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # 全タスクを投入
            future_to_file = {
                executor.submit(process_func, file_path): file_path 
                for file_path in files
            }
            
            # 完了順に結果を取得
            completed_count = 0
            for future in as_completed(future_to_file):
                file_path = future_to_file[future]
                completed_count += 1
                
                try:
                    result = future.result()
                    results.append(result)
                    
                    if self.logger:
                        self.logger.info(f"並列処理進捗: {completed_count}/{len(files)} - {file_path.name}")
                        
                except Exception as e:
                    if self.logger:
                        self.logger.error(f"並列処理エラー: {file_path.name} - {str(e)}")
                    results.append(None)
        
        end_time = time.time()
        duration = end_time - start_time
        
        if self.logger:
            self.logger.info(f"並列処理完了: {duration:.3f}秒, 平均{duration/len(files):.3f}秒/ファイル")
        
        return results
    
    def optimize_dataframe_operations(self, df: pd.DataFrame) -> pd.DataFrame:
        """DataFrameの操作を最適化"""
        if df.empty:
            return df
        
        # データタイプの最適化
        optimized_df = df.copy()
        
        # 数値列の最適化
        for col in optimized_df.select_dtypes(include=['int64']).columns:
            col_min = optimized_df[col].min()
            col_max = optimized_df[col].max()
            
            if col_min >= 0:
                if col_max < 255:
                    optimized_df[col] = optimized_df[col].astype('uint8')
                elif col_max < 65535:
                    optimized_df[col] = optimized_df[col].astype('uint16')
                elif col_max < 4294967295:
                    optimized_df[col] = optimized_df[col].astype('uint32')
            else:
                if col_min > -128 and col_max < 127:
                    optimized_df[col] = optimized_df[col].astype('int8')
                elif col_min > -32768 and col_max < 32767:
                    optimized_df[col] = optimized_df[col].astype('int16')
                elif col_min > -2147483648 and col_max < 2147483647:
                    optimized_df[col] = optimized_df[col].astype('int32')
        
        # float列の最適化
        for col in optimized_df.select_dtypes(include=['float64']).columns:
            optimized_df[col] = pd.to_numeric(optimized_df[col], downcast='float')
        
        # object列をcategoryに変換（重複値が多い場合）
        for col in optimized_df.select_dtypes(include=['object']).columns:
            if optimized_df[col].nunique() / len(optimized_df) < 0.5:
                optimized_df[col] = optimized_df[col].astype('category')
        
        # メモリ使用量の比較ログ
        if self.logger:
            original_memory = df.memory_usage(deep=True).sum() / 1024 / 1024
            optimized_memory = optimized_df.memory_usage(deep=True).sum() / 1024 / 1024
            reduction = (1 - optimized_memory / original_memory) * 100
            
            self.logger.info(f"DataFrame最適化: {original_memory:.2f}MB → {optimized_memory:.2f}MB "
                           f"({reduction:.1f}%削減)")
        
        return optimized_df
    
    def batch_process_data(self, data: List[Any], process_func: Callable, 
                          batch_size: int = 1000) -> List[Any]:
        """大量データをバッチ処理"""
        results = []
        total_batches = (len(data) + batch_size - 1) // batch_size
        
        if self.logger:
            self.logger.info(f"バッチ処理開始: {len(data)}件, バッチサイズ{batch_size}, {total_batches}バッチ")
        
        for i in range(0, len(data), batch_size):
            batch = data[i:i + batch_size]
            batch_num = (i // batch_size) + 1
            
            start_time = time.time()
            batch_results = process_func(batch)
            end_time = time.time()
            
            results.extend(batch_results)
            
            if self.logger:
                duration = end_time - start_time
                self.logger.info(f"バッチ処理進捗: {batch_num}/{total_batches} "
                               f"({len(batch)}件, {duration:.3f}秒)")
        
        return results
    
    def get_cache_stats(self) -> Dict[str, Any]:
        """キャッシュ統計を取得"""
        with self._cache_lock:
            return {
                'cache_size': len(self._cache),
                'cache_keys': list(self._cache.keys()),
                'memory_usage_estimate': sum(
                    len(str(key)) + len(str(value))
                    for key, (_, value) in self._cache.items()
                )
            }
    
    def clear_cache(self) -> None:
        """キャッシュをクリア"""
        with self._cache_lock:
            cleared_count = len(self._cache)
            self._cache.clear()
            
        if self.logger:
            self.logger.info(f"キャッシュクリア: {cleared_count}エントリを削除")
    
    def profile_memory_usage(self, func: Callable) -> Callable:
        """メモリ使用量をプロファイルするデコレータ"""
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            try:
                import psutil
                import os
                
                process = psutil.Process(os.getpid())
                memory_before = process.memory_info().rss / 1024 / 1024
                
                result = func(*args, **kwargs)
                
                memory_after = process.memory_info().rss / 1024 / 1024
                memory_diff = memory_after - memory_before
                
                if self.logger:
                    self.logger.info(f"Memory: {func.__name__} used {memory_diff:.2f}MB "
                                   f"(before: {memory_before:.2f}MB, after: {memory_after:.2f}MB)")
                
                return result
                
            except ImportError:
                if self.logger:
                    self.logger.warning("psutilが利用できません。メモリプロファイルをスキップします。")
                return func(*args, **kwargs)
            
        return wrapper


# グローバルインスタンス
_global_optimizer = None

def get_performance_optimizer(logger=None) -> PerformanceOptimizer:
    """グローバルパフォーマンス最適化インスタンスを取得"""
    global _global_optimizer
    if _global_optimizer is None:
        _global_optimizer = PerformanceOptimizer(logger)
    return _global_optimizer