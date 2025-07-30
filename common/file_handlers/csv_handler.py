"""
統一CSVハンドラー
"""
import pandas as pd
from pathlib import Path
from typing import List, Optional, Dict, Any
from ..utils.encoding_detector import EncodingDetector
from ..error_handling.exceptions import FileProcessingError, EncodingDetectionError


class CSVHandler:
    """CSVファイルの統一処理クラス"""
    
    def __init__(self, logger=None, error_handler=None):
        self.logger = logger
        self.error_handler = error_handler
        self.encoding_detector = EncodingDetector(logger)
    
    def read_csv_with_encoding_detection(self, file_path: Path, **kwargs) -> pd.DataFrame:
        """エンコーディング自動検出でCSVファイルを読み込み"""
        try:
            # まずエンコーディングを検出
            encoding = self.encoding_detector.detect_encoding(file_path)
            return self._read_csv_with_encoding(file_path, encoding, **kwargs)
            
        except (EncodingDetectionError, FileProcessingError):
            # 検出失敗時は複数エンコーディングを試行
            return self.try_multiple_encodings(file_path, **kwargs)
    
    def try_multiple_encodings(self, file_path: Path, encodings: Optional[List[str]] = None, **kwargs) -> pd.DataFrame:
        """複数のエンコーディングを順次試行してCSVを読み込み"""
        if encodings is None:
            encodings = ['utf-8', 'shift_jis', 'cp932', 'euc-jp']
        
        last_error = None
        
        for encoding in encodings:
            try:
                df = self._read_csv_with_encoding(file_path, encoding, **kwargs)
                if self.logger:
                    self.logger.info(f"CSV読み込み成功: {file_path.name} ({encoding})")
                return df
                
            except Exception as e:
                last_error = e
                if self.logger:
                    self.logger.debug(f"CSV読み込み失敗: {file_path.name} ({encoding}) - {str(e)}")
                continue
        
        # すべて失敗
        error_msg = f"すべてのエンコーディングでCSV読み込みに失敗: {file_path.name}"
        if self.logger:
            self.logger.error(error_msg)
        raise FileProcessingError(f"{error_msg} - 最後のエラー: {str(last_error)}")
    
    def _read_csv_with_encoding(self, file_path: Path, encoding: str, **kwargs) -> pd.DataFrame:
        """指定されたエンコーディングでCSVを読み込み"""
        try:
            df = pd.read_csv(file_path, encoding=encoding, **kwargs)
            return df
        except Exception as e:
            raise FileProcessingError(f"CSV読み込みエラー: {file_path.name} ({encoding}) - {str(e)}")
    
    def validate_csv_structure(self, df: pd.DataFrame, required_columns: Optional[int] = None, 
                             required_column_names: Optional[List[str]] = None) -> bool:
        """CSVの構造を検証"""
        try:
            # 最小列数チェック
            if required_columns and len(df.columns) < required_columns:
                if self.logger:
                    self.logger.error(f"列数不足: 必要{required_columns}列、実際{len(df.columns)}列")
                return False
            
            # 必須列名チェック
            if required_column_names:
                missing_columns = [col for col in required_column_names if col not in df.columns]
                if missing_columns:
                    if self.logger:
                        self.logger.error(f"必須列が不足: {missing_columns}")
                    return False
            
            # 空のデータフレームチェック
            if df.empty:
                if self.logger:
                    self.logger.warning("空のCSVファイル")
                return False
                
            return True
            
        except Exception as e:
            if self.logger:
                self.logger.error(f"CSV構造検証エラー: {str(e)}")
            return False
    
    def read_csv_safe(self, file_path: Path, **kwargs) -> Optional[pd.DataFrame]:
        """安全なCSV読み込み（エラー時はNoneを返す）"""
        try:
            return self.read_csv_with_encoding_detection(file_path, **kwargs)
        except Exception as e:
            if self.error_handler:
                self.error_handler.handle_file_processing_error(e, file_path)
            elif self.logger:
                self.logger.error(f"CSV読み込みエラー: {file_path.name} - {str(e)}")
            return None
    
    def get_file_info(self, file_path: Path) -> Dict[str, Any]:
        """CSVファイルの基本情報を取得"""
        try:
            df = self.read_csv_with_encoding_detection(file_path, nrows=0)  # ヘッダーのみ
            encoding = self.encoding_detector.detect_encoding(file_path)
            
            return {
                'file_name': file_path.name,
                'encoding': encoding,
                'columns': list(df.columns),
                'column_count': len(df.columns),
                'file_size': file_path.stat().st_size
            }
        except Exception as e:
            if self.logger:
                self.logger.error(f"CSVファイル情報取得エラー: {file_path.name} - {str(e)}")
            return {}