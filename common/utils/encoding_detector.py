"""
エンコーディング検出ユーティリティ
"""
import chardet
from pathlib import Path
from typing import List, Optional
from ..error_handling.exceptions import EncodingDetectionError


class EncodingDetector:
    """ファイルのエンコーディングを検出するユーティリティクラス"""
    
    DEFAULT_ENCODINGS = ['utf-8', 'shift_jis', 'cp932', 'euc-jp', 'iso-2022-jp']
    
    def __init__(self, logger=None):
        self.logger = logger
    
    def detect_encoding(self, file_path: Path) -> str:
        """ファイルのエンコーディングを検出"""
        try:
            with open(file_path, 'rb') as f:
                raw_data = f.read()
                result = chardet.detect(raw_data)
                
                if result['encoding']:
                    detected_encoding = result['encoding'].lower()
                    if self.logger:
                        self.logger.info(f"エンコーディング検出: {file_path.name} -> {detected_encoding} (信頼度: {result['confidence']:.2f})")
                    return detected_encoding
                else:
                    if self.logger:
                        self.logger.warning(f"エンコーディング検出失敗: {file_path.name}")
                    return 'utf-8'  # デフォルト
                    
        except Exception as e:
            if self.logger:
                self.logger.error(f"エンコーディング検出エラー: {file_path.name} - {str(e)}")
            raise EncodingDetectionError(f"エンコーディング検出に失敗: {str(e)}")
    
    def try_encodings(self, file_path: Path, encodings: Optional[List[str]] = None) -> str:
        """複数のエンコーディングを順次試行して最初に成功したものを返す"""
        if encodings is None:
            encodings = self.DEFAULT_ENCODINGS
        
        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    f.read(1024)  # 少し読んでみてエラーが出ないかチェック
                
                if self.logger:
                    self.logger.info(f"エンコーディング試行成功: {file_path.name} -> {encoding}")
                return encoding
                
            except (UnicodeDecodeError, UnicodeError):
                continue
            except Exception as e:
                if self.logger:
                    self.logger.warning(f"エンコーディング試行エラー: {file_path.name}, {encoding} - {str(e)}")
                continue
        
        # 全て失敗した場合
        raise EncodingDetectionError(f"すべてのエンコーディングで読み込みに失敗: {file_path.name}")
    
    def validate_encoding(self, file_path: Path, encoding: str) -> bool:
        """指定されたエンコーディングでファイルが読み込み可能かチェック"""
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                f.read()
            return True
        except (UnicodeDecodeError, UnicodeError):
            return False