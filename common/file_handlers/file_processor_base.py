"""
ファイル処理の基底クラス
"""
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Dict, Any


class FileProcessorBase(ABC):
    """ファイル処理の基底クラス"""
    
    def __init__(self, logger=None, error_handler=None):
        self.logger = logger
        self.error_handler = error_handler
    
    @abstractmethod
    def process_file(self, file_path: Path) -> Dict[str, Any]:
        """ファイルを処理する（サブクラスで実装）"""
        pass
    
    def validate_file_format(self, file_path: Path) -> bool:
        """ファイル形式を検証する"""
        return file_path.exists() and file_path.is_file()
    
    def extract_metadata(self, file_path: Path) -> Dict[str, str]:
        """ファイルのメタデータを抽出する"""
        if not file_path.exists():
            return {}
        
        return {
            'file_name': file_path.name,
            'file_size': str(file_path.stat().st_size),
            'last_modified': str(file_path.stat().st_mtime),
            'file_extension': file_path.suffix
        }