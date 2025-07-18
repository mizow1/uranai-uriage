"""
ファイル処理モジュール
"""

import os
import shutil
import logging
from typing import Optional, Dict, Any
from datetime import date
from pathlib import Path
import time


class FileProcessor:
    """ファイル処理クラス"""
    
    def __init__(self, base_path: str, retry_count: int = 3, retry_delay: int = 5):
        """
        ファイル処理を初期化
        
        Args:
            base_path: ファイル保存のベースパス
            retry_count: 再試行回数
            retry_delay: 再試行間隔（秒）
        """
        self.base_path = Path(base_path)
        self.retry_count = retry_count
        self.retry_delay = retry_delay
        self.logger = logging.getLogger(__name__)
        
    def create_directory_structure(self, target_date: date) -> Optional[Path]:
        """
        年/月のディレクトリ構造を作成
        
        Args:
            target_date: 対象日付
            
        Returns:
            Path: 作成されたディレクトリパス、失敗した場合はNone
        """
        try:
            year_str = target_date.strftime("%Y")
            month_str = target_date.strftime("%Y%m")
            
            # パス構築: base_path/yyyy/yyyymm/
            target_dir = self.base_path / year_str / month_str
            
            # ディレクトリ作成（存在しない場合のみ）
            target_dir.mkdir(parents=True, exist_ok=True)
            
            self.logger.info(f"ディレクトリ構造を作成しました: {target_dir}")
            return target_dir
            
        except Exception as e:
            self.logger.error(f"ディレクトリ作成中にエラーが発生しました: {e}")
            return None
    
    def rename_file(self, original_filename: str, target_date: date) -> str:
        """
        日付を含めるようにファイルをリネーム
        
        Args:
            original_filename: 元のファイル名
            target_date: 対象日付
            
        Returns:
            str: リネーム後のファイル名
        """
        try:
            # ファイル名と拡張子を分離
            name, ext = os.path.splitext(original_filename)
            
            # 日付を含む新しいファイル名を生成
            date_str = target_date.strftime("%Y-%m-%d")
            new_filename = f"{name}_{date_str}{ext}"
            
            self.logger.info(f"ファイル名を変更しました: {original_filename} -> {new_filename}")
            return new_filename
            
        except Exception as e:
            self.logger.error(f"ファイル名変更中にエラーが発生しました: {e}")
            return original_filename
    
    def save_file(self, file_content: bytes, filename: str, directory: Path) -> bool:
        """
        指定されたディレクトリにファイルを保存
        
        Args:
            file_content: ファイルの内容
            filename: ファイル名
            directory: 保存先ディレクトリ
            
        Returns:
            bool: 保存が成功した場合True
        """
        if not isinstance(directory, Path):
            directory = Path(directory)
            
        file_path = directory / filename
        
        for attempt in range(self.retry_count):
            try:
                # ディレクトリの存在確認
                if not directory.exists():
                    directory.mkdir(parents=True, exist_ok=True)
                
                # ファイルを保存
                with open(file_path, 'wb') as f:
                    f.write(file_content)
                
                self.logger.info(f"ファイルを保存しました: {file_path}")
                return True
                
            except Exception as e:
                self.logger.warning(f"ファイル保存の試行 {attempt + 1} 回目が失敗しました: {e}")
                
                if attempt < self.retry_count - 1:
                    self.logger.info(f"{self.retry_delay}秒後に再試行します")
                    time.sleep(self.retry_delay)
                else:
                    self.logger.error(f"ファイル保存に失敗しました: {file_path}")
                    
        return False
    
    def file_exists(self, filename: str, directory: Path) -> bool:
        """
        ファイルの存在確認
        
        Args:
            filename: ファイル名
            directory: 検索するディレクトリ
            
        Returns:
            bool: ファイルが存在する場合True
        """
        try:
            if not isinstance(directory, Path):
                directory = Path(directory)
                
            file_path = directory / filename
            return file_path.exists()
            
        except Exception as e:
            self.logger.error(f"ファイル存在確認中にエラーが発生しました: {e}")
            return False
    
    def get_file_size(self, filename: str, directory: Path) -> Optional[int]:
        """
        ファイルサイズを取得
        
        Args:
            filename: ファイル名
            directory: ファイルのディレクトリ
            
        Returns:
            int: ファイルサイズ（バイト）、取得失敗時はNone
        """
        try:
            if not isinstance(directory, Path):
                directory = Path(directory)
                
            file_path = directory / filename
            
            if file_path.exists():
                return file_path.stat().st_size
            else:
                return None
                
        except Exception as e:
            self.logger.error(f"ファイルサイズ取得中にエラーが発生しました: {e}")
            return None
    
    def backup_file(self, filename: str, directory: Path) -> bool:
        """
        ファイルのバックアップを作成
        
        Args:
            filename: ファイル名
            directory: ファイルのディレクトリ
            
        Returns:
            bool: バックアップが成功した場合True
        """
        try:
            if not isinstance(directory, Path):
                directory = Path(directory)
                
            source_path = directory / filename
            
            if not source_path.exists():
                self.logger.warning(f"バックアップ対象ファイルが存在しません: {source_path}")
                return False
            
            # バックアップファイル名を生成
            name, ext = os.path.splitext(filename)
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            backup_filename = f"{name}_backup_{timestamp}{ext}"
            backup_path = directory / backup_filename
            
            # ファイルをコピー
            shutil.copy2(source_path, backup_path)
            
            self.logger.info(f"ファイルをバックアップしました: {backup_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"ファイルバックアップ中にエラーが発生しました: {e}")
            return False
    
    def cleanup_old_files(self, directory: Path, days_to_keep: int = 30) -> int:
        """
        古いファイルを削除
        
        Args:
            directory: 対象ディレクトリ
            days_to_keep: 保持する日数
            
        Returns:
            int: 削除されたファイル数
        """
        try:
            if not isinstance(directory, Path):
                directory = Path(directory)
                
            if not directory.exists():
                return 0
                
            current_time = time.time()
            cutoff_time = current_time - (days_to_keep * 24 * 60 * 60)
            
            deleted_count = 0
            
            for file_path in directory.glob("*"):
                if file_path.is_file():
                    try:
                        file_mtime = file_path.stat().st_mtime
                        if file_mtime < cutoff_time:
                            file_path.unlink()
                            deleted_count += 1
                            self.logger.info(f"古いファイルを削除しました: {file_path}")
                    except Exception as e:
                        self.logger.warning(f"ファイル削除中にエラーが発生しました: {file_path}, {e}")
                        
            self.logger.info(f"古いファイルを {deleted_count} 個削除しました")
            return deleted_count
            
        except Exception as e:
            self.logger.error(f"古いファイル削除中にエラーが発生しました: {e}")
            return 0