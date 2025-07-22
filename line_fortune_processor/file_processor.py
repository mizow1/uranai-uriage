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

from .error_handler import retry_on_error, handle_errors, ErrorHandler, RetryableError, FatalError, ErrorType
from .constants import FileConstants
from .messages import MessageFormatter


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
        self.error_handler = ErrorHandler(self.logger)
        
    @retry_on_error(max_retries=3, base_delay=1.0)
    def create_directory_structure(self, target_date: date) -> Path:
        """
        年/月のディレクトリ構造を作成
        要件: base_path/yyyy/yyyymm/ 形式でディレクトリを作成
        
        Args:
            target_date: 対象日付
            
        Returns:
            Path: 作成されたディレクトリパス
            
        Raises:
            RetryableError: ディレクトリ作成に失敗した場合
        """
        year_str = target_date.strftime("%Y")
        month_str = target_date.strftime("%Y%m")
        
        # 要件に基づいたパス構築: base_path/yyyy/yyyymm/
        target_dir = self.base_path / year_str / month_str
        
        try:
            # ディレクトリ作成（存在しない場合のみ）
            target_dir.mkdir(parents=True, exist_ok=True)
            
            # 作成後のアクセス権限チェック
            if not target_dir.exists() or not os.access(target_dir, os.W_OK):
                raise RetryableError(f"ディレクトリの作成またはアクセスに失敗しました: {target_dir}", ErrorType.FILE_SYSTEM)
            
            self.logger.info(MessageFormatter.get_file_message(
            "directory_created", path=target_dir
        ))
            return target_dir
            
        except PermissionError as e:
            raise RetryableError(f"ディレクトリ作成の権限がありません: {target_dir}", ErrorType.FILE_SYSTEM, e)
        except OSError as e:
            raise RetryableError(f"ディレクトリ作成中にOSエラーが発生しました: {e}", ErrorType.FILE_SYSTEM, e)
        except Exception as e:
            error_type = self.error_handler.classify_error(e)
            raise RetryableError(f"ディレクトリ作成中にエラーが発生しました: {e}", error_type, e)
    
    @handle_errors("ファイル名変更")
    def rename_file(self, original_filename: str, target_date: date) -> str:
        """
        日付を含めるようにファイルをリネーム
        要件: ファイル名にyyyy-mm-dd形式の日付を含める
        
        Args:
            original_filename: 元のファイル名
            target_date: 対象日付
            
        Returns:
            str: リネーム後のファイル名
        """
        # ファイル名と拡張子を分離
        name, ext = os.path.splitext(original_filename)
        
        # 要件に基づく日付フォーマット：yyyy-mm-dd
        date_str = target_date.strftime(FileConstants.DATE_FORMAT)
        new_filename = f"{date_str}_{name}{ext}"
        
        self.logger.info(MessageFormatter.get_file_message(
            "file_renamed", old_name=original_filename, new_name=new_filename
        ))
        return new_filename
    
    @retry_on_error(max_retries=3, base_delay=2.0)
    def save_file(self, file_content: bytes, filename: str, directory: Path) -> bool:
        """
        指定されたディレクトリにファイルを保存
        要件: 対象ディレクトリが存在しない場合は作成する
        
        Args:
            file_content: ファイルの内容
            filename: ファイル名
            directory: 保存先ディレクトリ
            
        Returns:
            bool: 保存が成功した場合True
            
        Raises:
            RetryableError: ファイル保存に失敗した場合
        """
        if not isinstance(directory, Path):
            directory = Path(directory)
            
        file_path = directory / filename
        
        try:
            # 要件: ディレクトリが存在しない場合は作成
            if not directory.exists():
                directory.mkdir(parents=True, exist_ok=True)
                self.logger.info(f"保存先ディレクトリを作成しました: {directory}")
            
            # ファイルを保存
            with open(file_path, 'wb') as f:
                f.write(file_content)
            
            # 保存後の検証
            if not file_path.exists() or file_path.stat().st_size == 0:
                raise RetryableError(f"ファイルの保存に失敗しました: {file_path}", ErrorType.FILE_SYSTEM)
            
            self.logger.info(MessageFormatter.get_file_message(
                "file_saved", path=file_path, size=len(file_content)
            ))
            return True
            
        except PermissionError as e:
            raise RetryableError(f"ファイル保存の権限がありません: {file_path}", ErrorType.FILE_SYSTEM, e)
        except OSError as e:
            if "No space left" in str(e):
                raise FatalError(f"ディスク容量不足: {file_path}", ErrorType.FILE_SYSTEM, e)
            else:
                raise RetryableError(f"ファイル保存中にOSエラーが発生: {e}", ErrorType.FILE_SYSTEM, e)
        except Exception as e:
            error_type = self.error_handler.classify_error(e)
            raise RetryableError(f"ファイル保存中にエラーが発生: {e}", error_type, e)
    
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