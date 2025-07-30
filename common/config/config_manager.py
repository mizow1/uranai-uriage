"""
中央集約設定管理システム
"""
import json
from pathlib import Path
from typing import Dict, Any, Optional, List
from ..error_handling.exceptions import ConfigurationError


class ConfigManager:
    """設定管理の統一クラス"""
    
    DEFAULT_CONFIG_FILES = [
        'line_fortune_config.json',
        'config.json',
        'settings.json'
    ]
    
    def __init__(self, config_path: Optional[Path] = None, logger=None):
        self.logger = logger
        self.config_path = config_path
        self.config_data = {}
        self.load_config(config_path)
    
    def load_config(self, config_path: Optional[Path] = None) -> Dict[str, Any]:
        """設定ファイルを読み込み"""
        if config_path:
            self.config_data = self._load_single_config(config_path)
        else:
            # デフォルトの設定ファイルを順次試行
            for config_file in self.DEFAULT_CONFIG_FILES:
                try:
                    config_path = Path(config_file)
                    if config_path.exists():
                        self.config_data = self._load_single_config(config_path)
                        self.config_path = config_path
                        break
                except Exception as e:
                    if self.logger:
                        self.logger.debug(f"設定ファイル読み込み失敗: {config_file} - {str(e)}")
                    continue
            
            if not self.config_data:
                if self.logger:
                    self.logger.warning("設定ファイルが見つかりません。デフォルト設定を使用します。")
                self.config_data = self._get_default_config()
        
        return self.config_data
    
    def _load_single_config(self, config_path: Path) -> Dict[str, Any]:
        """単一の設定ファイルを読み込み"""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
            
            if self.logger:
                self.logger.info(f"設定ファイル読み込み成功: {config_path.name}")
            
            return config_data
            
        except FileNotFoundError:
            raise ConfigurationError(f"設定ファイルが見つかりません: {config_path}")
        except json.JSONDecodeError as e:
            raise ConfigurationError(f"設定ファイルの形式が無効です: {config_path} - {str(e)}")
        except Exception as e:
            raise ConfigurationError(f"設定ファイル読み込みエラー: {config_path} - {str(e)}")
    
    def _get_default_config(self) -> Dict[str, Any]:
        """デフォルト設定を取得"""
        return {
            'base_path': str(Path.cwd()),
            'encoding': 'utf-8',
            'excel_passwords': ['', 'password', '123456'],
            'log_level': 'INFO',
            'max_retries': 3,
            'timeout_seconds': 30
        }
    
    def get(self, key: str, default: Any = None) -> Any:
        """設定値を取得"""
        return self.config_data.get(key, default)
    
    def get_file_paths(self, year: str, month: str) -> Dict[str, Path]:
        """年月に基づいてファイルパスを構築"""
        base_path = Path(self.get('base_path', '.'))
        
        year_month_path = base_path / year / f"{year}{month.zfill(2)}"
        
        return {
            'base_path': base_path,
            'year_path': base_path / year,
            'month_path': year_month_path,
            'line_path': year_month_path / 'line',
            'ameba_path': year_month_path,
            'rakuten_path': year_month_path,
            'au_path': year_month_path,
            'excite_path': year_month_path
        }
    
    def get_processing_settings(self) -> Dict[str, Any]:
        """処理関連の設定を取得"""
        return {
            'encoding': self.get('encoding', 'utf-8'),
            'excel_passwords': self.get('excel_passwords', ['', 'password']),
            'max_retries': self.get('max_retries', 3),
            'timeout_seconds': self.get('timeout_seconds', 30),
            'chunk_size': self.get('chunk_size', 1000),
            'parallel_processing': self.get('parallel_processing', False)
        }
    
    def get_logging_settings(self) -> Dict[str, Any]:
        """ログ関連の設定を取得"""
        return {
            'log_level': self.get('log_level', 'INFO'),
            'log_file': self.get('log_file'),
            'log_format': self.get('log_format'),
            'log_rotation': self.get('log_rotation', True),
            'max_log_files': self.get('max_log_files', 5)
        }
    
    def get_email_settings(self) -> Dict[str, Any]:
        """メール関連の設定を取得"""
        return {
            'outlook_profile': self.get('outlook_profile'),
            'target_folder': self.get('target_folder', 'Inbox'),
            'sender_filter': self.get('sender_filter', []),
            'subject_filter': self.get('subject_filter', []),
            'max_emails': self.get('max_emails', 100)
        }
    
    def get_platform_settings(self, platform: str) -> Dict[str, Any]:
        """プラットフォーム固有の設定を取得"""
        platform_configs = self.get('platforms', {})
        return platform_configs.get(platform, {})
    
    def validate_configuration(self) -> bool:
        """設定の妥当性を検証"""
        required_fields = ['base_path']
        missing_fields = []
        
        for field in required_fields:
            if not self.get(field):
                missing_fields.append(field)
        
        if missing_fields:
            error_msg = f"必須設定項目が不足: {missing_fields}"
            if self.logger:
                self.logger.error(error_msg)
            raise ConfigurationError(error_msg)
        
        # base_pathの存在確認
        base_path = Path(self.get('base_path'))
        if not base_path.exists():
            error_msg = f"base_pathが存在しません: {base_path}"
            if self.logger:
                self.logger.error(error_msg)
            raise ConfigurationError(error_msg)
        
        if self.logger:
            self.logger.info("設定の妥当性検証完了")
        
        return True
    
    def save_config(self, config_path: Optional[Path] = None) -> None:
        """設定をファイルに保存"""
        if config_path is None:
            config_path = self.config_path or Path('config.json')
        
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config_data, f, ensure_ascii=False, indent=2)
            
            if self.logger:
                self.logger.info(f"設定ファイル保存完了: {config_path}")
                
        except Exception as e:
            error_msg = f"設定ファイル保存エラー: {config_path} - {str(e)}"
            if self.logger:
                self.logger.error(error_msg)
            raise ConfigurationError(error_msg)
    
    def update_config(self, updates: Dict[str, Any]) -> None:
        """設定を更新"""
        self.config_data.update(updates)
        
        if self.logger:
            self.logger.info(f"設定更新: {list(updates.keys())}")
    
    def get_all_settings(self) -> Dict[str, Any]:
        """すべての設定を取得"""
        return self.config_data.copy()