"""
設定管理モジュール
"""

import os
import json
import logging
from typing import Dict, Any, Optional, List
from pathlib import Path
from .error_handler import handle_errors, FatalError, ErrorType


class Config:
    """設定管理クラス"""
    
    def __init__(self, config_file: str = None):
        """
        設定を初期化
        
        Args:
            config_file: 設定ファイルのパス
        """
        self.config_file = config_file or "line_fortune_config.json"
        self.logger = logging.getLogger(__name__)
        self.config = self._load_config()
        self._validation_errors = []
        
    def _load_config(self) -> Dict[str, Any]:
        """設定ファイルから設定を読み込み"""
        default_config = {
            "email": {
                "server": "imap.gmail.com",
                "port": 993,
                "username": os.getenv("EMAIL_USERNAME", "mizoguchi@outward.jp"),
                "password": os.getenv("EMAIL_PASSWORD", ""),
                "use_ssl": True
            },
            "base_path": r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書",
            "log_file": "line_fortune_processor.log",
            "log_level": "INFO",
            "sender": "dev-line-fortune@linecorp.com",
            "recipient": "mizoguchi@outward.jp",
            "subject_pattern": "LineFortune Daily Report",
            "search_days": 3,
            "retry_count": 3,
            "retry_delay": 5
        }
        
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    file_config = json.load(f)
                    default_config.update(file_config)
            except Exception as e:
                logging.warning(f"設定ファイルの読み込みに失敗しました: {e}")
                
        return default_config
    
    def get(self, key: str, default=None):
        """設定値を取得"""
        keys = key.split('.')
        value = self.config
        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return default
        return value
    
    def set(self, key: str, value: Any):
        """設定値を設定"""
        keys = key.split('.')
        current = self.config
        
        # ネストした辞書を作成
        for k in keys[:-1]:
            if k not in current:
                current[k] = {}
            current = current[k]
        
        # 最後のキーに値を設定
        current[keys[-1]] = value
    
    def validate(self) -> bool:
        """設定の検証"""
        required_fields = [
            "email.username",
            "email.password",
            "base_path",
            "sender",
            "recipient",
            "subject_pattern"
        ]
        
        self._validation_errors = []
        
        for field in required_fields:
            if not self.get(field):
                error_msg = f"必須設定項目が不足しています: {field}"
                logging.error(error_msg)
                self._validation_errors.append(error_msg)
                
        return len(self._validation_errors) == 0
    
    def get_validation_errors(self) -> List[str]:
        """検証エラーのリストを取得"""
        return self._validation_errors.copy()
    
    def create_template(self, file_path: str = None):
        """設定ファイルテンプレートを作成"""
        template_path = file_path or "line_fortune_config_template.json"
        
        template = {
            "email": {
                "server": "imap.gmail.com",
                "port": 993,
                "username": "your_email@example.com",
                "password": "your_app_password",
                "use_ssl": True
            },
            "base_path": r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書",
            "log_file": "line_fortune_processor.log",
            "log_level": "INFO",
            "sender": "dev-line-fortune@linecorp.com",
            "recipient": "mizoguchi@outward.jp",
            "subject_pattern": "LineFortune Daily Report",
            "search_days": 3,
            "retry_count": 3,
            "retry_delay": 5
        }
        
        try:
            with open(template_path, 'w', encoding='utf-8') as f:
                json.dump(template, f, indent=2, ensure_ascii=False)
            logging.info(f"設定テンプレートファイルを作成しました: {template_path}")
            return True
        except Exception as e:
            logging.error(f"設定テンプレートファイルの作成に失敗しました: {e}")
            return False