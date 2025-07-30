"""
設定管理モジュール - 統一ConfigManagerへの移行

このモジュールは非推奨です。代わりにcommon.ConfigManagerを使用してください。
"""

# 統一ConfigManagerをインポート
from common.config.config_manager import ConfigManager as UnifiedConfigManager
import warnings
import logging
from pathlib import Path
from typing import Dict, List, Optional

# 後方互換性のためのラッパークラス
class ConfigManager(UnifiedConfigManager):
    """content_payment_statement_generator用ConfigManagerラッパー（非推奨）"""
    
    def __init__(self):
        """統一ConfigManagerを初期化し、既存の設定を追加"""
        warnings.warn(
            "content_payment_statement_generator.config_manager.ConfigManagerは非推奨です。"
            "common.ConfigManagerを直接使用してください。",
            DeprecationWarning,
            stacklevel=2
        )
        
        # 統一ConfigManagerを初期化
        super().__init__()
        
        # 既存の設定を統一フォーマットで追加
        legacy_config = {
            'base_paths': {
                'sales_data': r'C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書',
                'template_dir': r'C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ロイヤリティ\コンテンツ関連支払明細書フォーマット',
                'output_base': r'C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ロイヤリティ',
                'current_dir': r'C:\Users\OW\Dropbox\disk2とローカルの同期\溝口\miz\uriage'
            },
            'files': {
                'monthly_sales': '月別ISP別コンテンツ別売上.csv',
                'contents_mapping': 'contents_mapping.csv',
                'rate_data': 'rate.csv'
            }
        }
        
        self.update_config(legacy_config)
        
        # 後方互換性のためのプロパティ
        self.base_paths = legacy_config['base_paths']
        self.files = legacy_config['files']
        
        self._validate_paths()
        
    def _validate_paths(self) -> None:
        """設定されたパスの存在確認"""
        for path_name, path_value in self.base_paths.items():
            if not Path(path_value).exists():
                logging.warning(f"設定されたパス '{path_name}' が存在しません: {path_value}")
    
    def get_monthly_sales_file(self) -> str:
        """月別売上ファイルのパスを取得"""
        return str(Path(self.base_paths['current_dir']) / self.files['monthly_sales'])
    
    def get_line_contents_file(self, year: str, month: str) -> str:
        """LINE用コンテンツファイルのパスを取得"""
        filename = f"line-contents-{year}-{month.zfill(2)}.csv"
        return str(Path(self.base_paths['current_dir']) / filename)
    
    def get_contents_mapping_file(self) -> str:
        """コンテンツマッピングファイルのパスを取得"""
        return str(Path(self.base_paths['current_dir']) / self.files['contents_mapping'])
    
    def get_rate_data_file(self) -> str:
        """レートデータファイルのパスを取得"""
        return str(Path(self.base_paths['current_dir']) / self.files['rate_data'])
    
    def get_template_files(self) -> List[str]:
        """テンプレートファイルのリストを取得"""
        template_dir = Path(self.base_paths['template_dir'])
        if not template_dir.exists():
            logging.error(f"テンプレートディレクトリが存在しません: {template_dir}")
            return []
        
        excel_files = list(template_dir.glob("*.xlsx")) + list(template_dir.glob("*.xls"))
        return [str(f) for f in excel_files]
    
    def get_output_directory(self, year: str, month: str) -> str:
        """出力ディレクトリのパスを取得"""
        output_dir = Path(self.base_paths['output_base']) / f"{year}{month.zfill(2)}"
        output_dir.mkdir(parents=True, exist_ok=True)
        return str(output_dir)
    
    def get_template_file_by_name(self, template_name: str) -> Optional[str]:
        """テンプレート名からファイルパスを取得"""
        template_dir = Path(self.base_paths['template_dir'])
        
        # 完全一致を試行
        exact_match = template_dir / template_name
        if exact_match.exists():
            return str(exact_match)
        
        # 拡張子を追加して試行
        for ext in ['.xlsx', '.xls']:
            with_ext = template_dir / f"{template_name}{ext}"
            if with_ext.exists():
                return str(with_ext)
        
        logging.warning(f"テンプレートファイルが見つかりません: {template_name}")
        return None
    
    def validate_required_files(self, year: str, month: str) -> Dict[str, bool]:
        """必要なファイルの存在確認"""
        validation_results = {}
        
        # 月別売上ファイル
        monthly_sales_path = self.get_monthly_sales_file()
        validation_results['monthly_sales'] = Path(monthly_sales_path).exists()
        
        # LINEコンテンツファイル（オプション）
        line_contents_path = self.get_line_contents_file(year, month)
        line_contents_exists = Path(line_contents_path).exists()
        validation_results['line_contents'] = line_contents_exists
        if not line_contents_exists:
            logging.warning(f"LINEコンテンツファイルが見つかりません（オプション）: {line_contents_path}")
        
        # マッピングファイル
        mapping_path = self.get_contents_mapping_file()
        validation_results['contents_mapping'] = Path(mapping_path).exists()
        
        # レートファイル
        rate_path = self.get_rate_data_file()
        validation_results['rate_data'] = Path(rate_path).exists()
        
        # テンプレートディレクトリ
        template_dir = Path(self.base_paths['template_dir'])
        validation_results['template_dir'] = template_dir.exists()
        
        return validation_results