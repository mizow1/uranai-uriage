"""
定数定義モジュール
設計書要件7.1: ハードコードされた値を定数として定義
"""

from typing import List

# メール処理関連定数
class MailConstants:
    """メール処理に関する定数"""
    DEFAULT_SENDER = "dev-line-fortune@linecorp.com"
    DEFAULT_RECIPIENT = "mizoguchi@outward.jp"
    DEFAULT_SUBJECT_PATTERN = "LineFortune Daily Report"
    DEFAULT_SEARCH_DAYS = 7
    DEFAULT_MAX_EMAILS_PER_FOLDER = 5000
    
    # メール検索優先フォルダ
    PRIORITY_FOLDERS: List[str] = [
        'LINE/LINE&,wiB6lLV,wk-',  # LINE（自動）のエンコード済み名前
        'LINE（自動）',
        'LINE (自動)',
        'LINE',
        'INBOX',
        '[Gmail]/All Mail',
        '[Gmail]/すべてのメール',
        'All Mail'
    ]
    
    # 日付抽出パターン
    DATE_PATTERNS: List[str] = [
        r'(\d{4})-(\d{2})-(\d{2})',  # yyyy-mm-dd
        r'(\d{4})(\d{2})(\d{2})'     # yyyymmdd
    ]


# ファイル処理関連定数
class FileConstants:
    """ファイル処理に関する定数"""
    DEFAULT_CSV_EXTENSION = ".csv"
    DEFAULT_RETRY_COUNT = 3
    DEFAULT_RETRY_DELAY = 5  # 秒
    DEFAULT_CLEANUP_DAYS = 30
    
    # ファイル名形式
    DATE_FORMAT = "%Y-%m-%d"
    YEAR_FORMAT = "%Y"
    MONTH_FORMAT = "%Y%m"
    
    # ディレクトリ名
    LINE_SUBDIR = "line"
    
    # 除外ファイルパターン
    EXCLUDE_PATTERNS: List[str] = [
        'line-menu-',
        'line-contents-',
        'output.',
        'backup_'
    ]


# 統合処理関連定数
class ConsolidationConstants:
    """CSV統合処理に関する定数"""
    MONTHLY_FILENAME_FORMAT = "line-menu-{year:04d}-{month:02d}.csv"
    BACKUP_SUFFIX = ".backup_{timestamp}"
    
    # 集計対象列
    NUMERIC_AGGREGATION = 'sum'
    NON_NUMERIC_AGGREGATION = 'first'
    
    # 除外ファイルパターン（統合処理用）
    CONSOLIDATION_EXCLUDE_PATTERNS: List[str] = [
        'line-menu-',
        'line-contents-',
        'output.',
        'consolidated_',
        'backup_'
    ]


# エラー処理関連定数
class ErrorConstants:
    """エラー処理に関する定数"""
    DEFAULT_MAX_RETRIES = 3
    DEFAULT_BASE_DELAY = 1.0  # 秒
    DEFAULT_MAX_DELAY = 60.0  # 秒
    DEFAULT_BACKOFF_FACTOR = 2.0
    
    # エラー分類キーワード
    NETWORK_ERROR_KEYWORDS: List[str] = [
        'connection', 'network', 'timeout', 'socket'
    ]
    
    AUTHENTICATION_ERROR_KEYWORDS: List[str] = [
        'authentication', 'login', 'password', 'credential'
    ]
    
    FILE_SYSTEM_ERROR_KEYWORDS: List[str] = [
        'permission', 'access', 'file', 'directory'
    ]
    
    PARSING_ERROR_KEYWORDS: List[str] = [
        'parse', 'decode', 'format'
    ]


# ログ処理関連定数
class LogConstants:
    """ログ処理に関する定数"""
    DEFAULT_LOG_FILE = "line_fortune_processor.log"
    DEFAULT_LOG_LEVEL = "INFO"
    DEFAULT_MAX_BYTES = 10 * 1024 * 1024  # 10MB
    DEFAULT_BACKUP_COUNT = 5
    
    # 機密情報マスキング対象
    SENSITIVE_KEYWORDS: List[str] = [
        'password', 'token', 'secret', 'key', 'credential'
    ]
    
    # ログフォーマット
    DEFAULT_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s'
    CONSOLE_FORMAT = '%(asctime)s - %(levelname)s - %(message)s'
    DATE_FORMAT = '%Y-%m-%d %H:%M:%S'


# 設定関連定数
class ConfigConstants:
    """設定に関する定数"""
    DEFAULT_CONFIG_FILE = "line_fortune_config.json"
    TEMPLATE_CONFIG_FILE = "line_fortune_config_template.json"
    
    # デフォルト設定値
    DEFAULT_EMAIL_SERVER = "imap.gmail.com"
    DEFAULT_EMAIL_PORT = 993
    DEFAULT_EMAIL_USE_SSL = True
    DEFAULT_BASE_PATH = r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書"
    
    # 必須設定項目
    REQUIRED_CONFIG_FIELDS: List[str] = [
        "email.username",
        "email.password", 
        "base_path",
        "sender",
        "recipient",
        "subject_pattern"
    ]
    
    # 数値設定項目
    NUMERIC_CONFIG_FIELDS: List[str] = [
        "search_days",
        "retry_count",
        "retry_delay", 
        "max_emails_per_folder",
        "email.port"
    ]


# アプリケーション全体の定数
class AppConstants:
    """アプリケーション全体に関する定数"""
    APP_NAME = "LINE Fortune Email Processor"
    VERSION = "1.0.0"
    
    # 処理モード
    MODE_NORMAL = "normal"
    MODE_DRY_RUN = "dry_run"
    MODE_CLEANUP = "cleanup"
    
    # 統計情報のキー
    STATS_EMAILS_PROCESSED = 'emails_processed'
    STATS_EMAILS_SUCCESS = 'emails_success'
    STATS_EMAILS_ERROR = 'emails_error'
    STATS_FILES_SAVED = 'files_saved'
    STATS_CONSOLIDATIONS_CREATED = 'consolidations_created'