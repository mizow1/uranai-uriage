"""
メッセージテンプレートモジュール
設計書要件7.2: ログメッセージとエラーメッセージをテンプレート化
"""

from typing import Dict, Any
from enum import Enum


class MessageCategory(Enum):
    """メッセージカテゴリ"""
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"
    SUCCESS = "success"


class MessageTemplates:
    """メッセージテンプレート定義クラス"""
    
    # メール処理メッセージ
    EMAIL_MESSAGES = {
        "connection_success": "メールサーバーに接続しました: {server}:{port}",
        "connection_failed": "メールサーバーへの接続に失敗しました: {error}",
        "search_start": "メール検索を開始します - 送信者: {sender}, 件名パターン: {pattern}",
        "emails_found": "条件に一致するメールを {count} 件発見しました",
        "emails_not_found": "処理対象のメールが見つかりませんでした",
        "email_processing_start": "メール処理開始: {subject}",
        "email_processing_complete": "メール処理完了: {subject} (添付ファイル: {files} 件)",
        "email_processing_error": "メール処理中にエラーが発生しました: {subject}",
        "attachment_extracted": "添付ファイルを {count} 件抽出しました",
        "attachment_not_found": "CSV添付ファイルが見つかりませんでした",
        "date_extraction_failed": "件名から日付を抽出できませんでした: {subject}",
        "date_extraction_success": "件名から日付を抽出しました: {date}",
    }
    
    # ファイル処理メッセージ
    FILE_MESSAGES = {
        "directory_created": "ディレクトリ構造を作成しました: {path}",
        "directory_creation_failed": "ディレクトリの作成に失敗しました: {path}",
        "file_saved": "ファイルを保存しました: {path} ({size} bytes)",
        "file_save_failed": "ファイル保存に失敗しました: {path}",
        "file_renamed": "ファイル名を変更しました: {old_name} -> {new_name}",
        "file_exists": "ファイルが既に存在します: {path}",
        "file_backup_created": "ファイルをバックアップしました: {backup_path}",
        "file_cleanup_start": "古いファイルの削除を開始します: {days} 日より古いファイル",
        "file_cleanup_complete": "古いファイルを {count} 個削除しました",
        "permission_denied": "ファイル操作の権限がありません: {path}",
        "disk_space_full": "ディスク容量が不足しています: {path}",
    }
    
    # CSV統合処理メッセージ
    CONSOLIDATION_MESSAGES = {
        "consolidation_start": "CSV統合処理を開始します: {directory}",
        "consolidation_complete": "CSV統合処理が完了しました: {output_file} ({files} ファイル統合)",
        "consolidation_failed": "CSV統合処理に失敗しました: {directory}",
        "csv_files_found": "統合対象CSVファイル: {count} 個",
        "csv_file_processed": "CSVファイルを処理しました: {file} ({rows} 行)",
        "csv_file_empty": "空のCSVファイルです: {file}",
        "csv_file_error": "CSVファイルの処理中にエラーが発生しました: {file}",
        "header_mismatch": "ヘッダーが一致しません: {file}",
        "monthly_file_generated": "月次統合ファイルを生成しました: {filename}",
        "backup_created": "既存ファイルをバックアップしました: {backup_path}",
    }
    
    # エラー処理メッセージ
    ERROR_MESSAGES = {
        "retry_attempt": "{operation}の試行 {attempt}/{max_attempts} 回目が失敗しました: {error}",
        "retry_success": "{operation}が {attempt} 回目の試行で成功しました",
        "retry_exhausted": "{operation}が最大試行回数 {max_attempts} 回に達しました",
        "fatal_error": "致命的エラーが発生しました: {error}",
        "network_error": "ネットワークエラーが発生しました: {error}",
        "authentication_error": "認証エラーが発生しました: {error}",
        "file_system_error": "ファイルシステムエラーが発生しました: {error}",
        "parsing_error": "データ解析エラーが発生しました: {error}",
        "unknown_error": "予期しないエラーが発生しました: {error}",
    }
    
    # 設定関連メッセージ
    CONFIG_MESSAGES = {
        "config_loaded": "設定ファイルを読み込みました: {file}",
        "config_not_found": "設定ファイルが見つかりません。デフォルト設定を使用します: {file}",
        "config_invalid": "設定ファイルのJSON形式が不正です: {error}",
        "config_validation_success": "設定の検証が完了しました",
        "config_validation_failed": "設定の検証に失敗しました",
        "config_field_missing": "必須設定項目が不足: {field}",
        "config_field_invalid": "無効な設定値: {field} = {value}",
        "config_saved": "設定を保存しました: {file}",
        "template_created": "設定テンプレートファイルを作成しました: {file}",
    }
    
    # セッション管理メッセージ
    SESSION_MESSAGES = {
        "session_start": "セッション開始",
        "session_end_success": "セッション終了: 成功",
        "session_end_failure": "セッション終了: 失敗",
        "processing_results": "処理結果 - 処理数: {processed}, 成功: {success}, エラー: {error}",
        "statistics": "統計情報 - メール: {emails}, ファイル: {files}, 統合: {consolidations}",
    }
    
    # 一般的なシステムメッセージ
    SYSTEM_MESSAGES = {
        "app_start": "{app_name} を開始します (バージョン: {version})",
        "app_end": "{app_name} を終了します",
        "dry_run_mode": "ドライランモードで実行します",
        "cleanup_mode": "クリーンアップモードで実行します",
        "interrupt_received": "ユーザーによって処理が中断されました",
        "unexpected_error": "予期しないエラーが発生しました: {error}",
        "operation_timeout": "操作がタイムアウトしました: {operation}",
        "resource_exhausted": "リソースが枯渇しました: {resource}",
    }


class MessageFormatter:
    """メッセージフォーマッタークラス"""
    
    @staticmethod
    def format_message(template: str, **kwargs) -> str:
        """
        テンプレートをフォーマットしてメッセージを生成
        
        Args:
            template: メッセージテンプレート
            **kwargs: テンプレート変数
            
        Returns:
            str: フォーマットされたメッセージ
        """
        try:
            return template.format(**kwargs)
        except KeyError as e:
            return f"メッセージテンプレートエラー: 変数 {e} が見つかりません - {template}"
        except Exception as e:
            return f"メッセージフォーマットエラー: {e} - {template}"
    
    @staticmethod
    def get_email_message(key: str, **kwargs) -> str:
        """メール関連メッセージを取得"""
        template = MessageTemplates.EMAIL_MESSAGES.get(key, f"Unknown email message: {key}")
        return MessageFormatter.format_message(template, **kwargs)
    
    @staticmethod
    def get_file_message(key: str, **kwargs) -> str:
        """ファイル関連メッセージを取得"""
        template = MessageTemplates.FILE_MESSAGES.get(key, f"Unknown file message: {key}")
        return MessageFormatter.format_message(template, **kwargs)
    
    @staticmethod
    def get_consolidation_message(key: str, **kwargs) -> str:
        """統合処理関連メッセージを取得"""
        template = MessageTemplates.CONSOLIDATION_MESSAGES.get(key, f"Unknown consolidation message: {key}")
        return MessageFormatter.format_message(template, **kwargs)
    
    @staticmethod
    def get_error_message(key: str, **kwargs) -> str:
        """エラー関連メッセージを取得"""
        template = MessageTemplates.ERROR_MESSAGES.get(key, f"Unknown error message: {key}")
        return MessageFormatter.format_message(template, **kwargs)
    
    @staticmethod
    def get_config_message(key: str, **kwargs) -> str:
        """設定関連メッセージを取得"""
        template = MessageTemplates.CONFIG_MESSAGES.get(key, f"Unknown config message: {key}")
        return MessageFormatter.format_message(template, **kwargs)
    
    @staticmethod
    def get_session_message(key: str, **kwargs) -> str:
        """セッション関連メッセージを取得"""
        template = MessageTemplates.SESSION_MESSAGES.get(key, f"Unknown session message: {key}")
        return MessageFormatter.format_message(template, **kwargs)
    
    @staticmethod
    def get_system_message(key: str, **kwargs) -> str:
        """システム関連メッセージを取得"""
        template = MessageTemplates.SYSTEM_MESSAGES.get(key, f"Unknown system message: {key}")
        return MessageFormatter.format_message(template, **kwargs)


# 便利関数
def get_message(category: str, key: str, **kwargs) -> str:
    """
    カテゴリとキーに基づいてメッセージを取得
    
    Args:
        category: メッセージカテゴリ
        key: メッセージキー
        **kwargs: テンプレート変数
        
    Returns:
        str: フォーマットされたメッセージ
    """
    formatters = {
        'email': MessageFormatter.get_email_message,
        'file': MessageFormatter.get_file_message,
        'consolidation': MessageFormatter.get_consolidation_message,
        'error': MessageFormatter.get_error_message,
        'config': MessageFormatter.get_config_message,
        'session': MessageFormatter.get_session_message,
        'system': MessageFormatter.get_system_message,
    }
    
    formatter = formatters.get(category.lower())
    if formatter:
        return formatter(key, **kwargs)
    else:
        return f"Unknown message category: {category}.{key}"