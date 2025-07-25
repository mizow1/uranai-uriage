"""
カスタム例外クラス定義

システムで使用するカスタム例外を定義します。
"""


class ContentPaymentStatementError(Exception):
    """コンテンツ関連支払い明細書システムの基本例外クラス"""
    pass


class FileNotFoundError(ContentPaymentStatementError):
    """ファイルが見つからない場合の例外"""
    pass


class DataValidationError(ContentPaymentStatementError):
    """データ検証エラーの例外"""
    pass


class MappingError(ContentPaymentStatementError):
    """コンテンツマッピングエラーの例外"""
    pass


class CalculationError(ContentPaymentStatementError):
    """計算処理エラーの例外"""
    pass


class ExcelProcessingError(ContentPaymentStatementError):
    """Excel処理エラーの例外"""
    pass


class PDFConversionError(ContentPaymentStatementError):
    """PDF変換エラーの例外"""
    pass


class EmailAuthenticationError(ContentPaymentStatementError):
    """メール認証エラーの例外"""
    pass


class EmailSendError(ContentPaymentStatementError):
    """メール送信エラーの例外"""
    pass


class SchedulingError(ContentPaymentStatementError):
    """予約送信エラーの例外"""
    pass


class ConfigurationError(ContentPaymentStatementError):
    """設定エラーの例外"""
    pass