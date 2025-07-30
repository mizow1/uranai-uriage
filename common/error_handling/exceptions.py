"""
統一例外クラス定義
"""


class FileProcessingError(Exception):
    """ファイル処理関連のエラー"""
    pass


class DataValidationError(Exception):
    """データ検証関連のエラー"""
    pass


class ConfigurationError(Exception):
    """設定関連のエラー"""
    pass


class EncodingDetectionError(Exception):
    """エンコーディング検出関連のエラー"""
    pass


class NetworkError(Exception):
    """ネットワーク関連のエラー"""
    pass