"""
共通コンポーネントパッケージ
"""

from .file_handlers.csv_handler import CSVHandler
from .file_handlers.excel_handler import ExcelHandler
from .file_handlers.file_processor_base import FileProcessorBase
from .error_handling.exceptions import (
    FileProcessingError, 
    DataValidationError, 
    ConfigurationError, 
    EncodingDetectionError,
    NetworkError
)
from .error_handling.error_handler import ErrorHandler
from .logging.unified_logger import UnifiedLogger
from .config.config_manager import ConfigManager
from .utils.encoding_detector import EncodingDetector
from .data_models import (
    ProcessingResult,
    ContentDetail,
    FileMetadata,
    ProcessingSummary,
    EmailMetadata
)

__all__ = [
    'CSVHandler',
    'ExcelHandler',
    'FileProcessorBase',
    'FileProcessingError',
    'DataValidationError',
    'ConfigurationError',
    'EncodingDetectionError',
    'NetworkError',
    'ErrorHandler',
    'UnifiedLogger',
    'ConfigManager',
    'EncodingDetector',
    'ProcessingResult',
    'ContentDetail',
    'FileMetadata',
    'ProcessingSummary',
    'EmailMetadata'
]