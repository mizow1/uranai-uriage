"""
標準化されたデータモデル
"""
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any
from pathlib import Path
from datetime import datetime


@dataclass
class ProcessingResult:
    """処理結果の統一データモデル"""
    platform: str
    file_name: str
    success: bool
    total_performance: float = 0.0
    total_information_fee: float = 0.0
    details: List['ContentDetail'] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    processing_time: Optional[float] = None
    metadata: Dict[str, Any] = field(default_factory=dict)
    
    def add_error(self, error: str) -> None:
        """エラーを追加"""
        self.errors.append(error)
        self.success = False
    
    def add_detail(self, detail: 'ContentDetail') -> None:
        """詳細情報を追加"""
        self.details.append(detail)
    
    def calculate_totals(self) -> None:
        """詳細情報から合計を計算"""
        self.total_performance = sum(detail.performance for detail in self.details)
        self.total_information_fee = sum(detail.information_fee for detail in self.details)
    
    def to_dict(self) -> Dict[str, Any]:
        """辞書形式で出力"""
        return {
            'platform': self.platform,
            'file_name': self.file_name,
            'success': self.success,
            'total_performance': self.total_performance,
            'total_information_fee': self.total_information_fee,
            'details_count': len(self.details),
            'errors_count': len(self.errors),
            'processing_time': self.processing_time,
            'metadata': self.metadata
        }


@dataclass
class ContentDetail:
    """コンテンツ詳細の統一データモデル"""
    content_group: str
    performance: float = 0.0
    information_fee: float = 0.0
    sales_count: int = 0
    additional_data: Dict[str, Any] = field(default_factory=dict)
    
    def to_dict(self) -> Dict[str, Any]:
        """辞書形式で出力"""
        return {
            'content_group': self.content_group,
            'performance': self.performance,
            'information_fee': self.information_fee,
            'sales_count': self.sales_count,
            **self.additional_data
        }


@dataclass
class FileMetadata:
    """ファイルメタデータの統一データモデル"""
    file_path: Path
    file_size: int = 0
    last_modified: Optional[datetime] = None
    encoding: Optional[str] = None
    format_type: Optional[str] = None
    sheet_names: List[str] = field(default_factory=list)
    column_names: List[str] = field(default_factory=list)
    row_count: int = 0
    
    def __post_init__(self):
        """初期化後の処理"""
        if self.file_path.exists():
            stat = self.file_path.stat()
            if self.file_size == 0:
                self.file_size = stat.st_size
            if self.last_modified is None:
                self.last_modified = datetime.fromtimestamp(stat.st_mtime)
            if self.format_type is None:
                self.format_type = self.file_path.suffix.lower()
    
    def to_dict(self) -> Dict[str, Any]:
        """辞書形式で出力"""
        return {
            'file_path': str(self.file_path),
            'file_name': self.file_path.name,
            'file_size': self.file_size,
            'last_modified': self.last_modified.isoformat() if self.last_modified else None,
            'encoding': self.encoding,
            'format_type': self.format_type,
            'sheet_names': self.sheet_names,
            'column_names': self.column_names,
            'row_count': self.row_count
        }


@dataclass
class ProcessingSummary:
    """処理サマリーの統一データモデル"""
    total_files: int = 0
    successful_files: int = 0
    failed_files: int = 0
    total_performance: float = 0.0
    total_information_fee: float = 0.0
    processing_start: Optional[datetime] = None
    processing_end: Optional[datetime] = None
    platform_results: Dict[str, ProcessingResult] = field(default_factory=dict)
    errors: List[str] = field(default_factory=list)
    
    @property
    def success_rate(self) -> float:
        """成功率を計算"""
        if self.total_files == 0:
            return 0.0
        return (self.successful_files / self.total_files) * 100
    
    @property
    def processing_duration(self) -> Optional[float]:
        """処理時間を計算（秒）"""
        if self.processing_start and self.processing_end:
            return (self.processing_end - self.processing_start).total_seconds()
        return None
    
    def add_result(self, result: ProcessingResult) -> None:
        """処理結果を追加"""
        self.platform_results[result.platform] = result
        self.total_files += 1
        
        if result.success:
            self.successful_files += 1
            self.total_performance += result.total_performance
            self.total_information_fee += result.total_information_fee
        else:
            self.failed_files += 1
            self.errors.extend(result.errors)
    
    def to_dict(self) -> Dict[str, Any]:
        """辞書形式で出力"""
        return {
            'total_files': self.total_files,
            'successful_files': self.successful_files,
            'failed_files': self.failed_files,
            'success_rate': self.success_rate,
            'total_performance': self.total_performance,
            'total_information_fee': self.total_information_fee,
            'processing_duration': self.processing_duration,
            'processing_start': self.processing_start.isoformat() if self.processing_start else None,
            'processing_end': self.processing_end.isoformat() if self.processing_end else None,
            'platform_count': len(self.platform_results),
            'total_errors': len(self.errors)
        }


@dataclass
class EmailMetadata:
    """メールメタデータの統一データモデル"""
    subject: str
    sender: str
    received_time: datetime
    has_attachments: bool = False
    attachment_count: int = 0
    attachment_names: List[str] = field(default_factory=list)
    processed: bool = False
    processing_errors: List[str] = field(default_factory=list)
    
    def to_dict(self) -> Dict[str, Any]:
        """辞書形式で出力"""
        return {
            'subject': self.subject,
            'sender': self.sender,
            'received_time': self.received_time.isoformat(),
            'has_attachments': self.has_attachments,
            'attachment_count': self.attachment_count,
            'attachment_names': self.attachment_names,
            'processed': self.processed,
            'error_count': len(self.processing_errors)
        }