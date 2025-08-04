"""
データモデル定義

システムで使用するデータクラスを定義します。
"""

from dataclasses import dataclass
from datetime import datetime
from typing import List, Optional


@dataclass
class SalesRecord:
    """売上レコードデータクラス"""
    platform: str
    content_name: str
    performance: float
    information_fee: float
    target_month: str
    template_file: str
    rate: float
    recipient_email: str
    sales_count: int = 0  # 売上件数（ameba、mediba、rakutenのみ）


@dataclass
class PaymentStatement:
    """支払い明細書データクラス"""
    content_name: str
    template_file: str
    sales_records: List[SalesRecord]
    total_performance: float
    total_information_fee: float
    payment_date: datetime
    recipient_email: str