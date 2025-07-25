# Design Document

## Overview

コンテンツ関連支払い明細書自動生成システムは、複数のデータソースから売上情報を収集し、Excelテンプレートを使用して支払い明細書を生成、PDF化してメール送信する統合システムです。

システムは以下の主要コンポーネントで構成されます：
- データ収集・処理モジュール
- Excelファイル操作モジュール
- PDF変換モジュール
- メール送信モジュール
- 設定管理モジュール

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                    Main Controller                          │
├─────────────────────────────────────────────────────────────┤
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────┐ │
│  │   Data Loader   │  │  Excel Processor │  │ PDF & Email │ │
│  │                 │  │                  │  │  Processor  │ │
│  │ - CSV Reader    │  │ - Template Copy  │  │ - PDF Conv. │ │
│  │ - Data Merger   │  │ - Data Writing   │  │ - Gmail API │ │
│  │ - Validation    │  │ - Formula Calc   │  │ - Scheduling│ │
│  └─────────────────┘  └─────────────────┘  └─────────────┘ │
├─────────────────────────────────────────────────────────────┤
│                   Configuration Manager                     │
│  - File Paths      - Rate Mapping      - Email Settings    │
└─────────────────────────────────────────────────────────────┘
```

## Components and Interfaces

### 1. Configuration Manager
**責任**: システム設定とファイルパスの管理

```python
class ConfigManager:
    def __init__(self):
        self.base_paths = {
            'sales_data': r'C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書',
            'template_dir': r'C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ロイヤリティ\コンテンツ関連支払明細書フォーマット',
            'output_base': r'C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書'
        }
    
    def get_monthly_sales_file(self) -> str
    def get_line_contents_file(self, year: str, month: str) -> str
    def get_template_files(self) -> List[str]
    def get_output_directory(self, year: str, month: str) -> str
```

### 2. Data Loader
**責任**: 各種CSVファイルからのデータ読み込みと統合

```python
class SalesDataLoader:
    def load_monthly_sales(self, target_month: str) -> pd.DataFrame
    def load_line_contents(self, year: str, month: str) -> pd.DataFrame
    def load_content_mapping(self) -> pd.DataFrame
    def load_rate_data(self) -> pd.DataFrame
    def merge_sales_data(self, monthly_data: pd.DataFrame, line_data: pd.DataFrame) -> pd.DataFrame
```

### 3. Excel Processor
**責任**: Excelテンプレートの複製と明細データの書き込み

```python
class ExcelProcessor:
    def copy_template(self, template_name: str, output_path: str, target_month: str) -> str
    def write_payment_date(self, workbook_path: str, target_month: str) -> None
    def write_statement_details(self, workbook_path: str, sales_data: List[Dict]) -> None
    def calculate_payment_amount(self, performance: float, rate: float) -> float
```

### 4. PDF Converter
**責任**: ExcelファイルのPDF変換

```python
class PDFConverter:
    def convert_excel_to_pdf(self, excel_path: str) -> str
    def validate_pdf_output(self, pdf_path: str) -> bool
```

### 5. Email Processor
**責任**: Gmail APIを使用したメール送信と予約

```python
class EmailProcessor:
    def __init__(self):
        self.gmail_service = self._setup_gmail_service()
    
    def send_payment_notification(self, recipient: str, pdf_path: str, target_month: str) -> bool
    def schedule_email(self, email_data: Dict, send_date: datetime) -> str
    def _setup_gmail_service(self) -> Any
```

## Data Models

### SalesRecord
```python
@dataclass
class SalesRecord:
    platform: str
    content_name: str
    performance: float
    information_fee: float
    target_month: str
    template_file: str
    rate: float
    recipient_email: str
```

### PaymentStatement
```python
@dataclass
class PaymentStatement:
    content_name: str
    template_file: str
    sales_records: List[SalesRecord]
    total_performance: float
    total_information_fee: float
    payment_date: datetime
    recipient_email: str
```

## Error Handling

### ファイル関連エラー
- **FileNotFoundError**: 必要なCSVファイルやテンプレートファイルが見つからない場合
- **PermissionError**: ファイルアクセス権限がない場合
- **データ形式エラー**: CSVファイルの形式が期待と異なる場合

### データ処理エラー
- **DataValidationError**: 売上データの値が不正な場合
- **MappingError**: コンテンツマッピングで対応するテンプレートが見つからない場合
- **CalculationError**: 料率計算でエラーが発生した場合

### Excel/PDF処理エラー
- **ExcelProcessingError**: Excelファイルの読み書きでエラーが発生した場合
- **PDFConversionError**: PDF変換に失敗した場合

### メール送信エラー
- **EmailAuthenticationError**: Gmail認証に失敗した場合
- **EmailSendError**: メール送信に失敗した場合
- **SchedulingError**: 予約送信の設定に失敗した場合

## Testing Strategy

### Unit Tests
- 各コンポーネントの個別機能テスト
- データ変換ロジックのテスト
- エラーハンドリングのテスト

### Integration Tests
- データフロー全体のテスト
- ファイル操作の統合テスト
- メール送信機能のテスト（テスト環境）

### End-to-End Tests
- 実際のデータファイルを使用した完全なワークフローテスト
- 異常系シナリオのテスト

### Test Data Management
- テスト用のサンプルCSVファイル作成
- モックExcelテンプレートの準備
- Gmail APIのテスト環境設定

## Dependencies

### Python Libraries
- `pandas`: CSVデータ処理
- `openpyxl`: Excel ファイル操作
- `xlwings`: Excel自動化（PDF変換用）
- `google-api-python-client`: Gmail API
- `google-auth`: Google認証
- `pathlib`: ファイルパス操作
- `datetime`: 日付処理
- `logging`: ログ出力

### External Dependencies
- Microsoft Excel: PDF変換に必要
- Gmail API: メール送信機能
- Google Cloud Console: API認証設定

## Security Considerations

- Gmail API認証情報の安全な管理
- ファイルアクセス権限の適切な設定
- メールアドレスの検証
- 機密データの適切な処理

## Performance Considerations

- 大量のコンテンツデータの効率的な処理
- Excel操作の最適化
- メール送信の並列処理
- ファイルI/Oの最適化