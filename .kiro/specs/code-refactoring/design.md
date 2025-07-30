# Design Document

## Overview

This design document outlines the refactoring approach for the current codebase to eliminate redundant processing, consolidate duplicate code, and improve overall maintainability. The analysis reveals significant code duplication across file processing, error handling, logging, and configuration management.

## Architecture

### Current Issues Identified

1. **File Processing Duplication**: Multiple similar file processing functions across different modules
2. **CSV Reading Patterns**: Repeated CSV reading logic with different encoding handling
3. **Error Handling Inconsistency**: Mix of print statements and proper logging for errors
4. **Configuration Scattered**: Multiple configuration approaches across modules
5. **Logging Inconsistency**: Mix of print statements and structured logging

### Proposed Architecture

```
common/
├── file_handlers/
│   ├── csv_handler.py          # Unified CSV reading with encoding detection
│   ├── excel_handler.py        # Unified Excel file processing
│   └── file_processor_base.py  # Base class for file processors
├── error_handling/
│   ├── exceptions.py           # Custom exception classes
│   └── error_handler.py        # Centralized error handling
├── logging/
│   └── unified_logger.py       # Standardized logging system
├── config/
│   └── config_manager.py       # Centralized configuration management
└── utils/
    ├── data_validator.py       # Common data validation functions
    └── encoding_detector.py    # Encoding detection utilities
```

## Components and Interfaces

### 1. Unified File Handler System

#### CSVHandler Class
```python
class CSVHandler:
    def read_csv_with_encoding_detection(self, file_path: Path, **kwargs) -> pd.DataFrame
    def try_multiple_encodings(self, file_path: Path, encodings: List[str]) -> pd.DataFrame
    def validate_csv_structure(self, df: pd.DataFrame, required_columns: int) -> bool
```

#### ExcelHandler Class
```python
class ExcelHandler:
    def read_excel_with_password_handling(self, file_path: Path, **kwargs) -> pd.DataFrame
    def try_multiple_engines(self, file_path: Path, engines: List[str]) -> pd.DataFrame
    def handle_encrypted_files(self, file_path: Path, passwords: List[str]) -> pd.DataFrame
```

#### FileProcessorBase Class
```python
class FileProcessorBase:
    def process_file(self, file_path: Path) -> Dict[str, Any]
    def validate_file_format(self, file_path: Path) -> bool
    def extract_metadata(self, file_path: Path) -> Dict[str, str]
```

### 2. Unified Error Handling System

#### ErrorHandler Class
```python
class ErrorHandler:
    def handle_file_processing_error(self, error: Exception, file_path: Path) -> None
    def handle_data_validation_error(self, error: Exception, data_context: str) -> None
    def log_and_continue(self, error: Exception, context: str) -> None
    def log_and_raise(self, error: Exception, context: str) -> None
```

#### Custom Exception Classes
```python
class FileProcessingError(Exception): pass
class DataValidationError(Exception): pass
class ConfigurationError(Exception): pass
class EncodingDetectionError(Exception): pass
```

### 3. Unified Logging System

#### UnifiedLogger Class
```python
class UnifiedLogger:
    def setup_logger(self, name: str, level: str = "INFO") -> logging.Logger
    def log_file_operation(self, operation: str, file_path: Path, success: bool) -> None
    def log_processing_progress(self, current: int, total: int, item: str) -> None
    def log_error_with_context(self, error: Exception, context: Dict[str, Any]) -> None
```

### 4. Centralized Configuration Management

#### ConfigManager Class
```python
class ConfigManager:
    def load_config(self, config_path: Optional[Path] = None) -> Dict[str, Any]
    def get_file_paths(self, year: str, month: str) -> Dict[str, Path]
    def get_processing_settings(self) -> Dict[str, Any]
    def validate_configuration(self) -> bool
```

### 5. Data Processing Standardization

#### DataProcessor Classes
```python
class AmebaProcessor(FileProcessorBase):
    def process_file(self, file_path: Path) -> Dict[str, Any]

class RakutenProcessor(FileProcessorBase):
    def process_file(self, file_path: Path) -> Dict[str, Any]

class AuProcessor(FileProcessorBase):
    def process_file(self, file_path: Path) -> Dict[str, Any]

class ExciteProcessor(FileProcessorBase):
    def process_file(self, file_path: Path) -> Dict[str, Any]

class LineProcessor(FileProcessorBase):
    def process_file(self, file_path: Path) -> Dict[str, Any]
```

## Data Models

### Standardized Data Models

```python
@dataclass
class ProcessingResult:
    platform: str
    file_name: str
    success: bool
    total_performance: float
    total_information_fee: float
    details: List[ContentDetail]
    errors: List[str]

@dataclass
class ContentDetail:
    content_group: str
    performance: float
    information_fee: float

@dataclass
class FileMetadata:
    file_path: Path
    file_size: int
    last_modified: datetime
    encoding: str
    format_type: str
```

## Error Handling

### Centralized Error Handling Strategy

1. **Replace print statements** with structured logging
2. **Implement consistent error recovery** patterns
3. **Standardize error reporting** across all modules
4. **Add proper exception chaining** for better debugging

### Error Categories

- **FileProcessingError**: File reading, parsing, format issues
- **DataValidationError**: Data integrity, missing columns, invalid values
- **ConfigurationError**: Missing config files, invalid settings
- **NetworkError**: Email processing, external service calls

## Testing Strategy

### Unit Testing Approach

1. **Test each unified component** independently
2. **Mock file system operations** for consistent testing
3. **Test error handling paths** explicitly
4. **Validate data transformation** accuracy

### Integration Testing

1. **Test end-to-end processing** with sample files
2. **Verify backward compatibility** with existing data
3. **Test configuration loading** from different sources
4. **Validate logging output** format and content

### Test Data Management

1. **Create sample files** for each platform type
2. **Include edge cases** (encrypted files, malformed data)
3. **Test encoding variations** (UTF-8, Shift-JIS, CP932)
4. **Validate error scenarios** (missing files, permission issues)

## Migration Strategy

### Phase 1: Foundation Components
- Create common utilities (file handlers, error handling, logging)
- Implement base classes and interfaces
- Set up unified configuration management

### Phase 2: Processor Refactoring
- Refactor each platform processor to use common components
- Standardize data models and return formats
- Implement consistent error handling

### Phase 3: Integration and Testing
- Update main processing scripts to use refactored components
- Comprehensive testing with existing data
- Performance validation and optimization

### Phase 4: Cleanup
- Remove duplicate code and unused functions
- Update documentation and comments
- Final validation and deployment