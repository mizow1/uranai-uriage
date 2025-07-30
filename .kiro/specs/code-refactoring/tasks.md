# Implementation Plan

- [ ] 1. Create foundation directory structure and base components
  - Create common/ directory with subdirectories for file_handlers, error_handling, logging, config, and utils
  - Implement base exception classes in common/error_handling/exceptions.py
  - Create FileProcessorBase abstract class in common/file_handlers/file_processor_base.py
  - _Requirements: 1.1, 4.1, 6.1_

- [ ] 2. Implement unified file handling system
- [ ] 2.1 Create CSVHandler class with encoding detection
  - Implement CSVHandler class in common/file_handlers/csv_handler.py
  - Add methods for encoding detection and CSV reading with fallback encodings
  - Include validation for CSV structure and data integrity
  - _Requirements: 1.1, 2.1, 6.2_

- [ ] 2.2 Create ExcelHandler class with password handling
  - Implement ExcelHandler class in common/file_handlers/excel_handler.py
  - Add support for multiple Excel engines and encrypted file handling
  - Include methods for trying different password combinations
  - _Requirements: 1.1, 2.1, 6.2_

- [ ] 2.3 Create EncodingDetector utility
  - Implement EncodingDetector class in common/utils/encoding_detector.py
  - Add methods for automatic encoding detection and validation
  - Include support for Japanese encodings (Shift-JIS, CP932, UTF-8)
  - _Requirements: 1.1, 6.2_

- [ ] 3. Implement unified error handling and logging system
- [ ] 3.1 Create centralized error handler
  - Implement ErrorHandler class in common/error_handling/error_handler.py
  - Add methods for different error handling strategies (log and continue, log and raise)
  - Include context-aware error reporting with file and operation details
  - _Requirements: 3.1, 3.2_

- [ ] 3.2 Create unified logging system
  - Implement UnifiedLogger class in common/logging/unified_logger.py
  - Replace all print statements with structured logging calls
  - Add methods for progress logging, error logging, and operation logging
  - _Requirements: 3.1, 3.2, 3.3_

- [ ] 4. Implement centralized configuration management
  - Create ConfigManager class in common/config/config_manager.py
  - Consolidate all configuration loading logic into single interface
  - Add validation methods for configuration completeness and correctness
  - _Requirements: 4.1, 4.2, 4.3_

- [ ] 5. Create standardized data models
  - Implement ProcessingResult and ContentDetail dataclasses in common/data_models.py
  - Add FileMetadata dataclass for consistent file information handling
  - Include validation methods and serialization support for data models
  - _Requirements: 6.1, 6.3_

- [ ] 6. Refactor platform-specific processors
- [ ] 6.1 Refactor AmebaProcessor using common components
  - Update sales_aggregator.py process_ameba_file method to use FileProcessorBase
  - Replace direct file reading with CSVHandler and ExcelHandler
  - Implement unified error handling and logging throughout the processor
  - _Requirements: 1.1, 2.1, 3.1, 6.1_

- [ ] 6.2 Refactor RakutenProcessor using common components
  - Update sales_aggregator.py process_rakuten_file method to use FileProcessorBase
  - Replace direct file reading with unified file handlers
  - Standardize error handling and logging patterns
  - _Requirements: 1.1, 2.1, 3.1, 6.1_

- [ ] 6.3 Refactor AuProcessor using common components
  - Update sales_aggregator.py process_au_file method to use FileProcessorBase
  - Replace encoding detection logic with EncodingDetector utility
  - Implement consistent error handling and logging
  - _Requirements: 1.1, 2.1, 3.1, 6.1_

- [ ] 6.4 Refactor ExciteProcessor using common components
  - Update sales_aggregator.py process_excite_file method to use FileProcessorBase
  - Replace manual CSV reading with CSVHandler
  - Standardize error handling and logging patterns
  - _Requirements: 1.1, 2.1, 3.1, 6.1_

- [ ] 6.5 Refactor LineProcessor using common components
  - Update sales_aggregator.py process_line_file method to use FileProcessorBase
  - Replace direct file reading with unified file handlers
  - Implement consistent error handling and logging
  - _Requirements: 1.1, 2.1, 3.1, 6.1_

- [ ] 7. Update main processing scripts
- [ ] 7.1 Update line_fortune_email_processor.py to use common components
  - Replace direct CSV reading with CSVHandler throughout the script
  - Update error handling to use ErrorHandler and UnifiedLogger
  - Consolidate configuration loading using ConfigManager
  - _Requirements: 1.1, 3.1, 4.1_

- [ ] 7.2 Update mediba_sales_processor.py to use common components
  - Replace manual CSV reading with CSVHandler
  - Update error handling and logging to use unified system
  - Standardize data processing patterns
  - _Requirements: 1.1, 3.1, 6.1_

- [ ] 7.3 Update content_payment_statement_generator components
  - Update sales_data_loader.py to use CSVHandler for all CSV operations
  - Replace scattered error handling with unified ErrorHandler
  - Consolidate logging using UnifiedLogger
  - _Requirements: 1.1, 3.1, 4.1_

- [ ] 8. Remove duplicate and unused code
- [ ] 8.1 Identify and remove duplicate functions
  - Scan codebase for duplicate function implementations
  - Remove redundant file processing functions after refactoring
  - Clean up unused imports and variables
  - _Requirements: 5.1, 5.2_

- [ ] 8.2 Remove dead code paths and unused utilities
  - Identify unreachable code paths in conditional statements
  - Remove unused utility functions and helper methods
  - Clean up commented-out code and debug statements
  - _Requirements: 5.1, 5.3_

- [ ] 9. Create comprehensive unit tests
- [ ] 9.1 Create tests for common components
  - Write unit tests for CSVHandler, ExcelHandler, and EncodingDetector
  - Create tests for ErrorHandler and UnifiedLogger functionality
  - Add tests for ConfigManager and data model validation
  - _Requirements: 1.1, 3.1, 4.1_

- [ ] 9.2 Create integration tests for refactored processors
  - Write integration tests for each platform processor
  - Create test data files for different scenarios (normal, encrypted, malformed)
  - Add tests for error handling and recovery scenarios
  - _Requirements: 1.1, 2.1, 3.1_

- [ ] 10. Update documentation and finalize refactoring
- [ ] 10.1 Update code documentation and comments
  - Add docstrings to all new classes and methods
  - Update existing comments to reflect refactored code structure
  - Create usage examples for common components
  - _Requirements: 7.1, 7.2_

- [ ] 10.2 Validate backward compatibility and performance
  - Run comprehensive tests with existing data files
  - Verify that all processing results match original outputs
  - Measure and validate performance improvements
  - _Requirements: 1.3, 2.2, 7.3_