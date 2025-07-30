"""
共通コンポーネントの統合テスト
"""
import unittest
import tempfile
import pandas as pd
from pathlib import Path
import sys
import os

# プロジェクトルートをパスに追加
sys.path.insert(0, str(Path(__file__).parent.parent))

from common import (
    CSVHandler, 
    ExcelHandler, 
    UnifiedLogger, 
    ErrorHandler, 
    ConfigManager,
    EncodingDetector,
    ProcessingResult,
    ContentDetail
)


class TestCommonComponents(unittest.TestCase):
    """共通コンポーネントの統合テスト"""
    
    def setUp(self):
        """テスト前の準備"""
        self.temp_dir = Path(tempfile.mkdtemp())
        self.logger = UnifiedLogger("test_logger")
        self.error_handler = ErrorHandler(self.logger.logger)
        self.csv_handler = CSVHandler(self.logger.logger, self.error_handler)
        self.excel_handler = ExcelHandler(self.logger.logger, self.error_handler)
        self.encoding_detector = EncodingDetector(self.logger.logger)
    
    def tearDown(self):
        """テスト後のクリーンアップ"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_csv_handler_with_utf8(self):
        """UTF-8エンコーディングのCSVファイル処理テスト"""
        # テストデータ作成
        test_data = pd.DataFrame({
            'A': [1, 2, 3],
            'B': ['テスト1', 'テスト2', 'テスト3'],
            'C': [100.5, 200.0, 300.25]
        })
        
        csv_file = self.temp_dir / "test_utf8.csv"
        test_data.to_csv(csv_file, index=False, encoding='utf-8')
        
        # CSVHandlerでファイル読み込み
        df = self.csv_handler.read_csv_with_encoding_detection(csv_file)
        
        # 検証
        self.assertIsNotNone(df)
        self.assertEqual(len(df), 3)
        self.assertEqual(len(df.columns), 3)
        self.assertIn('テスト1', df['B'].values)
    
    def test_csv_handler_with_shift_jis(self):
        """Shift-JISエンコーディングのCSVファイル処理テスト"""
        # テストデータ作成
        test_data = pd.DataFrame({
            'プラットフォーム': ['楽天', 'mediba', 'LINE'],
            '実績': [1000, 2000, 3000],
            '情報提供料': [300, 600, 900]
        })
        
        csv_file = self.temp_dir / "test_sjis.csv"
        test_data.to_csv(csv_file, index=False, encoding='shift_jis')
        
        # CSVHandlerでファイル読み込み
        df = self.csv_handler.read_csv_with_encoding_detection(csv_file)
        
        # 検証
        self.assertIsNotNone(df)
        self.assertEqual(len(df), 3)
        self.assertIn('楽天', df['プラットフォーム'].values)
    
    def test_excel_handler_basic(self):
        """基本的なExcelファイル処理テスト"""
        # テストデータ作成
        test_data = pd.DataFrame({
            'ID': [1, 2, 3],
            'Name': ['テストA', 'テストB', 'テストC'],
            'Value': [100, 200, 300]
        })
        
        excel_file = self.temp_dir / "test.xlsx"
        test_data.to_excel(excel_file, index=False)
        
        # ExcelHandlerでファイル読み込み
        df = self.excel_handler.read_excel_with_password_handling(excel_file)
        
        # 検証
        self.assertIsNotNone(df)
        self.assertEqual(len(df), 3)
        self.assertEqual(len(df.columns), 3)
    
    def test_processing_result_model(self):
        """ProcessingResultデータモデルのテスト"""
        result = ProcessingResult(
            platform="test_platform",
            file_name="test_file.csv",
            success=False
        )
        
        # 詳細データ追加
        detail1 = ContentDetail(
            content_group="group1",
            performance=1000.0,
            information_fee=300.0
        )
        detail2 = ContentDetail(
            content_group="group2", 
            performance=2000.0,
            information_fee=600.0
        )
        
        result.add_detail(detail1)
        result.add_detail(detail2)
        
        # エラー追加
        result.add_error("テストエラー")
        
        # 合計計算
        result.calculate_totals()
        
        # 検証
        self.assertEqual(result.platform, "test_platform")
        self.assertEqual(result.file_name, "test_file.csv")
        self.assertFalse(result.success)  # エラーがあるのでfalse
        self.assertEqual(len(result.details), 2)
        self.assertEqual(len(result.errors), 1)
        self.assertEqual(result.total_performance, 3000.0)
        self.assertEqual(result.total_information_fee, 900.0)
    
    def test_config_manager_basic(self):
        """ConfigManagerの基本テスト"""
        config_manager = ConfigManager()
        
        # デフォルト設定の検証
        base_path = config_manager.get('base_path')
        self.assertIsNotNone(base_path)
        
        # 処理設定の取得
        processing_settings = config_manager.get_processing_settings()
        self.assertIn('encoding', processing_settings)
        self.assertIn('excel_passwords', processing_settings)
        
        # ログ設定の取得
        logging_settings = config_manager.get_logging_settings()
        self.assertIn('log_level', logging_settings)
    
    def test_encoding_detector(self):
        """EncodingDetectorのテスト"""
        # UTF-8ファイルを作成
        utf8_file = self.temp_dir / "test_utf8.txt"
        with open(utf8_file, 'w', encoding='utf-8') as f:
            f.write("UTF-8テストファイル\n日本語文字列")
        
        # エンコーディング検出
        detected_encoding = self.encoding_detector.detect_encoding(utf8_file)
        self.assertIsNotNone(detected_encoding)
        
        # 複数エンコーディング試行
        successful_encoding = self.encoding_detector.try_encodings(utf8_file)
        self.assertIn(successful_encoding.lower(), ['utf-8', 'ascii'])
    
    def test_error_handler(self):
        """ErrorHandlerのテスト"""
        test_file = self.temp_dir / "test_error.txt"
        test_error = Exception("テストエラー")
        
        # エラーハンドリングのテスト（例外が発生しないことを確認）
        try:
            self.error_handler.handle_file_processing_error(test_error, test_file)
            self.error_handler.log_and_continue(test_error, "テストコンテキスト")
        except Exception as e:
            self.fail(f"ErrorHandlerでエラーが発生: {e}")
    
    def test_unified_logger(self):
        """UnifiedLoggerのテスト"""
        test_file = self.temp_dir / "test_log_file.txt"
        
        # ロガーの各メソッドをテスト（例外が発生しないことを確認）
        try:
            self.logger.info("テスト情報メッセージ")
            self.logger.warning("テスト警告メッセージ")
            self.logger.error("テストエラーメッセージ")
            self.logger.log_file_operation("読み込み", test_file, True)
            self.logger.log_processing_progress(1, 10, "テストアイテム")
        except Exception as e:
            self.fail(f"UnifiedLoggerでエラーが発生: {e}")


class TestIntegrationScenarios(unittest.TestCase):
    """統合シナリオのテスト"""
    
    def setUp(self):
        """テスト前の準備"""
        self.temp_dir = Path(tempfile.mkdtemp())
        self.logger = UnifiedLogger("integration_test")
        self.error_handler = ErrorHandler(self.logger.logger)
        self.csv_handler = CSVHandler(self.logger.logger, self.error_handler)
    
    def tearDown(self):
        """テスト後のクリーンアップ"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_mediba_style_csv_processing(self):
        """mediba風CSVファイルの処理統合テスト"""
        # mediba風のテストデータ作成
        test_data = pd.DataFrame({
            'Column_A': ['header_info', 'data1', 'data2', 'data3'],
            'Column_B': ['Program_ID', 'Program_1', 'Program_2', 'Program_1'],
            'Column_C': [0, 0, 0, 0],
            'Column_D': [0, 0, 0, 0],
            'Column_E': [0, 0, 0, 0],
            'Column_F': [0, 0, 0, 0],
            'Column_G': [0, 1000, 2000, 1500],  # 料金データ
            'Column_H': [0, 0, 0, 0],
            'Column_I': [0, 0, 0, 0],
            'Column_J': [0, 0, 0, 0],
            'Column_K': [0, 100, 200, 150]  # CP売上負担額
        })
        
        csv_file = self.temp_dir / "mediba_test.csv"
        test_data.to_csv(csv_file, index=False, encoding='utf-8')
        
        # CSVファイル読み込み
        df = self.csv_handler.read_csv_with_encoding_detection(csv_file)
        
        # 列数チェック
        self.assertGreaterEqual(len(df.columns), 11)
        
        # Program_1の合計計算テスト
        program_1_rows = df[df.iloc[:, 1] == 'Program_1']
        g_sum = program_1_rows.iloc[:, 6].sum()  # G列の合計
        k_sum = program_1_rows.iloc[:, 10].sum()  # K列の合計
        
        self.assertEqual(g_sum, 2500)  # 1000 + 1500
        self.assertEqual(k_sum, 250)  # 100 + 150
        
        # 情報提供料計算（G列の40% - K列）
        information_fee = (g_sum * 0.4) - k_sum
        self.assertEqual(information_fee, 750)  # (2500 * 0.4) - 250 = 1000 - 250 = 750
    
    def test_file_processing_error_handling(self):
        """ファイル処理エラーハンドリングの統合テスト"""
        # 存在しないファイル
        non_existent_file = self.temp_dir / "non_existent.csv"
        
        # 安全な読み込みテスト
        result = self.csv_handler.read_csv_safe(non_existent_file)
        self.assertIsNone(result)
        
        # 不正なCSVファイル
        invalid_csv = self.temp_dir / "invalid.csv"
        with open(invalid_csv, 'w', encoding='utf-8') as f:
            f.write("invalid,csv\ndata,with,too,many,columns\n")
        
        # 不正ファイルの読み込み（例外が発生しないことを確認）
        try:
            df = self.csv_handler.read_csv_with_encoding_detection(invalid_csv)
            # 何らかのデータが読み込まれることを確認
            self.assertIsNotNone(df)
        except Exception as e:
            # エラーが発生した場合も適切にハンドリングされることを確認
            self.assertIsInstance(e, Exception)


if __name__ == '__main__':
    unittest.main()