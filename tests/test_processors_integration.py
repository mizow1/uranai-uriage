"""
プロセッサー統合テスト
"""
import unittest
import tempfile
import pandas as pd
from pathlib import Path
import sys
import os

# プロジェクトルートをパスに追加
sys.path.insert(0, str(Path(__file__).parent.parent))

from mediba_sales_processor import MedibaSalesProcessor
from sales_aggregator import SalesAggregator
from common import ProcessingResult, ContentDetail


class TestMedibaSalesProcessor(unittest.TestCase):
    """MedibaSalesProcessorの統合テスト"""
    
    def setUp(self):
        """テスト前の準備"""
        self.temp_dir = Path(tempfile.mkdtemp())
        self.processor = MedibaSalesProcessor()
    
    def tearDown(self):
        """テスト後のクリーンアップ"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_mediba_processor_with_valid_csv(self):
        """有効なCSVファイルでのMedibaプロセッサーテスト"""
        # テストデータ作成（mediba形式）
        test_data = pd.DataFrame({
            'Column_A': ['header'],
            'Column_B': ['Program_A', 'Program_B', 'Program_A', 'Program_C'],
            'Column_C': [0, 0, 0, 0],
            'Column_D': [0, 0, 0, 0],
            'Column_E': [0, 0, 0, 0],
            'Column_F': [0, 0, 0, 0],
            'Column_G': [0, 1000, 1500, 2000],  # 料金データ
            'Column_H': [0, 0, 0, 0],
            'Column_I': [0, 0, 0, 0],
            'Column_J': [0, 0, 0, 0],
            'Column_K': [0, 100, 150, 200],  # CP売上負担額
        })
        
        test_file = self.temp_dir / "SalesSummary_test.csv"
        test_data.to_csv(test_file, index=False, encoding='utf-8')
        
        # プロセッサー実行
        result = self.processor.process_sales_data(test_file)
        
        # 検証
        self.assertIsInstance(result, ProcessingResult)
        self.assertEqual(result.platform, "mediba")
        self.assertTrue(result.success)
        self.assertGreater(len(result.details), 0)
        self.assertGreater(result.total_performance, 0)
        self.assertGreater(result.total_information_fee, 0)
    
    def test_mediba_processor_with_insufficient_columns(self):
        """列数不足のCSVファイルでのテスト"""
        # 不完全なテストデータ
        test_data = pd.DataFrame({
            'A': [1, 2, 3],
            'B': ['Program_A', 'Program_B', 'Program_C']
        })
        
        test_file = self.temp_dir / "incomplete.csv"
        test_data.to_csv(test_file, index=False, encoding='utf-8')
        
        # プロセッサー実行
        result = self.processor.process_sales_data(test_file)
        
        # 検証
        self.assertIsInstance(result, ProcessingResult)
        self.assertFalse(result.success)
        self.assertGreater(len(result.errors), 0)
        self.assertIn("列数が不足", result.errors[0])


class TestSalesAggregator(unittest.TestCase):
    """SalesAggregatorの統合テスト"""
    
    def setUp(self):
        """テスト前の準備"""
        self.temp_dir = Path(tempfile.mkdtemp())
        self.aggregator = SalesAggregator(self.temp_dir)
    
    def tearDown(self):
        """テスト後のクリーンアップ"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_ameba_processor_basic(self):
        """Amebaプロセッサーの基本テスト"""
        # テスト用のExcelファイルを作成
        test_data_従量 = pd.DataFrame({
            'A': [0, 0, 0],
            'B': [0, 0, 0],
            'C': ['Content_A', 'Content_B', 'Content_A'],
            'D': [0, 0, 0], 'E': [0, 0, 0], 'F': [0, 0, 0], 
            'G': [0, 0, 0], 'H': [0, 0, 0], 'I': [0, 0, 0],
            'J': [1000, 2000, 1500]  # 情報提供料
        })
        
        test_file = self.temp_dir / "satori実績_test.xlsx"
        
        # 複数シートを持つExcelファイルを作成
        with pd.ExcelWriter(test_file, engine='openpyxl') as writer:
            test_data_従量.to_excel(writer, sheet_name='従量実績', index=False)
        
        # プロセッサー実行
        result = self.aggregator.process_ameba_file(test_file)
        
        # 検証
        self.assertIsInstance(result, ProcessingResult)
        self.assertEqual(result.platform, "ameba")
        self.assertTrue(result.success)
        self.assertGreater(len(result.details), 0)
    
    def test_rakuten_processor_with_rcms_file(self):
        """楽天プロセッサーのRCMSファイルテスト"""
        # RCMS形式のテストデータ
        test_data = pd.DataFrame({
            'A': [0, 0, 0, 0],
            'B': [0, 0, 0, 0], 'C': [0, 0, 0, 0], 'D': [0, 0, 0, 0],
            'E': [0, 0, 0, 0], 'F': [0, 0, 0, 0], 'G': [0, 0, 0, 0],
            'H': [0, 0, 0, 0], 'I': [0, 0, 0, 0], 'J': [0, 0, 0, 0],
            'K': [0, 0, 0, 0],
            'L': ['hoge_001', 'fuga_002', 'hoge_003', 'piyo_004'],  # L列
            'M': [0, 0, 0, 0],
            'N': [1100, 2200, 1650, 3300]  # N列（税込み金額）
        })
        
        test_file = self.temp_dir / "rcms_test.csv"
        test_data.to_csv(test_file, index=False, encoding='utf-8')
        
        # プロセッサー実行
        result = self.aggregator.process_rakuten_file(test_file)
        
        # 検証
        self.assertIsInstance(result, ProcessingResult)
        self.assertEqual(result.platform, "rakuten")
        self.assertTrue(result.success)
        self.assertGreater(len(result.details), 0)
        
        # hoge グループの計算確認
        hoge_detail = next((d for d in result.details if d.content_group == 'hoge'), None)
        self.assertIsNotNone(hoge_detail)
        # (1100 + 1650) / 1.1 = 2500 (実績)
        # 2500 * 0.725 = 1812.5 (情報提供料)
        self.assertEqual(hoge_detail.performance, 2500)
        self.assertEqual(hoge_detail.information_fee, 1813)  # 四捨五入
    
    def test_au_processor_with_sales_summary(self):
        """medibaプロセッサーのSalesSummaryファイルテスト"""
        # SalesSummary形式のテストデータ
        test_data = pd.DataFrame({
            'A': [0, 0, 0],
            'B': ['Program_X', 'Program_Y', 'Program_X'],  # B列（番組ID）
            'C': [0, 0, 0], 'D': [0, 0, 0], 'E': [0, 0, 0], 'F': [0, 0, 0],
            'G': [1000, 2000, 1500],  # G列（料金）
            'H': [0, 0, 0], 'I': [0, 0, 0], 'J': [0, 0, 0],
            'K': [100, 200, 150]  # K列（CP売上負担額）
        })
        
        test_file = self.temp_dir / "SalesSummary_au_test.csv"
        test_data.to_csv(test_file, index=False, encoding='utf-8')
        
        # プロセッサー実行
        result = self.aggregator.process_au_file(test_file)
        
        # 検証
        self.assertIsInstance(result, ProcessingResult)
        self.assertEqual(result.platform, "mediba") 
        self.assertTrue(result.success)
        self.assertGreater(len(result.details), 0)
        
        # Program_Xの計算確認
        program_x_detail = next((d for d in result.details if d.content_group == 'Program_X'), None)
        self.assertIsNotNone(program_x_detail)
        # G列合計: 1000 + 1500 = 2500 (実績)
        # 情報提供料: (2500 * 0.4) - (100 + 150) = 1000 - 250 = 750
        self.assertEqual(program_x_detail.performance, 2500)
        self.assertEqual(program_x_detail.information_fee, 750)
    
    def test_line_processor_with_contents_file(self):
        """LINEプロセッサーのコンテンツファイルテスト"""
        # LINE contents形式のテストデータ
        test_data = pd.DataFrame({
            'A': [0, 0, 0],
            'B': [0, 0, 0],
            'C': [0, 0, 0],
            'D': ['service_A', 'service_B', 'service_A'],  # D列（service_name）
            'E': [1000, 2000, 1500],  # E列（revenue）
            'F': [0, 0, 0]
        })
        
        test_file = self.temp_dir / "line-contents-test.csv"
        test_data.to_csv(test_file, index=False, encoding='utf-8')
        
        # プロセッサー実行
        result = self.aggregator.process_line_file(test_file)
        
        # 検証
        self.assertIsInstance(result, ProcessingResult)
        self.assertEqual(result.platform, "line")
        self.assertTrue(result.success)
        self.assertGreater(len(result.details), 0)
        
        # service_Aの計算確認
        service_a_detail = next((d for d in result.details if d.content_group == 'service_A'), None)
        self.assertIsNotNone(service_a_detail)
        # revenue合計: 1000 + 1500 = 2500 (実績)
        # 情報提供料: 2500 * 0.6 = 1500
        self.assertEqual(service_a_detail.performance, 2500)
        self.assertEqual(service_a_detail.information_fee, 1500)


class TestEndToEndIntegration(unittest.TestCase):
    """エンドツーエンド統合テスト"""
    
    def setUp(self):
        """テスト前の準備"""
        self.temp_dir = Path(tempfile.mkdtemp())
    
    def tearDown(self):
        """テスト後のクリーンアップ"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_complete_processing_workflow(self):
        """完全な処理ワークフローのテスト"""
        # 複数のプロセッサーを使用した統合テスト
        processor = MedibaSalesProcessor()
        
        # テストデータファイルを複数作成
        files_and_results = []
        
        # Mediba形式ファイル1
        mediba_data1 = pd.DataFrame({
            'A': ['header'], 'B': ['Prog_1', 'Prog_2'], 'C': [0, 0], 'D': [0, 0],
            'E': [0, 0], 'F': [0, 0], 'G': [0, 1000, 2000], 'H': [0, 0, 0],
            'I': [0, 0, 0], 'J': [0, 0, 0], 'K': [0, 100, 200]
        })
        file1 = self.temp_dir / "SalesSummary_1.csv"
        mediba_data1.to_csv(file1, index=False, encoding='utf-8')
        
        # Mediba形式ファイル2
        mediba_data2 = pd.DataFrame({
            'A': ['header'], 'B': ['Prog_3', 'Prog_4'], 'C': [0, 0], 'D': [0, 0],
            'E': [0, 0], 'F': [0, 0], 'G': [0, 1500, 2500], 'H': [0, 0, 0],
            'I': [0, 0, 0], 'J': [0, 0, 0], 'K': [0, 150, 250]
        })
        file2 = self.temp_dir / "SalesSummary_2.csv"
        mediba_data2.to_csv(file2, index=False, encoding='utf-8')
        
        # 各ファイルを処理
        result1 = processor.process_sales_data(file1)
        result2 = processor.process_sales_data(file2)
        
        # 結果の検証
        self.assertTrue(result1.success)
        self.assertTrue(result2.success)
        self.assertGreater(result1.total_performance, 0)
        self.assertGreater(result2.total_performance, 0)
        
        # 合計処理
        total_performance = result1.total_performance + result2.total_performance
        total_information_fee = result1.total_information_fee + result2.total_information_fee
        
        self.assertGreater(total_performance, 0)
        self.assertGreater(total_information_fee, 0)
        
        # 保存処理のテスト
        results = [result1, result2]
        save_success = processor.save_results(results, str(self.temp_dir / "test_output.csv"))
        self.assertTrue(save_success)
        
        # 保存されたファイルの確認
        output_file = self.temp_dir / "test_output.csv"
        self.assertTrue(output_file.exists())
        
        # 保存されたデータの読み込みと確認
        saved_df = pd.read_csv(output_file)
        self.assertGreater(len(saved_df), 0)
        self.assertIn('番組ID', saved_df.columns)
        self.assertIn('実績', saved_df.columns)
        self.assertIn('情報提供料合計', saved_df.columns)


if __name__ == '__main__':
    # テストスイートの実行
    unittest.main(verbosity=2)