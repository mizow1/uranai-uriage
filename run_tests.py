#!/usr/bin/env python3
"""
統合テスト実行スクリプト
"""
import unittest
import sys
import os
from pathlib import Path
import time

# プロジェクトルートをパスに追加
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def run_all_tests():
    """すべてのテストを実行"""
    print("=" * 60)
    print("リファクタリング統合テスト実行")
    print("=" * 60)
    
    start_time = time.time()
    
    # テストディスカバリー
    loader = unittest.TestLoader()
    start_dir = project_root / 'tests'
    suite = loader.discover(start_dir, pattern='test_*.py')
    
    # テストランナー設定
    runner = unittest.TextTestRunner(
        verbosity=2,
        stream=sys.stdout,
        buffer=True
    )
    
    # テスト実行
    print(f"\nテスト実行開始: {time.strftime('%Y-%m-%d %H:%M:%S')}")
    print("-" * 60)
    
    result = runner.run(suite)
    
    end_time = time.time()
    duration = end_time - start_time
    
    # 結果サマリー
    print("-" * 60)
    print("テスト結果サマリー:")
    print(f"実行時間: {duration:.2f}秒")
    print(f"実行テスト数: {result.testsRun}")
    print(f"成功: {result.testsRun - len(result.failures) - len(result.errors)}")
    print(f"失敗: {len(result.failures)}")
    print(f"エラー: {len(result.errors)}")
    
    if result.failures:
        print("\n失敗したテスト:")
        for test, traceback in result.failures:
            print(f"  - {test}: {traceback.split('\\n')[-2] if traceback else 'Unknown error'}")
    
    if result.errors:
        print("\nエラーが発生したテスト:")
        for test, traceback in result.errors:
            print(f"  - {test}: {traceback.split('\\n')[-2] if traceback else 'Unknown error'}")
    
    # 終了コード
    success_rate = (result.testsRun - len(result.failures) - len(result.errors)) / result.testsRun * 100 if result.testsRun > 0 else 0
    print(f"\n成功率: {success_rate:.1f}%")
    
    if result.failures or result.errors:
        print("\\n⚠️  テストに失敗またはエラーがあります")
        return 1
    else:
        print("\\n✅ すべてのテストが成功しました")
        return 0

def run_specific_test(test_name):
    """特定のテストモジュールを実行"""
    print(f"テスト実行: {test_name}")
    
    try:
        module = __import__(f'tests.{test_name}', fromlist=[test_name])
        suite = unittest.TestLoader().loadTestsFromModule(module)
        runner = unittest.TextTestRunner(verbosity=2)
        result = runner.run(suite)
        
        return 0 if not (result.failures or result.errors) else 1
    except ImportError as e:
        print(f"テストモジュールが見つかりません: {test_name}")
        print(f"エラー: {e}")
        return 1

def main():
    """メイン関数"""
    if len(sys.argv) > 1:
        test_name = sys.argv[1]
        return run_specific_test(test_name)
    else:
        return run_all_tests()

if __name__ == '__main__':
    exit_code = main()
    sys.exit(exit_code)