#!/usr/bin/env python3
"""
SoftBank単一ファイルのテストスクリプト
"""

from pathlib import Path
from sales_aggregator import SalesAggregator

def test_single_softbank_file():
    """単一のSoftBankファイルをテスト"""
    
    # テスト対象のファイル
    test_file = Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\2023\202306\OID_PAY_9ATI_202306.PDF")
    
    if not test_file.exists():
        print(f"エラー: テストファイルが存在しません: {test_file}")
        return
    
    print(f"=== SoftBank単一ファイルテスト ===")
    print(f"テストファイル: {test_file.name}")
    
    # Sales Aggregatorを初期化
    aggregator = SalesAggregator(test_file.parent.parent.parent)
    
    # SoftBankファイルを処理
    result = aggregator.process_softbank_file(test_file)
    
    print(f"\n=== 処理結果 ===")
    print(f"成功: {result.success}")
    print(f"エラー: {result.errors}")
    print(f"プラットフォーム: {result.platform}")
    print(f"ファイル名: {result.file_name}")
    
    if result.success and result.details:
        for detail in result.details:
            print(f"\nコンテンツ: {detail.content_group}")
            print(f"実績: {detail.performance:,}円")
            print(f"情報提供料: {detail.information_fee:,}円")
            print(f"件数: {detail.sales_count}")
    
    if result.metadata:
        print(f"\n=== メタデータ ===")
        for key, value in result.metadata.items():
            print(f"{key}: {value}")
    
    # 期待値との比較
    if result.success and result.details:
        detail = result.details[0]
        expected_performance = round(10754 / 1.1)  # 9776円
        expected_info_fee = round(7217 / 1.1)      # 6561円
        
        print(f"\n=== 期待値との比較 ===")
        print(f"実績: 期待値={expected_performance:,}円, 実際={detail.performance:,}円, {'✓' if detail.performance == expected_performance else '✗'}")
        print(f"情報提供料: 期待値={expected_info_fee:,}円, 実際={detail.information_fee:,}円, {'✓' if detail.information_fee == expected_info_fee else '✗'}")

if __name__ == "__main__":
    test_single_softbank_file()