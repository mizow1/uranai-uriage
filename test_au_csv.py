#!/usr/bin/env python3
"""
auのCSVファイルテスト
"""

import sys
import os
from pathlib import Path
from sales_aggregator import SalesAggregator

def main():
    print("=== auのCSVファイルテスト ===")
    
    # auのCSVファイルをテスト
    test_file = Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\2025\202503\au202503\202503cp02お支払い明細書.csv")
    
    if not test_file.exists():
        print(f"ファイルが存在しません: {test_file}")
        return
    
    print(f"テストファイル: {test_file}")
    
    try:
        # SalesAggregatorでauファイルを処理
        aggregator = SalesAggregator(test_file.parent.parent.parent)
        result = aggregator.process_au_new_file(test_file)
        
        print(f"処理成功: {result.success}")
        print(f"エラー: {result.errors}")
        print(f"詳細数: {len(result.details) if result.details else 0}")
        
        if result.details:
            for detail in result.details:
                print(f"  - {detail.content_group}: 実績={detail.performance}, 情報提供料={detail.information_fee}")
        
        if result.metadata:
            print(f"メタデータ: {result.metadata}")
            
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()