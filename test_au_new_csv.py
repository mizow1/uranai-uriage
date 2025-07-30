#!/usr/bin/env python3
"""
auの新しいCSVファイルテスト
"""

import sys
import os
from pathlib import Path
from sales_aggregator import SalesAggregator

def main():
    print("=== auの新しいCSVファイルテスト ===")
    
    # 2025年の他のauCSVファイルをテスト
    test_files = [
        Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\2025\202501\202501cp02お支払い明細書.csv"),
        Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\2025\202504\202504cp02お支払い明細書.csv"),
        Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\2025\202505\202505cp02お支払い明細書.csv"),
    ]
    
    for test_file in test_files:
        if not test_file.exists():
            print(f"ファイルが存在しません: {test_file}")
            continue
        
        print(f"\nテストファイル: {test_file.name}")
        
        try:
            # SalesAggregatorでauファイルを処理
            aggregator = SalesAggregator(test_file.parent.parent.parent)
            result = aggregator.process_au_new_file(test_file)
            
            print(f"処理成功: {result.success}")
            if result.errors:
                print(f"エラー: {result.errors}")
            print(f"詳細数: {len(result.details) if result.details else 0}")
            
            if result.details:
                for detail in result.details:
                    print(f"  - {detail.content_group}: 実績={detail.performance}, 情報提供料={detail.information_fee}")
            
            if result.success:
                break  # 成功したらループ終了
                
        except Exception as e:
            print(f"エラーが発生しました: {str(e)}")

if __name__ == "__main__":
    main()