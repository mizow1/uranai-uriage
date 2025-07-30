#!/usr/bin/env python3
"""
処理のテストスクリプト
"""

import sys
import os
from pathlib import Path
from sales_aggregator import SalesAggregator

def main():
    print("=== 処理テスト ===")
    
    # デフォルトのパス設定
    base_path = r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書"
    
    # パスの存在確認
    if not Path(base_path).exists():
        print(f"エラー: 指定されたパスが存在しません: {base_path}")
        return
    
    print(f"データフォルダ: {base_path}")
    
    try:
        # 売上集計処理の実行
        aggregator = SalesAggregator(base_path)
        
        # 1つのauファイルをテスト処理
        files_by_platform = aggregator.find_files_in_yearmonth_folders()
        if files_by_platform['au']:
            test_au_file = files_by_platform['au'][0]
            print(f"\n=== auファイル処理テスト ===")
            print(f"ファイル: {test_au_file}")
            result = aggregator.process_au_new_file(test_au_file)
            print(f"成功: {result.success}")
            print(f"エラー: {result.errors}")
            print(f"詳細数: {len(result.details) if result.details else 0}")
            if result.details:
                for detail in result.details:
                    print(f"  - {detail.content_group}: 実績={detail.performance}, 情報提供料={detail.information_fee}")
        
        # 1つのsoftbankファイルをテスト処理
        if files_by_platform['softbank']:
            test_sb_file = files_by_platform['softbank'][0]
            print(f"\n=== softbankファイル処理テスト ===")
            print(f"ファイル: {test_sb_file}")
            result = aggregator.process_softbank_file(test_sb_file)
            print(f"成功: {result.success}")
            print(f"エラー: {result.errors}")
            print(f"詳細数: {len(result.details) if result.details else 0}")
            if result.details:
                for detail in result.details:
                    print(f"  - {detail.content_group}: 実績={detail.performance}, 情報提供料={detail.information_fee}")
            
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()