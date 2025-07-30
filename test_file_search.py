#!/usr/bin/env python3
"""
ファイル検索のテストスクリプト
"""

import sys
import os
from pathlib import Path
from sales_aggregator import SalesAggregator

def main():
    print("=== ファイル検索テスト ===")
    
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
        files_by_platform = aggregator.find_files_in_yearmonth_folders()
        
        print("\n=== 検索されたファイル ===")
        for platform, files in files_by_platform.items():
            print(f"\n{platform}: {len(files)}件")
            for file in files[:3]:  # 最初の3件だけ表示
                print(f"  - {file}")
            if len(files) > 3:
                print(f"  ... 他{len(files)-3}件")
        
        # auとsoftbankファイルを特に詳しく確認
        print(f"\n=== au ファイル詳細 ===")
        for file in files_by_platform['au']:
            print(f"  {file}")
            
        print(f"\n=== softbank ファイル詳細 ===")
        for file in files_by_platform['softbank']:
            print(f"  {file}")
            
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()