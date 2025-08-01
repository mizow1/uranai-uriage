#!/usr/bin/env python3
"""
完全な売上集計のテスト
"""

import sys
import os
from pathlib import Path
from sales_aggregator import SalesAggregator

def main():
    print("=== 完全売上集計テスト ===")
    
    # デフォルトのパス設定
    base_path = r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書"
    output_path = "月別ISP別コンテンツ別売上_テスト.csv"
    
    # パスの存在確認
    if not Path(base_path).exists():
        print(f"エラー: 指定されたパスが存在しません: {base_path}")
        return
    
    print(f"データフォルダ: {base_path}")
    print(f"出力ファイル: {output_path}")
    
    try:
        # 売上集計処理の実行
        aggregator = SalesAggregator(base_path)
        aggregator.process_all_files()
        
        if not aggregator.results:
            print("警告: 処理対象のファイルが見つかりませんでした。")
            return
        
        # CSV出力（現在のディレクトリに出力）
        full_output_path = Path.cwd() / output_path
        aggregator.export_to_csv(str(full_output_path))
        
        print(f"\n=== 処理結果 ===")
        print(f"処理済みファイル数: {len(aggregator.results)}")
        
        # プラットフォーム別の結果表示
        platform_totals = {}
        for result in aggregator.results:
            platform = result['platform']
            if platform not in platform_totals:
                platform_totals[platform] = 0
            platform_totals[platform] += result['情報提供料']
        
        print("\nプラットフォーム別分配額:")
        for platform, total in platform_totals.items():
            print(f"  {platform}: {total:,.0f}円")
        
        overall_total = sum(platform_totals.values())
        print(f"\n全体合計: {overall_total:,.0f}円")
        
        print(f"\n結果は '{full_output_path}' に保存されました。")
        
        # au/softbankが含まれているかチェック
        if 'au' in platform_totals:
            print(f"✓ auが含まれています: {platform_totals['au']:,.0f}円")
        else:
            print("✗ auが含まれていません")
            
        if 'softbank' in platform_totals:
            print(f"✓ softbankが含まれています: {platform_totals['softbank']:,.0f}円")
        else:
            print("✗ softbankが含まれていません")
        
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()