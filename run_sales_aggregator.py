#!/usr/bin/env python3
"""
売上集計システムの実行スクリプト
"""

import sys
import os
from pathlib import Path
from sales_aggregator import SalesAggregator

def main():
    print("=== 売上集計システム ===")
    
    # デフォルトのパス設定
    default_base_path = r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書"
    default_output_path = "sales_distribution_summary.csv"
    
    # パス設定の確認
    base_path = input(f"データフォルダのパス (デフォルト: {default_base_path}): ").strip()
    if not base_path:
        base_path = default_base_path
    
    output_path = input(f"出力CSVファイル名 (デフォルト: {default_output_path}): ").strip()
    if not output_path:
        output_path = default_output_path
    
    # パスの存在確認
    if not Path(base_path).exists():
        print(f"エラー: 指定されたパスが存在しません: {base_path}")
        return
    
    print(f"\n処理開始...")
    print(f"データフォルダ: {base_path}")
    print(f"出力ファイル: {output_path}")
    
    try:
        # 売上集計処理の実行
        aggregator = SalesAggregator(base_path)
        aggregator.process_all_files()
        
        if not aggregator.results:
            print("警告: 処理対象のファイルが見つかりませんでした。")
            print("以下のファイル名パターンを確認してください:")
            print("- ameba占い: '【株式会社アウトワード御中】satori実績_'を含むファイル")
            print("- 楽天占い: 'rcms'を含むファイル")
            print("- au占い: 'salessummary'を含むファイル")
            print("- excite占い: 'excite'を含むファイル")
            print("- LINE占い: 'line'を含むファイル")
            return
        
        # CSV出力
        aggregator.export_to_csv(output_path)
        
        print(f"\n=== 処理結果 ===")
        print(f"処理済みファイル数: {len(aggregator.results)}")
        
        # プラットフォーム別の結果表示
        platform_totals = {}
        for result in aggregator.results:
            platform = result['platform']
            if platform not in platform_totals:
                platform_totals[platform] = 0
            platform_totals[platform] += result['total_amount']
        
        print("\nプラットフォーム別分配額:")
        for platform, total in platform_totals.items():
            print(f"  {platform}: {total:,.0f}円")
        
        overall_total = sum(platform_totals.values())
        print(f"\n全体合計: {overall_total:,.0f}円")
        
        print(f"\n結果は '{output_path}' に保存されました。")
        
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()