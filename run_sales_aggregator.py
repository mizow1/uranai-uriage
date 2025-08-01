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
    output_path = "月別ISP別コンテンツ別売上.csv"  # 固定ファイル名
    
    # パス設定（入力不要でデフォルトパスを使用）
    base_path = default_base_path
    print(f"使用するデータフォルダのパス: {base_path}")
    
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
            print("- mediba占い: 'salessummary'を含むファイル")
            print("- excite占い: 'excite'を含むファイル")
            print("- LINE占い: 'line'を含むファイル")
            print("- docomo占い: 'bp40000746'を含むCSVファイル")
            print("- au占い: 'cp02お支払い明細書.csv'を含むファイル")
            print("- softbank占い: 'OID_PAY_9ATI'を含むPDFファイル")
            return
        
        # CSV出力（指定されたデータフォルダに出力）
        full_output_path = Path(base_path) / output_path
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
        
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()