#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LINEの重複データ修正テスト
"""

import sys
import os
from pathlib import Path
from sales_aggregator import SalesAggregator

def main():
    print("=== LINE重複データ修正テスト ===")
    
    base_path = r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書"
    
    if not Path(base_path).exists():
        print(f"エラー: 指定されたパスが存在しません: {base_path}")
        return
    
    try:
        # 修正された売上集計処理を実行
        aggregator = SalesAggregator(base_path)
        
        # LINEファイルのみを対象にテスト
        files_by_platform = aggregator.find_files_in_yearmonth_folders()
        line_files = files_by_platform.get('line', [])
        
        print(f"検出されたLINEファイル数: {len(line_files)}")
        
        # 各LINEファイルを処理してログを確認
        for file_path in line_files[:5]:  # 最初の5ファイルのみテスト
            print(f"\n処理中: {file_path.name}")
            result = aggregator.process_line_file(file_path)
            
            if result.success:
                print(f"  成功: {len(result.details)}コンテンツ")
                for detail in result.details[:3]:  # 最初の3コンテンツのみ表示
                    print(f"    - {detail.content_group}: 実績={detail.performance}, 情報提供料={detail.information_fee}")
            else:
                print(f"  エラー/スキップ: {', '.join(result.errors)}")
        
        # 全ファイル処理（重複除去機能付き）
        print(f"\n全ファイル処理を開始...")
        aggregator.process_all_files()
        
        if aggregator.results:
            # テスト用の出力ファイル
            test_output = Path(base_path) / "test_月別ISP別コンテンツ別売上.csv"
            aggregator.export_to_csv(str(test_output))
            
            print(f"\n=== 処理結果 ===")
            print(f"処理済みファイル数: {len(aggregator.results)}")
            
            # LINEデータの統計
            line_results = [r for r in aggregator.results if r['platform'] == 'line']
            if line_results:
                print(f"LINEファイル処理結果: {len(line_results)}件")
                line_total = sum(r['情報提供料'] for r in line_results)
                print(f"LINE情報提供料合計: {line_total:,.0f}円")
            
            print(f"\nテスト結果は '{test_output}' に保存されました。")
        else:
            print("処理対象のファイルが見つかりませんでした。")
        
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()