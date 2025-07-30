#!/usr/bin/env python3
"""
docomoファイルのテストスクリプト
"""

import sys
import os
import pandas as pd
from pathlib import Path
from sales_aggregator import SalesAggregator

def main():
    print("=== docomoファイル処理テスト ===")
    
    # テスト用のdocomoファイル
    test_file = Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\2024\202405\bp40000746shiharai20250730143312.csv")
    
    if not test_file.exists():
        print(f"ファイルが存在しません: {test_file}")
        return
    
    print(f"ファイル: {test_file}")
    
    try:
        # 様々なエンコーディングで読み込みテスト
        encodings = ['utf-8', 'shift_jis', 'cp932', 'utf-8-sig']
        
        for encoding in encodings:
            try:
                print(f"\n=== {encoding}エンコーディングでの読み込み ===")
                
                # 最初の10行を読み込み（ヘッダー含む）
                df_full = pd.read_csv(test_file, encoding=encoding, nrows=10)
                print(f"全体（最初10行）: {len(df_full)}行, {len(df_full.columns)}列")
                print("列名:", list(df_full.columns)[:5], "...")
                
                # 現在のskiprows=3での読み込み
                df_skip3 = pd.read_csv(test_file, encoding=encoding, skiprows=3, nrows=5)
                print(f"skiprows=3: {len(df_skip3)}行, {len(df_skip3.columns)}列")
                
                # 以前のskiprows=4での読み込み
                df_skip4 = pd.read_csv(test_file, encoding=encoding, skiprows=4, nrows=5)
                print(f"skiprows=4: {len(df_skip4)}行, {len(df_skip4.columns)}列")
                
                # R列（18列目）の内容を確認
                if len(df_skip3.columns) >= 18:
                    r_col_skip3 = df_skip3.iloc[:, 17]  # R列
                    print("R列（skiprows=3）:", r_col_skip3.tolist()[:3])
                
                if len(df_skip4.columns) >= 18:
                    r_col_skip4 = df_skip4.iloc[:, 17]  # R列
                    print("R列（skiprows=4）:", r_col_skip4.tolist()[:3])
                
                break  # 成功したエンコーディングで処理を続行
                
            except Exception as e:
                print(f"{encoding}での読み込み失敗: {str(e)}")
                continue
        
        # 実際のSalesAggregatorでの処理テスト
        print(f"\n=== SalesAggregatorでの処理テスト ===")
        aggregator = SalesAggregator(test_file.parent.parent.parent)
        result = aggregator.process_docomo_file(test_file)
        print(f"成功: {result.success}")
        print(f"エラー: {result.errors}")
        print(f"詳細数: {len(result.details) if result.details else 0}")
        if result.details:
            for detail in result.details[:3]:  # 最初の3件
                print(f"  - {detail.content_group}: 実績={detail.performance}, 情報提供料={detail.information_fee}")
            
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()