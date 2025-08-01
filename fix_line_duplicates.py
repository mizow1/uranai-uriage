#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LINEの売上データの重複を除去するスクリプト
"""

import pandas as pd
import shutil
from datetime import datetime
import os

def fix_line_duplicates(csv_file_path):
    """
    LINEの売上データの重複を除去する
    
    Args:
        csv_file_path (str): CSVファイルのパス
    """
    print(f"処理開始: {csv_file_path}")
    
    # バックアップを作成
    backup_path = csv_file_path.replace('.csv', f'_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv')
    shutil.copy2(csv_file_path, backup_path)
    print(f"バックアップ作成: {backup_path}")
    
    # CSVファイルを読み込み
    df = pd.read_csv(csv_file_path)
    print(f"全データ数: {len(df)}")
    
    # LINEのデータのみ抽出
    line_data = df[df['プラットフォーム'].str.lower() == 'line']
    other_data = df[df['プラットフォーム'].str.lower() != 'line']
    
    print(f"LINE売上データ数: {len(line_data)}")
    print(f"その他のデータ数: {len(other_data)}")
    
    # 重複チェック
    duplicates = line_data.duplicated(subset=['プラットフォーム', 'コンテンツ', '年月'], keep=False)
    print(f"重複の可能性があるLINEデータ: {duplicates.sum()}")
    
    if duplicates.sum() > 0:
        print("\n重複データの詳細:")
        dup_data = line_data[duplicates].sort_values(['コンテンツ', '年月'])
        print(dup_data[['プラットフォーム', 'コンテンツ', '実績', '情報提供料', '年月']])
        
        # 重複除去（最初の行を保持）
        line_data_cleaned = line_data.drop_duplicates(subset=['プラットフォーム', 'コンテンツ', '年月'], keep='first')
        print(f"\n重複除去後のLINEデータ数: {len(line_data_cleaned)}")
        print(f"除去された重複データ数: {len(line_data) - len(line_data_cleaned)}")
        
        # データを再結合
        df_cleaned = pd.concat([other_data, line_data_cleaned], ignore_index=True)
        
        # 元のファイルを上書き
        df_cleaned.to_csv(csv_file_path, index=False)
        print(f"\n修正完了: {csv_file_path}")
        print(f"最終データ数: {len(df_cleaned)}")
        
        return {
            'original_count': len(df),
            'line_original_count': len(line_data),
            'line_cleaned_count': len(line_data_cleaned),
            'duplicates_removed': len(line_data) - len(line_data_cleaned),
            'final_count': len(df_cleaned)
        }
    else:
        print("重複データは見つかりませんでした。")
        return None

if __name__ == "__main__":
    csv_file = r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\月別ISP別コンテンツ別売上.csv"
    
    if os.path.exists(csv_file):
        result = fix_line_duplicates(csv_file)
        if result:
            print(f"\n処理結果:")
            print(f"  元のデータ数: {result['original_count']}")
            print(f"  元のLINEデータ数: {result['line_original_count']}")
            print(f"  修正後のLINEデータ数: {result['line_cleaned_count']}")
            print(f"  除去された重複数: {result['duplicates_removed']}")
            print(f"  最終データ数: {result['final_count']}")
    else:
        print(f"ファイルが見つかりません: {csv_file}")