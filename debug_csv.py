#!/usr/bin/env python3
"""
CSVファイルの詳細確認
"""

import pandas as pd
from pathlib import Path

def main():
    test_file = Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\2024\202405\bp40000746shiharai20250730143312.csv")
    
    print("=== CSVファイル詳細確認 ===")
    print(f"ファイル: {test_file}")
    
    try:
        # shift_jisで読み込み、区切り文字やクォート文字を調整
        df_skip3 = pd.read_csv(
            test_file, 
            encoding='shift_jis', 
            skiprows=3,
            low_memory=False,
            on_bad_lines='skip'
        )
        print(f"skiprows=3: {len(df_skip3)}行処理")
        
        df_skip4 = pd.read_csv(
            test_file, 
            encoding='shift_jis', 
            skiprows=4,
            low_memory=False,
            on_bad_lines='skip'
        )
        print(f"skiprows=4: {len(df_skip4)}行処理") 
        
        # R列の内容比較（18列目、0ベースで17）
        if len(df_skip3.columns) >= 18:
            print("\n=== skiprows=3でのR列内容 ===")
            r_col_skip3 = df_skip3.iloc[:, 17]
            for i, val in enumerate(r_col_skip3.head()):
                if pd.notna(val) and str(val).strip():
                    print(f"行{i+1}: {val}")
                    
        if len(df_skip4.columns) >= 18:
            print("\n=== skiprows=4でのR列内容 ===")
            r_col_skip4 = df_skip4.iloc[:, 17]
            for i, val in enumerate(r_col_skip4.head()):
                if pd.notna(val) and str(val).strip():
                    print(f"行{i+1}: {val}")
        
        # 違いがある行を特定
        print(f"\n=== 行数の違い ===")
        print(f"skiprows=3: {len(df_skip3)}行")
        print(f"skiprows=4: {len(df_skip4)}行")
        print(f"差分: {len(df_skip3) - len(df_skip4)}行")
        
    except Exception as e:
        print(f"エラー: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()