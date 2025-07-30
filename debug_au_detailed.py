#!/usr/bin/env python3
"""
auファイルの詳細デバッグ
"""

import pandas as pd
import chardet
from pathlib import Path

def analyze_au_csv():
    test_file = Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\2025\202503\au202503\202503cp02お支払い明細書.csv")
    
    print(f"=== auCSVファイル詳細分析 ===")
    print(f"ファイル: {test_file}")
    
    if not test_file.exists():
        print("ファイルが存在しません")
        return
        
    # エンコーディング検出
    with open(test_file, 'rb') as f:
        raw_data = f.read()
        encoding_result = chardet.detect(raw_data)
        print(f"検出されたエンコーディング: {encoding_result}")
    
    # 複数のエンコーディングで読み込み試行
    encodings = ['utf-8', 'shift_jis', 'cp932', 'iso-2022-jp', 'euc-jp']
    
    for encoding in encodings:
        try:
            print(f"\n=== {encoding}での読み込み試行 ===")
            
            # まず最初の5行をテキストとして読み込み
            with open(test_file, 'r', encoding=encoding) as f:
                lines = []
                for i in range(10):  # 最初の10行
                    try:
                        line = f.readline()
                        if not line:
                            break
                        lines.append(line.strip())
                    except UnicodeDecodeError:
                        lines.append(f"[{i+1}行目: 読み込みエラー]")
                        
            print("ファイル内容（最初の10行）:")
            for i, line in enumerate(lines, 1):
                print(f"{i:2d}: {line}")
            
            # pandasで読み込み試行
            try:
                df = pd.read_csv(test_file, encoding=encoding, nrows=10)
                print(f"\nDataFrame読み込み成功:")
                print(f"  形状: {df.shape}")
                print(f"  列名: {list(df.columns)}")
                
                # 数値列を探す
                for col in df.columns:
                    print(f"\n  列 '{col}':")
                    for idx, val in enumerate(df[col]):
                        print(f"    [{idx}] {repr(val)} (型: {type(val)})")
                        if idx >= 3:  # 最初の4行だけ
                            break
                
                # 成功した場合はこのエンコーディングで詳細分析
                print(f"\n✓ {encoding}での読み込みに成功 - 詳細分析を実行")
                detailed_analysis(test_file, encoding)
                break
                
            except Exception as e:
                print(f"  DataFrame読み込みエラー: {e}")
                
        except Exception as e:
            print(f"  {encoding}読み込みエラー: {e}")

def detailed_analysis(file_path, encoding):
    """詳細分析"""
    try:
        df = pd.read_csv(file_path, encoding=encoding)
        print(f"\n=== 詳細分析 ===")
        print(f"全体形状: {df.shape}")
        
        # 全データから数値を探す
        numeric_values = []
        
        for col in df.columns:
            for val in df[col].dropna():
                # 数値変換を試行
                try:
                    num_val = float(val)
                    if num_val > 0:
                        numeric_values.append((col, val, num_val))
                except:
                    # 文字列内の数字を抽出
                    import re
                    str_val = str(val)
                    matches = re.findall(r'(\d{1,3}(?:,\d{3})*|\d+)', str_val)
                    for match in matches:
                        try:
                            clean_num = float(match.replace(',', ''))
                            if clean_num > 0:
                                numeric_values.append((col, val, clean_num))
                        except:
                            pass
        
        # 数値をソートして表示
        numeric_values.sort(key=lambda x: x[2], reverse=True)
        
        print(f"\n見つかった数値（上位20件）:")
        for i, (col, original, num) in enumerate(numeric_values[:20]):
            print(f"{i+1:2d}. 列'{col}': {original} → {num}")
            
        if numeric_values:
            max_val = max(numeric_values, key=lambda x: x[2])
            print(f"\n最大値: {max_val[2]} (列: {max_val[0]}, 元値: {max_val[1]})")
        
    except Exception as e:
        print(f"詳細分析エラー: {e}")

if __name__ == "__main__":
    analyze_au_csv()