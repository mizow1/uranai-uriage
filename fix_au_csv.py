#!/usr/bin/env python3
"""
auCSVファイルの構造問題を解決する専用読み込み処理
"""

import pandas as pd
from pathlib import Path

def read_au_csv_properly():
    test_file = Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\2025\202503\au202503\202503cp02お支払い明細書.csv")
    
    print(f"=== auCSV専用読み込み処理 ===")
    
    try:
        # 手動でCSVを解析
        with open(test_file, 'r', encoding='shift_jis') as f:
            lines = f.readlines()
        
        print(f"総行数: {len(lines)}")
        
        # ヘッダー部分（最初の2行）
        print("\n=== ヘッダー部分 ===")
        for i in range(min(3, len(lines))):
            print(f"{i+1}: {lines[i].strip()}")
        
        # データ部分を抽出（3行目以降）
        data_lines = []
        for i, line in enumerate(lines[3:], 4):  # 4行目から
            fields = line.strip().split(',')
            # クォートを除去
            fields = [field.strip('"') for field in fields]
            if len(fields) >= 10 and fields[0]:  # 行番号がある行のみ
                data_lines.append(fields)
                if len(data_lines) <= 5:  # 最初の5行のみ表示
                    print(f"{i}: {fields}")
        
        print(f"\nデータ行数: {len(data_lines)}")
        
        # データをDataFrameに変換
        if data_lines:
            # 列名を設定（3行目のヘッダーから）
            header_line = lines[2].strip().split(',')
            headers = [h.strip('"') for h in header_line if h.strip('"')]
            
            print(f"\n列名: {headers}")
            
            # 最大列数を確認
            max_cols = max(len(row) for row in data_lines)
            print(f"最大列数: {max_cols}")
            
            # DataFrameを作成（列数を統一）
            df_data = []
            for row in data_lines:
                # 列数を統一（不足分は空文字で埋める）
                while len(row) < len(headers):
                    row.append('')
                df_data.append(row[:len(headers)])  # 余分な列は切り取り
            
            df = pd.DataFrame(df_data, columns=headers)
            print(f"\nDataFrame作成成功: {df.shape}")
            
            # 金額関連の列を特定
            amount_columns = []
            for col in df.columns:
                if any(keyword in col for keyword in ['金額', '件数', '料金', '合計', '支払']):
                    amount_columns.append(col)
            
            print(f"\n金額関連列: {amount_columns}")
            
            # 金額データを抽出
            total_amounts = []
            for col in amount_columns:
                print(f"\n=== 列 '{col}' の内容 ===")
                for idx, val in enumerate(df[col].head()):
                    print(f"  [{idx}] {repr(val)}")
                    try:
                        if val and val != '':
                            num_val = float(str(val).replace(',', ''))
                            if num_val > 0:
                                total_amounts.append((col, num_val))
                                print(f"      → 数値: {num_val}")
                    except:
                        pass
            
            # 合計金額を計算
            if total_amounts:
                print(f"\n=== 抽出された金額 ===")
                for col, amount in sorted(total_amounts, key=lambda x: x[1], reverse=True):
                    print(f"  {col}: {amount:,.0f}円")
                
                max_amount = max(total_amounts, key=lambda x: x[1])
                print(f"\n最大金額: {max_amount[1]:,.0f}円 (列: {max_amount[0]})")
                
                # 「お支払い対象金額合計」列があるかチェック
                target_col = 'お支払い対象金額合計'
                if target_col in df.columns:
                    target_values = df[target_col].dropna()
                    target_sum = 0
                    for val in target_values:
                        try:
                            if val and val != '':
                                target_sum += float(str(val).replace(',', ''))
                        except:
                            pass
                    print(f"\n{target_col}の合計: {target_sum:,.0f}円")
            
            return df
                
    except Exception as e:
        print(f"エラー: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    read_au_csv_properly()