#!/usr/bin/env python3
"""
月別ISP別コンテンツ別売上.csvのプラットフォーム名を修正
auプラットフォームのOWD_*コンテンツをmedibaに変更
"""

import pandas as pd
import csv
from pathlib import Path

def fix_platform_names():
    """auプラットフォームのOWD_*コンテンツをmedibaに変更"""
    
    csv_file = Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\月別ISP別コンテンツ別売上.csv")
    backup_file = csv_file.with_suffix('.csv.backup')
    
    print(f"CSVファイルを処理中: {csv_file}")
    
    # バックアップを作成
    if csv_file.exists():
        import shutil
        shutil.copy2(csv_file, backup_file)
        print(f"バックアップ作成: {backup_file}")
    
    # CSVファイルを読み込み
    try:
        df = pd.read_csv(csv_file, encoding='utf-8-sig')
        print(f"CSVファイル読み込み完了: {len(df)}行")
        
        # 修正前の統計
        au_owd_count = len(df[(df['プラットフォーム'] == 'au') & (df['コンテンツ'].str.startswith('OWD_'))])
        print(f"修正対象: au プラットフォームの OWD_* コンテンツ {au_owd_count}件")
        
        # auプラットフォームでOWD_で始まるコンテンツをmedibaに変更
        mask = (df['プラットフォーム'] == 'au') & (df['コンテンツ'].str.startswith('OWD_'))
        df.loc[mask, 'プラットフォーム'] = 'mediba'
        
        # 修正後の統計
        modified_count = mask.sum()
        remaining_au_owd = len(df[(df['プラットフォーム'] == 'au') & (df['コンテンツ'].str.startswith('OWD_'))])
        mediba_count = len(df[df['プラットフォーム'] == 'mediba'])
        
        print(f"修正完了: {modified_count}件のプラットフォーム名を au → mediba に変更")
        print(f"残りのau OWD_*コンテンツ: {remaining_au_owd}件")
        print(f"medibaプラットフォーム合計: {mediba_count}件")
        
        # CSVファイルに保存
        df.to_csv(csv_file, index=False, encoding='utf-8-sig')
        print(f"修正済みCSVファイル保存完了: {csv_file}")
        
        return True
        
    except Exception as e:
        print(f"エラー: {e}")
        return False

if __name__ == '__main__':
    fix_platform_names()