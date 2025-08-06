#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
各プラットフォームの売上ファイルが存在するかを年月ごとに確認するプログラム
"""

import os
import glob
from datetime import datetime, timedelta
from pathlib import Path
import pandas as pd

# 基本パス
BASE_PATH = r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書"

# プラットフォーム別ファイルパターン
PLATFORMS = {
    "楽天占い": {"pattern": "*rcms*.csv", "path": ""},
    "excite占い": {"pattern": "*excite*.csv", "path": ""},
    "ameba占い": {"pattern": "*SATORI実績*.xlsx", "path": ""},
    "mediba占い": {"pattern": "*SalesSummary*.csv", "path": ""},
    "LINE占い": {"pattern": "*line-contents*.csv", "path": "line"},
    "docomo占い": {"pattern": "*bp40000746shiharai*.csv", "path": ""},
    "softbank占い": {"pattern": "OID_PAY_9ATI_{yyyymm}.pdf", "path": "{yyyymm}\\oidshiharai"},
    "au占い": {"pattern": "*cp02お支払い明細書*.csv", "path": ""}
}

def get_year_months(start_year=2020, start_month=1):
    """指定した年月から現在までの年月リストを生成"""
    start_date = datetime(start_year, start_month, 1)
    current_date = datetime.now()
    
    year_months = []
    current = start_date
    
    while current <= current_date:
        year_months.append({
            'year': current.year,
            'month': current.month,
            'yyyymm': f"{current.year:04d}{current.month:02d}"
        })
        
        # 次の月へ
        if current.month == 12:
            current = current.replace(year=current.year + 1, month=1)
        else:
            current = current.replace(month=current.month + 1)
    
    return year_months

def check_file_exists(platform, year, month, yyyymm):
    """特定のプラットフォームのファイルが存在するかチェック"""
    year_dir = os.path.join(BASE_PATH, str(year))
    month_dir = os.path.join(year_dir, yyyymm)
    
    if not os.path.exists(month_dir):
        return False, f"月ディレクトリが存在しません: {month_dir}"
    
    platform_info = PLATFORMS[platform]
    pattern = platform_info["pattern"]
    sub_path = platform_info["path"]
    
    # softbank占いの特殊処理
    if platform == "softbank占い":
        pattern = pattern.replace("{yyyymm}", yyyymm)
        sub_path = sub_path.replace("{yyyymm}", yyyymm)
    
    # 検索ディレクトリを決定
    if sub_path:
        search_dir = os.path.join(month_dir, sub_path)
    else:
        search_dir = month_dir
    
    if not os.path.exists(search_dir):
        return False, f"検索ディレクトリが存在しません: {search_dir}"
    
    # ファイルを検索
    search_pattern = os.path.join(search_dir, pattern)
    files = glob.glob(search_pattern, recursive=False)
    
    if files:
        return True, f"ファイル見つかりました: {', '.join([os.path.basename(f) for f in files])}"
    else:
        return False, f"ファイルが見つかりません: {search_pattern}"

def main():
    print("=== 各プラットフォーム売上ファイル存在確認 ===")
    print()
    
    # 開始年月を入力
    while True:
        try:
            start_year = int(input("開始年を入力してください (例: 2020): "))
            start_month = int(input("開始月を入力してください (例: 1): "))
            if 1 <= start_month <= 12:
                break
            else:
                print("月は1-12の範囲で入力してください。")
        except ValueError:
            print("数値で入力してください。")
    
    # 年月リストを生成
    year_months = get_year_months(start_year, start_month)
    
    # 結果を格納するリスト
    results = []
    
    print(f"\n{start_year}年{start_month}月から現在までの期間をチェックします...")
    print()
    
    # 各年月をチェック
    for ym in year_months:
        year = ym['year']
        month = ym['month']
        yyyymm = ym['yyyymm']
        
        print(f"=== {year}年{month}月 ({yyyymm}) ===")
        
        month_results = {
            'year': year,
            'month': month,
            'yyyymm': yyyymm
        }
        
        for platform in PLATFORMS.keys():
            exists, message = check_file_exists(platform, year, month, yyyymm)
            month_results[platform] = '○' if exists else '×'
            
            status = '○' if exists else '×'
            print(f"{platform:12}: {status} {message}")
        
        results.append(month_results)
        print()
    
    # 結果をCSVファイルに出力
    df = pd.DataFrame(results)
    output_file = "sales_files_check_result.csv"
    df.to_csv(output_file, index=False, encoding='utf-8-sig')
    
    print(f"結果をCSVファイルに保存しました: {output_file}")
    
    # サマリーを表示
    print("\n=== サマリー ===")
    total_months = len(results)
    
    for platform in PLATFORMS.keys():
        exists_count = sum(1 for r in results if r[platform] == '○')
        missing_count = total_months - exists_count
        print(f"{platform:12}: 存在 {exists_count:2d}/{total_months:2d}, 欠損 {missing_count:2d}")

if __name__ == "__main__":
    main()