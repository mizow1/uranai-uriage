#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
LINEファイル検出の詳細デバッグ
"""

import sys
import os
from pathlib import Path
from sales_aggregator import SalesAggregator

def main():
    print("=== LINEファイル検出デバッグ ===")
    
    base_path = r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書"
    
    if not Path(base_path).exists():
        print(f"エラー: 指定されたパスが存在しません: {base_path}")
        return
    
    try:
        aggregator = SalesAggregator(base_path)
        
        # 年月フォルダ内のLINEファイルを詳細に調査
        files_by_platform = aggregator.find_files_in_yearmonth_folders()
        line_files = files_by_platform.get('line', [])
        
        print(f"検出されたLINEファイル数: {len(line_files)}")
        print("\n=== LINEファイル一覧 ===")
        
        for i, file_path in enumerate(line_files):
            print(f"{i+1}. {file_path}")
            print(f"   ディレクトリ: {file_path.parent}")
            print(f"   ファイル名: {file_path.name}")
            
            # ファイル名パターン分析
            filename_lower = file_path.name.lower()
            if 'line-contents-' in filename_lower:
                print(f"   → 集計済みファイル（スキップ対象）")
            elif 'line-menu-' in filename_lower:
                print(f"   → メニューファイル（処理対象）")
            elif filename_lower.startswith('line-'):
                print(f"   → その他のLINEファイル")
            else:
                print(f"   → 'line'を含むファイル")
        
        # 実際の処理テスト
        print(f"\n=== 処理テスト ===")
        for i, file_path in enumerate(line_files[:5]):  # 最初の5ファイルのみテスト
            print(f"\n{i+1}. {file_path.name} の処理テスト:")
            result = aggregator.process_line_file(file_path)
            
            if result.success:
                print(f"   成功: {len(result.details)}コンテンツ")
            else:
                print(f"   失敗/スキップ: {', '.join(result.errors)}")
        
        # line-menu-ファイルの検索
        print(f"\n=== line-menu-ファイルの検索 ===")
        menu_files = []
        for file_path in line_files:
            if 'line-menu-' in file_path.name.lower():
                menu_files.append(file_path)
        
        print(f"line-menu-ファイル数: {len(menu_files)}")
        for file_path in menu_files[:3]:  # 最初の3ファイル
            print(f"  - {file_path}")
        
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()