#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
xlwingsテスト
"""
import xlwings as xw
import os

def test_xlwings():
    # テスト用のファイルパスを探す
    test_dir = r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ロイヤリティ"
    
    # 最新のyyyymmフォルダを探す
    folders = []
    for item in os.listdir(test_dir):
        if len(item) == 6 and item.isdigit():
            folder_path = os.path.join(test_dir, item)
            if os.path.isdir(folder_path):
                folders.append(item)
    
    if not folders:
        print("テスト用フォルダが見つかりません")
        return
    
    folders.sort()
    latest_folder = folders[-1]
    folder_path = os.path.join(test_dir, latest_folder)
    
    # フォルダ内のxlsxファイルを探す
    xlsx_files = []
    for file in os.listdir(folder_path):
        if file.endswith('.xlsx'):
            xlsx_files.append(os.path.join(folder_path, file))
    
    if not xlsx_files:
        print(f"{latest_folder}にxlsxファイルが見つかりません")
        return
    
    print(f"テスト対象: {latest_folder}")
    print(f"ファイル数: {len(xlsx_files)}")
    
    # 最初の3ファイルをテスト
    for i, file_path in enumerate(xlsx_files[:3]):
        try:
            print(f"\n{i+1}. {os.path.basename(file_path)}")
            
            app = xw.App(visible=False)
            wb = app.books.open(file_path)
            ws = wb.sheets[0]
            
            # AE19セルの値を取得
            ae19_value = ws.range('AE19').value
            print(f"   AE19: {ae19_value}")
            
            wb.close()
            app.quit()
            
        except Exception as e:
            print(f"   エラー: {e}")
            try:
                if 'wb' in locals():
                    wb.close()
                if 'app' in locals():
                    app.quit()
            except:
                pass

if __name__ == "__main__":
    test_xlwings()