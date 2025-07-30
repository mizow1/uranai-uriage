#!/usr/bin/env python3
"""
keikoファイルの計算結果確認
"""

import openpyxl
import win32com.client
from pathlib import Path

def test_keiko_calculation():
    """keikoファイルの計算結果を確認"""
    excel_path = r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ロイヤリティ\202505\202505_keiko.xlsx"
    
    if not Path(excel_path).exists():
        print(f"ファイルが見つかりません: {excel_path}")
        return
    
    try:
        # COMアプリケーションでExcelを開いて強制計算
        excel_app = win32com.client.Dispatch("Excel.Application")
        excel_app.Visible = False
        
        workbook = excel_app.Workbooks.Open(excel_path)
        worksheet = workbook.ActiveSheet
        
        # 計算を強制実行
        excel_app.Calculate()
        
        # M19セルの値を取得
        m19_value = worksheet.Range("M19").Value
        print(f"KEIKO M19セルの計算結果: {m19_value}")
        
        # AE列の値も確認
        print("KEIKO AE列の計算結果:")
        ae_total = 0
        for row in range(23, 35):  # データがありそうな範囲
            ae_value = worksheet.Range(f"AE{row}").Value
            y_value = worksheet.Range(f"Y{row}").Value
            ac_value = worksheet.Range(f"AC{row}").Value
            
            if ae_value and ae_value != 0:
                print(f"  AE{row}: {ae_value} (Y{row}={y_value}, AC{row}={ac_value})")
                ae_total += ae_value
        
        print(f"AE列の合計: {ae_total}")
        
        # 整数値かどうか確認
        if m19_value and isinstance(m19_value, (int, float)):
            if m19_value == int(m19_value):
                print("✅ M19セルは整数値です（ROUND関数が正常に動作）")
            else:
                print("❌ M19セルに小数点があります")
        
        workbook.Save()
        workbook.Close()
        excel_app.Quit()
        
    except Exception as e:
        print(f"エラー: {e}")
        try:
            excel_app.Quit()
        except:
            pass

if __name__ == '__main__':
    test_keiko_calculation()