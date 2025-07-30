#!/usr/bin/env python3
"""
修正されたテンプレートファイルの検証
"""

import openpyxl
from pathlib import Path
import random

def verify_template_fixes():
    """修正されたテンプレートファイルを検証"""
    template_dir = Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ロイヤリティ\コンテンツ関連支払明細書フォーマット")
    
    # いくつかのファイルをランダムに選んで検証
    excel_files = list(template_dir.glob("*.xlsx"))
    sample_files = random.sample(excel_files, min(10, len(excel_files)))
    
    print("=== 修正内容の検証 ===")
    
    for template_path in sample_files:
        print(f"\n検証中: {template_path.name}")
        
        try:
            wb = openpyxl.load_workbook(template_path)
            ws = wb.active
            
            # AE23-AE27のセルを確認
            modified_count = 0
            unmodified_count = 0
            
            for row in range(23, 28):  # 最初の5行をチェック
                cell_ref = f"AE{row}"
                cell = ws[cell_ref]
                
                if cell.value and str(cell.value).startswith('='):
                    formula = str(cell.value)
                    
                    if "ROUND(" in formula:
                        modified_count += 1
                        if row == 23:  # 最初の行の詳細を表示
                            print(f"  {cell_ref}: {formula}")
                    elif f"Y{row}*AC{row}*0.01" in formula:
                        unmodified_count += 1
                        print(f"  未修正: {cell_ref}: {formula}")
            
            if modified_count > 0 and unmodified_count == 0:
                print(f"  ✅ 修正完了: {modified_count}個のセルでROUND関数を確認")
            elif unmodified_count > 0:
                print(f"  ❌ 未修正あり: {unmodified_count}個のセルが未修正")
            else:
                print(f"  ⚠️  該当する数式が見つかりません")
            
            wb.close()
            
        except Exception as e:
            print(f"  エラー: {e}")

if __name__ == '__main__':
    verify_template_fixes()