#!/usr/bin/env python3
"""
PDFファイルの内容確認
"""

import PyPDF2
import re
from pathlib import Path

def test_au_pdf():
    test_file = Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\2023\202303\202303cp02お支払い明細書.pdf")
    
    print("=== auファイル内容確認 ===")
    print(f"ファイル: {test_file}")
    
    try:
        with open(test_file, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            text_content = ""
            for page in pdf_reader.pages:
                text_content += page.extract_text()
        
        print("PDF内容:")
        print(text_content[:1000])  # 最初の1000文字
        
        # 様々な金額パターンを試す
        patterns = [
            r'[合計金額|金額|合計][:：\s]*([\d,]+)',
            r'合計[:：\s]*([\d,]+)',
            r'金額[:：\s]*([\d,]+)',
            r'(\d{1,3}(?:,\d{3})*)',  # カンマ区切りの数字
            r'(\d+)',  # 数字のみ
        ]
        
        for i, pattern in enumerate(patterns):
            matches = re.findall(pattern, text_content)
            print(f"パターン{i+1}: {matches[:5] if matches else '見つからず'}")
            
    except Exception as e:
        print(f"エラー: {str(e)}")

def test_softbank_pdf():
    test_file = Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\2023\202303\SB202303\oidshiharai\OID_PAY_9ATI_202303.PDF")
    
    print("\n=== softbankファイル内容確認 ===")
    print(f"ファイル: {test_file}")
    
    try:
        with open(test_file, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            text_content = ""
            for page in pdf_reader.pages:
                text_content += page.extract_text()
        
        print("PDF内容:")
        print(text_content[:1000])  # 最初の1000文字
        
        # 様々な金額パターンを試す
        patterns = [
            r'合計金額[:\s]*([\d,]+)',
            r'お支払い金額[:\s]*([\d,]+)',
            r'(\d{1,3}(?:,\d{3})*)',  # カンマ区切りの数字
            r'(\d+)',  # 数字のみ
        ]
        
        for i, pattern in enumerate(patterns):
            matches = re.findall(pattern, text_content)
            print(f"パターン{i+1}: {matches[:5] if matches else '見つからず'}")
            
    except Exception as e:
        print(f"エラー: {str(e)}")

def main():
    test_au_pdf()
    test_softbank_pdf()

if __name__ == "__main__":
    main()