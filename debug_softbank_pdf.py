#!/usr/bin/env python3
"""
SoftBank PDFファイルの金額抽出のデバッグスクリプト
"""

import sys
import re
from pathlib import Path

def debug_pdf_parsing(pdf_path):
    """PDFファイルから金額抽出をデバッグ"""
    try:
        import PyPDF2
        
        print(f"=== PDFファイル解析: {pdf_path} ===")
        
        # PDFからテキストを抽出
        text_content = ""
        with open(pdf_path, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            print(f"ページ数: {len(pdf_reader.pages)}")
            
            for i, page in enumerate(pdf_reader.pages):
                try:
                    page_text = page.extract_text()
                    text_content += page_text
                    print(f"ページ {i+1} のテキスト文字数: {len(page_text)}")
                except Exception as e:
                    print(f"ページ {i+1} の読み込みエラー: {str(e)}")
        
        print(f"\n総テキスト文字数: {len(text_content)}")
        
        # テキストの一部を表示（デバッグ用）
        print(f"\n=== 抽出されたテキストの先頭500文字 ===")
        print(text_content[:500])
        print("\n" + "="*50)
        
        # 金額抽出パターンのテスト
        print(f"\n=== 金額抽出パターンのテスト ===")
        
        # 1. 合計金額の抽出パターン
        total_amount_patterns = [
            r'＜④合計＞[^0-9]*合計金額[^0-9]*(\d{1,3}(?:,\d{3})*)',
            r'合計金額[^0-9]*(\d{1,3}(?:,\d{3})*)',
            r'④合計[^0-9]*(\d{1,3}(?:,\d{3})*)',
            r'＜④合計＞.*?(\d{1,3}(?:,\d{3})*)',
            r'④合計.*?(\d{1,3}(?:,\d{3})*)'
        ]
        
        print("合計金額抽出パターンのテスト:")
        for i, pattern in enumerate(total_amount_patterns):
            matches = re.findall(pattern, text_content, re.MULTILINE | re.DOTALL)
            print(f"  パターン {i+1}: {pattern}")
            print(f"    結果: {matches[:5]}")  # 最初の5件のみ表示
        
        # 2. お支払い金額の抽出パターン
        payment_amount_patterns = [
            r'お支払い金額[^0-9]*(\d{1,3}(?:,\d{3})*)',
            r'支払い金額[^0-9]*(\d{1,3}(?:,\d{3})*)',
            r'お支払い.*?(\d{1,3}(?:,\d{3})*)',
            r'支払い.*?(\d{1,3}(?:,\d{3})*)'
        ]
        
        print("\nお支払い金額抽出パターンのテスト:")
        for i, pattern in enumerate(payment_amount_patterns):
            matches = re.findall(pattern, text_content, re.MULTILINE | re.DOTALL)
            print(f"  パターン {i+1}: {pattern}")
            print(f"    結果: {matches[:5]}")
        
        # 3. 一般的な金額パターン
        general_patterns = [
            r'(\d{1,3}(?:,\d{3})*)[円～\s]*',
            r'(\d+)[円]',
            r'(\d{1,3}(?:,\d{3})*)'
        ]
        
        print("\n一般的な金額パターンのテスト:")
        for i, pattern in enumerate(general_patterns):
            matches = re.findall(pattern, text_content, re.MULTILINE)
            # 1000以上の数値のみフィルタリング
            valid_matches = [m for m in matches if int(str(m).replace(',', '')) >= 1000]
            print(f"  パターン {i+1}: {pattern}")
            print(f"    結果（1000以上）: {valid_matches[:10]}")
        
        # 4. 特定キーワード周辺のテキスト検索
        keywords = ['＜④合計＞', '④合計', '合計金額', 'お支払い金額', '支払い金額']
        print(f"\n=== 特定キーワード周辺のテキスト ===")
        for keyword in keywords:
            pos = text_content.find(keyword)
            if pos != -1:
                # キーワード前後100文字を表示
                start = max(0, pos - 50)
                end = min(len(text_content), pos + len(keyword) + 100)
                context = text_content[start:end]
                print(f"キーワード '{keyword}' 周辺:")
                print(f"  {context}")
                print()
            else:
                print(f"キーワード '{keyword}' は見つかりませんでした")
        
    except ImportError:
        print("エラー: PyPDF2が必要です。pip install PyPDF2でインストールしてください")
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        import traceback
        traceback.print_exc()

def main():
    # デフォルトのテストファイル
    default_pdf = r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\2023\202306\OID_PAY_9ATI_202306.PDF"
    
    # コマンドライン引数から取得、なければデフォルトを使用
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
    else:
        pdf_path = default_pdf
    
    pdf_file = Path(pdf_path)
    if not pdf_file.exists():
        print(f"エラー: ファイルが存在しません: {pdf_path}")
        return
    
    debug_pdf_parsing(pdf_file)

if __name__ == "__main__":
    main()