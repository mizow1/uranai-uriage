#!/usr/bin/env python3
"""
auファイルの詳細確認
"""

try:
    import PyPDF2
    import re
    from pathlib import Path

    def test_au_detailed():
        test_file = Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\2023\202303\202303cp02お支払い明細書.pdf")
        
        print("=== auファイル詳細解析 ===")
        print(f"ファイル: {test_file}")
        
        try:
            with open(test_file, 'rb') as pdf_file:
                pdf_reader = PyPDF2.PdfReader(pdf_file)
                print(f"ページ数: {len(pdf_reader.pages)}")
                
                all_text = ""
                for i, page in enumerate(pdf_reader.pages):
                    try:
                        text = page.extract_text()
                        all_text += text
                        print(f"\nページ{i+1}のテキスト: {len(text)}文字")
                        if text:
                            # 数字を含む行のみ表示
                            lines = text.split('\n')
                            for line in lines:
                                if re.search(r'\d+', line):
                                    print(f"  数字含む行: {line[:100]}")
                    except Exception as e:
                        print(f"ページ{i+1}の読み込みエラー: {e}")
                
                # 全ての数字パターンを抽出
                print(f"\n=== 全数字パターン抽出 ===")
                patterns = [
                    r'(\d{4,})',  # 4桁以上の数字
                    r'(\d{1,3}(?:,\d{3})+)',  # カンマ区切り
                    r'(\d+\.\d+)',  # 小数点
                    r'(\d+)',  # 全ての数字
                ]
                
                for pattern in patterns:
                    matches = re.findall(pattern, all_text)
                    valid_numbers = []
                    for match in matches:
                        try:
                            num = float(match.replace(',', ''))
                            if num > 50:  # 50以上の数字
                                valid_numbers.append((match, num))
                        except ValueError:
                            continue
                    
                    if valid_numbers:
                        print(f"パターン {pattern}: {len(valid_numbers)}件")
                        # 最大10件まで表示
                        for text, num in sorted(valid_numbers, key=lambda x: x[1], reverse=True)[:10]:
                            print(f"  {text} -> {num}")
                        break
                
        except Exception as e:
            print(f"エラー: {str(e)}")
            
            # 代替手段でファイル存在確認
            if test_file.exists():
                print(f"ファイルは存在します。サイズ: {test_file.stat().st_size} bytes")
            else:
                print("ファイルが存在しません")

    if __name__ == "__main__":
        test_au_detailed()
        
except ImportError as e:
    print(f"必要なライブラリが不足: {e}")
except Exception as e:
    print(f"予期しないエラー: {e}")