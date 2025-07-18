#!/usr/bin/env python3
"""
シンプルなメールテストスクリプト
実際のメールの送信者と件名を確認
"""

import sys
from pathlib import Path
import imaplib
import email
from email.header import decode_header

# パッケージのパスを追加
sys.path.insert(0, str(Path(__file__).parent))

from line_fortune_processor.config import Config


def decode_mime_header(header):
    """MIMEヘッダーをデコード"""
    if header is None:
        return ""
    try:
        decoded = decode_header(header)
        result = []
        for text, charset in decoded:
            if isinstance(text, bytes):
                for encoding in [charset, 'utf-8', 'iso-2022-jp', 'shift_jis', 'euc-jp']:
                    if encoding:
                        try:
                            decoded_text = text.decode(encoding)
                            result.append(decoded_text)
                            break
                        except (UnicodeDecodeError, LookupError):
                            continue
                else:
                    result.append(text.decode('utf-8', errors='ignore'))
            else:
                result.append(text)
        return ''.join(result)
    except Exception as e:
        return str(header) if header else ""


def main():
    """メイン関数"""
    config = Config()
    email_config = config.get('email', {})
    
    print("=== シンプルメールテスト ===")
    print(f"期待する送信者: '{config.get('sender')}'")
    print(f"期待する件名: '{config.get('subject_pattern')}'")
    print()
    
    try:
        # IMAP接続
        connection = imaplib.IMAP4_SSL(email_config.get('server'), email_config.get('port'))
        connection.login(email_config.get('username'), email_config.get('password'))
        connection.select('INBOX')
        
        # 最新20件のメールを取得
        typ, data = connection.search(None, 'ALL')
        all_ids = data[0].split()
        recent_ids = all_ids[-20:] if len(all_ids) > 20 else all_ids
        
        print(f"最新20件のメール:")
        for i, email_id in enumerate(reversed(recent_ids), 1):
            try:
                typ, msg_data = connection.fetch(email_id, '(RFC822)')
                if typ == 'OK':
                    raw_email = msg_data[0][1]
                    email_message = email.message_from_bytes(raw_email)
                    
                    sender = decode_mime_header(email_message.get('From', ''))
                    subject = decode_mime_header(email_message.get('Subject', ''))
                    date = email_message.get('Date', '')
                    
                    print(f"{i:2d}. 送信者: {sender}")
                    print(f"    件名: {subject}")
                    print(f"    日付: {date}")
                    
                    # 条件チェック
                    expected_sender = config.get('sender')
                    expected_subject = config.get('subject_pattern')
                    
                    sender_match = expected_sender.lower() in sender.lower()
                    subject_match = expected_subject.lower() in subject.lower()
                    
                    if sender_match and subject_match:
                        print(f"    *** 条件に一致! ***")
                    elif sender_match:
                        print(f"    → 送信者は一致")
                    elif subject_match:
                        print(f"    → 件名は一致")
                    
                    print()
                    
            except Exception as e:
                print(f"メール処理エラー: {e}")
                continue
        
        # LineFortune関連のメールを専用検索
        print("LineFortune関連のメールを検索:")
        fortune_queries = [
            'SUBJECT "LineFortune"',
            'SUBJECT "Line Fortune"',
            'SUBJECT "LINE Fortune"',
            'SUBJECT "Fortune"',
            'FROM "linecorp.com"',
            'FROM "line.me"',
            'FROM "dev-line-fortune"'
        ]
        
        for query in fortune_queries:
            print(f"  クエリ: {query}")
            typ, data = connection.search(None, query)
            if typ == 'OK' and data[0]:
                ids = data[0].split()
                print(f"    結果: {len(ids)} 件")
                
                if ids:
                    # 最新の1件を表示
                    email_id = ids[-1]
                    typ, msg_data = connection.fetch(email_id, '(RFC822)')
                    if typ == 'OK':
                        raw_email = msg_data[0][1]
                        email_message = email.message_from_bytes(raw_email)
                        sender = decode_mime_header(email_message.get('From', ''))
                        subject = decode_mime_header(email_message.get('Subject', ''))
                        print(f"    例: {sender} | {subject}")
            else:
                print(f"    結果: 0 件")
            print()
        
        connection.close()
        connection.logout()
        
    except Exception as e:
        print(f"エラー: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()