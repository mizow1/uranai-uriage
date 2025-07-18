#!/usr/bin/env python3
"""
メール検索のテストスクリプト
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
    decoded = decode_header(header)
    return ''.join([
        text.decode(charset or 'utf-8') if isinstance(text, bytes) else text
        for text, charset in decoded
    ])


def test_imap_search():
    """IMAP検索のテスト"""
    config = Config()
    email_config = config.get('email', {})
    
    print("=== IMAP検索テスト ===")
    print(f"サーバー: {email_config.get('server')}")
    print(f"ユーザー: {email_config.get('username')}")
    print()
    
    try:
        # IMAP接続
        connection = imaplib.IMAP4_SSL(email_config.get('server'), email_config.get('port'))
        connection.login(email_config.get('username'), email_config.get('password'))
        connection.select('INBOX')
        
        print("1. 全メール数を確認")
        typ, data = connection.search(None, 'ALL')
        all_ids = data[0].split()
        print(f"   全メール数: {len(all_ids)}")
        
        print("2. 最新10件のメール")
        recent_ids = all_ids[-10:] if len(all_ids) > 10 else all_ids
        for i, email_id in enumerate(reversed(recent_ids), 1):
            typ, msg_data = connection.fetch(email_id, '(ENVELOPE)')
            if typ == 'OK':
                typ, msg_data = connection.fetch(email_id, '(RFC822)')
                if typ == 'OK':
                    raw_email = msg_data[0][1]
                    email_message = email.message_from_bytes(raw_email)
                    sender = decode_mime_header(email_message.get('From', ''))
                    subject = decode_mime_header(email_message.get('Subject', ''))
                    print(f"   {i}. {sender[:50]}... | {subject[:50]}...")
        
        print()
        print("3. 設定された条件での検索")
        sender = config.get('sender')
        subject_pattern = config.get('subject_pattern')
        
        # 個別に検索してみる
        print(f"   送信者検索: FROM \"{sender}\"")
        typ, data = connection.search(None, f'FROM "{sender}"')
        sender_ids = data[0].split()
        print(f"   結果: {len(sender_ids)} 件")
        
        print(f"   件名検索: SUBJECT \"{subject_pattern}\"")
        typ, data = connection.search(None, f'SUBJECT "{subject_pattern}"')
        subject_ids = data[0].split()
        print(f"   結果: {len(subject_ids)} 件")
        
        # 組み合わせ検索
        combined_query = f'FROM "{sender}" SUBJECT "{subject_pattern}"'
        print(f"   組み合わせ検索: {combined_query}")
        typ, data = connection.search(None, combined_query)
        combined_ids = data[0].split()
        print(f"   結果: {len(combined_ids)} 件")
        
        # 見つかったメールの詳細を表示
        if combined_ids:
            print("   見つかったメールの詳細:")
            for email_id in combined_ids[-5:]:  # 最新5件
                typ, msg_data = connection.fetch(email_id, '(RFC822)')
                if typ == 'OK':
                    raw_email = msg_data[0][1]
                    email_message = email.message_from_bytes(raw_email)
                    sender = decode_mime_header(email_message.get('From', ''))
                    subject = decode_mime_header(email_message.get('Subject', ''))
                    date = email_message.get('Date', '')
                    print(f"     ID: {email_id.decode()}")
                    print(f"     送信者: {sender}")
                    print(f"     件名: {subject}")
                    print(f"     日付: {date}")
                    print()
        
        print("4. LineFortune関連のメール検索")
        fortune_queries = [
            'SUBJECT "LineFortune"',
            'SUBJECT "Line Fortune"',
            'SUBJECT "LINE Fortune"',
            'SUBJECT "Fortune"',
            'FROM "linecorp.com"',
            'FROM "line.me"'
        ]
        
        for query in fortune_queries:
            print(f"   クエリ: {query}")
            typ, data = connection.search(None, query)
            if typ == 'OK':
                ids = data[0].split()
                print(f"   結果: {len(ids)} 件")
                
                if ids:
                    # 最新の1件を表示
                    email_id = ids[-1]
                    typ, msg_data = connection.fetch(email_id, '(RFC822)')
                    if typ == 'OK':
                        raw_email = msg_data[0][1]
                        email_message = email.message_from_bytes(raw_email)
                        sender = decode_mime_header(email_message.get('From', ''))
                        subject = decode_mime_header(email_message.get('Subject', ''))
                        print(f"     最新メール例: {sender} | {subject}")
            print()
        
        connection.close()
        connection.logout()
        
    except Exception as e:
        print(f"エラー: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    test_imap_search()