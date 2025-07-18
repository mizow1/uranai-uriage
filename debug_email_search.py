#!/usr/bin/env python3
"""
メール検索デバッグスクリプト
実際のメールボックスの内容を確認するためのツール
"""

import sys
from pathlib import Path

# パッケージのパスを追加
sys.path.insert(0, str(Path(__file__).parent))

from line_fortune_processor.config import Config
from line_fortune_processor.email_processor import EmailProcessor
from line_fortune_processor.logger import get_logger
import imaplib
import email
from email.header import decode_header


def decode_mime_header(header):
    """MIMEヘッダーをデコード"""
    if header is None:
        return ""
    decoded = decode_header(header)
    return ''.join([
        text.decode(charset or 'utf-8') if isinstance(text, bytes) else text
        for text, charset in decoded
    ])


def search_all_emails(email_processor, limit=20):
    """最新のメールを検索"""
    if not email_processor.connection:
        return []
    
    try:
        # INBOXを選択
        email_processor.connection.select('INBOX')
        
        # 最新のメールを取得
        typ, data = email_processor.connection.search(None, 'ALL')
        
        if typ != 'OK':
            return []
        
        email_ids = data[0].split()
        # 最新のメールから取得
        recent_ids = email_ids[-limit:] if len(email_ids) > limit else email_ids
        
        emails = []
        for email_id in reversed(recent_ids):  # 新しい順
            try:
                typ, msg_data = email_processor.connection.fetch(email_id, '(RFC822)')
                if typ == 'OK':
                    raw_email = msg_data[0][1]
                    email_message = email.message_from_bytes(raw_email)
                    
                    emails.append({
                        'id': email_id.decode(),
                        'sender': decode_mime_header(email_message.get('From', '')),
                        'recipient': decode_mime_header(email_message.get('To', '')),
                        'subject': decode_mime_header(email_message.get('Subject', '')),
                        'date': email_message.get('Date', ''),
                    })
            except Exception as e:
                print(f"メール処理エラー: {e}")
                continue
        
        return emails
        
    except Exception as e:
        print(f"メール検索エラー: {e}")
        return []


def search_by_sender(email_processor, sender_pattern):
    """送信者で検索"""
    if not email_processor.connection:
        return []
    
    try:
        email_processor.connection.select('INBOX')
        typ, data = email_processor.connection.search(None, f'FROM "{sender_pattern}"')
        
        if typ != 'OK':
            return []
        
        email_ids = data[0].split()
        emails = []
        
        for email_id in email_ids:
            try:
                typ, msg_data = email_processor.connection.fetch(email_id, '(RFC822)')
                if typ == 'OK':
                    raw_email = msg_data[0][1]
                    email_message = email.message_from_bytes(raw_email)
                    
                    emails.append({
                        'id': email_id.decode(),
                        'sender': decode_mime_header(email_message.get('From', '')),
                        'subject': decode_mime_header(email_message.get('Subject', '')),
                        'date': email_message.get('Date', ''),
                    })
            except Exception as e:
                continue
        
        return emails
        
    except Exception as e:
        print(f"送信者検索エラー: {e}")
        return []


def main():
    """メイン関数"""
    try:
        # 設定の読み込み
        config = Config()
        logger = get_logger()
        
        # メール処理の初期化
        email_processor = EmailProcessor(config.get('email', {}))
        
        # メールサーバーに接続
        if not email_processor.connect():
            print("メールサーバーへの接続に失敗しました")
            return
        
        print("=== メール検索デバッグ ===")
        print(f"設定:")
        print(f"  送信者パターン: {config.get('sender')}")
        print(f"  受信者パターン: {config.get('recipient')}")
        print(f"  件名パターン: {config.get('subject_pattern')}")
        print()
        
        # 1. 最新のメールを表示
        print("1. 最新のメール20件:")
        recent_emails = search_all_emails(email_processor, 20)
        for i, email_info in enumerate(recent_emails, 1):
            print(f"  {i}. 送信者: {email_info['sender']}")
            print(f"     件名: {email_info['subject']}")
            print(f"     日付: {email_info['date']}")
            print()
        
        # 2. LINE関連のメールを検索
        print("2. LINE関連のメールを検索:")
        line_emails = search_by_sender(email_processor, "linecorp.com")
        if line_emails:
            for email_info in line_emails:
                print(f"  送信者: {email_info['sender']}")
                print(f"  件名: {email_info['subject']}")
                print(f"  日付: {email_info['date']}")
                print()
        else:
            print("  LINE関連のメールは見つかりませんでした")
        
        # 3. Fortune関連のメールを検索
        print("3. Fortune関連のメールを検索:")
        fortune_emails = search_by_sender(email_processor, "Fortune")
        if fortune_emails:
            for email_info in fortune_emails:
                print(f"  送信者: {email_info['sender']}")
                print(f"  件名: {email_info['subject']}")
                print(f"  日付: {email_info['date']}")
                print()
        else:
            print("  Fortune関連のメールは見つかりませんでした")
        
        # 4. 現在の設定での検索結果
        print("4. 現在の設定での検索結果:")
        matching_emails = email_processor.fetch_matching_emails(
            config.get('sender'),
            config.get('recipient'),
            config.get('subject_pattern')
        )
        
        if matching_emails:
            for email_info in matching_emails:
                print(f"  送信者: {email_info['sender']}")
                print(f"  件名: {email_info['subject']}")
                print(f"  日付: {email_info['date']}")
                print()
        else:
            print("  条件に一致するメールは見つかりませんでした")
        
        # メールサーバーから切断
        email_processor.disconnect()
        
    except Exception as e:
        print(f"エラーが発生しました: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()