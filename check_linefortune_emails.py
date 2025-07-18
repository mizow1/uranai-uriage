#!/usr/bin/env python3
"""
LineFortune関連メールの送信者を確認するスクリプト
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
    
    print("=== LineFortune メール送信者確認 ===")
    
    try:
        # IMAP接続
        connection = imaplib.IMAP4_SSL(email_config.get('server'), email_config.get('port'))
        connection.login(email_config.get('username'), email_config.get('password'))
        connection.select('INBOX')
        
        # LineFortune関連のメールを検索
        print("SUBJECT 'LineFortune' で検索...")
        typ, data = connection.search(None, 'SUBJECT "LineFortune"')
        
        if typ == 'OK' and data[0]:
            ids = data[0].split()
            print(f"見つかったメール数: {len(ids)} 件")
            
            # 最新10件を詳細確認
            latest_ids = ids[-10:] if len(ids) > 10 else ids
            print(f"\n最新{len(latest_ids)}件の詳細:")
            
            for i, email_id in enumerate(reversed(latest_ids), 1):
                try:
                    typ, msg_data = connection.fetch(email_id, '(RFC822)')
                    if typ == 'OK':
                        raw_email = msg_data[0][1]
                        email_message = email.message_from_bytes(raw_email)
                        
                        sender = decode_mime_header(email_message.get('From', ''))
                        subject = decode_mime_header(email_message.get('Subject', ''))
                        date = email_message.get('Date', '')
                        
                        print(f"\n{i}. 送信者: {sender}")
                        print(f"   件名: {subject}")
                        print(f"   日付: {date}")
                        
                        # 転送メールかどうかを判定
                        if subject.startswith('Fwd:') or subject.startswith('転送:'):
                            print("   → 転送メール")
                        else:
                            print("   → 直接メール")
                            
                except Exception as e:
                    print(f"メール処理エラー: {e}")
                    continue
            
            # 送信者の統計を取る
            print(f"\n送信者の統計（最新100件）:")
            sender_stats = {}
            recent_ids = ids[-100:] if len(ids) > 100 else ids
            
            for email_id in recent_ids:
                try:
                    typ, msg_data = connection.fetch(email_id, '(RFC822)')
                    if typ == 'OK':
                        raw_email = msg_data[0][1]
                        email_message = email.message_from_bytes(raw_email)
                        sender = decode_mime_header(email_message.get('From', ''))
                        
                        # メールアドレス部分を抽出
                        if '<' in sender and '>' in sender:
                            email_part = sender.split('<')[1].split('>')[0]
                        else:
                            email_part = sender
                            
                        sender_stats[email_part] = sender_stats.get(email_part, 0) + 1
                        
                except Exception as e:
                    continue
            
            # 送信者別の件数を表示
            for sender, count in sorted(sender_stats.items(), key=lambda x: x[1], reverse=True):
                print(f"  {sender}: {count} 件")
                
        else:
            print("LineFortune関連のメールが見つかりませんでした")
        
        connection.close()
        connection.logout()
        
    except Exception as e:
        print(f"エラー: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()