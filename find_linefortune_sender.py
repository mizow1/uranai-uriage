#!/usr/bin/env python3
"""
LineFortune関連メールから実際の送信者を調べるスクリプト
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
    
    print("=== LineFortune メール送信者特定 ===")
    
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
            
            # "LineFortune Daily Report" を含むメールを探す
            daily_report_emails = []
            
            print("\n'LineFortune Daily Report' を含むメールを検索中...")
            
            # 最新100件をチェック
            recent_ids = ids[-100:] if len(ids) > 100 else ids
            
            for email_id in reversed(recent_ids):
                try:
                    typ, msg_data = connection.fetch(email_id, '(RFC822)')
                    if typ == 'OK':
                        raw_email = msg_data[0][1]
                        email_message = email.message_from_bytes(raw_email)
                        
                        sender = decode_mime_header(email_message.get('From', ''))
                        subject = decode_mime_header(email_message.get('Subject', ''))
                        date = email_message.get('Date', '')
                        
                        # "LineFortune Daily Report" を含むかチェック
                        if "LineFortune Daily Report" in subject:
                            daily_report_emails.append({
                                'id': email_id,
                                'sender': sender,
                                'subject': subject,
                                'date': date
                            })
                            
                            if len(daily_report_emails) >= 10:  # 最新10件で十分
                                break
                        
                except Exception as e:
                    print(f"メール処理エラー: {e}")
                    continue
            
            print(f"\n'LineFortune Daily Report' を含むメール: {len(daily_report_emails)} 件")
            
            if daily_report_emails:
                print("\n最新の Daily Report メール:")
                for i, email_info in enumerate(daily_report_emails, 1):
                    print(f"\n{i}. 送信者: {email_info['sender']}")
                    print(f"   件名: {email_info['subject']}")
                    print(f"   日付: {email_info['date']}")
                    
                    # dev-line-fortune@linecorp.com を含むかチェック
                    if "dev-line-fortune@linecorp.com" in email_info['sender']:
                        print("   ✓ 期待する送信者と一致!")
                    else:
                        print("   ✗ 期待する送信者と異なる")
                        
                    # 転送メールかチェック
                    if email_info['subject'].startswith(('Fwd:', '転送:', 'Re:')):
                        print("   → 転送/返信メール")
                    else:
                        print("   → 直接メール")
            else:
                print("\n'LineFortune Daily Report' を含むメールが見つかりませんでした")
                
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