#!/usr/bin/env python3
"""
LINEフォルダの内容を確認するスクリプト
"""

import sys
from pathlib import Path
import imaplib
import email
from email.header import decode_header
from datetime import datetime, timedelta

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
    
    print("=== LINEフォルダ内容確認 ===")
    
    try:
        # IMAP接続
        connection = imaplib.IMAP4_SSL(email_config.get('server'), email_config.get('port'))
        connection.login(email_config.get('username'), email_config.get('password'))
        
        # LINEフォルダを検索
        line_folder = 'LINE/LINE&,wiB6lLV,wk-'
        
        print(f"フォルダ '{line_folder}' を選択中...")
        try:
            connection.select(f'"{line_folder}"')
            
            # 最新1週間のメールを検索
            end_date = datetime.now()
            start_date = end_date - timedelta(days=7)
            start_date_str = start_date.strftime("%d-%b-%Y")
            
            print(f"検索期間: {start_date_str} から現在まで")
            
            # 日付範囲で検索
            typ, data = connection.search(None, f'SINCE "{start_date_str}"')
            
            if typ == 'OK' and data[0]:
                email_ids = data[0].split()
                print(f"見つかったメール数: {len(email_ids)} 件")
                
                # 最新20件を表示
                recent_ids = email_ids[-20:] if len(email_ids) > 20 else email_ids
                
                print(f"\n最新{len(recent_ids)}件のメール:")
                for i, email_id in enumerate(reversed(recent_ids), 1):
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
                            
                            # LineFortune Daily Report を含むかチェック
                            if "LineFortune Daily Report" in subject:
                                print("   ✓ LineFortune Daily Report を含む!")
                            elif "LineFortune" in subject:
                                print("   → LineFortune を含む")
                            
                            # 設定されている送信者と一致するかチェック
                            expected_sender = config.get('sender')
                            if expected_sender.lower() in sender.lower():
                                print(f"   ✓ 期待する送信者と一致: {expected_sender}")
                            else:
                                print(f"   ✗ 期待する送信者と異なる: {expected_sender}")
                                
                    except Exception as e:
                        print(f"メール処理エラー: {e}")
                        continue
                        
                # LineFortune Daily Report を含むメールの統計
                print(f"\n=== LineFortune Daily Report 統計 ===")
                fortune_count = 0
                senders = {}
                
                for email_id in email_ids:
                    try:
                        typ, msg_data = connection.fetch(email_id, '(RFC822)')
                        if typ == 'OK':
                            raw_email = msg_data[0][1]
                            email_message = email.message_from_bytes(raw_email)
                            
                            sender = decode_mime_header(email_message.get('From', ''))
                            subject = decode_mime_header(email_message.get('Subject', ''))
                            
                            if "LineFortune Daily Report" in subject:
                                fortune_count += 1
                                senders[sender] = senders.get(sender, 0) + 1
                                
                    except Exception as e:
                        continue
                        
                print(f"LineFortune Daily Report を含むメール: {fortune_count} 件")
                print("送信者別統計:")
                for sender, count in sorted(senders.items(), key=lambda x: x[1], reverse=True):
                    print(f"  {sender}: {count} 件")
                    
            else:
                print("指定期間内にメールが見つかりませんでした")
                
        except Exception as e:
            print(f"フォルダ選択エラー: {e}")
            
        connection.close()
        connection.logout()
        
    except Exception as e:
        print(f"エラー: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()