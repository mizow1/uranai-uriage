"""
メール処理モジュール
"""

import imaplib
import email
import re
import logging
from typing import List, Dict, Any, Optional, Tuple
from datetime import datetime, date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import decode_header
import time


class EmailProcessor:
    """メール処理クラス"""
    
    def __init__(self, config: Dict[str, Any]):
        """
        メール処理を初期化
        
        Args:
            config: メール設定情報
        """
        self.config = config
        self.connection = None
        self.logger = logging.getLogger(__name__)
        
    def connect(self) -> bool:
        """
        メールサーバーに接続
        
        Returns:
            bool: 接続が成功した場合True
        """
        try:
            server = self.config.get('server', 'imap.gmail.com')
            port = self.config.get('port', 993)
            use_ssl = self.config.get('use_ssl', True)
            
            if use_ssl:
                self.connection = imaplib.IMAP4_SSL(server, port)
            else:
                self.connection = imaplib.IMAP4(server, port)
            
            username = self.config.get('username')
            password = self.config.get('password')
            
            if not username or not password:
                self.logger.error("メール認証情報が設定されていません")
                return False
                
            self.connection.login(username, password)
            self.logger.info(f"メールサーバーに接続しました: {server}:{port}")
            return True
            
        except Exception as e:
            self.logger.error(f"メールサーバーへの接続に失敗しました: {e}")
            return False
    
    def disconnect(self):
        """メールサーバーから切断"""
        if self.connection:
            try:
                self.connection.close()
                self.connection.logout()
                self.logger.info("メールサーバーから切断しました")
            except Exception as e:
                self.logger.warning(f"メールサーバーの切断中にエラーが発生しました: {e}")
            finally:
                self.connection = None
    
    def fetch_matching_emails(self, sender: str, recipient: str, subject_pattern: str, days_back: int = 30) -> List[Dict[str, Any]]:
        """
        条件に一致するメールを取得
        
        Args:
            sender: 送信者のメールアドレス
            recipient: 受信者のメールアドレス
            subject_pattern: 件名のパターン
            days_back: 検索対象の日数（デフォルト30日）
            
        Returns:
            List[Dict]: 一致するメールのリスト
        """
        if not self.connection:
            self.logger.error("メールサーバーに接続していません")
            return []
            
        try:
            # INBOXを選択
            self.connection.select('INBOX')
            
            # 日付範囲を計算
            from datetime import datetime, timedelta
            end_date = datetime.now()
            start_date = end_date - timedelta(days=days_back)
            
            # IMAP日付フォーマット
            start_date_str = start_date.strftime("%d-%b-%Y")
            end_date_str = end_date.strftime("%d-%b-%Y")
            
            self.logger.info(f"検索範囲: {start_date_str} から {end_date_str} ({days_back}日間)")
            
            # 日付範囲での検索
            date_query = f'SINCE "{start_date_str}"'
            
            matching_emails = []
            
            # 戦略1: 日付範囲 + 送信者で検索してから件名をフィルタリング
            if sender:
                self.logger.info(f"送信者での検索: {sender}")
                try:
                    # 送信者 + 日付範囲での検索
                    search_query = f'FROM "{sender}" {date_query}'
                    self.logger.info(f"検索クエリ: {search_query}")
                    
                    typ, data = self.connection.search(None, search_query)
                    self.logger.info(f"送信者+日付検索結果: typ={typ}, data={data}")
                    
                    if typ == 'OK' and data[0]:
                        email_ids = data[0].split()
                        self.logger.info(f"送信者+日付で見つかったメールID数: {len(email_ids)}")
                        
                        # 件名でフィルタリング（最新50件まで）
                        for email_id in reversed(email_ids[-50:]):  # 最新50件のみ処理
                            try:
                                typ, msg_data = self.connection.fetch(email_id, '(RFC822)')
                                if typ == 'OK':
                                    raw_email = msg_data[0][1]
                                    email_message = email.message_from_bytes(raw_email)
                                    
                                    email_info = self._extract_email_info(email_message)
                                    email_info['id'] = email_id.decode()
                                    
                                    # 件名のフィルタリング
                                    subject = email_info.get('subject', '')
                                    if not subject_pattern or subject_pattern.lower() in subject.lower():
                                        matching_emails.append(email_info)
                                        self.logger.info(f"一致するメール: {subject}")
                                    else:
                                        self.logger.debug(f"件名が一致しないメール: {subject}")
                                        
                            except Exception as e:
                                self.logger.warning(f"メールID {email_id} の処理中にエラーが発生しました: {e}")
                                continue
                                
                except Exception as e:
                    self.logger.error(f"送信者検索中にエラーが発生しました: {e}")
            
            # 戦略2: 日付範囲 + 件名で検索してから送信者をフィルタリング
            if not matching_emails and subject_pattern:
                self.logger.info(f"件名での検索: {subject_pattern}")
                try:
                    # 件名 + 日付範囲での検索
                    search_query = f'SUBJECT "{subject_pattern}" {date_query}'
                    self.logger.info(f"検索クエリ: {search_query}")
                    
                    typ, data = self.connection.search(None, search_query)
                    
                    if typ == 'OK' and data[0]:
                        email_ids = data[0].split()
                        self.logger.info(f"件名+日付検索結果: {len(email_ids)} 件")
                        
                        # 送信者でフィルタリング（最新50件まで）
                        for email_id in reversed(email_ids[-50:]):  # 最新50件のみ処理
                            try:
                                typ, msg_data = self.connection.fetch(email_id, '(RFC822)')
                                if typ == 'OK':
                                    raw_email = msg_data[0][1]
                                    email_message = email.message_from_bytes(raw_email)
                                    
                                    email_info = self._extract_email_info(email_message)
                                    email_info['id'] = email_id.decode()
                                    
                                    # 送信者のフィルタリング
                                    if not sender or sender.lower() in email_info.get('sender', '').lower():
                                        matching_emails.append(email_info)
                                        self.logger.info(f"一致するメール: {email_info.get('subject', '')}")
                                        
                            except Exception as e:
                                self.logger.warning(f"メールID {email_id} の処理中にエラーが発生しました: {e}")
                                continue
                                
                except Exception as e:
                    self.logger.error(f"件名検索中にエラーが発生しました: {e}")
            
            # 戦略3: 日付範囲のみで検索してから条件をフィルタリング
            if not matching_emails:
                self.logger.info(f"日付範囲での全メール検索: {date_query}")
                try:
                    typ, data = self.connection.search(None, date_query)
                    if typ == 'OK' and data[0]:
                        all_ids = data[0].split()
                        recent_ids = all_ids[-100:] if len(all_ids) > 100 else all_ids
                        
                        self.logger.info(f"日付範囲内の最新 {len(recent_ids)} 件のメールから条件に一致するものを検索")
                        
                        for email_id in reversed(recent_ids):  # 新しい順
                            try:
                                typ, msg_data = self.connection.fetch(email_id, '(RFC822)')
                                if typ == 'OK':
                                    raw_email = msg_data[0][1]
                                    email_message = email.message_from_bytes(raw_email)
                                    
                                    email_info = self._extract_email_info(email_message)
                                    email_info['id'] = email_id.decode()
                                    
                                    # 条件チェック（デバッグ情報付き）
                                    actual_sender = email_info.get('sender', '')
                                    actual_subject = email_info.get('subject', '')
                                    
                                    sender_match = not sender or sender.lower() in actual_sender.lower()
                                    subject_match = not subject_pattern or subject_pattern.lower() in actual_subject.lower()
                                    
                                    # デバッグ情報を出力
                                    self.logger.debug(f"チェック中のメール:")
                                    self.logger.debug(f"  実際の送信者: '{actual_sender}'")
                                    self.logger.debug(f"  期待する送信者: '{sender}'")
                                    self.logger.debug(f"  送信者マッチ: {sender_match}")
                                    self.logger.debug(f"  実際の件名: '{actual_subject}'")
                                    self.logger.debug(f"  期待する件名: '{subject_pattern}'")
                                    self.logger.debug(f"  件名マッチ: {subject_match}")
                                    
                                    if sender_match and subject_match:
                                        matching_emails.append(email_info)
                                        self.logger.info(f"一致するメール発見: {actual_subject}")
                                    else:
                                        # マッチしない理由を記録
                                        if not sender_match:
                                            self.logger.debug(f"送信者が一致しません: '{actual_sender}' に '{sender}' が含まれていません")
                                        if not subject_match:
                                            self.logger.debug(f"件名が一致しません: '{actual_subject}' に '{subject_pattern}' が含まれていません")
                                        
                            except Exception as e:
                                self.logger.warning(f"メールID {email_id} の処理中にエラーが発生しました: {e}")
                                continue
                                
                except Exception as e:
                    self.logger.error(f"日付範囲検索中にエラーが発生しました: {e}")
            
            # 重複を除去
            unique_emails = []
            seen_ids = set()
            for email_info in matching_emails:
                if email_info['id'] not in seen_ids:
                    unique_emails.append(email_info)
                    seen_ids.add(email_info['id'])
            
            self.logger.info(f"条件に一致するメールを {len(unique_emails)} 件見つけました")
            return unique_emails
            
        except Exception as e:
            self.logger.error(f"メール取得中にエラーが発生しました: {e}")
            return []
    
    def _extract_email_info(self, email_message) -> Dict[str, Any]:
        """メールから情報を抽出"""
        def decode_mime_header(header):
            """MIMEヘッダーをデコード（エンコーディング対応強化）"""
            if header is None:
                return ""
            try:
                decoded = decode_header(header)
                result = []
                for text, charset in decoded:
                    if isinstance(text, bytes):
                        # 複数のエンコーディングを試す
                        for encoding in [charset, 'utf-8', 'iso-2022-jp', 'shift_jis', 'euc-jp']:
                            if encoding:
                                try:
                                    decoded_text = text.decode(encoding)
                                    result.append(decoded_text)
                                    break
                                except (UnicodeDecodeError, LookupError):
                                    continue
                        else:
                            # どのエンコーディングでも失敗した場合は、エラーを無視して強制デコード
                            result.append(text.decode('utf-8', errors='ignore'))
                    else:
                        result.append(text)
                return ''.join(result)
            except Exception as e:
                self.logger.warning(f"ヘッダーデコードエラー: {e}")
                return str(header) if header else ""
        
        return {
            'sender': decode_mime_header(email_message.get('From', '')),
            'recipient': decode_mime_header(email_message.get('To', '')),
            'subject': decode_mime_header(email_message.get('Subject', '')),
            'date': email_message.get('Date', ''),
            'message': email_message
        }
    
    def extract_attachments(self, email_info: Dict[str, Any], file_type: str = ".csv") -> List[Dict[str, Any]]:
        """
        メールから添付ファイルを抽出
        
        Args:
            email_info: メール情報
            file_type: 抽出するファイルタイプ
            
        Returns:
            List[Dict]: 添付ファイルのリスト
        """
        attachments = []
        email_message = email_info.get('message')
        
        if not email_message:
            return attachments
            
        try:
            for part in email_message.walk():
                if part.get_content_disposition() == 'attachment':
                    filename = part.get_filename()
                    if filename and filename.lower().endswith(file_type.lower()):
                        content = part.get_payload(decode=True)
                        if content:
                            attachments.append({
                                'filename': filename,
                                'content': content,
                                'content_type': part.get_content_type()
                            })
                            
            self.logger.info(f"添付ファイルを {len(attachments)} 件抽出しました")
            
        except Exception as e:
            self.logger.error(f"添付ファイル抽出中にエラーが発生しました: {e}")
            
        return attachments
    
    def extract_date_from_subject(self, subject: str) -> Optional[date]:
        """
        件名から日付を抽出
        
        Args:
            subject: メール件名
            
        Returns:
            date: 抽出された日付、見つからない場合はNone
        """
        try:
            # "LineFortune Daily Report for yyyy-mm-dd" 形式から日付を抽出
            date_pattern = r'(\d{4})-(\d{2})-(\d{2})'
            match = re.search(date_pattern, subject)
            
            if match:
                year, month, day = match.groups()
                extracted_date = date(int(year), int(month), int(day))
                self.logger.info(f"件名から日付を抽出しました: {extracted_date}")
                return extracted_date
            else:
                self.logger.warning(f"件名から日付を抽出できませんでした: {subject}")
                return None
                
        except Exception as e:
            self.logger.error(f"日付抽出中にエラーが発生しました: {e}")
            return None
    
    def mark_as_read(self, email_id: str) -> bool:
        """
        メールを既読にマーク
        
        Args:
            email_id: メールID
            
        Returns:
            bool: 成功した場合True
        """
        try:
            if not self.connection:
                return False
                
            self.connection.store(email_id, '+FLAGS', '\\Seen')
            return True
            
        except Exception as e:
            self.logger.error(f"メール既読マーク中にエラーが発生しました: {e}")
            return False