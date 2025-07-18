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
    
    def fetch_matching_emails(self, sender: str, recipient: str, subject_pattern: str) -> List[Dict[str, Any]]:
        """
        条件に一致するメールを取得
        
        Args:
            sender: 送信者のメールアドレス
            recipient: 受信者のメールアドレス
            subject_pattern: 件名のパターン
            
        Returns:
            List[Dict]: 一致するメールのリスト
        """
        if not self.connection:
            self.logger.error("メールサーバーに接続していません")
            return []
            
        try:
            # INBOXを選択
            self.connection.select('INBOX')
            
            # 検索クエリを構築
            search_criteria = []
            if sender:
                search_criteria.append(f'FROM "{sender}"')
            if subject_pattern:
                # 完全一致ではなく部分一致で検索
                search_criteria.append(f'SUBJECT "{subject_pattern}"')
            # 受信者での検索は除外（自分のメールボックスなので不要）
            # if recipient:
            #     search_criteria.append(f'TO "{recipient}"')
                
            search_query = ' '.join(search_criteria)
            
            self.logger.info(f"メール検索クエリ: {search_query}")
            
            # 未読メールを検索
            typ, data = self.connection.search(None, search_query)
            
            if typ != 'OK':
                self.logger.error("メール検索に失敗しました")
                return []
                
            email_ids = data[0].split()
            matching_emails = []
            
            for email_id in email_ids:
                try:
                    # メールを取得
                    typ, msg_data = self.connection.fetch(email_id, '(RFC822)')
                    if typ == 'OK':
                        raw_email = msg_data[0][1]
                        email_message = email.message_from_bytes(raw_email)
                        
                        # メール情報を抽出
                        email_info = self._extract_email_info(email_message)
                        email_info['id'] = email_id.decode()
                        
                        matching_emails.append(email_info)
                        
                except Exception as e:
                    self.logger.warning(f"メールID {email_id} の処理中にエラーが発生しました: {e}")
                    continue
                    
            self.logger.info(f"条件に一致するメールを {len(matching_emails)} 件見つけました")
            return matching_emails
            
        except Exception as e:
            self.logger.error(f"メール取得中にエラーが発生しました: {e}")
            return []
    
    def _extract_email_info(self, email_message) -> Dict[str, Any]:
        """メールから情報を抽出"""
        def decode_mime_header(header):
            """MIMEヘッダーをデコード"""
            if header is None:
                return ""
            decoded = decode_header(header)
            return ''.join([
                text.decode(charset or 'utf-8') if isinstance(text, bytes) else text
                for text, charset in decoded
            ])
        
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