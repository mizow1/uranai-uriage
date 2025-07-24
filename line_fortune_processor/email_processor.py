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
try:
    from tqdm import tqdm
except ImportError:
    # tqdmが利用できない場合のダミークラス
    class tqdm:
        def __init__(self, iterable=None, total=None, desc=None, **kwargs):
            self.iterable = iterable or []
            self.total = total or 0
            self.desc = desc
            self.current = 0
            if desc:
                print(f"{desc}: 開始")
        
        def __iter__(self):
            return iter(self.iterable)
        
        def __enter__(self):
            return self
        
        def __exit__(self, *args):
            if self.desc:
                print(f"{self.desc}: 完了")
        
        def update(self, n=1):
            self.current += n
            if self.desc and self.total > 0:
                progress = (self.current / self.total) * 100
                print(f"{self.desc}: {self.current}/{self.total} ({progress:.1f}%)")
        
        def set_description(self, desc):
            self.desc = desc
            print(f"{desc}")

from .error_handler import retry_on_error, handle_errors, ErrorHandler, RetryableError, FatalError, ErrorType
from .constants import MailConstants, ErrorConstants
from .messages import MessageFormatter


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
        self.error_handler = ErrorHandler(self.logger)
        
    @retry_on_error(max_retries=3, base_delay=2.0)
    def connect(self) -> bool:
        """
        メールサーバーに接続
        
        Returns:
            bool: 接続が成功した場合True
        """
        server = self.config.get('server', 'imap.gmail.com')
        port = self.config.get('port', 993)
        use_ssl = self.config.get('use_ssl', True)
        
        username = self.config.get('username')
        password = self.config.get('password')
        
        if not username or not password:
            raise FatalError("メール認証情報が設定されていません", ErrorType.AUTHENTICATION)
        
        try:
            if use_ssl:
                self.connection = imaplib.IMAP4_SSL(server, port)
            else:
                self.connection = imaplib.IMAP4(server, port)
            
            self.connection.login(username, password)
            self.logger.info(MessageFormatter.get_email_message(
                "connection_success", server=server, port=port
            ))
            return True
            
        except imaplib.IMAP4.error as e:
            if 'authentication' in str(e).lower():
                raise FatalError(f"認証に失敗しました: {e}", ErrorType.AUTHENTICATION, e)
            else:
                raise RetryableError(f"IMAPエラー: {e}", ErrorType.NETWORK, e)
        except (ConnectionError, OSError) as e:
            raise RetryableError(f"ネットワークエラー: {e}", ErrorType.NETWORK, e)
        except Exception as e:
            error_type = self.error_handler.classify_error(e)
            if self.error_handler.is_retryable(e, error_type):
                raise RetryableError(f"メールサーバーへの接続に失敗: {e}", error_type, e)
            else:
                raise FatalError(f"メールサーバーへの接続に失敗: {e}", error_type, e)
    
    def disconnect(self):
        """メールサーバーから切断"""
        if self.connection:
            try:
                import concurrent.futures
                import threading
                
                # タイムアウト付きで切断処理を実行
                def disconnect_impl():
                    self.connection.close()
                    self.connection.logout()
                
                # ThreadPoolExecutorでタイムアウト処理
                with concurrent.futures.ThreadPoolExecutor() as executor:
                    future = executor.submit(disconnect_impl)
                    try:
                        future.result(timeout=30)  # 30秒でタイムアウト
                        self.logger.info("メールサーバーから切断しました")
                    except concurrent.futures.TimeoutError:
                        self.logger.warning("メールサーバー切断がタイムアウトしました")
                        future.cancel()
                    
            except Exception as e:
                self.logger.warning(f"メールサーバーの切断中にエラーが発生しました: {e}")
            finally:
                self.connection = None
    
    @handle_errors("メール検索")
    def fetch_matching_emails(self, sender: str, recipient: str, subject_pattern: str, start_date: date = None, end_date: date = None, days_back: int = 30) -> List[Dict[str, Any]]:
        """
        条件に一致するメールを取得
        
        Args:
            sender: 送信者のメールアドレス
            recipient: 受信者のメールアドレス
            subject_pattern: 件名のパターン
            start_date: 検索開始日
            end_date: 検索終了日
            days_back: 検索対象の日数（start_date/end_dateが指定されていない場合のみ使用）
            
        Returns:
            List[Dict]: 一致するメールのリスト
        """
        if not self.connection:
            raise FatalError("メールサーバーに接続していません", ErrorType.NETWORK)
            
        try:
            # 全てのメールフォルダを検索
            matching_emails = []
            
            # 利用可能なフォルダを取得
            typ, folders = self.connection.list()
            folder_names = []
            
            if typ == 'OK':
                for folder in folders:
                    folder_str = folder.decode('utf-8')
                    # フォルダ名を抽出
                    folder_name = folder_str.split('"')[-2] if '"' in folder_str else folder_str.split()[-1]
                    folder_names.append(folder_name)
                    
                self.logger.info(f"利用可能なフォルダ: {folder_names}")
            
            # 主要なフォルダを優先的に検索（LINEラベルを最優先）
            priority_folders = [
                'LINE/LINE&,wiB6lLV,wk-',  # LINE（自動）のエンコード済み名前
                'LINE（自動）',
                'LINE (自動)',
                'LINE',
                'INBOX', 
                '[Gmail]/All Mail', 
                '[Gmail]/すべてのメール', 
                'All Mail'
            ]
            search_folders = []
            
            for folder in priority_folders:
                if folder in folder_names:
                    search_folders.append(folder)
            
            # 他のフォルダも追加（システムフォルダは除外）
            for folder in folder_names:
                if folder not in search_folders and not folder.startswith('[Gmail]/'):
                    search_folders.append(folder)
            
            self.logger.info(f"検索対象フォルダ: {search_folders}")
            
            # 各フォルダで検索（進捗表示付き）
            search_folders_limited = search_folders[:5]  # 最大5フォルダまで
            with tqdm(total=len(search_folders_limited), desc="フォルダ検索進捗", unit="フォルダ") as folder_pbar:
                for folder in search_folders_limited:
                    try:
                        folder_pbar.set_description(f"検索中: {folder}")
                        self.logger.info(f"フォルダ '{folder}' を検索中...")
                        self.connection.select(f'"{folder}"')
                        
                        folder_emails = self._search_in_current_folder(sender, recipient, subject_pattern, days_back, start_date, end_date)
                        matching_emails.extend(folder_emails)
                        
                        if matching_emails:
                            self.logger.info(f"フォルダ '{folder}' で {len(folder_emails)} 件見つかりました")
                            folder_pbar.update(1)
                            break  # 見つかったら他のフォルダは検索しない
                            
                    except Exception as e:
                        self.logger.warning(f"フォルダ '{folder}' の検索中にエラーが発生しました: {e}")
                    finally:
                        folder_pbar.update(1)
                    
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
    
    def _search_in_current_folder(self, sender: str, recipient: str, subject_pattern: str, days_back: int, start_date: date = None, end_date: date = None) -> List[Dict[str, Any]]:
        """現在選択されているフォルダ内で検索"""
        try:
            date_query = self._build_date_query(start_date, end_date, days_back)
            
            # 最適化された検索戦略を順次実行
            search_strategies = [
                lambda: self._search_by_sender_and_date(sender, date_query, subject_pattern),
                lambda: self._search_by_subject_and_date(subject_pattern, date_query, sender),
                lambda: self._search_by_date_only(date_query, sender, subject_pattern)
            ]
            
            for strategy in search_strategies:
                try:
                    emails = strategy()
                    if emails:
                        return self._remove_duplicates(emails)
                except Exception as e:
                    self.logger.warning(f"検索戦略の実行中にエラーが発生しました: {e}")
                    continue
            
            self.logger.info("条件に一致するメールが見つかりませんでした")
            return []
            
        except Exception as e:
            self.logger.error(f"フォルダ内検索中にエラーが発生しました: {e}")
            return []
    
    def _build_date_query(self, start_date: date, end_date: date, days_back: int) -> str:
        """日付検索クエリを構築"""
        from datetime import datetime, timedelta
        
        if start_date and end_date:
            query_start_date = datetime.combine(start_date, datetime.min.time())
            # 終了日を含めるため、翌日の00:00:00を設定
            query_end_date = datetime.combine(end_date, datetime.min.time()) + timedelta(days=1)
        else:
            query_end_date = datetime.now()
            query_start_date = query_end_date - timedelta(days=days_back)
        
        start_date_str = query_start_date.strftime("%d-%b-%Y")
        end_date_str = query_end_date.strftime("%d-%b-%Y")
        
        query = f'SINCE "{start_date_str}" BEFORE "{end_date_str}"'
        self.logger.info(f"生成された日付クエリ: {query}")
        self.logger.info(f"指定された日付範囲: {start_date} ～ {end_date}")
        self.logger.info(f"クエリの実際の日付範囲: {query_start_date} ～ {query_end_date}")
        
        return query
    
    def _search_by_sender_and_date(self, sender: str, date_query: str, subject_pattern: str) -> List[Dict[str, Any]]:
        """送信者と日付で検索し、件名でフィルタリング"""
        if not sender:
            return []
        
        self.logger.info(f"送信者と日付での検索: {sender}")
        search_query = f'FROM "{sender}" {date_query}'
        
        typ, data = self.connection.search(None, search_query)
        if typ != 'OK' or not data[0]:
            return []
        
        email_ids = data[0].split()
        self.logger.debug(f"送信者+日付で見つかったメールID数: {len(email_ids)}")
        
        return self._process_email_ids(email_ids, subject_pattern=subject_pattern)
    
    def _search_by_subject_and_date(self, subject_pattern: str, date_query: str, sender: str) -> List[Dict[str, Any]]:
        """件名と日付で検索し、送信者でフィルタリング"""
        if not subject_pattern:
            return []
        
        self.logger.info(f"件名と日付での検索: {subject_pattern}")
        search_query = f'SUBJECT "{subject_pattern}" {date_query}'
        
        typ, data = self.connection.search(None, search_query)
        if typ != 'OK' or not data[0]:
            return []
        
        email_ids = data[0].split()
        self.logger.debug(f"件名+日付で見つかったメールID数: {len(email_ids)}")
        
        return self._process_email_ids(email_ids, sender_filter=sender)
    
    def _search_by_date_only(self, date_query: str, sender: str, subject_pattern: str) -> List[Dict[str, Any]]:
        """日付のみで検索し、送信者と件名でフィルタリング"""
        self.logger.info("日付範囲での全メール検索")
        
        typ, data = self.connection.search(None, date_query)
        if typ != 'OK' or not data[0]:
            return []
        
        all_ids = data[0].split()
        recent_ids = all_ids[-100:] if len(all_ids) > 100 else all_ids
        
        self.logger.debug(f"日付範囲内の最新 {len(recent_ids)} 件のメールから条件に一致するものを検索")
        return self._process_email_ids(recent_ids, sender_filter=sender, subject_pattern=subject_pattern)
    
    def _process_email_ids(self, email_ids: list, sender_filter: str = None, subject_pattern: str = None, limit: int = 5000) -> List[Dict[str, Any]]:
        """メールIDリストを処理してメール情報を抽出"""
        matching_emails = []
        email_ids_limited = email_ids[-limit:]
        
        # メール処理の進捗表示
        with tqdm(total=len(email_ids_limited), desc="メール処理進捗", unit="件") as email_pbar:
            for email_id in reversed(email_ids_limited):
                try:
                    email_pbar.set_description(f"処理中: メールID {email_id.decode() if isinstance(email_id, bytes) else email_id}")
                    
                    typ, msg_data = self.connection.fetch(email_id, '(RFC822)')
                    if typ != 'OK':
                        email_pbar.update(1)
                        continue
                    
                    raw_email = msg_data[0][1]
                    email_message = email.message_from_bytes(raw_email)
                    
                    email_info = self._extract_email_info(email_message)
                    email_info['id'] = email_id.decode() if isinstance(email_id, bytes) else email_id
                    
                    if self._matches_filters(email_info, sender_filter, subject_pattern):
                        matching_emails.append(email_info)
                        email_pbar.set_description(f"マッチしたメール: {len(matching_emails)}件")
                        
                except Exception as e:
                    self.logger.warning(f"メールID {email_id} の処理中にエラーが発生しました: {e}")
                finally:
                    email_pbar.update(1)
        
        return matching_emails
    
    def _matches_filters(self, email_info: Dict[str, Any], sender_filter: str = None, subject_pattern: str = None) -> bool:
        """メールが指定されたフィルター条件に一致するかチェック"""
        if sender_filter:
            actual_sender = email_info.get('sender', '')
            if sender_filter.lower() not in actual_sender.lower():
                return False
        
        if subject_pattern:
            actual_subject = email_info.get('subject', '')
            if subject_pattern.lower() not in actual_subject.lower():
                return False
        
        return True
    
    def _remove_duplicates(self, emails: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """重複したメールを除去"""
        unique_emails = []
        seen_ids = set()
        
        for email_info in emails:
            email_id = email_info.get('id')
            if email_id and email_id not in seen_ids:
                unique_emails.append(email_info)
                seen_ids.add(email_id)
        
        self.logger.info(f"条件に一致するメールを {len(unique_emails)} 件見つけました")
        return unique_emails
    
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
    
    @handle_errors("添付ファイル抽出")
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
        
        # デバッグ用：メール構造をログ出力
        self._log_email_structure(email_message)
            
        for part in email_message.walk():
            # より包括的な添付ファイル検出
            content_disposition = part.get_content_disposition()
            content_type = part.get_content_type()
            filename = self._get_attachment_filename(part)
            
            # デバッグ情報を出力
            self.logger.debug(f"Part: content_disposition={content_disposition}, content_type={content_type}, filename={filename}")
            
            # 添付ファイルの判定条件を拡張
            is_attachment = (
                content_disposition == 'attachment' or
                content_disposition == 'inline' or
                (filename and content_type in ['application/octet-stream', 'text/csv', 'application/csv', 'text/plain']) or
                (filename and filename.lower().endswith(file_type.lower())) or
                (content_type == 'text/csv' and filename) or
                # さらに包括的な条件：ファイル名があり、MIMEタイプがテキスト系の場合
                (filename and content_type.startswith('text/') and filename.lower().endswith(file_type.lower())) or
                # Content-Dispositionヘッダーが存在し、filenameパラメータが含まれている場合
                (content_disposition and 'filename=' in str(content_disposition).lower() and filename)
            )
            
            if is_attachment and filename and filename.lower().endswith(file_type.lower()):
                content = part.get_payload(decode=True)
                if content:
                    attachments.append({
                        'filename': filename,
                        'content': content,
                        'content_type': content_type
                    })
                    self.logger.info(f"添付ファイルを検出: {filename} (type: {content_type}, disposition: {content_disposition})")
                        
        self.logger.info(f"添付ファイルを {len(attachments)} 件抽出しました")
        return attachments
    
    def _get_attachment_filename(self, part):
        """添付ファイル名を取得（複数の方法を試行）"""
        def decode_mime_filename(filename_str):
            """MIMEエンコードされたファイル名をデコード"""
            if not filename_str:
                return None
            try:
                # すでにデコードされているかチェック
                if not filename_str.startswith('=?'):
                    return filename_str
                
                # MIMEヘッダーをデコード
                decoded = decode_header(filename_str)
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
                self.logger.warning(f"ファイル名デコードエラー: {e} (filename: {filename_str})")
                return filename_str
        
        # 通常の方法でファイル名を取得
        filename = part.get_filename()
        if filename:
            decoded_filename = decode_mime_filename(filename)
            self.logger.debug(f"ファイル名デコード: '{filename}' -> '{decoded_filename}'")
            return decoded_filename
        
        # Content-Dispositionヘッダーから直接取得
        content_disposition = part.get('Content-Disposition', '')
        if content_disposition:
            import re
            # filename="..." または filename*=... の形式を検索
            filename_match = re.search(r'filename[*]?=([^;]+)', content_disposition, re.IGNORECASE)
            if filename_match:
                filename = filename_match.group(1).strip('"\'')
                decoded_filename = decode_mime_filename(filename)
                self.logger.debug(f"Content-Dispositionからファイル名デコード: '{filename}' -> '{decoded_filename}'")
                return decoded_filename
        
        # Content-Typeヘッダーからname属性を取得
        content_type = part.get('Content-Type', '')
        if content_type:
            import re
            name_match = re.search(r'name=([^;]+)', content_type, re.IGNORECASE)
            if name_match:
                filename = name_match.group(1).strip('"\'')
                decoded_filename = decode_mime_filename(filename)
                self.logger.debug(f"Content-Typeからファイル名デコード: '{filename}' -> '{decoded_filename}'")
                return decoded_filename
        
        return None
    
    def _log_email_structure(self, email_message):
        """デバッグ用：メールの構造をログ出力"""
        self.logger.debug("=== メール構造の詳細 ===")
        for i, part in enumerate(email_message.walk()):
            content_type = part.get_content_type()
            content_disposition = part.get_content_disposition()
            filename = part.get_filename()
            decoded_filename = self._get_attachment_filename(part)
            is_multipart = part.is_multipart()
            
            self.logger.debug(f"Part {i}: "
                            f"content_type={content_type}, "
                            f"disposition={content_disposition}, "
                            f"filename={filename}, "
                            f"decoded_filename={decoded_filename}, "
                            f"multipart={is_multipart}")
            
            # CSVファイルかどうかをチェック
            if decoded_filename and decoded_filename.lower().endswith('.csv'):
                self.logger.debug(f"  *** CSVファイルを検出: {decoded_filename} ***")
                            
            # ヘッダー情報も出力
            if hasattr(part, 'items') and part.items():
                for header_name, header_value in part.items():
                    if header_name.lower() in ['content-type', 'content-disposition', 'content-transfer-encoding']:
                        self.logger.debug(f"  Header {header_name}: {header_value}")
        self.logger.debug("=== メール構造の詳細終了 ===")
    
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
                self.logger.info(f"件名から日付を抽出しました: {extracted_date} (件名: {subject[:50]}...)")
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