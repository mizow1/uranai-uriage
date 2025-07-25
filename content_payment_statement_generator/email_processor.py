"""
メール送信処理モジュール

Gmail APIを使用したメール送信と予約機能を提供します。
"""

import os
import base64
from pathlib import Path
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Any
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

try:
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    GOOGLE_APIS_AVAILABLE = True
except ImportError:
    GOOGLE_APIS_AVAILABLE = False
    logging.warning("Google API libraries not available. Email functionality will be limited.")


class EmailProcessor:
    """メール送信処理クラス"""
    
    # Gmail API のスコープ
    SCOPES = ['https://www.googleapis.com/auth/gmail.send']
    
    def __init__(self, credentials_path: str = 'credentials.json', token_path: str = 'token.json'):
        """メール送信処理クラスを初期化"""
        self.logger = logging.getLogger(__name__)
        self.credentials_path = credentials_path
        self.token_path = token_path
        self.gmail_service = None
        
        if GOOGLE_APIS_AVAILABLE:
            try:
                self.gmail_service = self._setup_gmail_service()
            except Exception as e:
                self.logger.error(f"Gmail API初期化エラー: {e}")
        else:
            self.logger.warning("Google API libraries not available")
    
    def _setup_gmail_service(self) -> Optional[Any]:
        """Gmail APIサービスを設定"""
        try:
            creds = None
            
            # トークンファイルが存在する場合は読み込み
            if os.path.exists(self.token_path):
                creds = Credentials.from_authorized_user_file(self.token_path, self.SCOPES)
            
            # 有効な認証情報がない場合は認証フローを実行
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    if not os.path.exists(self.credentials_path):
                        raise FileNotFoundError(f"認証情報ファイルが見つかりません: {self.credentials_path}")
                    
                    flow = InstalledAppFlow.from_client_secrets_file(self.credentials_path, self.SCOPES)
                    creds = flow.run_local_server(port=0)
                
                # トークンを保存
                with open(self.token_path, 'w') as token:
                    token.write(creds.to_json())
            
            # Gmail APIサービスを構築
            service = build('gmail', 'v1', credentials=creds)
            self.logger.info("Gmail APIサービスが正常に初期化されました")
            return service
            
        except Exception as e:
            self.logger.error(f"Gmail APIサービス初期化エラー: {e}")
            return None
    
    def create_message_with_attachment(
        self,
        sender: str,
        recipient: str,
        cc: Optional[str] = None,
        bcc: Optional[str] = None,
        subject: str = "",
        body: str = "",
        attachment_path: Optional[str] = None
    ) -> Dict[str, str]:
        """添付ファイル付きメールメッセージを作成"""
        try:
            # MIMEメッセージを作成
            message = MIMEMultipart()
            message['from'] = sender
            message['to'] = recipient
            
            if cc:
                message['cc'] = cc
            if bcc:
                message['bcc'] = bcc
                
            message['subject'] = subject
            
            # メール本文を追加
            message.attach(MIMEText(body, 'plain', 'utf-8'))
            
            # 添付ファイルを追加
            if attachment_path and Path(attachment_path).exists():
                with open(attachment_path, 'rb') as f:
                    attachment = MIMEApplication(f.read())
                    attachment.add_header(
                        'Content-Disposition',
                        'attachment',
                        filename=Path(attachment_path).name
                    )
                    message.attach(attachment)
                    
                self.logger.debug(f"添付ファイルを追加しました: {attachment_path}")
            
            # Base64エンコード
            raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
            
            return {'raw': raw_message}
            
        except Exception as e:
            self.logger.error(f"メールメッセージ作成エラー: {e}")
            raise
    
    def send_message(self, message: Dict[str, str]) -> bool:
        """メールメッセージを送信"""
        try:
            if not self.gmail_service:
                self.logger.error("Gmail APIサービスが初期化されていません")
                return False
            
            # メールを送信
            result = self.gmail_service.users().messages().send(
                userId='me',
                body=message
            ).execute()
            
            self.logger.info(f"メール送信完了: Message ID {result['id']}")
            return True
            
        except HttpError as error:
            self.logger.error(f"Gmail API エラー: {error}")
            return False
        except Exception as e:
            self.logger.error(f"メール送信エラー: {e}")
            return False
    
    def send_payment_notification(
        self,
        recipient: str,
        pdf_path: str,
        target_month: str,
        content_name: str = ""
    ) -> bool:
        """支払い通知メールを送信"""
        try:
            # メール設定
            sender = "mizoguchi@outward.jp"
            cc = "ow-fortune@ml.outward.jp"
            bcc = "mizoguchi@outward.jp"
            subject = "今月のお支払額のご連絡"
            
            # メール本文を作成
            body = self._create_payment_notification_body(target_month, content_name)
            
            # メールメッセージを作成
            message = self.create_message_with_attachment(
                sender=sender,
                recipient=recipient,
                cc=cc,
                bcc=bcc,
                subject=subject,
                body=body,
                attachment_path=pdf_path
            )
            
            # メールを送信
            success = self.send_message(message)
            
            if success:
                self.logger.info(f"支払い通知メール送信完了: {recipient}")
            else:
                self.logger.error(f"支払い通知メール送信失敗: {recipient}")
            
            return success
            
        except Exception as e:
            self.logger.error(f"支払い通知メール送信エラー: {e}")
            return False
    
    def _create_payment_notification_body(self, target_month: str, content_name: str = "") -> str:
        """支払い通知メールの本文を作成"""
        try:
            # 年月を表示用にフォーマット
            year = target_month[:4]
            month = target_month[4:]
            formatted_month = f"{year}年{int(month)}月"
            
            body = f"""いつもお世話になっております。

{formatted_month}のお支払額をご連絡いたします。

詳細につきましては、添付の支払明細書をご確認ください。

何かご不明な点がございましたら、お気軽にお問い合わせください。

よろしくお願いいたします。

---
溝口
mizoguchi@outward.jp"""
            
            return body
            
        except Exception as e:
            self.logger.error(f"メール本文作成エラー: {e}")
            return "今月のお支払額をご連絡いたします。"
    
    def schedule_email(self, email_data: Dict, send_date: datetime) -> Optional[str]:
        """メール予約送信を設定（注意: Gmail APIは予約送信をサポートしていません）"""
        try:
            # Gmail APIは予約送信をネイティブサポートしていないため、
            # ここでは即座に送信するか、別のスケジューリングシステムを使用する必要があります
            
            current_time = datetime.now()
            
            if send_date <= current_time:
                # 指定時刻が過去または現在の場合は即座に送信
                self.logger.info("指定時刻が過去のため、即座にメールを送信します")
                success = self.send_message(email_data)
                return "immediate_send" if success else None
            else:
                # 将来の日時の場合は警告を出力
                self.logger.warning(f"Gmail APIは予約送信をサポートしていません。指定日時: {send_date}")
                self.logger.warning("代替案として、外部スケジューラーの使用を検討してください")
                return None
                
        except Exception as e:
            self.logger.error(f"メール予約設定エラー: {e}")
            return None
    
    def validate_email_address(self, email: str) -> bool:
        """メールアドレスの形式を検証"""
        try:
            import re
            
            # 基本的なメールアドレス形式の正規表現
            pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
            
            if re.match(pattern, email):
                return True
            else:
                self.logger.warning(f"無効なメールアドレス形式: {email}")
                return False
                
        except Exception as e:
            self.logger.error(f"メールアドレス検証エラー: {e}")
            return False
    
    def test_connection(self) -> bool:
        """Gmail API接続をテスト"""
        try:
            if not self.gmail_service:
                self.logger.error("Gmail APIサービスが初期化されていません")
                return False
            
            # プロフィール情報を取得してテスト
            profile = self.gmail_service.users().getProfile(userId='me').execute()
            
            self.logger.info(f"Gmail API接続テスト成功: {profile.get('emailAddress')}")
            return True
            
        except Exception as e:
            self.logger.error(f"Gmail API接続テストエラー: {e}")
            return False