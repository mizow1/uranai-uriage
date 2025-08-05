"""
メール送信処理モジュール

Gmail APIを使用したメール送信と予約機能を提供します。
"""

import os
import base64
import csv
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
    
    # Gmail API のスコープ（下書き作成に必要なスコープを追加）
    SCOPES = [
        'https://www.googleapis.com/auth/gmail.send',
        'https://www.googleapis.com/auth/gmail.compose',
        'https://www.googleapis.com/auth/gmail.modify'
    ]
    
    def __init__(self, credentials_path: str = 'credentials.json', token_path: str = 'token.json'):
        """メール送信処理クラスを初期化"""
        self.logger = logging.getLogger(__name__)
        self.credentials_path = credentials_path
        self.token_path = token_path
        self.gmail_service = None
        self.contents_mapping = {}
        
        # コンテンツマッピングを読み込み
        self._load_contents_mapping()
        
        if GOOGLE_APIS_AVAILABLE:
            try:
                self.gmail_service = self._setup_gmail_service()
            except Exception as e:
                self.logger.error(f"Gmail API初期化エラー: {e}")
        else:
            self.logger.warning("Google API libraries not available")
    
    def _load_contents_mapping(self) -> None:
        """contents_mapping.csvからコンテンツマッピングを読み込み"""
        try:
            csv_path = Path(__file__).parent.parent / 'contents_mapping.csv'
            if not csv_path.exists():
                self.logger.warning(f"contents_mapping.csvが見つかりません: {csv_path}")
                return
            
            with open(csv_path, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                headers = next(reader)
                
                for row in reader:
                    if len(row) >= 7:  # 最低7列必要（A-G列）
                        content_id = row[0]  # A列
                        mediba_name = row[3]  # D列
                        ameba_name = row[5]   # F列
                        rakuten_name = row[6] # G列
                        
                        self.contents_mapping[content_id] = {
                            'mediba': mediba_name,
                            'ameba': ameba_name,
                            'rakuten': rakuten_name
                        }
                        
                        # デバッグ用ログ（最初の5件のみ）
                        if len(self.contents_mapping) <= 5:
                            self.logger.debug(f"コンテンツID '{content_id}': mediba='{mediba_name}', ameba='{ameba_name}', rakuten='{rakuten_name}'")
            
            self.logger.info(f"コンテンツマッピングを読み込みました: {len(self.contents_mapping)}件")
            
        except Exception as e:
            self.logger.error(f"コンテンツマッピング読み込みエラー: {e}")
    
    def _get_content_name_for_subject(self, content_id: str) -> str:
        """件名用のコンテンツ名を取得（優先順位：ameba > mediba > rakuten）"""
        try:
            self.logger.debug(f"コンテンツ名を検索中: '{content_id}'")
            self.logger.debug(f"利用可能なコンテンツID: {list(self.contents_mapping.keys())[:5]}...")  # 最初の5件のみ表示
            
            if content_id not in self.contents_mapping:
                self.logger.warning(f"コンテンツID '{content_id}' がマッピングに見つかりません")
                return ""
            
            mapping = self.contents_mapping[content_id]
            self.logger.debug(f"マッピング内容: {mapping}")
            
            # 優先順位：F列（ameba）> D列（mediba）> G列（rakuten）
            if mapping['ameba'] and mapping['ameba'].strip():
                result = mapping['ameba'].strip()
                self.logger.debug(f"amebaから取得: '{result}'")
                return result
            elif mapping['mediba'] and mapping['mediba'].strip():
                result = mapping['mediba'].strip()
                self.logger.debug(f"medibaから取得: '{result}'")
                return result
            elif mapping['rakuten'] and mapping['rakuten'].strip():
                result = mapping['rakuten'].strip()
                self.logger.debug(f"rakutenから取得: '{result}'")
                return result
            else:
                self.logger.debug("すべての列が空でした")
                return ""
                
        except Exception as e:
            self.logger.error(f"コンテンツ名取得エラー: {e}")
            return ""
    
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
            import traceback
            self.logger.error(f"トレースバック: {traceback.format_exc()}")
            self.logger.error("認証情報ファイル(credentials.json)とトークンファイル(token.json)を確認してください")
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
    
    def save_as_draft(self, message: Dict[str, str]) -> bool:
        """メッセージを下書きとして保存"""
        try:
            if not self.gmail_service:
                self.logger.error("Gmail APIサービスが初期化されていません")
                self.logger.error("1. credentials.jsonファイルが存在するか確認してください")
                self.logger.error("2. token.jsonファイルを削除して再認証を試してください")
                self.logger.error("3. Gmail APIが有効になっているか確認してください")
                return False
            
            # Gmail APIで下書きを作成
            self.logger.debug("下書き作成リクエストを準備中...")
            draft_request = {
                'message': message
            }
            
            self.logger.debug("Gmail APIに下書き作成リクエストを送信中...")
            result = self.gmail_service.users().drafts().create(
                userId='me',
                body=draft_request
            ).execute()
            
            draft_id = result.get('id', 'Unknown')
            self.logger.info(f"下書き作成完了: Draft ID {draft_id}")
            self.logger.info("Gmailの下書きフォルダに下書きメールが作成されました")
            return True
            
        except HttpError as error:
            error_details = getattr(error, 'error_details', str(error))
            self.logger.error(f"Gmail API エラー: {error}")
            self.logger.error(f"エラー詳細: {error_details}")
            if hasattr(error, 'resp') and hasattr(error.resp, 'status'):
                self.logger.error(f"HTTPステータス: {error.resp.status}")
            return False
        except Exception as e:
            self.logger.error(f"下書き保存エラー: {e}")
            import traceback
            self.logger.error(f"トレースバック: {traceback.format_exc()}")
            return False

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
    
    def create_payment_notification_draft(
        self,
        recipient: str,
        pdf_path: str,
        target_month: str,
        content_name: str = "",
        addressee_name: str = "",
        content_id: str = ""
    ) -> bool:
        """支払い通知メールを下書きとして作成"""
        try:
            # メール設定
            sender = "mizoguchi@outward.jp"
            cc = "ow-fortune@ml.outward.jp"
            bcc = "mizoguchi@outward.jp"
            
            # 複数アドレスの場合の処理
            if ',' in recipient:
                # 複数アドレスをTOに設定（カンマ区切りのまま）
                recipient_to = recipient
                self.logger.info(f"複数メールアドレスをTOに設定: {recipient_to}")
            else:
                recipient_to = recipient
            
            # メール件名を作成（コンテンツ名を含む）
            subject = self._create_payment_notification_subject(target_month, addressee_name, content_id)
            
            # メール本文を作成（A6セルの値を使用）
            body = self._create_payment_notification_body(target_month, content_name, addressee_name)
            
            # メールメッセージを作成
            message = self.create_message_with_attachment(
                sender=sender,
                recipient=recipient_to,
                cc=cc,
                bcc=bcc,
                subject=subject,
                body=body,
                attachment_path=pdf_path
            )
            
            # 下書きとして保存
            success = self.save_as_draft(message)
            
            if success:
                self.logger.info(f"支払い通知メール下書き作成完了: {recipient}")
            else:
                self.logger.error(f"支払い通知メール下書き作成失敗: {recipient}")
            
            return success
            
        except Exception as e:
            self.logger.error(f"支払い通知メール下書き作成エラー: {e}")
            return False

    def send_payment_notification(
        self,
        recipient: str,
        pdf_path: str,
        target_month: str,
        content_name: str = "",
        addressee_name: str = "",
        content_id: str = ""
    ) -> bool:
        """支払い通知メールを下書きとして作成（注意：送信はしません）"""
        try:
            # メール設定
            sender = "mizoguchi@outward.jp"
            cc = "ow-fortune@ml.outward.jp"
            bcc = "mizoguchi@outward.jp"
            
            # メール件名を作成（コンテンツ名を含む）
            subject = self._create_payment_notification_subject(target_month, addressee_name, content_id)
            
            # メール本文を作成（A6セルの値を使用）
            body = self._create_payment_notification_body(target_month, content_name, addressee_name)
            
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
            
            # 下書きとして保存（送信ではなく）
            success = self.save_as_draft(message)
            
            if success:
                self.logger.info(f"支払い通知メール下書き作成完了: {recipient}")
            else:
                self.logger.error(f"支払い通知メール下書き作成失敗: {recipient}")
            
            return success
            
        except Exception as e:
            self.logger.error(f"支払い通知メール下書き作成エラー: {e}")
            return False
    
    def _create_payment_notification_subject(self, target_month: str, addressee_name: str = "", content_id: str = "") -> str:
        """支払い通知メールの件名を作成"""
        try:
            # 年月を表示用にフォーマット
            year = target_month[:4]
            month = target_month[4:]
            
            # 基本の件名
            base_subject = f"{year}年{int(month)}月お支払いのご連絡"
            
            # コンテンツ名を取得
            content_name_for_subject = ""
            if content_id and content_id.strip():
                content_name_for_subject = self._get_content_name_for_subject(content_id.strip())
                self.logger.info(f"コンテンツID '{content_id}' からコンテンツ名 '{content_name_for_subject}' を取得")
            else:
                self.logger.warning(f"コンテンツIDが空またはNone: '{content_id}'")
            
            # 件名にコンテンツ名を追加
            if content_name_for_subject:
                subject = f"{base_subject}【{content_name_for_subject}】"
                self.logger.info(f"件名にコンテンツ名を追加: {subject}")
            else:
                subject = base_subject
                self.logger.info(f"コンテンツ名なしの件名: {subject}")
            
            return subject
            
        except Exception as e:
            self.logger.error(f"メール件名作成エラー: {e}")
            return "今月のお支払額のご連絡"

    def _create_payment_notification_body(self, target_month: str, content_name: str = "", addressee_name: str = "") -> str:
        """支払い通知メールの本文を作成"""
        try:
            # 年月を表示用にフォーマット
            year = target_month[:4]
            month = target_month[4:]
            formatted_month = f"{year}年{int(month)}月"
            
            # 宛名部分
            greeting = f"{addressee_name}様" if addressee_name else ""
            if greeting:
                greeting += "\n\n"
            
            body = f"""{greeting}お世話になっております。アウトワード溝口です。

{formatted_month}のロイヤリティ明細書をお送りいたします。

今回のロイヤリティにつきましては、10,000円未満のため、誠に恐縮ながら次回分と合算してのお振込みとさせていただきます。何卒ご了承いただけますようお願い申し上げます。

お振込みやロイヤリティに関して、ご不明な点やご不備などがございましたら、ご遠慮なくご連絡ください。

今後ともどうぞよろしくお願い申し上げます。


---
────────────────────────────────
溝口　洋輔

MAIL: mizoguchi@outward.jp

株式会社アウトワード

本社／福岡市西区福重3-36-6　〒819-0022
TEL：092-885-1364　FAX：092-885-1459
URL: http://www.outward.jp
────────────────────────────────
"""
            
            return body
            
        except Exception as e:
            self.logger.error(f"メール本文作成エラー: {e}")
            return "今月のお支払額をご連絡いたします。"
    
    def schedule_email(self, email_data: Dict, send_date: datetime) -> Optional[str]:
        """メール下書き作成を設定（注意: 送信は行わず下書きのみ作成します）"""
        try:
            # Gmail APIは予約送信をネイティブサポートしていないため、
            # ここでは下書きとして保存します
            
            current_time = datetime.now()
            
            if send_date <= current_time:
                # 指定時刻が過去または現在の場合は下書きとして保存
                self.logger.info("指定時刻が過去のため、下書きとして保存します")
                success = self.save_as_draft(email_data)
                return "draft_created" if success else None
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