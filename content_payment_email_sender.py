#!/usr/bin/env python3
"""
コンテンツ関連支払明細書メール送信処理

rate.csvファイルを読み込み、各コンテンツの支払明細書を対応する複数のメールアドレスに下書き状態で送信します。
"""

import csv
import os
import sys
from pathlib import Path
from typing import List, Dict, Optional
import logging

# プロジェクトルートをパスに追加
sys.path.insert(0, str(Path(__file__).parent))

from content_payment_statement_generator.email_processor import EmailProcessor


class ContentPaymentEmailSender:
    """コンテンツ関連支払明細書メール送信クラス"""
    
    def __init__(self, rate_csv_path: str = "rate.csv", 
                 statements_dir: str = "output/content_payment_statements/"):
        """初期化"""
        self.logger = logging.getLogger(__name__)
        self.rate_csv_path = rate_csv_path
        self.statements_dir = Path(statements_dir)
        self.email_processor = EmailProcessor()
        
        # ログ設定
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
    
    def load_rate_csv(self) -> List[Dict]:
        """rate.csvファイルを読み込んでメール送信データを作成"""
        try:
            if not os.path.exists(self.rate_csv_path):
                raise FileNotFoundError(f"rate.csvファイルが見つかりません: {self.rate_csv_path}")
            
            email_data = []
            
            with open(self.rate_csv_path, 'r', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                headers = next(reader)  # ヘッダー行をスキップ
                
                for row in reader:
                    if len(row) < 3:  # 最低限A列、B列、C列が必要
                        continue
                    
                    content_name = row[0].strip()  # A列: コンテンツ名（ファイル名）
                    if not content_name:
                        continue
                    
                    # C列以降のメールアドレスを収集
                    email_addresses = []
                    for i in range(2, len(row)):  # C列（インデックス2）以降
                        email = row[i].strip()
                        if email and self.email_processor.validate_email_address(email):
                            email_addresses.append(email)
                    
                    if email_addresses:  # メールアドレスがある場合のみ追加
                        email_data.append({
                            'content_name': content_name,
                            'email_addresses': email_addresses
                        })
                        self.logger.info(f"コンテンツ '{content_name}' のメールアドレス {len(email_addresses)}件を読み込みました")
                    else:
                        self.logger.warning(f"コンテンツ '{content_name}' にはメールアドレスが設定されていません")
            
            self.logger.info(f"rate.csvから {len(email_data)} 件のコンテンツデータを読み込みました")
            return email_data
            
        except Exception as e:
            self.logger.error(f"rate.csv読み込みエラー: {e}")
            raise
    
    def find_statement_file(self, content_name: str) -> Optional[Path]:
        """コンテンツ名に対応する支払明細書ファイルを検索"""
        try:
            # PDFファイルを優先して検索
            for extension in ['.pdf', '.xlsx']:
                # 直接的な名前マッチング
                statement_file = self.statements_dir / f"{content_name}{extension}"
                if statement_file.exists():
                    return statement_file
                
                # パターンマッチング（コンテンツ名を含むファイル）
                for file_path in self.statements_dir.glob(f"*{content_name}*{extension}"):
                    if file_path.is_file():
                        return file_path
            
            self.logger.warning(f"コンテンツ '{content_name}' の支払明細書ファイルが見つかりません")
            return None
            
        except Exception as e:
            self.logger.error(f"支払明細書ファイル検索エラー: {e}")
            return None
    
    def create_draft_email(self, recipient: str, content_name: str, 
                          statement_file: Path, target_month: str = "") -> bool:
        """下書きメールを作成"""
        try:
            # Gmail APIは下書き作成をサポートしているが、send_messageの代わりに
            # drafts().create()を使用する必要がある
            
            sender = "mizoguchi@outward.jp"
            cc = "ow-fortune@ml.outward.jp"
            bcc = "mizoguchi@outward.jp"
            subject = "今月のお支払額のご連絡"
            
            # メール本文を作成
            body = self._create_payment_notification_body(target_month, content_name)
            
            # メールメッセージを作成
            message = self.email_processor.create_message_with_attachment(
                sender=sender,
                recipient=recipient,
                cc=cc,
                bcc=bcc,
                subject=subject,
                body=body,
                attachment_path=str(statement_file)
            )
            
            # 下書きとして保存
            success = self._save_as_draft(message)
            
            if success:
                self.logger.info(f"下書きメール作成完了: {recipient} (コンテンツ: {content_name})")
            else:
                self.logger.error(f"下書きメール作成失敗: {recipient} (コンテンツ: {content_name})")
            
            return success
            
        except Exception as e:
            self.logger.error(f"下書きメール作成エラー: {e}")
            return False
    
    def _save_as_draft(self, message: Dict[str, str]) -> bool:
        """メッセージを下書きとして保存"""
        return self.email_processor.save_as_draft(message)
    
    def _create_payment_notification_body(self, target_month: str = "", content_name: str = "") -> str:
        """支払い通知メールの本文を作成"""
        try:
            if target_month:
                # 年月を表示用にフォーマット
                year = target_month[:4]
                month = target_month[4:]
                formatted_month = f"{year}年{int(month)}月"
            else:
                formatted_month = "今月"
            
            content_info = f"（{content_name}）" if content_name else ""
            
            body = f"""いつもお世話になっております。

{formatted_month}のお支払額{content_info}をご連絡いたします。

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
    
    def send_all_content_emails(self, target_month: str = "") -> bool:
        """すべてのコンテンツの支払明細書メールを下書き状態で作成"""
        try:
            # rate.csvを読み込み
            email_data = self.load_rate_csv()
            
            if not email_data:
                self.logger.warning("送信対象のデータがありません")
                return False
            
            total_emails = 0
            successful_emails = 0
            
            # 各コンテンツに対してメール作成
            for data in email_data:
                content_name = data['content_name']
                email_addresses = data['email_addresses']
                
                # 支払明細書ファイルを検索
                statement_file = self.find_statement_file(content_name)
                if not statement_file:
                    self.logger.warning(f"コンテンツ '{content_name}' の支払明細書ファイルが見つかりません。スキップします。")
                    continue
                
                self.logger.info(f"コンテンツ '{content_name}' の明細書ファイル: {statement_file}")
                
                # 各メールアドレスに対してメール作成
                for email_address in email_addresses:
                    total_emails += 1
                    
                    if self.create_draft_email(email_address, content_name, statement_file, target_month):
                        successful_emails += 1
                    else:
                        self.logger.error(f"メール作成失敗: {email_address} (コンテンツ: {content_name})")
            
            self.logger.info(f"メール作成処理完了: 成功 {successful_emails}/{total_emails}")
            
            return successful_emails == total_emails
            
        except Exception as e:
            self.logger.error(f"メール一括作成エラー: {e}")
            return False
    
    def send_specific_content_emails(self, content_name: str, target_month: str = "") -> bool:
        """特定のコンテンツの支払明細書メールを下書き状態で作成"""
        try:
            # rate.csvを読み込み
            email_data = self.load_rate_csv()
            
            # 指定されたコンテンツを検索
            target_data = None
            for data in email_data:
                if data['content_name'] == content_name:
                    target_data = data
                    break
            
            if not target_data:
                self.logger.error(f"コンテンツ '{content_name}' がrate.csvに見つかりません")
                return False
            
            # 支払明細書ファイルを検索
            statement_file = self.find_statement_file(content_name)
            if not statement_file:
                self.logger.error(f"コンテンツ '{content_name}' の支払明細書ファイルが見つかりません")
                return False
            
            self.logger.info(f"コンテンツ '{content_name}' の明細書ファイル: {statement_file}")
            
            total_emails = 0
            successful_emails = 0
            
            # 各メールアドレスに対してメール作成
            for email_address in target_data['email_addresses']:
                total_emails += 1
                
                if self.create_draft_email(email_address, content_name, statement_file, target_month):
                    successful_emails += 1
                else:
                    self.logger.error(f"メール作成失敗: {email_address}")
            
            self.logger.info(f"コンテンツ '{content_name}' のメール作成完了: 成功 {successful_emails}/{total_emails}")
            
            return successful_emails == total_emails
            
        except Exception as e:
            self.logger.error(f"特定コンテンツメール作成エラー: {e}")
            return False


def main():
    """メイン関数"""
    import argparse
    
    parser = argparse.ArgumentParser(description="コンテンツ関連支払明細書メール送信（下書き作成）")
    parser.add_argument('--content', type=str, help='特定のコンテンツ名を指定')
    parser.add_argument('--month', type=str, help='対象月（YYYYMM形式）')
    parser.add_argument('--rate-csv', type=str, default='rate.csv', help='rate.csvファイルのパス')
    parser.add_argument('--statements-dir', type=str, default='output/content_payment_statements/', 
                       help='支払明細書ファイルのディレクトリ')
    
    args = parser.parse_args()
    
    try:
        sender = ContentPaymentEmailSender(
            rate_csv_path=args.rate_csv,
            statements_dir=args.statements_dir
        )
        
        if args.content:
            # 特定のコンテンツのみ処理
            success = sender.send_specific_content_emails(args.content, args.month or "")
        else:
            # すべてのコンテンツを処理
            success = sender.send_all_content_emails(args.month or "")
        
        if success:
            print("下書きメール作成が正常に完了しました。")
            sys.exit(0)
        else:
            print("下書きメール作成中にエラーが発生しました。")
            sys.exit(1)
            
    except Exception as e:
        print(f"予期しないエラーが発生しました: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()