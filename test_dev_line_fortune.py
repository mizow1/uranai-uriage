#!/usr/bin/env python3
"""
dev-line-fortune@linecorp.com の存在を直接確認
"""

import sys
from pathlib import Path
import imaplib

# パッケージのパスを追加
sys.path.insert(0, str(Path(__file__).parent))

from line_fortune_processor.config import Config


def main():
    """メイン関数"""
    config = Config()
    email_config = config.get('email', {})
    
    print("=== dev-line-fortune@linecorp.com 存在確認 ===")
    
    try:
        # IMAP接続
        connection = imaplib.IMAP4_SSL(email_config.get('server'), email_config.get('port'))
        connection.login(email_config.get('username'), email_config.get('password'))
        connection.select('INBOX')
        
        # 直接的な検索テスト
        test_queries = [
            'FROM "dev-line-fortune@linecorp.com"',
            'FROM "dev-line-fortune"',
            'FROM "linecorp.com"',
            'FROM "@linecorp.com"'
        ]
        
        for query in test_queries:
            print(f"\nクエリ: {query}")
            typ, data = connection.search(None, query)
            
            if typ == 'OK':
                ids = data[0].split()
                print(f"  結果: {len(ids)} 件")
            else:
                print(f"  検索エラー: {typ}")
        
        connection.close()
        connection.logout()
        
    except Exception as e:
        print(f"エラー: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()