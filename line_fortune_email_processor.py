#!/usr/bin/env python3
"""
LINE Fortune Email Processor - メインスクリプト

LINE Fortuneの日次レポートメールを自動処理するシステム
"""

import sys
import argparse
import os
from pathlib import Path

# パッケージのパスを追加
sys.path.insert(0, str(Path(__file__).parent))

from line_fortune_processor.main_processor import LineFortuneProcessor
from line_fortune_processor.config import Config
from line_fortune_processor.logger import get_logger


def main():
    """メイン関数"""
    parser = argparse.ArgumentParser(
        description="LINE Fortune Email Processor - メール処理自動化システム"
    )
    
    parser.add_argument(
        "--config",
        "-c",
        type=str,
        help="設定ファイルのパス (デフォルト: line_fortune_config.json)"
    )
    
    parser.add_argument(
        "--dry-run",
        "-d",
        action="store_true",
        help="ドライランモード（実際の処理は行わない）"
    )
    
    parser.add_argument(
        "--create-config",
        action="store_true",
        help="設定ファイルのテンプレートを作成"
    )
    
    parser.add_argument(
        "--cleanup",
        type=int,
        metavar="DAYS",
        help="指定した日数より古いファイルを削除"
    )
    
    parser.add_argument(
        "--log-level",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        default="INFO",
        help="ログレベル (デフォルト: INFO)"
    )
    
    
    args = parser.parse_args()
    
    # 日付範囲の対話形式での取得
    from datetime import datetime, date
    
    # 設定ファイルテンプレートを作成
    if args.create_config:
        config = Config()
        if config.create_template():
            print("設定ファイルテンプレートを作成しました: line_fortune_config_template.json")
            print("必要な設定を入力して、line_fortune_config.jsonとして保存してください。")
            return 0
        else:
            print("設定ファイルテンプレートの作成に失敗しました。")
            return 1
    
    def get_date_input(prompt, default_date=None):
        """対話形式で日付を取得"""
        while True:
            if default_date:
                user_input = input(f"{prompt} (デフォルト: {default_date}, Enterキーで確定): ").strip()
                if not user_input:
                    return default_date
            else:
                user_input = input(f"{prompt} (YYYY-MM-DD形式): ").strip()
            
            if not user_input:
                continue
                
            try:
                return datetime.strptime(user_input, '%Y-%m-%d').date()
            except ValueError:
                print("エラー: 日付は YYYY-MM-DD 形式で入力してください（例: 2025-07-22）")
    
    # 対話形式で日付範囲を取得
    print("\n" + "="*50)
    print("LINE Fortune Email Processor")
    print("="*50)
    print("処理対象メールの日付範囲を指定してください")
    
    # 開始日の入力
    today = date.today()
    start_date = get_date_input("開始年月日を入力", today)
    
    # 終了日の入力
    end_date = get_date_input("終了年月日を入力", start_date)
    
    # 日付範囲の妥当性チェック
    if start_date > end_date:
        print("\nエラー: 開始日は終了日より前である必要があります")
        return 1
    
    # 確認表示
    if start_date == end_date:
        print(f"\n処理対象: {start_date} のメール")
    else:
        print(f"\n処理対象: {start_date} ～ {end_date} のメール")
    
    confirm = input("この設定で実行しますか？ (y/N): ").strip().lower()
    if confirm not in ['y', 'yes']:
        print("処理を中止しました")
        return 0

    try:
        # プロセッサーの初期化（日付範囲指定を渡す）
        processor = LineFortuneProcessor(args.config, start_date=start_date, end_date=end_date)
        
        # ログレベルの設定
        logger = get_logger()
        logger.set_level(args.log_level)
        
        # 古いファイルの削除
        if args.cleanup:
            # cleanup処理の場合は日付入力をスキップして直接処理を実行
            cleanup_processor = LineFortuneProcessor(args.config)
            logger = get_logger()
            logger.set_level(args.log_level)
            logger.info(f"古いファイルの削除を開始します: {args.cleanup} 日より古いファイル")
            if cleanup_processor.cleanup_old_files(args.cleanup):
                logger.info("古いファイルの削除が完了しました")
            else:
                logger.error("古いファイルの削除に失敗しました")
                return 1
            return 0
        
        # ドライランモード
        if args.dry_run:
            logger.info("ドライランモードで実行します")
            success = processor.dry_run()
        else:
            logger.info("メール処理を開始します")
            success = processor.process()
        
        # 結果の表示
        stats = processor.get_stats()
        if not args.dry_run:
            print(f"処理結果:")
            print(f"  処理したメール数: {stats['emails_processed']}")
            print(f"  成功したメール数: {stats['emails_success']}")
            print(f"  エラーが発生したメール数: {stats['emails_error']}")
            print(f"  保存したファイル数: {stats['files_saved']}")
            print(f"  作成した統合ファイル数: {stats['consolidations_created']}")
        
        return 0 if success else 1
        
    except KeyboardInterrupt:
        logger = get_logger()
        logger.info("ユーザーによって処理が中断されました")
        return 1
        
    except Exception as e:
        logger = get_logger()
        logger.error(f"予期しないエラーが発生しました: {e}", exception=e)
        return 1


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)