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
    
    try:
        # プロセッサーの初期化
        processor = LineFortuneProcessor(args.config)
        
        # ログレベルの設定
        logger = get_logger()
        logger.set_level(args.log_level)
        
        # 古いファイルの削除
        if args.cleanup:
            logger.info(f"古いファイルの削除を開始します: {args.cleanup} 日より古いファイル")
            if processor.cleanup_old_files(args.cleanup):
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