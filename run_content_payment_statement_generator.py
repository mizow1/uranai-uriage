#!/usr/bin/env python3
"""
コンテンツ関連支払い明細書生成システム メイン実行スクリプト

使用方法:
    python run_content_payment_statement_generator.py 2024 12
    python run_content_payment_statement_generator.py 2024 12 aiga
    python run_content_payment_statement_generator.py 2024 12 --no-email
    python run_content_payment_statement_generator.py 2024 12 aiga --test
"""

import sys
import argparse
from pathlib import Path

# プロジェクトルートをパスに追加
sys.path.insert(0, str(Path(__file__).parent))

from content_payment_statement_generator.main_controller import MainController


def parse_arguments():
    """コマンドライン引数を解析"""
    parser = argparse.ArgumentParser(
        description="コンテンツ関連支払い明細書自動生成システム",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  %(prog)s 2024 12                    # 2024年12月の明細書を生成・送信
  %(prog)s 2024 12 aiga               # 2024年12月のaigaコンテンツのみ生成・送信
  %(prog)s 2024 12 --no-email         # メール送信せずに明細書のみ生成
  %(prog)s 2024 12 aiga --test        # aigaコンテンツのシステムテストを実行
  %(prog)s 2024 12 --log-level DEBUG  # デバッグレベルでログ出力
        """
    )
    
    parser.add_argument(
        'year',
        type=str,
        help='処理対象年（例: 2024）'
    )
    
    parser.add_argument(
        'month',
        type=str,
        help='処理対象月（例: 12）'
    )
    
    parser.add_argument(
        'content',
        type=str,
        nargs='?',
        help='処理対象コンテンツ名（例: aiga）。指定した場合、そのコンテンツのみ処理'
    )
    
    parser.add_argument(
        '--no-email',
        action='store_true',
        help='メール送信をスキップ（明細書生成のみ）'
    )
    
    parser.add_argument(
        '--test',
        action='store_true',
        help='システムテストを実行'
    )
    
    parser.add_argument(
        '--log-level',
        choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
        default='INFO',
        help='ログレベル（デフォルト: INFO）'
    )
    
    parser.add_argument(
        '--template-filter',
        type=str,
        help='特定のテンプレートファイルのみ処理（例: aiga.xlsx）'
    )
    
    return parser.parse_args()


def validate_input(year: str, month: str) -> bool:
    """入力値を検証"""
    try:
        # 年の検証
        year_int = int(year)
        if year_int < 2020 or year_int > 2030:
            print(f"エラー: 年は2020-2030の範囲で指定してください: {year}")
            return False
        
        # 月の検証
        month_int = int(month)
        if month_int < 1 or month_int > 12:
            print(f"エラー: 月は1-12の範囲で指定してください: {month}")
            return False
        
        return True
        
    except ValueError:
        print(f"エラー: 年月は数値で指定してください: {year}, {month}")
        return False


def run_system_test(controller: MainController) -> bool:
    """システムテストを実行"""
    print("=" * 60)
    print("システムテストを実行中...")
    print("=" * 60)
    
    test_results = controller.test_system_components()
    
    print("\nテスト結果:")
    all_passed = True
    
    for component, result in test_results.items():
        status = "成功" if result else "失敗"
        print(f"  {component}: {status}")
        if not result:
            all_passed = False
    
    print("\n" + "=" * 60)
    if all_passed:
        print("すべてのテストが成功しました。システムは正常に動作します。")
    else:
        print("一部のテストが失敗しました。設定を確認してください。")
    print("=" * 60)
    
    return all_passed


def main():
    """メイン関数"""
    try:
        # コマンドライン引数を解析
        args = parse_arguments()
        
        # 入力値を検証
        if not validate_input(args.year, args.month):
            sys.exit(1)
        
        print("=" * 60)
        print("コンテンツ関連支払い明細書生成システム")
        print(f"処理対象: {args.year}年{args.month}月")
        if args.content:
            print(f"対象コンテンツ: {args.content}")
        print(f"メール送信: {'無効' if args.no_email else '有効'}")
        print(f"ログレベル: {args.log_level}")
        print("=" * 60)
        
        # メインコントローラーを初期化
        controller = MainController(log_level=args.log_level)
        
        # システムテストを実行
        if args.test:
            success = run_system_test(controller)
            sys.exit(0 if success else 1)
        
        # メイン処理を実行
        send_emails = not args.no_email
        
        # コンテンツ指定がある場合はtemplate_filterとして使用
        template_filter = args.template_filter
        if args.content and not template_filter:
            template_filter = f"{args.content}.xlsx"
        
        success = controller.process_payment_statements(
            args.year, 
            args.month, 
            send_emails,
            template_filter=template_filter,
            content_name=args.content
        )
        
        if success:
            print("\n処理が正常に完了しました。")
            sys.exit(0)
        else:
            print("\n処理中にエラーが発生しました。ログを確認してください。")
            sys.exit(1)
            
    except KeyboardInterrupt:
        print("\n\n処理が中断されました。")
        sys.exit(1)
    except Exception as e:
        print(f"\n予期しないエラーが発生しました: {e}")
        sys.exit(1)


if __name__ == '__main__':
    main()