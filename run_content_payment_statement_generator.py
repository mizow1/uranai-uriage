#!/usr/bin/env python3
"""
コンテンツ関連支払い明細書生成システム メイン実行スクリプト

使用方法:
    python run_content_payment_statement_generator.py 2024 12
    python run_content_payment_statement_generator.py 2024 12 aiga
    python run_content_payment_statement_generator.py 2024 12 --no-email
    python run_content_payment_statement_generator.py 2024 12 aiga --test
    python run_content_payment_statement_generator.py 2025 06 2025 08  # 期間指定
"""

import sys
import argparse
from pathlib import Path
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

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
  %(prog)s 2025 06 2025 08            # 2025年6月から8月までの期間を処理
        """
    )
    
    parser.add_argument(
        'year',
        type=str,
        help='処理対象年（例: 2024）または開始年'
    )
    
    parser.add_argument(
        'month',
        type=str,
        help='処理対象月（例: 12）または開始月'
    )
    
    parser.add_argument(
        'content',
        type=str,
        nargs='?',
        help='処理対象コンテンツ名（例: aiga）または終了年（期間指定時）'
    )
    
    parser.add_argument(
        'end_month',
        type=str,
        nargs='?',
        help='終了月（期間指定時）'
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


def parse_content_and_dates(args):
    """引数からコンテンツ名と期間を解析"""
    content = None
    end_year = None
    end_month = None
    
    # 3番目の引数（content位置）が数値なら終了年、そうでなければコンテンツ名
    if args.content:
        try:
            # 数値として解析できれば終了年
            int(args.content)
            end_year = args.content
            end_month = args.end_month
        except ValueError:
            # 数値でなければコンテンツ名
            content = args.content
    
    return content, end_year, end_month


def validate_input(year: str, month: str, end_year: str = None, end_month: str = None) -> bool:
    """入力値を検証"""
    try:
        # 開始年月の検証
        year_int = int(year)
        if year_int < 2020 or year_int > 2030:
            print(f"エラー: 年は2020-2030の範囲で指定してください: {year}")
            return False
        
        month_int = int(month)
        if month_int < 1 or month_int > 12:
            print(f"エラー: 月は1-12の範囲で指定してください: {month}")
            return False
        
        # 終了年月の検証（指定されている場合）
        if end_year is not None and end_month is not None:
            end_year_int = int(end_year)
            if end_year_int < 2020 or end_year_int > 2030:
                print(f"エラー: 終了年は2020-2030の範囲で指定してください: {end_year}")
                return False
            
            end_month_int = int(end_month)
            if end_month_int < 1 or end_month_int > 12:
                print(f"エラー: 終了月は1-12の範囲で指定してください: {end_month}")
                return False
            
            # 開始日 <= 終了日の検証
            start_date = datetime(year_int, month_int, 1)
            end_date = datetime(end_year_int, end_month_int, 1)
            if start_date > end_date:
                print(f"エラー: 開始年月は終了年月より前にしてください: {year}/{month} > {end_year}/{end_month}")
                return False
        
        return True
        
    except ValueError:
        print(f"エラー: 年月は数値で指定してください")
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


def generate_month_range(start_year: int, start_month: int, end_year: int, end_month: int):
    """指定された期間の年月リストを生成"""
    months = []
    current_date = datetime(start_year, start_month, 1)
    end_date = datetime(end_year, end_month, 1)
    
    while current_date <= end_date:
        months.append((str(current_date.year), str(current_date.month).zfill(2)))
        current_date += relativedelta(months=1)
    
    return months


def main():
    """メイン関数"""
    try:
        # コマンドライン引数を解析
        args = parse_arguments()
        
        # コンテンツ名と期間を解析
        content, end_year, end_month = parse_content_and_dates(args)
        
        # 入力値を検証
        if not validate_input(args.year, args.month, end_year, end_month):
            sys.exit(1)
        
        # 処理対象月を決定
        if end_year and end_month:
            # 期間指定の場合
            target_months = generate_month_range(
                int(args.year), int(args.month), 
                int(end_year), int(end_month)
            )
            period_str = f"{args.year}年{args.month}月から{end_year}年{end_month}月まで（{len(target_months)}ヶ月）"
        else:
            # 単月指定の場合
            target_months = [(args.year, args.month)]
            period_str = f"{args.year}年{args.month}月"
        
        print("=" * 60)
        print("コンテンツ関連支払い明細書生成システム")
        print(f"処理対象: {period_str}")
        if content:
            print(f"対象コンテンツ: {content}")
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
        if content and not template_filter:
            template_filter = f"{content}.xlsx"
        
        # 各月に対して処理を実行
        overall_success = True
        for i, (year, month) in enumerate(target_months, 1):
            print(f"\n{'='*20} {i}/{len(target_months)}: {year}年{month}月の処理開始 {'='*20}")
            
            success = controller.process_payment_statements(
                year, 
                month, 
                send_emails,
                template_filter=template_filter,
                content_name=content
            )
            
            if success:
                print(f"✓ {year}年{month}月の処理が完了しました。")
            else:
                print(f"✗ {year}年{month}月の処理でエラーが発生しました。")
                overall_success = False
        
        success = overall_success
        
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