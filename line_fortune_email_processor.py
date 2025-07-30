#!/usr/bin/env python3
"""
LINE Fortune Email Processor - メインスクリプト

LINE Fortuneの日次レポートメールを自動処理するシステム
"""

import sys
import argparse
import os
import re
import csv
from pathlib import Path
from datetime import datetime

# パッケージのパスを追加
sys.path.insert(0, str(Path(__file__).parent))

from line_fortune_processor.main_processor import LineFortuneProcessor
from line_fortune_processor.config import Config
from line_fortune_processor.logger import get_logger


def extract_date_from_filename(filename):
    """ファイル名から年月日を抽出する"""
    # 新しい形式: yyyy-mm-dd_元ファイル名.csv（先頭からの日付）
    date_pattern = r'^(\d{4})-(\d{1,2})-(\d{1,2})_'
    match = re.search(date_pattern, filename)
    if match:
        year, month, day = match.groups()
        return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
    
    # 旧形式: 元ファイル名_yyyy-mm-dd.csv（従来の検索）
    date_pattern2 = r'(\d{4})-(\d{1,2})-(\d{1,2})'
    match2 = re.search(date_pattern2, filename)
    if match2:
        year, month, day = match2.groups()
        return f"{year}-{month.zfill(2)}-{day.zfill(2)}"
    
    # 20250722のような日付パターンを検索
    date_pattern3 = r'(\d{4})(\d{2})(\d{2})'
    match3 = re.search(date_pattern3, filename)
    if match3:
        year, month, day = match3.groups()
        return f"{year}-{month}-{day}"
    
    return ""


def aggregate_all_service_data():
    """全年月のline-menu-yyyy-mmファイルからservice_nameごとの合計を一括計算してline_contents_yyyy_mm.csvを作成する"""
    try:
        from line_contents_aggregator import LineContentsAggregator
        import glob
        
        # 設定ファイルから保存場所を取得
        config = Config()
        
        # ベースパスを取得
        base_path = Path(config.get('base_path'))
        
        if not base_path.exists():
            print(f"ベースパスが見つかりません: {base_path}")
            return 1
        
        # yyyy/yyyymm/line パターンのディレクトリを検索
        line_directories = []
        year_pattern = base_path / "*" / "*" / "line"
        
        for line_dir in glob.glob(str(year_pattern)):
            line_path = Path(line_dir)
            if line_path.is_dir():
                # パスから年月を抽出
                parts = line_path.parts
                if len(parts) >= 3:
                    try:
                        year_month_str = parts[-2]  # yyyymm部分
                        year_str = parts[-3]        # yyyy部分
                        
                        # yyyymmフォーマットの検証
                        if (len(year_month_str) == 6 and year_month_str.isdigit() and
                            len(year_str) == 4 and year_str.isdigit()):
                            year = int(year_str)
                            month = int(year_month_str[4:6])
                            
                            # line-menu-yyyy-mm.csvファイルの存在確認
                            menu_filename = f"line-menu-{year}-{month:02d}.csv"
                            menu_file = line_path / menu_filename
                            
                            if menu_file.exists():
                                line_directories.append({
                                    'path': line_path,
                                    'year': year,
                                    'month': month,
                                    'menu_file': menu_file,
                                    'menu_filename': menu_filename
                                })
                    except (ValueError, IndexError):
                        continue
        
        if not line_directories:
            print("処理可能なline-menu-yyyy-mm.csvファイルが見つかりませんでした。")
            return 1
        
        # 年月順でソート
        line_directories.sort(key=lambda x: (x['year'], x['month']))
        
        print("\n" + "="*60)
        print("全年月一括サービス別売上集計機能")
        print("="*60)
        print(f"見つかったline-menu-yyyy-mm.csvファイル: {len(line_directories)}個")
        
        for dir_info in line_directories:
            print(f"  {dir_info['year']}-{dir_info['month']:02d}: {dir_info['menu_filename']}")
        
        # 実行確認
        confirm = input(f"\n{len(line_directories)}個のファイルを一括でコンテンツ別集計しますか？ (y/N): ").strip().lower()
        if confirm not in ['y', 'yes']:
            print("処理を中止しました。")
            return 0
        
        # 進捗表示のためのtqdmインポート
        try:
            from tqdm import tqdm
        except ImportError:
            # tqdmが利用できない場合のダミークラス
            class tqdm:
                def __init__(self, iterable=None, total=None, desc=None, **kwargs):
                    self.iterable = iterable or []
                    self.desc = desc
                    self.current = 0
                    self.total = total or len(self.iterable)
                    if desc:
                        print(f"{desc}: 開始")
                
                def __iter__(self):
                    return iter(self.iterable)
                
                def __enter__(self):
                    return self
                
                def __exit__(self, *args):
                    if self.desc:
                        print(f"{self.desc}: 完了")
                
                def set_description(self, desc):
                    pass
                
                def update(self, n=1):
                    self.current += n
        
        # LineContentsAggregatorの初期化
        aggregator = LineContentsAggregator()
        
        success_count = 0
        error_count = 0
        
        with tqdm(total=len(line_directories), desc="一括集計進捗", unit="ファイル") as pbar:
            for dir_info in line_directories:
                try:
                    pbar.set_description(f"処理中: {dir_info['year']}-{dir_info['month']:02d}")
                    
                    # 出力ファイル名
                    output_filename = f"line-contents-{dir_info['year']}-{dir_info['month']:02d}.csv"
                    output_path = dir_info['path'] / output_filename
                    
                    # 既存ファイルがある場合は上書き
                    if output_path.exists():
                        print(f"\n  {output_filename} は既に存在します。上書きします。")
                    
                    # LineContentsAggregatorを使用して処理
                    df = aggregator.process_line_menu_file(str(dir_info['menu_file']))
                    
                    if df.empty:
                        print(f"\n  {dir_info['menu_filename']}: 集計可能なデータが見つかりませんでした。")
                        continue
                    
                    # ファイルを保存
                    if aggregator.save_contents_file(df, str(output_path)):
                        success_count += 1
                        print(f"\n  ✓ {output_filename} - コンテンツ数: {len(df)}, 総実績: {df['実績'].sum():,}円")
                    else:
                        error_count += 1
                        print(f"\n  ✗ {output_filename}: ファイル保存に失敗しました。")
                        
                except Exception as e:
                    error_count += 1
                    print(f"\n  ✗ {dir_info['menu_filename']}: エラー - {e}")
                finally:
                    pbar.update(1)
        
        # 結果サマリー
        print("\n" + "="*60)
        print("一括集計完了")
        print("="*60)
        print(f"  処理対象ファイル数: {len(line_directories)}個")
        print(f"  成功: {success_count}個")
        print(f"  エラー: {error_count}個")
        
        return 0 if error_count == 0 else 1
        
    except Exception as e:
        print(f"一括サービス別集計でエラーが発生しました: {e}")
        return 1


def aggregate_service_data(auto_mode=False, target_year=None, target_month=None):
    """line-menu-yyyy-mmファイルからservice_nameごとの合計を計算してline_contents_yyyy_mm.csvを作成する"""
    try:
        from line_contents_aggregator import LineContentsAggregator
        
        # 設定ファイルから保存場所を取得
        config = Config()
        
        # 年月を取得（引数で指定されていればそれを使用、なければ現在の年月）
        if target_year and target_month:
            year = target_year
            month = target_month
        else:
            now = datetime.now()
            year = now.year
            month = now.month
        
        # LINE CSVファイルの保存ディレクトリを構築
        base_path = Path(config.get('base_path'))
        line_dir = base_path / str(year) / f"{year}{month:02d}" / "line"
        
        if not line_dir.exists():
            print(f"LINE CSVファイルの保存ディレクトリが見つかりません: {line_dir}")
            return 1
        
        # line-menu-yyyy-mm.csvファイルを探す
        menu_filename = f"line-menu-{year}-{month:02d}.csv"
        menu_file = line_dir / menu_filename
        
        if not menu_file.exists():
            print(f"line-menu-yyyy-mm.csvファイルが見つかりません: {menu_filename}")
            return 1
        
        print("\n" + "="*50)
        print("サービス別売上集計機能")
        print("="*50)
        print(f"集計対象: {menu_filename}")
        
        # 実行確認（自動モード時はスキップ）
        if not auto_mode:
            confirm = input(f"\nline-menu-yyyy-mmファイルからサービス別集計を実行しますか？ (y/N): ").strip().lower()
            if confirm not in ['y', 'yes']:
                print("処理を中止しました。")
                return 0
        
        # デフォルトの出力ファイル名
        default_output_filename = f"line-contents-{year}-{month:02d}.csv"
        
        # ファイル名の確認（自動モード時はデフォルトを使用）
        if not auto_mode:
            print(f"\n出力ファイル名 (デフォルト: {default_output_filename})")
            output_input = input("カスタムファイル名を入力 (Enterでデフォルト使用): ").strip()
            
            if output_input:
                if not output_input.endswith('.csv'):
                    output_input += '.csv'
                output_filename = output_input
            else:
                output_filename = default_output_filename
        else:
            output_filename = default_output_filename
            
        output_path = line_dir / output_filename
        
        # 既存ファイルの確認（自動モード時は自動上書き）
        if output_path.exists() and not auto_mode:
            overwrite = input(f"\n'{output_filename}'は既に存在します。上書きしますか？ (y/N): ").strip().lower()
            if overwrite not in ['y', 'yes']:
                print("処理を中止しました。")
                return 0
        
        # LineContentsAggregatorを使用して処理
        aggregator = LineContentsAggregator()
        df = aggregator.process_line_menu_file(str(menu_file))
        
        if df.empty:
            print("集計可能なデータが見つかりませんでした。")
            return 1
        
        # ファイルを保存
        if aggregator.save_contents_file(df, str(output_path)):
            print(f"\n集計完了: {output_filename}")
            print(f"  処理したファイル: {menu_filename}")
            print(f"  コンテンツ数: {len(df)}")
            print(f"  総実績: {df['実績'].sum():,}円")
            print(f"  総情報提供料: {df['情報提供料'].sum():,}円")
            
            # 集計結果のサマリー表示
            print("\n=== 集計結果サマリー ===")
            for _, row in df.iterrows():
                print(f"  {row['コンテンツ名']}: 実績={row['実績']:,}, 情報提供料={row['情報提供料']:,}")
            
            return 0
        else:
            print("ファイル保存に失敗しました。")
            return 1
        
    except Exception as e:
        print(f"サービス別集計でエラーが発生しました: {e}")
        return 1


def merge_csv_files():
    """設定ファイルで指定された保存場所のCSVファイルを統合する"""
    try:
        # 設定ファイルから保存場所を取得
        config = Config()
        
        # 現在の年月を取得
        now = datetime.now()
        year = now.year
        month = now.month
        
        # LINE CSVファイルの保存ディレクトリを構築
        base_path = Path(config.get('base_path'))
        line_dir = base_path / str(year) / f"{year}{month:02d}" / "line"
        
        if not line_dir.exists():
            print(f"LINE CSVファイルの保存ディレクトリが見つかりません: {line_dir}")
            return 1
        
        # yyyy-mm-dd_コンテンツ名.csvパターンのファイルのみを対象とする
        csv_files = []
        for file in line_dir.glob("*.csv"):
            filename = file.name
            # output.csv や line-menu-*.csv などの統合ファイルは除外
            # 新しいファイル名形式: yyyy-mm-dd_元ファイル名.csv
            import re
            date_pattern = r'^\d{4}-\d{2}-\d{2}_.*\.csv$'
            if (not filename.startswith("output") and 
                not filename.startswith("line-menu-") and 
                not filename.startswith("line-contents-") and
                re.match(date_pattern, filename)):
                csv_files.append(file)
        
        if not csv_files:
            print("CSVファイルが見つかりませんでした。")
            return 1
        
        print("\n" + "="*50)
        print("CSV統合機能")
        print("="*50)
        print(f"見つかったCSVファイル: {len(csv_files)}個")
        for csv_file in csv_files:
            date_str = extract_date_from_filename(csv_file.name)
            print(f"  {csv_file.name} (抽出日付: {date_str if date_str else '不明'})")
        
        # デフォルトの出力ファイル名
        default_output_filename = f"line-menu-{year}-{month:02d}.csv"
        
        # ファイル名の確認
        print(f"\n出力ファイル名 (デフォルト: {default_output_filename})")
        output_input = input("カスタムファイル名を入力 (Enterでデフォルト使用): ").strip()
        
        if output_input:
            if not output_input.endswith('.csv'):
                output_input += '.csv'
            output_filename = output_input
        else:
            output_filename = default_output_filename
            
        output_path = line_dir / output_filename
        
        # 既存ファイルの確認（自動モード時は自動上書き）
        if output_path.exists() and not auto_mode:
            overwrite = input(f"\n'{output_filename}'は既に存在します。上書きしますか？ (y/N): ").strip().lower()
            if overwrite not in ['y', 'yes']:
                print("処理を中止しました。")
                return 0
        
        # 実行確認
        print(f"\n統合設定:")
        print(f"  対象ファイル数: {len(csv_files)}個")
        print(f"  出力ファイル名: {output_filename}")
        confirm = input("\nこの設定でCSV統合を実行しますか？ (y/N): ").strip().lower()
        if confirm not in ['y', 'yes']:
            print("処理を中止しました。")
            return 0
        
        all_rows = []
        
        # ヘッダー行を準備
        header_added = False
        
        # 進捗表示のためのtqdmインポート
        try:
            from tqdm import tqdm
        except ImportError:
            # tqdmが利用できない場合のダミークラス
            class tqdm:
                def __init__(self, iterable=None, total=None, desc=None, **kwargs):
                    self.iterable = iterable or []
                    self.desc = desc
                    if desc:
                        print(f"{desc}: 開始")
                
                def __iter__(self):
                    return iter(self.iterable)
                
                def __enter__(self):
                    return self
                
                def __exit__(self, *args):
                    if self.desc:
                        print(f"{self.desc}: 完了")
        
        with tqdm(total=len(csv_files), desc="CSVファイル統合進捗", unit="ファイル") as pbar:
            for csv_file in csv_files:
                try:
                    pbar.set_description(f"処理中: {csv_file.name}")
                    date_str = extract_date_from_filename(csv_file.name)
                    
                    with open(csv_file, 'r', encoding='utf-8') as f:
                        reader = csv.reader(f)
                        rows = list(reader)
                        
                        # ヘッダー行の処理
                        if rows and not header_added:
                            # 最初のファイルの場合、ヘッダー行に日付列を追加
                            header_row = ["日付"] + rows[0]
                            all_rows.append(header_row)
                            header_added = True
                        
                        # データ行の処理（1行目のヘッダーを除外）
                        data_rows = rows[1:] if rows else []
                        for row in data_rows:
                            # 元の行の先頭に日付を追加
                            enhanced_row = [date_str] + row
                            all_rows.append(enhanced_row)
                            
                        print(f"  {csv_file.name}: {len(data_rows)}行を読み込み（ヘッダー除外）")
                    
                except Exception as e:
                    print(f"  エラー: {csv_file.name}の読み込みに失敗 - {e}")
                finally:
                    pbar.update(1)
        
        if not all_rows:
            print("統合可能なデータが見つかりませんでした。")
            return 1
        
        # 統合ファイルを書き込み
        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerows(all_rows)
        
        print(f"\n統合完了: {output_filename}")
        print(f"  統合したファイル数: {len(csv_files)}個")
        print(f"  総行数: {len(all_rows)}行（ヘッダー1行 + データ{len(all_rows)-1}行）")
        
        return 0
        
    except Exception as e:
        print(f"CSV統合でエラーが発生しました: {e}")
        return 1


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
    
    parser.add_argument(
        "--merge-csvs",
        action="store_true",
        help="同フォルダ内の全CSVファイルを統合してline-menu-yyyy-mmファイルを作成"
    )
    
    parser.add_argument(
        "--aggregate-services",
        action="store_true",
        help="同フォルダのline-menu-yyyy-mm.csvファイルからservice_nameごとの合計を計算してline-contents-yyyy-mm.csvを作成"
    )
    
    parser.add_argument(
        "--year",
        "-y",
        type=int,
        help="処理対象の年を指定 (例: 2025)"
    )
    
    parser.add_argument(
        "--month",
        "-m",
        type=int,
        help="処理対象の月を指定 (例: 7)"
    )
    
    parser.add_argument(
        "--all",
        "-a",
        action="store_true",
        help="全年月のline-menu-yyyy-mm.csvファイルを一括でコンテンツ別集計処理"
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
    
    # CSV統合機能
    if args.merge_csvs:
        return merge_csv_files()
    
    # サービス別集計機能
    if args.aggregate_services:
        # --allオプションが指定されている場合は全年月を一括処理
        if args.all:
            # --allと--year/--monthの併用をチェック
            if args.year or args.month:
                print("エラー: --allオプションは--year/-yや--month/-mと併用できません")
                return 1
            return aggregate_all_service_data()
        
        # 年月が指定されている場合は検証
        if args.year and not (1900 <= args.year <= 2100):
            print("エラー: 年は1900-2100の範囲で指定してください")
            return 1
        if args.month and not (1 <= args.month <= 12):
            print("エラー: 月は1-12の範囲で指定してください")
            return 1
        
        return aggregate_service_data(auto_mode=False, target_year=args.year, target_month=args.month)
    
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

    # CSV統合確認
    merge_confirm = input("\nメール処理完了後、同フォルダ内のCSVファイルを統合しますか？ (y/N): ").strip().lower()
    do_merge_csvs = merge_confirm in ['y', 'yes']
    
    # サービス別集計確認
    aggregate_confirm = input("CSVファイル統合後、line-menu-yyyy-mmファイルからサービス別売上集計を行いますか？ (y/N): ").strip().lower()
    do_aggregate_services = aggregate_confirm in ['y', 'yes']

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
            logger.info("メール処理が完了しました")
        
        # 結果の表示
        logger.info("統計情報を取得中...")
        stats = processor.get_stats()
        logger.info("統計情報を取得完了")
        
        if not args.dry_run:
            logger.info("処理結果を表示中...")
            print(f"処理結果:")
            print(f"  処理したメール数: {stats['emails_processed']}")
            print(f"  成功したメール数: {stats['emails_success']}")
            print(f"  エラーが発生したメール数: {stats['emails_error']}")
            print(f"  保存したファイル数: {stats['files_saved']}")
            print(f"  作成した統合ファイル数: {stats['consolidations_created']}")
            logger.info("処理結果の表示完了")
        
        # メール処理後のCSV統合実行（既に内部で完了している場合はスキップ）
        if success and do_merge_csvs and not args.dry_run:
            if stats.get('consolidations_created', 0) > 0:
                print("\n" + "="*50)
                print("CSV統合処理は既に完了しています")
                print(f"作成した統合ファイル数: {stats['consolidations_created']}")
                print("="*50)
            else:
                print("\n" + "="*50)
                print("CSV統合処理を開始します")
                print("="*50)
                merge_result = merge_csv_files()
                if merge_result != 0:
                    print("CSV統合処理でエラーが発生しました")
                    return 1
        
        # CSV統合後のサービス別集計実行
        if success and do_aggregate_services and not args.dry_run:
            print("\n" + "="*50)
            print("サービス別売上集計処理を開始します")
            print("="*50)
            aggregate_result = aggregate_service_data(auto_mode=True)
            if aggregate_result != 0:
                print("サービス別集計処理でエラーが発生しました")
                return 1
        
        logger.info("処理完了、プログラム終了")
        return 0 if success else 1
        
    except KeyboardInterrupt:
        logger = get_logger()
        logger.info("ユーザーによって処理が中断されました")
        return 1
        
    except Exception as e:
        logger = get_logger()
        logger.error(f"予期しないエラーが発生しました: {e}", exception=e)
        return 1
    
    finally:
        # 確実にプロセスを終了
        logger = get_logger()
        logger.info("プログラム終了処理を実行中...")


if __name__ == "__main__":
    try:
        exit_code = main()
        print(f"プログラムが正常終了しました (exit_code: {exit_code})")
        sys.exit(exit_code)
    except Exception as e:
        print(f"プログラム実行中にエラーが発生しました: {e}")
        sys.exit(1)