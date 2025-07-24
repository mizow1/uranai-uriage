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


def aggregate_service_data(auto_mode=False):
    """同フォルダのCSVデータからservice_nameごとの合計を計算してline_contents_yyyy_mm.csvを作成する"""
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
        
        # yyyy-mm-dd_サービス名.csvパターンのファイルのみを対象とする
        csv_files = []
        for file in line_dir.glob("*.csv"):
            filename = file.name
            # yyyy-mm-dd_サービス名.csvパターンのファイルのみを対象
            import re
            date_pattern = r'^\d{4}-\d{2}-\d{2}_.*\.csv$'
            if re.match(date_pattern, filename):
                csv_files.append(file)
        
        if not csv_files:
            print("yyyy-mm-dd_サービス名.csv形式のファイルが見つかりませんでした。")
            return 1
        
        print("\n" + "="*50)
        print("サービス別売上集計機能")
        print("="*50)
        print(f"集計対象: yyyy-mm-dd_サービス名.csv形式のファイル ({len(csv_files)}個)")
        
        # ファイル一覧表示
        for csv_file in csv_files:
            print(f"  - {csv_file.name}")
        
        # 実行確認（自動モード時はスキップ）
        if not auto_mode:
            confirm = input(f"\nこれらのファイルを統合してサービス別集計を実行しますか？ (y/N): ").strip().lower()
            if confirm not in ['y', 'yes']:
                print("処理を中止しました。")
                return 0
        
        # マッピングファイルを読み込み
        mapping_file = Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\line-reiwa-contents-menu.csv")
        content_mapping = {}
        
        if mapping_file.exists():
            try:
                with open(mapping_file, 'r', encoding='utf-8') as f:
                    for line in f:
                        parts = line.strip().split(',')
                        if len(parts) == 2:
                            content_mapping[parts[0]] = parts[1]
                print(f"マッピングファイルを読み込みました: {len(content_mapping)}件")
            except Exception as e:
                print(f"マッピングファイルの読み込みエラー: {e}")
        else:
            print("マッピングファイルが見つかりません。細分化機能は無効です。")
        
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
        
        # 集計処理を実行（全CSVファイルを統合）
        service_totals = {}
        total_rows = 0
        processed_files = 0
        
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
        
        with tqdm(total=len(csv_files), desc="CSV統合進捗", unit="ファイル") as pbar:
            for csv_file in csv_files:
                try:
                    pbar.set_description(f"処理中: {csv_file.name}")
                    print(f"\n処理中: {csv_file.name}")
                
                    with open(csv_file, 'r', encoding='utf-8') as f:
                        reader = csv.reader(f)
                        header = next(reader, None)
                        
                        if not header:
                            print(f"  スキップ: {csv_file.name} は空のファイルです")
                            continue
                        
                        # ヘッダーの確認（最初のファイルのみ）
                        if processed_files == 0:
                            print(f"CSVヘッダー: {header}")
                        
                        # LINE CSVの場合の列を特定
                        if 'service_name' in header:
                            # LINE CSVファイルの場合
                            try:
                                service_name_idx = header.index('service_name')
                                item_code_idx = header.index('item_code') if 'item_code' in header else -1
                                ios_amount_idx = header.index('ios_paid_amount') if 'ios_paid_amount' in header else -1
                                android_amount_idx = header.index('android_paid_amount') if 'android_paid_amount' in header else -1
                                web_amount_idx = header.index('web_paid_amount') if 'web_paid_amount' in header else -1
                            except ValueError as e:
                                print(f"  スキップ: {csv_file.name} - LINE CSV形式の必要な列が見つかりません: {e}")
                                continue
                        else:
                            # 従来のCSV形式の場合
                            try:
                                platform_idx = header.index('プラットフォーム') if 'プラットフォーム' in header else 0
                                content_idx = header.index('コンテンツ') if 'コンテンツ' in header else 2
                                amount_idx = header.index('実績') if '実績' in header else 3
                                fee_idx = header.index('情報提供料合計') if '情報提供料合計' in header else 4
                            except ValueError as e:
                                print(f"  スキップ: {csv_file.name} - 必要な列が見つかりません: {e}")
                                continue
                        
                        file_rows = 0
                        for row in reader:
                            if 'service_name' in header:
                                # LINE CSVファイルの処理
                                if len(row) <= service_name_idx:
                                    continue
                                
                                service_name = row[service_name_idx].strip()
                                item_code = row[item_code_idx].strip() if item_code_idx >= 0 else ""
                                
                                # 金額の合計計算（iOS + Android + Web）
                                try:
                                    ios_amount = int(row[ios_amount_idx]) if ios_amount_idx >= 0 and row[ios_amount_idx].strip() else 0
                                    android_amount = int(row[android_amount_idx]) if android_amount_idx >= 0 and row[android_amount_idx].strip() else 0  
                                    web_amount = int(row[web_amount_idx]) if web_amount_idx >= 0 and row[web_amount_idx].strip() else 0
                                    total_amount = ios_amount + android_amount + web_amount
                                except ValueError:
                                    continue
                                
                                # サービス名の決定
                                if service_name == '最強の姓名鑑定師軍団' and item_code:
                                    # マッピングファイルから細分化グループを検索
                                    group_name = content_mapping.get(item_code, 'その他')
                                    final_service_name = f"最強の姓名鑑定師軍団_{group_name}"
                                else:
                                    final_service_name = service_name
                                
                                # 集計
                                if final_service_name not in service_totals:
                                    service_totals[final_service_name] = {'amount': 0, 'fee': 0, 'count': 0}
                                
                                service_totals[final_service_name]['amount'] += total_amount
                                service_totals[final_service_name]['fee'] += 0  # LINE CSVには情報提供料の概念がない
                                service_totals[final_service_name]['count'] += 1
                                total_rows += 1
                                file_rows += 1
                            
                            else:
                                # 従来のCSV形式の処理
                                if len(row) <= max(platform_idx, content_idx, amount_idx, fee_idx):
                                    continue
                                
                                platform = row[platform_idx].strip()
                                content_name = row[content_idx].strip()
                                
                                try:
                                    amount = int(row[amount_idx]) if row[amount_idx].strip() else 0
                                    fee = int(row[fee_idx]) if row[fee_idx].strip() else 0
                                except ValueError:
                                    continue
                                
                                # サービス名の決定
                                service_name = content_name if content_name else platform
                                
                                # 集計
                                if service_name not in service_totals:
                                    service_totals[service_name] = {'amount': 0, 'fee': 0, 'count': 0}
                                
                                service_totals[service_name]['amount'] += amount
                                service_totals[service_name]['fee'] += fee
                                service_totals[service_name]['count'] += 1
                                total_rows += 1
                                file_rows += 1
                        
                        print(f"  完了: {file_rows}行を処理")
                        processed_files += 1
                    
                except Exception as e:
                    print(f"  エラー: {csv_file.name}の読み込みに失敗 - {e}")
                finally:
                    pbar.update(1)
        
        if not service_totals:
            print("集計可能なデータが見つかりませんでした。")
            return 1
        
        # 結果をCSVに書き込み
        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['サービス名', '売上合計', '情報提供料合計', 'レコード数'])
            
            # サービス名でソート
            for service_name in sorted(service_totals.keys()):
                data = service_totals[service_name]
                writer.writerow([
                    service_name,
                    data['amount'],
                    data['fee'],
                    data['count']
                ])
        
        print(f"\n集計完了: {output_filename}")
        print(f"  処理したCSVファイル数: {processed_files}")
        print(f"  処理したレコード数: {total_rows}")
        print(f"  サービス数: {len(service_totals)}")
        
        # 集計結果のサマリー表示
        print("\n=== 集計結果サマリー ===")
        for service_name in sorted(service_totals.keys()):
            data = service_totals[service_name]
            print(f"  {service_name}: 売上={data['amount']:,}, 情報提供料={data['fee']:,}, 件数={data['count']}")
        
        return 0
        
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
        help="同フォルダのyyyy-mm-dd_サービス名.csvファイルからservice_nameごとの合計を計算してline-contents-yyyy-mm.csvを作成"
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
        return aggregate_service_data()
    
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
    aggregate_confirm = input("CSVファイル統合後、サービス別売上集計を行いますか？ (y/N): ").strip().lower()
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