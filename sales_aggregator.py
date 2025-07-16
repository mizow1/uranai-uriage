import os
import pandas as pd
import openpyxl
from pathlib import Path
import re
import csv
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

class SalesAggregator:
    def __init__(self, base_path):
        self.base_path = Path(base_path)
        self.results = []
        
    def find_files_in_yearmonth_folders(self):
        """年月フォルダ内のファイルを検索"""
        files_by_platform = {
            'ameba': [],
            'rakuten': [],
            'au': [],
            'excite': [],
            'line': []
        }
        
        for folder in self.base_path.iterdir():
            if folder.is_dir() and re.match(r'\d{6}', folder.name):
                for file in folder.iterdir():
                    if file.is_file():
                        filename = file.name.lower()
                        if 'satori実績_' in filename or 'satori' in filename:
                            files_by_platform['ameba'].append(file)
                        elif 'rcms' in filename or '楽天' in filename:
                            files_by_platform['rakuten'].append(file)
                        elif 'salessummary' in filename:
                            files_by_platform['au'].append(file)
                        elif 'excite' in filename:
                            files_by_platform['excite'].append(file)
                        elif 'line' in filename and (file.suffix.lower() in ['.xls', '.xlsx']):
                            files_by_platform['line'].append(file)
        
        return files_by_platform
    
    def process_ameba_file(self, file_path):
        """ameba占い（SATORI実績）ファイルを処理"""
        try:
            # パスワード付きExcelファイルの場合、一旦スキップ
            # TODO: msoffcrypto-toolライブラリが必要
            try:
                wb = openpyxl.load_workbook(file_path)
            except:
                print(f"ameba占いファイル処理エラー: {file_path.name} - パスワード保護ファイルの可能性があります")
                return None
            
            # 「従量実績」シートを読み込み
            従量実績_sheet = wb['従量実績']
            従量実績_df = pd.DataFrame(従量実績_sheet.values)
            従量実績_df.columns = 従量実績_df.iloc[0]
            従量実績_df = 従量実績_df.drop(0).reset_index(drop=True)
            
            # 「docomo占い」シートを読み込み
            docomo占い_sheet = wb['docomo占い']
            docomo占い_df = pd.DataFrame(docomo占い_sheet.values)
            docomo占い_df.columns = docomo占い_df.iloc[0]
            docomo占い_df = docomo占い_df.drop(0).reset_index(drop=True)
            
            # C列の値が一致するもののJ列の合計額を算出
            matching_records = []
            for _, row in 従量実績_df.iterrows():
                c_value = row.iloc[2]  # C列
                matching_docomo = docomo占い_df[docomo占い_df.iloc[:, 2] == c_value]
                if not matching_docomo.empty:
                    j_sum = matching_docomo.iloc[:, 9].sum()  # J列
                    matching_records.append({
                        'content_id': c_value,
                        'amount': j_sum
                    })
            
            情報提供料合計 = sum(record['amount'] for record in matching_records)
            実績 = 情報提供料合計 / 0.3  # 30%を除算した値
            
            return {
                'platform': 'ameba',
                'file': file_path.name,
                '実績': round(実績),
                '情報提供料合計': round(情報提供料合計),
                'details': matching_records
            }
            
        except Exception as e:
            print(f"ameba占いファイル処理エラー: {file_path.name} - {str(e)}")
            return None
    
    def process_rakuten_file(self, file_path):
        """楽天占い（rcms・楽天明細）ファイルを処理"""
        try:
            # ファイル拡張子に応じてデータを読み込み
            if file_path.suffix.lower() == '.csv':
                df = pd.read_csv(file_path, encoding='utf-8')
            else:
                df = pd.read_excel(file_path)
            
            # ファイル名に応じて処理を分岐
            if 'rcms' in file_path.name.lower():
                # RCMSファイルの処理
                # L列の値「hoge_xxx」のhoge部分が一致するもののN列の値を0.7倍
                l_column = df.iloc[:, 11]  # L列
                n_column = df.iloc[:, 13]  # N列
                
                hoge_groups = {}
                for i, value in enumerate(l_column):
                    if pd.notna(value) and '_' in str(value):
                        hoge_part = str(value).split('_')[0]
                        if hoge_part not in hoge_groups:
                            hoge_groups[hoge_part] = []
                        hoge_groups[hoge_part].append(n_column.iloc[i])
                
                実績_amount = 0
                情報提供料_amount = 0
                details = []
                
                for hoge, values in hoge_groups.items():
                    group_sum = sum(pd.to_numeric(v, errors='coerce') for v in values if pd.notna(v))
                    実績_sum = group_sum / 1.1  # 1.1で除算
                    情報提供料_sum = 実績_sum * 0.7  # 70%
                    実績_amount += 実績_sum
                    情報提供料_amount += 情報提供料_sum
                    details.append({
                        'content_group': hoge,
                        '実績': 実績_sum,
                        '情報提供料': 情報提供料_sum
                    })
                
                # 「月額実績」シートの処理も追加（もしあれば）
                if file_path.suffix.lower() in ['.xlsx', '.xls']:
                    try:
                        monthly_df = pd.read_excel(file_path, sheet_name='月額実績')
                        # B列の値が一致するもののE列の合計額を追加
                        monthly_amount = monthly_df.iloc[:, 4].sum()  # E列
                        monthly_実績 = monthly_amount / 1.1
                        monthly_情報提供料 = monthly_実績 * 0.7
                        実績_amount += monthly_実績
                        情報提供料_amount += monthly_情報提供料
                        details.append({
                            'content_group': '月額実績',
                            '実績': monthly_実績,
                            '情報提供料': monthly_情報提供料
                        })
                    except:
                        pass
                        
            else:
                # 楽天明細ファイルの処理（シンプルな合計）
                # 数値列を探して合計を計算
                numeric_columns = df.select_dtypes(include=['number']).columns
                if len(numeric_columns) > 0:
                    # 最も右側の数値列を金額と仮定
                    amount_column = numeric_columns[-1]
                    total_amount = df[amount_column].sum()
                    実績_amount = total_amount / 1.1
                    情報提供料_amount = 実績_amount * 0.7
                    details = [{'content_group': '楽天明細合計', '実績': 実績_amount, '情報提供料': 情報提供料_amount}]
                else:
                    # 数値列が見つからない場合は0
                    実績_amount = 0
                    情報提供料_amount = 0
                    details = [{'content_group': '楽天明細合計', '実績': 0, '情報提供料': 0}]
            
            return {
                'platform': 'rakuten',
                'file': file_path.name,
                '実績': round(実績_amount),
                '情報提供料合計': round(情報提供料_amount),
                'details': details
            }
            
        except Exception as e:
            print(f"楽天占いファイル処理エラー: {file_path.name} - {str(e)}")
            return None
    
    def process_au_file(self, file_path):
        """au占い（SalesSummary）ファイルを処理"""
        try:
            # ファイル拡張子に応じてデータを読み込み
            if file_path.suffix.lower() == '.csv':
                # 複数のエンコーディングを試す
                for encoding in ['utf-8', 'shift_jis', 'cp932']:
                    try:
                        df = pd.read_csv(file_path, encoding=encoding)
                        break
                    except UnicodeDecodeError:
                        continue
                else:
                    print(f"au占いファイル処理エラー: {file_path.name} - エンコーディングが不明です")
                    return None
            else:
                df = pd.read_excel(file_path)
            
            # B列の値が一致するもののG列の40%の値からK列の値を引いた値を算出
            b_column = df.iloc[:, 1]  # B列（番組ID）
            g_column = df.iloc[:, 6]  # G列
            k_column = df.iloc[:, 10]  # K列
            
            # B列の値でグループ化してG列の値（実績）とG列の40%からK列を引いた値（情報提供料）を計算
            b_groups = {}
            for i, b_value in enumerate(b_column):
                if pd.notna(b_value):
                    if b_value not in b_groups:
                        b_groups[b_value] = {'g_values': [], 'k_values': []}
                    g_value = pd.to_numeric(g_column.iloc[i], errors='coerce')
                    k_value = pd.to_numeric(k_column.iloc[i], errors='coerce')
                    if pd.notna(g_value) and pd.notna(k_value):
                        b_groups[b_value]['g_values'].append(g_value)
                        b_groups[b_value]['k_values'].append(k_value)
            
            実績_amount = 0
            情報提供料_amount = 0
            details = []
            
            for b_value, values in b_groups.items():
                g_sum = sum(values['g_values'])
                k_sum = sum(values['k_values'])
                実績_sum = g_sum  # G列の値
                情報提供料_sum = (g_sum * 0.4) - k_sum  # G列の40%からK列を引いた値
                実績_amount += 実績_sum
                情報提供料_amount += 情報提供料_sum
                details.append({
                    'content_group': b_value,
                    '実績': 実績_sum,
                    '情報提供料': 情報提供料_sum
                })
            
            return {
                'platform': 'au',
                'file': file_path.name,
                '実績': round(実績_amount),
                '情報提供料合計': round(情報提供料_amount),
                'details': details
            }
            
        except Exception as e:
            print(f"au占いファイル処理エラー: {file_path.name} - {str(e)}")
            return None
    
    def process_excite_file(self, file_path):
        """excite占いファイルを処理"""
        try:
            # ファイル拡張子に応じてデータを読み込み
            if file_path.suffix.lower() == '.csv':
                # 複数のエンコーディングを試す
                for encoding in ['utf-8', 'shift_jis', 'cp932']:
                    try:
                        df = pd.read_csv(file_path, encoding=encoding)
                        break
                    except UnicodeDecodeError:
                        continue
                else:
                    print(f"excite占いファイル処理エラー: {file_path.name} - エンコーディングが不明です")
                    return None
            else:
                df = pd.read_excel(file_path)
            
            # B列の値が一致するもののF列の合計額を算出
            f_column = df.iloc[:, 5]  # F列
            実績_amount = pd.to_numeric(f_column, errors='coerce').sum()
            情報提供料_amount = 実績_amount * 0.6  # 60%
            
            return {
                'platform': 'excite',
                'file': file_path.name,
                '実績': round(実績_amount),
                '情報提供料合計': round(情報提供料_amount),
                'details': [{'content_group': 'excite合計', '実績': 実績_amount, '情報提供料': 情報提供料_amount}]
            }
            
        except Exception as e:
            print(f"excite占いファイル処理エラー: {file_path.name} - {str(e)}")
            return None
    
    def process_line_file(self, file_path):
        """LINE占いファイルを処理"""
        try:
            # 「内訳」シートを読み込み
            df = pd.read_excel(file_path, sheet_name='内訳')
            
            # F列の値が一致するもののJ列の合計額を1.1で除算
            f_column = df.iloc[:, 5]  # F列
            j_column = df.iloc[:, 9]  # J列
            
            # F列の値でグループ化してJ列とN列の合計を計算
            f_column = df.iloc[:, 5]  # F列
            j_column = df.iloc[:, 9]  # J列
            n_column = df.iloc[:, 13]  # N列
            
            f_groups = {}
            for i, f_value in enumerate(f_column):
                if pd.notna(f_value):
                    if f_value not in f_groups:
                        f_groups[f_value] = {'j_values': [], 'n_values': []}
                    f_groups[f_value]['j_values'].append(j_column.iloc[i])
                    f_groups[f_value]['n_values'].append(n_column.iloc[i])
            
            実績_amount = 0
            情報提供料_amount = 0
            details = []
            
            for f_value, values in f_groups.items():
                j_sum = sum(pd.to_numeric(v, errors='coerce') for v in values['j_values'] if pd.notna(v))
                n_sum = sum(pd.to_numeric(v, errors='coerce') for v in values['n_values'] if pd.notna(v))
                実績_sum = j_sum / 1.1
                情報提供料_sum = n_sum / 1.1
                実績_amount += 実績_sum
                情報提供料_amount += 情報提供料_sum
                details.append({
                    'content_group': f_value,
                    '実績': 実績_sum,
                    '情報提供料': 情報提供料_sum
                })
            
            return {
                'platform': 'line',
                'file': file_path.name,
                '実績': round(実績_amount),
                '情報提供料合計': round(情報提供料_amount),
                'details': details
            }
            
        except Exception as e:
            print(f"LINE占いファイル処理エラー: {file_path.name} - {str(e)}")
            return None
    
    def process_all_files(self):
        """すべてのファイルを処理"""
        files_by_platform = self.find_files_in_yearmonth_folders()
        
        # 各プラットフォームのファイルを処理
        for ameba_file in files_by_platform['ameba']:
            result = self.process_ameba_file(ameba_file)
            if result:
                self.results.append(result)
        
        for rakuten_file in files_by_platform['rakuten']:
            result = self.process_rakuten_file(rakuten_file)
            if result:
                self.results.append(result)
        
        for au_file in files_by_platform['au']:
            result = self.process_au_file(au_file)
            if result:
                self.results.append(result)
        
        for excite_file in files_by_platform['excite']:
            result = self.process_excite_file(excite_file)
            if result:
                self.results.append(result)
        
        for line_file in files_by_platform['line']:
            result = self.process_line_file(line_file)
            if result:
                self.results.append(result)
    
    def export_to_csv(self, output_path):
        """結果をCSVファイルに出力（コンテンツごと）"""
        if not self.results:
            print("処理結果がありません。")
            return
        
        # CSVファイルの作成
        with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            
            # ヘッダー
            writer.writerow(['プラットフォーム', 'ファイル名', 'コンテンツ', '実績', '情報提供料合計', '年月', '処理日時'])
            
            # データ行（コンテンツごと）
            total_実績 = 0
            total_情報提供料 = 0
            for result in self.results:
                # ファイル名から年月を抽出
                yearmonth = '不明'
                for folder in self.base_path.iterdir():
                    if folder.is_dir() and re.match(r'\d{6}', folder.name):
                        if any(result['file'] in str(f) for f in folder.iterdir()):
                            yearmonth = folder.name
                            break
                
                # 各コンテンツの詳細を出力
                for detail in result['details']:
                    実績 = detail.get('実績', 0)
                    情報提供料 = detail.get('情報提供料', 0)
                    writer.writerow([
                        result['platform'],
                        result['file'],
                        detail['content_group'],
                        実績,
                        情報提供料,
                        yearmonth,
                        datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    ])
                    total_実績 += 実績
                    total_情報提供料 += 情報提供料
                
                # プラットフォーム合計行も追加
                writer.writerow([
                    result['platform'] + '合計',
                    result['file'],
                    '合計',
                    result['実績'],
                    result['情報提供料合計'],
                    yearmonth,
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                ])
        
        print(f"結果をCSVファイルに出力しました: {output_path}")
        
        # 合計金額の表示
        print(f"全プラットフォーム合計実績: {total_実績:,.0f}円")
        print(f"全プラットフォーム合計情報提供料: {total_情報提供料:,.0f}円")

def main():
    # 設定
    base_path = r"/mnt/c/Users/OW/Dropbox/disk2とローカルの同期/占い/占い売上/履歴/ISP支払通知書"
    output_path = "sales_distribution_summary.csv"
    
    # 処理実行
    aggregator = SalesAggregator(base_path)
    aggregator.process_all_files()
    aggregator.export_to_csv(output_path)

if __name__ == "__main__":
    main()