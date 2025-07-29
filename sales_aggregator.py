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
        
        # 年フォルダ（4桁）内の月フォルダ（6桁）を検索
        for year_folder in self.base_path.iterdir():
            if year_folder.is_dir() and re.match(r'\d{4}', year_folder.name):
                for month_folder in year_folder.iterdir():
                    if month_folder.is_dir() and re.match(r'\d{6}', month_folder.name):
                        # 月フォルダ直下のファイルを検索
                        for file in month_folder.iterdir():
                            if file.is_file():
                                filename = file.name.lower()
                                if '【株式会社アウトワード御中】satori実績_' in file.name or 'satori実績_' in filename:
                                    files_by_platform['ameba'].append(file)
                                elif 'rcms' in filename:
                                    files_by_platform['rakuten'].append(file)
                                elif 'salessummary' in filename:
                                    files_by_platform['au'].append(file)
                                elif 'excite' in filename:
                                    files_by_platform['excite'].append(file)
                            # サブフォルダも検索（LINEファイル用）
                            elif file.is_dir():
                                for subfile in file.iterdir():
                                    if subfile.is_file():
                                        subfilename = subfile.name.lower()
                                        if subfilename.startswith('line-contents-') and (subfile.suffix.lower() in ['.xls', '.xlsx', '.csv']):
                                            files_by_platform['line'].append(subfile)
        
        return files_by_platform
    
    def process_ameba_file(self, file_path):
        """ameba占い（SATORI実績）ファイルを処理"""
        try:
            # パスワード保護されたファイルの解除を試行
            wb = None
            try:
                # 通常の読み込みを試行
                wb = openpyxl.load_workbook(file_path, data_only=True)
            except Exception as e:
                if "password" in str(e).lower() or "protected" in str(e).lower() or "zip file" in str(e).lower():
                    # パスワード保護解除を試行
                    try:
                        import msoffcrypto
                        import io
                        
                        # よく使われるパスワードを試行
                        common_passwords = ['', 'password', '123456', '000000', 'admin', 'user']
                        
                        for password in common_passwords:
                            try:
                                with open(file_path, 'rb') as file:
                                    office_file = msoffcrypto.OfficeFile(file)
                                    office_file.load_key(password=password if password else None)
                                    
                                    decrypted = io.BytesIO()
                                    office_file.save(decrypted)
                                    decrypted.seek(0)
                                    
                                    wb = openpyxl.load_workbook(decrypted, data_only=True)
                                    print(f"ameba占いファイル: {file_path.name} - パスワード '{password}' で解除成功")
                                    break
                            except:
                                continue
                        
                        if wb is None:
                            print(f"ameba占いファイル処理エラー: {file_path.name} - パスワード解除失敗")
                            return None
                            
                    except ImportError:
                        print(f"ameba占いファイル処理エラー: {file_path.name} - msoffcrypto-toolが必要です")
                        return None
                    except Exception as decrypt_error:
                        print(f"ameba占いファイル処理エラー: {file_path.name} - パスワード保護解除エラー: {str(decrypt_error)}")
                        return None
                else:
                    print(f"ameba占いファイル処理エラー: {file_path.name} - {str(e)}")
                    return None
            
            # 新仕様：各シートから情報提供料を集計し、同一コンテンツは合算する
            content_groups = {}
            
            # 1. 「従量実績」シートでC列の値が一致するもののJ列の合計額を算出
            try:
                従量実績_sheet = wb['従量実績']
                従量実績_df = pd.DataFrame(従量実績_sheet.values)
                従量実績_df.columns = 従量実績_df.iloc[0]
                従量実績_df = 従量実績_df.drop(0).reset_index(drop=True)
                
                for _, row in 従量実績_df.iterrows():
                    c_value = row.iloc[2]  # C列
                    j_value = row.iloc[9]  # J列
                    if pd.notna(c_value) and pd.notna(j_value):
                        j_numeric = pd.to_numeric(j_value, errors='coerce')
                        if pd.notna(j_numeric):
                            if c_value not in content_groups:
                                content_groups[c_value] = 0
                            content_groups[c_value] += j_numeric
                            
            except Exception as e:
                print(f"従量実績シート処理エラー: {e}")
            
            # 2. 「docomo占い」シートでC列の値が一致するもののJ列の合計額を算出
            try:
                docomo占い_sheet = wb['docomo占い']
                docomo占い_df = pd.DataFrame(docomo占い_sheet.values)
                docomo占い_df.columns = docomo占い_df.iloc[0]
                docomo占い_df = docomo占い_df.drop(0).reset_index(drop=True)
                
                for _, row in docomo占い_df.iterrows():
                    c_value = row.iloc[2]  # C列
                    j_value = row.iloc[9]  # J列
                    if pd.notna(c_value) and pd.notna(j_value):
                        j_numeric = pd.to_numeric(j_value, errors='coerce')
                        if pd.notna(j_numeric):
                            if c_value not in content_groups:
                                content_groups[c_value] = 0
                            content_groups[c_value] += j_numeric
                            
            except Exception as e:
                print(f"docomo占いシート処理エラー: {e}")
            
            # 3. 「月額実績」シートのB列の値が一致するもののE列の合計額を算出
            try:
                月額実績_sheet = wb['月額実績']
                月額実績_df = pd.DataFrame(月額実績_sheet.values)
                月額実績_df.columns = 月額実績_df.iloc[0]
                月額実績_df = 月額実績_df.drop(0).reset_index(drop=True)
                
                for _, row in 月額実績_df.iterrows():
                    b_value = row.iloc[1]  # B列
                    e_value = row.iloc[4]  # E列
                    if pd.notna(b_value) and pd.notna(e_value):
                        e_numeric = pd.to_numeric(e_value, errors='coerce')
                        if pd.notna(e_numeric):
                            if b_value not in content_groups:
                                content_groups[b_value] = 0
                            content_groups[b_value] += e_numeric
                            
            except Exception as e:
                print(f"月額実績シート処理エラー（存在しない可能性）: {e}")
            
            # 合計値を計算
            情報提供料合計 = sum(content_groups.values())
            実績合計 = 情報提供料合計 / 0.3 if 情報提供料合計 > 0 else 0  # 30%を除算した値
            
            # details配列の作成（コンテンツごと）
            details = []
            for content_name, 情報提供料 in content_groups.items():
                実績 = 情報提供料 / 0.3 if 情報提供料 > 0 else 0
                details.append({
                    'content_group': str(content_name),
                    '実績': round(実績),
                    '情報提供料': round(情報提供料)
                })
            
            return {
                'platform': 'ameba',
                'file': file_path.name,
                '実績': round(実績合計),
                '情報提供料合計': round(情報提供料合計),
                'details': details
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
            
            # ファイル名にrcmsを含むファイルのみ処理
            if 'rcms' in file_path.name.lower():
                # RCMSファイルの処理
                # L列の値「hoge_xxx」のhoge部分が一致するもののN列の値で計算
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
                    実績_sum = group_sum / 1.1  # N列の値の合計額を1.1で除算
                    情報提供料_sum = 実績_sum * 0.725  # 実績に0.725を乗算したものを「情報提供料合計」
                    実績_amount += 実績_sum
                    情報提供料_amount += 情報提供料_sum
                    details.append({
                        'content_group': hoge,
                        '実績': round(実績_sum),
                        '情報提供料': round(情報提供料_sum)
                    })
                
                # 「月額実績」シートの処理も追加（もしあれば）
                if file_path.suffix.lower() in ['.xlsx', '.xls']:
                    try:
                        monthly_df = pd.read_excel(file_path, sheet_name='月額実績')
                        # B列の値が一致するもののE列の合計額を追加
                        monthly_amount = monthly_df.iloc[:, 4].sum()  # E列
                        monthly_実績 = monthly_amount / 1.1
                        monthly_情報提供料 = monthly_実績 * 0.725
                        実績_amount += monthly_実績
                        情報提供料_amount += monthly_情報提供料
                        details.append({
                            'content_group': '月額実績',
                            '実績': round(monthly_実績),
                            '情報提供料': round(monthly_情報提供料)
                        })
                    except:
                        pass
                        
            else:
                # rcmsファイル以外はスキップ
                print(f"楽天占いファイル処理スキップ: {file_path.name} - rcmsファイルではありません")
                return None
            
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
            # Excelファイルの場合、暗号化チェック
            if file_path.suffix.lower() in ['.xlsx', '.xls']:
                import subprocess
                try:
                    result = subprocess.run(['file', str(file_path)], capture_output=True, text=True, timeout=10)
                    if result.stdout and ('encrypted' in result.stdout.lower() or 'cdfv2 encrypted' in result.stdout.lower()):
                        print(f"au占いファイルスキップ: {file_path.name} - まだ暗号化されています")
                        return None
                except (subprocess.TimeoutExpired, FileNotFoundError, UnicodeDecodeError):
                    # subprocessでエラーが発生した場合はそのまま続行
                    pass
            
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
            elif file_path.suffix.lower() in ['.xlsx', '.xls']:
                # Excelファイルの処理
                df = None
                try:
                    if file_path.suffix.lower() == '.xlsx':
                        df = pd.read_excel(file_path, engine='openpyxl')
                    else:
                        # .xlsファイルの場合は複数のエンジンを試す
                        engines_to_try = ['xlrd', 'openpyxl', 'calamine']
                        for engine in engines_to_try:
                            try:
                                df = pd.read_excel(file_path, engine=engine)
                                break
                            except Exception as engine_error:
                                continue
                        
                        if df is None:
                            print(f"au占いファイル処理エラー: {file_path.name} - .xlsファイルの読み込みエラー")
                            return None
                except Exception as e:
                    if "password" in str(e).lower() or "protected" in str(e).lower() or "zip file" in str(e).lower():
                        # パスワード保護解除を試行
                        try:
                            import msoffcrypto
                            import io
                            
                            common_passwords = ['', 'password', '123456', '000000', 'admin', 'user']
                            
                            for password in common_passwords:
                                try:
                                    with open(file_path, 'rb') as file:
                                        office_file = msoffcrypto.OfficeFile(file)
                                        office_file.load_key(password=password if password else None)
                                        
                                        decrypted = io.BytesIO()
                                        office_file.save(decrypted)
                                        decrypted.seek(0)
                                        
                                        df = pd.read_excel(decrypted, engine='openpyxl')
                                        print(f"au占いファイル: {file_path.name} - パスワード '{password}' で解除成功")
                                        break
                                except:
                                    continue
                            
                            if df is None:
                                print(f"au占いファイル処理エラー: {file_path.name} - パスワード解除失敗")
                                return None
                                
                        except ImportError:
                            print(f"au占いファイル処理エラー: {file_path.name} - msoffcrypto-toolが必要です")
                            return None
                        except Exception as decrypt_error:
                            print(f"au占いファイル処理エラー: {file_path.name} - パスワード保護解除エラー: {str(decrypt_error)}")
                            return None
                    else:
                        print(f"au占いファイル処理エラー: {file_path.name} - {str(e)}")
                        return None
            else:
                print(f"au占いファイル処理エラー: {file_path.name} - サポートされていないファイル形式")
                return None
            
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
            
            # OWD_*コンテンツが含まれているかチェック
            has_owd_content = any(str(b_value).startswith('OWD_') for b_value in b_groups.keys() if pd.notna(b_value))
            platform_name = 'mediba' if has_owd_content else 'au'
            
            for b_value, values in b_groups.items():
                g_sum = sum(values['g_values'])
                k_sum = sum(values['k_values'])
                実績_sum = g_sum  # G列の値
                情報提供料_sum = ((g_sum * 0.4) - k_sum) * 1.1  # G列の40%からK列を引いた値に1.1を乗算
                実績_amount += 実績_sum
                情報提供料_amount += 情報提供料_sum
                details.append({
                    'content_group': b_value,
                    '実績': round(実績_sum),
                    '情報提供料': round(情報提供料_sum)
                })
            
            return {
                'platform': platform_name,
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
                # 複数のエンコーディングと区切り文字を試す
                encodings = ['shift_jis', 'cp932', 'utf-8', 'euc-jp']  # 日本語ファイルなのでshift_jisを最初に
                separators = [',', ';', '\t', '|']
                
                df = None
                for encoding in encodings:
                    for sep in separators:
                        try:
                            # ヘッダー行をスキップして読み込み（説明文がある場合）
                            df = pd.read_csv(file_path, encoding=encoding, sep=sep, 
                                           skiprows=3, on_bad_lines='skip', engine='python')
                            # データが正常に読み込めたかチェック（列数が5以上で、データ行がある）
                            if df.shape[1] >= 5 and df.shape[0] > 0:
                                # さらに数値列があるかチェック
                                numeric_cols = df.select_dtypes(include=['number']).columns
                                if len(numeric_cols) > 0:
                                    break
                        except (UnicodeDecodeError, pd.errors.ParserError) as e:
                            continue
                    if df is not None and df.shape[1] >= 5:
                        break
                
                if df is None or df.shape[1] < 5:
                    print(f"excite占いファイル処理エラー: {file_path.name} - CSVファイルの読み込みエラー")
                    return None
            elif file_path.suffix.lower() in ['.xlsx', '.xls']:
                try:
                    if file_path.suffix.lower() == '.xlsx':
                        df = pd.read_excel(file_path, engine='openpyxl')
                    else:
                        try:
                            df = pd.read_excel(file_path, engine='openpyxl')
                        except:
                            print(f"excite占いファイル処理エラー: {file_path.name} - .xlsファイルの読み込みエラー")
                            return None
                except Exception as e:
                    print(f"excite占いファイル処理エラー: {file_path.name} - {str(e)}")
                    return None
            else:
                print(f"excite占いファイル処理エラー: {file_path.name} - サポートされていないファイル形式")
                return None
            
            # B列の値が一致するもののF列の合計額を算出
            b_column = df.iloc[:, 1]  # B列
            f_column = df.iloc[:, 5]  # F列
            
            # B列の値でグループ化してF列の合計を計算
            b_groups = {}
            for i, b_value in enumerate(b_column):
                if pd.notna(b_value):
                    if b_value not in b_groups:
                        b_groups[b_value] = []
                    f_value = pd.to_numeric(f_column.iloc[i], errors='coerce')
                    if pd.notna(f_value):
                        b_groups[b_value].append(f_value)
            
            実績_amount = 0
            情報提供料_amount = 0
            details = []
            
            for b_value, f_values in b_groups.items():
                f_sum = sum(f_values)
                実績_sum = f_sum
                情報提供料_sum = f_sum * 0.6  # 60%
                実績_amount += 実績_sum
                情報提供料_amount += 情報提供料_sum
                details.append({
                    'content_group': str(b_value),
                    '実績': round(実績_sum),
                    '情報提供料': round(情報提供料_sum)
                })
            
            return {
                'platform': 'excite',
                'file': file_path.name,
                '実績': round(実績_amount),
                '情報提供料合計': round(情報提供料_amount),
                'details': details
            }
            
        except Exception as e:
            print(f"excite占いファイル処理エラー: {file_path.name} - {str(e)}")
            return None
    
    def process_line_file(self, file_path):
        """LINE占いファイルを処理"""
        try:
            # CSVファイルとExcelファイルで処理を分ける
            if file_path.suffix.lower() == '.csv':
                # CSVファイルの処理（2022年形式）
                try:
                    df = pd.read_csv(file_path, encoding='utf-8-sig')
                except UnicodeDecodeError:
                    df = pd.read_csv(file_path, encoding='shift_jis')
                
                # CSVファイルは既に集計済みの形式（コンテンツ名,実績,情報提供料）
                if len(df.columns) >= 3:
                    content_column = df.iloc[:, 0]  # コンテンツ名列
                    jisseki_column = df.iloc[:, 1]  # 実績列
                    info_fee_column = df.iloc[:, 2]  # 情報提供料列
                    
                    実績_amount = 0
                    情報提供料_amount = 0
                    details = []
                    
                    for i, content_name in enumerate(content_column):
                        if pd.notna(content_name):
                            jisseki_value = pd.to_numeric(jisseki_column.iloc[i], errors='coerce')
                            info_fee_value = pd.to_numeric(info_fee_column.iloc[i], errors='coerce')
                            
                            if pd.notna(jisseki_value) and pd.notna(info_fee_value):
                                実績_amount += jisseki_value
                                情報提供料_amount += info_fee_value
                                details.append({
                                    'content_group': str(content_name),
                                    '実績': round(jisseki_value),
                                    '情報提供料': round(info_fee_value)
                                })
                    
                    return {
                        'platform': 'line',
                        'file': file_path.name,
                        '実績': round(実績_amount),
                        '情報提供料合計': round(情報提供料_amount),
                        'details': details
                    }
                else:
                    print(f"LINE占いCSVファイル処理エラー: {file_path.name} - 列数が不足しています")
                    return None
            
            else:
                # Excelファイルの処理（2023年2月以降形式）
                # 「内訳」シートを読み込み
                try:
                    if file_path.suffix.lower() == '.xlsx':
                        df = pd.read_excel(file_path, sheet_name='内訳', engine='openpyxl')
                    else:
                        # .xlsファイルの場合は最初にxlrdを試す
                        df = None
                        try:
                            # xlrdで試す（.xlsファイルの標準的なライブラリ）
                            df = pd.read_excel(file_path, sheet_name='内訳', engine='xlrd')
                        except Exception as xlrd_error:
                            try:
                                # openpyxlで試す
                                df = pd.read_excel(file_path, sheet_name='内訳', engine='openpyxl')
                            except Exception as openpyxl_error:
                                try:
                                    # calamine（新しいライブラリ）で試す
                                    df = pd.read_excel(file_path, sheet_name='内訳', engine='calamine')
                                except Exception as calamine_error:
                                    print(f"LINE占いファイル処理エラー: {file_path.name} - .xlsファイルの読み込みエラー (xlrd: {str(xlrd_error)[:50]}, openpyxl: {str(openpyxl_error)[:50]})")
                                    return None
                except Exception as e:
                    print(f"LINE占いファイル処理エラー: {file_path.name} - {str(e)}")
                    return None
                
                # 1行目で必要な列を特定
                item_name_column_index = None
                standard_amount_column_index = None
                rs_amount_column_index = None
                
                for i, col_name in enumerate(df.columns):
                    if pd.notna(col_name):
                        col_name_str = str(col_name)
                        if 'アイテム名' in col_name_str:
                            item_name_column_index = i
                        elif '基準額（税込）' in col_name_str:
                            standard_amount_column_index = i
                        elif 'RS金額' in col_name_str:
                            rs_amount_column_index = i
                
                if item_name_column_index is None:
                    print(f"LINE占いファイル処理エラー: {file_path.name} - 「アイテム名」列が見つかりません")
                    return None
                
                if standard_amount_column_index is None:
                    print(f"LINE占いファイル処理エラー: {file_path.name} - 「基準額（税込）」列が見つかりません")
                    return None
                
                if rs_amount_column_index is None:
                    print(f"LINE占いファイル処理エラー: {file_path.name} - 「RS金額」列が見つかりません")
                    return None
                
                item_name_column = df.iloc[:, item_name_column_index]  # アイテム名列
                standard_amount_column = df.iloc[:, standard_amount_column_index]  # 基準額（税込）列
                rs_amount_column = df.iloc[:, rs_amount_column_index]  # RS金額列
                
                # アイテム名でグループ化して基準額（税込）とRS金額の合計を計算
                item_groups = {}
                for i, item_name in enumerate(item_name_column):
                    if pd.notna(item_name):
                        if item_name not in item_groups:
                            item_groups[item_name] = {'standard_values': [], 'rs_values': []}
                        
                        standard_value = pd.to_numeric(standard_amount_column.iloc[i], errors='coerce')
                        rs_value = pd.to_numeric(rs_amount_column.iloc[i], errors='coerce')
                        
                        if pd.notna(standard_value):
                            item_groups[item_name]['standard_values'].append(standard_value)
                        if pd.notna(rs_value):
                            item_groups[item_name]['rs_values'].append(rs_value)
                
                実績_amount = 0
                情報提供料_amount = 0
                details = []
                
                for item_name, values in item_groups.items():
                    standard_sum = sum(values['standard_values'])
                    rs_sum = sum(values['rs_values'])
                    
                    実績_sum = standard_sum / 1.1 if standard_sum > 0 else 0
                    情報提供料_sum = rs_sum / 1.1 if rs_sum > 0 else 0
                    
                    実績_amount += 実績_sum
                    情報提供料_amount += 情報提供料_sum
                    details.append({
                        'content_group': str(item_name),
                        '実績': round(実績_sum) if pd.notna(実績_sum) else 0,
                        '情報提供料': round(情報提供料_sum) if pd.notna(情報提供料_sum) else 0
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
    
    def load_mediba_content_mapping(self):
        """contents_maping.csvからmedibaコンテンツのマッピングを読み込み"""
        try:
            mapping_file = Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\溝口\miz\uriage\contents_maping.csv")
            if not mapping_file.exists():
                print("contents_maping.csvが見つかりません。medibaマッピングは適用されません。")
                return set()
            
            mediba_contents = set()
            df = pd.read_csv(mapping_file, encoding='utf-8-sig')
            
            # medibaコンテンツ（D列）を取得
            if len(df.columns) > 2:  # D列が存在することを確認
                mediba_column = df.iloc[:, 2]  # D列（0-indexed）
                for content in mediba_column:
                    if pd.notna(content) and str(content).startswith('OWD_'):
                        mediba_contents.add(str(content))
            
            print(f"medibaコンテンツマッピング読み込み完了: {len(mediba_contents)}件")
            return mediba_contents
        except Exception as e:
            print(f"medibaマッピング読み込みエラー: {e}")
            return set()

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
                # 年フォルダ（4桁）内の月フォルダ（6桁）を検索
                for year_folder in self.base_path.iterdir():
                    if year_folder.is_dir() and re.match(r'\d{4}', year_folder.name):
                        for month_folder in year_folder.iterdir():
                            if month_folder.is_dir() and re.match(r'\d{6}', month_folder.name):
                                # 月フォルダ直下のファイルをチェック
                                if any(result['file'] in str(f.name) for f in month_folder.iterdir() if f.is_file()):
                                    yearmonth = month_folder.name
                                    break
                                # サブフォルダ内のファイルもチェック
                                for subfolder in month_folder.iterdir():
                                    if subfolder.is_dir():
                                        if any(result['file'] in str(f.name) for f in subfolder.iterdir() if f.is_file()):
                                            yearmonth = month_folder.name
                                            break
                                if yearmonth != '不明':
                                    break
                        if yearmonth != '不明':
                            break
                
                # 各コンテンツの詳細を出力
                for detail in result['details']:
                    実績 = round(detail.get('実績', 0))
                    情報提供料 = round(detail.get('情報提供料', 0))
                    
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
        
        print(f"結果をCSVファイルに出力しました: {output_path}")

def main():
    # 設定
    base_path = r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書"
    output_path = "sales_distribution_summary.csv"
    
    # 処理実行
    aggregator = SalesAggregator(base_path)
    aggregator.process_all_files()
    aggregator.export_to_csv(output_path)

if __name__ == "__main__":
    main()