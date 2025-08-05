#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ロイヤリティ累計推移表作成ツール

ロイヤリティフォルダ内のyyyymm形式フォルダを対象に、
各ファイルのAE19列の値を参照し、rate.csvの「名称」と「エージェント」で
マッチングして月次・累計売上を集計する。

累計は1万円以上になった翌月から0から再開。
"""

import os
import pandas as pd
import glob
from datetime import datetime
import openpyxl
import openpyxl.utils
from pathlib import Path
import re
import csv

class RoyaltyAggregator:
    def __init__(self, royalty_dir, rate_csv_path, output_path):
        self.royalty_dir = royalty_dir
        self.rate_csv_path = rate_csv_path
        self.output_path = output_path
        self.rate_data = None
        self.agent_groups = {}
        self.monthly_totals = {}
        self.cumulative_totals = {}
        
    def load_rate_data(self):
        """rate.csvを読み込み、エージェントグループを作成"""
        try:
            self.rate_data = pd.read_csv(self.rate_csv_path, encoding='utf-8')
            print(f"Rate data loaded: {len(self.rate_data)} records")
            print(f"Columns: {list(self.rate_data.columns)}")
            
            # エージェントグループを作成（エージェント列でグループ化）
            for _, row in self.rate_data.iterrows():
                name = row['名称']
                agent = row['エージェント']
                if pd.notna(agent) and agent.strip():
                    if agent not in self.agent_groups:
                        self.agent_groups[agent] = []
                    self.agent_groups[agent].append(name)
            
            print(f"Agent groups created: {list(self.agent_groups.keys())}")
            return True
        except Exception as e:
            print(f"Error loading rate data: {e}")
            return False
    
    def get_yyyymm_folders(self):
        """yyyymm形式のフォルダリストを取得"""
        folders = []
        pattern = re.compile(r'^\d{6}$')  # 6桁の数字
        
        for item in os.listdir(self.royalty_dir):
            if pattern.match(item):
                folder_path = os.path.join(self.royalty_dir, item)
                if os.path.isdir(folder_path):
                    folders.append(item)
        
        folders.sort()  # 昇順ソート
        print(f"Found YYYYMM folders: {folders}")
        return folders
    
    def read_sales_from_xlsx(self, file_path):
        """xlsxファイルから売上データを読み取り（AE19の数式を評価）"""
        try:
            # まず数式を含んだ状態でブックを読み込み
            workbook_formula = openpyxl.load_workbook(file_path, data_only=False)
            worksheet_formula = workbook_formula.active
            
            # 次に計算値のみでブックを読み込み
            workbook_data = openpyxl.load_workbook(file_path, data_only=True)
            worksheet_data = workbook_data.active
            
            # AE19セル（31列目、19行目）の処理
            ae19_cell = worksheet_formula.cell(row=19, column=31)
            ae19_value = worksheet_data.cell(row=19, column=31).value
            
            print(f"  AE19セル - 数式: {ae19_cell.value}, 計算値: {ae19_value}")
            
            # AE19が数式の場合、数式の参照先を直接評価
            if ae19_cell.value and isinstance(ae19_cell.value, str) and ae19_cell.value.startswith('='):
                print(f"  AE19は数式です: {ae19_cell.value}")
                
                # M19が実際には=SUM(AE23:AJ58)のような数式なので、直接AE23:AJ58を計算
                # M19の数式を確認
                m19_cell = worksheet_formula.cell(row=19, column=13)
                print(f"    M19の数式: {m19_cell.value}")
                
                # AE23:AE58の数式を手動で計算する
                # AE{row} = IF(OR(Y{row}="", AC{row}=""), "", IF(Y{row}*AC{row}*0.01=0, "", Y{row}*AC{row}*0.01))
                # AC{row} = IF(Y{row}>0,10,"")
                ae_sum = 0
                found_values = []
                
                for row in range(23, 59):  # 23-58行目
                    y_value = worksheet_data.cell(row=row, column=25).value  # Y列 = 25列目
                    
                    if isinstance(y_value, (int, float)) and y_value > 0:
                        # AC列の数式を取得して評価
                        ac_cell = worksheet_formula.cell(row=row, column=29)  # AC列の数式
                        ac_value_data = worksheet_data.cell(row=row, column=29).value  # AC列の計算値
                        
                        # AC列の数式がある場合は数式を解析、ない場合は計算値を使用
                        if ac_cell.value and isinstance(ac_cell.value, str) and ac_cell.value.startswith('='):
                            # 数式の場合、手動で評価
                            formula = ac_cell.value
                            print(f"      AC{row}の数式: {formula}")
                            
                            if 'IF(' in formula.upper():
                                # IF文の解析
                                import re
                                # IF(condition, true_value, false_value) のパターンを抽出
                                if_match = re.search(r'IF\(([^,]+),([^,]+),([^)]+)\)', formula.upper())
                                if if_match:
                                    condition = if_match.group(1).strip()
                                    true_value = if_match.group(2).strip()
                                    false_value = if_match.group(3).strip()
                                    
                                    # 条件を評価（Y{row}>0 のような条件）
                                    if f'Y{row}' in condition.upper() and '>' in condition:
                                        if y_value > 0:
                                            # true_valueを数値に変換
                                            try:
                                                ac_value = float(true_value)
                                            except:
                                                ac_value = 10  # デフォルト値
                                        else:
                                            ac_value = 0
                                    else:
                                        # その他の条件の場合は計算値を使用
                                        ac_value = ac_value_data if isinstance(ac_value_data, (int, float)) else 0
                                else:
                                    ac_value = ac_value_data if isinstance(ac_value_data, (int, float)) else 0
                            else:
                                # IF文以外の数式の場合は計算値を使用
                                ac_value = ac_value_data if isinstance(ac_value_data, (int, float)) else 0
                        else:
                            # 数式がない場合は計算値を使用
                            ac_value = ac_value_data if isinstance(ac_value_data, (int, float)) else 0
                        
                        # AE{row}の計算: Y{row} * AC{row} * 0.01
                        ae_value = y_value * ac_value * 0.01
                        ae_sum += ae_value
                        found_values.append(f"Y{row}={y_value}, AC{row}={ac_value}(formula:{ac_cell.value}) -> AE{row}={ae_value}")
                
                if found_values:
                    print(f"    計算された売上データ: {found_values[:5]}{'...' if len(found_values) > 5 else ''} (合計{len(found_values)}個)")
                
                # 計算ロジック: AE23:AJ58の合計 - (10.21%の四捨五入) + (10%の四捨五入) 
                if ae_sum > 0:
                    # 10.21%を四捨五入 (S19の計算)
                    minus_value = round(ae_sum * 0.1021)
                    # 10%を四捨五入 (Y19の計算)
                    plus_value = round(ae_sum * 0.10)
                    # 最終計算値 (AE19の計算: M19 - S19 + Y19)
                    calculated_value = ae_sum - minus_value + plus_value
                    # 小数点第一位で四捨五入
                    final_value = round(calculated_value, 1)
                    
                    print(f"  {os.path.basename(file_path)}: AE合計={ae_sum}, -10.21%={minus_value}, +10%={plus_value}, 最終値={final_value}円")
                    
                    workbook_formula.close()
                    workbook_data.close()
                    return final_value
                else:
                    print(f"  AE23:AJ58範囲に有効な数値が見つかりませんでした")
                    
            elif ae19_value is not None and isinstance(ae19_value, (int, float)):
                # AE19が直接数値の場合
                print(f"  AE19は数値です: {ae19_value}")
                workbook_formula.close()
                workbook_data.close()
                return ae19_value
            else:
                print(f"  AE19に有効な値が見つかりませんでした")
            
            workbook_formula.close()
            workbook_data.close()
            return 0
            
        except Exception as e:
            print(f"Error reading {file_path}: {e}")
            return 0
    
    def extract_name_from_filename(self, filename):
        """ファイル名から名称を抽出"""
        # ファイル名のパターン: YYYYMM_name.xlsx または name_YYYYMM.xlsx
        base_name = os.path.splitext(filename)[0]
        
        # アンダースコアで分割
        parts = base_name.split('_')
        
        # YYYYMMパターンを除去
        yyyymm_pattern = re.compile(r'^\d{6}$')
        name_parts = [part for part in parts if not yyyymm_pattern.match(part)]
        
        if name_parts:
            return name_parts[0]  # 最初の非YYYYMM部分を名称とする
        
        return base_name
    
    def process_monthly_data(self, yyyymm):
        """指定月のデータを処理"""
        folder_path = os.path.join(self.royalty_dir, yyyymm)
        if not os.path.exists(folder_path):
            print(f"Folder not exists: {folder_path}")
            return
        
        print(f"Processing {yyyymm}...")
        monthly_agent_totals = {}
        
        # フォルダ内のxlsxファイルを処理
        xlsx_files = glob.glob(os.path.join(folder_path, "*.xlsx"))
        
        for file_path in xlsx_files:
            filename = os.path.basename(file_path)
            name = self.extract_name_from_filename(filename)
            sales_value = self.read_sales_from_xlsx(file_path)
            
            if sales_value > 0:
                # rate.csvで対応するエージェントを検索
                matching_row = self.rate_data[self.rate_data['名称'] == name]
                if not matching_row.empty:
                    agent = matching_row.iloc[0]['エージェント']
                    if pd.notna(agent) and agent.strip():
                        if agent not in monthly_agent_totals:
                            monthly_agent_totals[agent] = 0
                        monthly_agent_totals[agent] += sales_value
                        print(f"  {name} -> {agent}: {sales_value}")
                else:
                    print(f"  {name}: Not found in rate.csv")
        
        self.monthly_totals[yyyymm] = monthly_agent_totals
    
    def calculate_cumulative_totals(self):
        """累計売上を計算（1万円以上で翌月リセット）"""
        sorted_months = sorted(self.monthly_totals.keys())
        
        for agent in self.agent_groups.keys():
            cumulative = 0
            
            for i, month in enumerate(sorted_months):
                monthly_total = self.monthly_totals.get(month, {}).get(agent, 0)
                cumulative += monthly_total
                
                if month not in self.cumulative_totals:
                    self.cumulative_totals[month] = {}
                
                self.cumulative_totals[month][agent] = cumulative
                
                # 1万円以上になったら翌月リセット（今月の累計は記録してから次月でリセット）
                if cumulative >= 10000:
                    print(f"  {agent}: 累計{cumulative}円で1万円超過 -> 翌月リセット")
                    cumulative = 0
    
    def create_output_excel(self):
        """出力Excel作成"""
        sorted_months = sorted(self.monthly_totals.keys())
        sorted_agents = sorted(self.agent_groups.keys())
        
        # データフレーム用のデータを準備
        data = []
        
        for month in sorted_months:
            row = {'売上年月': month}
            
            for agent in sorted_agents:
                monthly_total = self.monthly_totals.get(month, {}).get(agent, 0)
                cumulative_total = self.cumulative_totals.get(month, {}).get(agent, 0)
                
                row[f'{agent}売上'] = monthly_total
                row[f'{agent}累計'] = cumulative_total
            
            data.append(row)
        
        # DataFrameを作成
        df = pd.DataFrame(data)
        
        # Excelファイルに出力
        df.to_excel(self.output_path, index=False, engine='openpyxl')
        print(f"Output created: {self.output_path}")
        
        # CSV版も作成
        csv_path = self.output_path.replace('.xlsx', '.csv')
        df.to_csv(csv_path, index=False, encoding='utf-8-sig')
        print(f"CSV output created: {csv_path}")
    
    def run(self):
        """メイン処理実行"""
        print("=== ロイヤリティ累計推移表作成開始 ===")
        
        # rate.csvの読み込み
        if not self.load_rate_data():
            return False
        
        # yyyymmフォルダの取得
        yyyymm_folders = self.get_yyyymm_folders()
        if not yyyymm_folders:
            print("No YYYYMM folders found")
            return False
        
        # 各月のデータを処理
        for yyyymm in yyyymm_folders:
            self.process_monthly_data(yyyymm)
        
        # 累計売上を計算
        self.calculate_cumulative_totals()
        
        # 出力Excel作成
        self.create_output_excel()
        
        print("=== 処理完了 ===")
        return True


def main():
    # パスの設定
    royalty_dir = r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ロイヤリティ"
    rate_csv_path = r"C:\Users\OW\pj\uriage\rate.csv"
    output_path = r"C:\Users\OW\pj\uriage\コンテンツ別累計推移表.xlsx"
    
    # 集計実行
    aggregator = RoyaltyAggregator(royalty_dir, rate_csv_path, output_path)
    success = aggregator.run()
    
    if success:
        print("\n集計処理が正常に完了しました。")
    else:
        print("\n集計処理でエラーが発生しました。")


if __name__ == "__main__":
    main()