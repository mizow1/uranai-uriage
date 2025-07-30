#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
mediba占い売上データ処理スクリプト
SalesSummaryファイルからB列（番組ID）ごとの売上集計を行い、
実績と情報提供料合計を算出する
"""

import pandas as pd
import glob
import os
from pathlib import Path
import logging

# ログ設定
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def find_sales_summary_files():
    """SalesSummaryを含むCSVファイルを検索"""
    base_path = Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書")
    pattern = "**/SalesSummary*.csv"
    files = list(base_path.glob(pattern))
    logger.info(f"見つかったSalesSummaryファイル数: {len(files)}")
    return files

def process_sales_data(csv_file_path):
    """
    CSVファイルから売上データを処理
    
    Args:
        csv_file_path: CSVファイルのパス
        
    Returns:
        dict: 処理結果
    """
    try:
        # CSVファイルを読み込み（文字コード自動判定）
        try:
            df = pd.read_csv(csv_file_path, encoding='utf-8')
        except UnicodeDecodeError:
            try:
                df = pd.read_csv(csv_file_path, encoding='shift_jis')
            except UnicodeDecodeError:
                df = pd.read_csv(csv_file_path, encoding='cp932')
        
        logger.info(f"ファイル読み込み完了: {csv_file_path}")
        logger.info(f"データ行数: {len(df)}")
        
        # 列のインデックスで参照（文字化け対応）
        # B列（インデックス1）: 番組ID
        # G列（インデックス6）: 料金（税込）
        # K列（インデックス10）: CP売上負担額（税込）
        
        if len(df.columns) < 11:
            logger.error(f"列数が不足しています: {len(df.columns)}列")
            return None
            
        # 列名を設定
        program_id_col = df.columns[1]  # B列
        revenue_col = df.columns[6]     # G列  
        cp_cost_col = df.columns[10]    # K列
        
        logger.info(f"使用する列: B列={program_id_col}, G列={revenue_col}, K列={cp_cost_col}")
        
        # 数値に変換（エラーハンドリング付き）
        df[revenue_col] = pd.to_numeric(df[revenue_col], errors='coerce').fillna(0)
        df[cp_cost_col] = pd.to_numeric(df[cp_cost_col], errors='coerce').fillna(0)
        
        # B列でグループ化してG列の合計を計算（実績）
        grouped = df.groupby(program_id_col)[revenue_col].sum().reset_index()
        grouped = grouped.rename(columns={revenue_col: '実績'})
        
        # K列の合計も計算
        cp_cost_sum = df.groupby(program_id_col)[cp_cost_col].sum().reset_index()
        cp_cost_sum = cp_cost_sum.rename(columns={cp_cost_col: 'CP売上負担額合計'})
        
        # 結果をマージ
        result = pd.merge(grouped, cp_cost_sum, on=program_id_col)
        
        # 情報提供料合計 = 実績の40% - K列の値
        result['情報提供料合計'] = result['実績'] * 0.4 - result['CP売上負担額合計']
        
        # 列名を整理
        result = result.rename(columns={program_id_col: '番組ID'})
        
        logger.info(f"処理完了: {len(result)}件の番組IDで集計")
        
        return {
            'file_path': csv_file_path,
            'result': result,
            'total_records': len(df),
            'unique_programs': len(result)
        }
        
    except Exception as e:
        logger.error(f"ファイル処理エラー: {csv_file_path}, エラー: {str(e)}")
        return None

def save_results(results, output_file="mediba_sales_summary.csv"):
    """結果をCSVファイルに保存"""
    try:
        if results and len(results) > 0:
            # 全結果を結合
            all_results = []
            for result_data in results:
                if result_data and 'result' in result_data:
                    df_temp = result_data['result'].copy()
                    df_temp['ファイル名'] = os.path.basename(result_data['file_path'])
                    all_results.append(df_temp)
            
            if all_results:
                final_df = pd.concat(all_results, ignore_index=True)
                final_df.to_csv(output_file, index=False, encoding='utf-8-sig')
                logger.info(f"結果を保存しました: {output_file}")
                return output_file
    except Exception as e:
        logger.error(f"結果保存エラー: {str(e)}")
    return None

def main():
    """メイン処理"""
    logger.info("mediba占い売上データ処理を開始")
    
    # SalesSummaryファイルを検索
    files = find_sales_summary_files()
    
    if not files:
        logger.warning("SalesSummaryファイルが見つかりませんでした")
        return
    
    # 各ファイルを処理
    results = []
    for file_path in files:
        logger.info(f"処理中: {file_path}")
        result = process_sales_data(file_path)
        if result:
            results.append(result)
            
            # 個別結果を表示
            print(f"\n=== {os.path.basename(file_path)} ===")
            print(f"総レコード数: {result['total_records']}")
            print(f"番組ID数: {result['unique_programs']}")
            print("\n集計結果（上位10件）:")
            print(result['result'].head(10).to_string(index=False))
    
    # 結果をファイルに保存
    output_file = save_results(results)
    
    if output_file:
        print(f"\n全ての結果が {output_file} に保存されました")
    
    logger.info("処理完了")

if __name__ == "__main__":
    main()