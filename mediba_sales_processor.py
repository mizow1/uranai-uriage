#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
mediba占い売上データ処理スクリプト
SalesSummaryファイルからB列（番組ID）ごとの売上集計を行い、
実績と情報提供料合計を算出する
"""

import pandas as pd
import os
from pathlib import Path
from datetime import datetime

# 共通コンポーネントをインポート
from common import (
    CSVHandler, 
    UnifiedLogger, 
    ErrorHandler, 
    ConfigManager,
    ProcessingResult,
    ContentDetail,
    ProcessingSummary
)

class MedibaSalesProcessor:
    """mediba占い売上データ処理クラス"""
    
    def __init__(self):
        self.logger = UnifiedLogger(__name__, log_file=Path("logs/mediba_sales_processor.log"))
        self.error_handler = ErrorHandler(self.logger.logger)
        self.csv_handler = CSVHandler(self.logger.logger, self.error_handler)
        self.config = ConfigManager(logger=self.logger.logger)
    
    def find_sales_summary_files(self):
        """SalesSummaryを含むCSVファイルを検索"""
        base_path = Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書")
        pattern = "**/SalesSummary*.csv"
        files = list(base_path.glob(pattern))
        self.logger.info(f"見つかったSalesSummaryファイル数: {len(files)}")
        return files

    def process_sales_data(self, csv_file_path: Path) -> ProcessingResult:
        """
        CSVファイルから売上データを処理
        
        Args:
            csv_file_path: CSVファイルのパス
            
        Returns:
            ProcessingResult: 処理結果
        """
        result = ProcessingResult(
            platform="mediba",
            file_name=csv_file_path.name,
            success=False
        )
        
        start_time = datetime.now()
        
        try:
            # CSVファイルを読み込み（統一ハンドラー使用）
            df = self.csv_handler.read_csv_with_encoding_detection(csv_file_path)
            
            self.logger.log_file_operation("読み込み", csv_file_path, True)
            self.logger.info(f"データ行数: {len(df)}")
            
            # 列数チェック
            if not self.csv_handler.validate_csv_structure(df, required_columns=11):
                result.add_error(f"列数が不足: 必要11列、実際{len(df.columns)}列")
                return result
                
            # 列名を設定（インデックスで参照）
            program_id_col = df.columns[1]  # B列
            revenue_col = df.columns[6]     # G列  
            cp_cost_col = df.columns[10]    # K列
            
            self.logger.info(f"使用する列: B列={program_id_col}, G列={revenue_col}, K列={cp_cost_col}")
            
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
            merged = pd.merge(grouped, cp_cost_sum, on=program_id_col)
            
            # 情報提供料合計 = 実績の40% - K列の値
            merged['情報提供料合計'] = merged['実績'] * 0.4 - merged['CP売上負担額合計']
            merged = merged.rename(columns={program_id_col: '番組ID'})
            
            # ContentDetailリストを作成
            for _, row in merged.iterrows():
                detail = ContentDetail(
                    content_group=str(row['番組ID']),
                    performance=float(row['実績']),
                    information_fee=float(row['情報提供料合計']),
                    additional_data={'CP売上負担額合計': float(row['CP売上負担額合計'])}
                )
                result.add_detail(detail)
            
            # 合計を計算
            result.calculate_totals()
            result.success = True
            result.metadata = {
                'total_records': len(df),
                'unique_programs': len(merged),
                'merged_result': merged
            }
            
            self.logger.info(f"処理完了: {len(merged)}件の番組IDで集計")
            
        except Exception as e:
            result.add_error(str(e))
            self.error_handler.handle_file_processing_error(e, csv_file_path)
        
        finally:
            end_time = datetime.now()
            result.processing_time = (end_time - start_time).total_seconds()
        
        return result

    def save_results(self, results: list, output_file: str = "mediba_sales_summary.csv") -> bool:
        """結果をCSVファイルに保存"""
        try:
            if not results:
                self.logger.warning("保存する結果がありません")
                return False
            
            # 全結果を結合
            all_results = []
            for result in results:
                if result.success and 'merged_result' in result.metadata:
                    df_temp = result.metadata['merged_result'].copy()
                    df_temp['ファイル名'] = result.file_name
                    df_temp['プラットフォーム'] = result.platform
                    all_results.append(df_temp)
            
            if all_results:
                final_df = pd.concat(all_results, ignore_index=True)
                final_df.to_csv(output_file, index=False, encoding='utf-8-sig')
                self.logger.log_file_operation("保存", Path(output_file), True)
                return True
            else:
                self.logger.warning("結合可能な結果がありません")
                return False
                
        except Exception as e:
            self.error_handler.log_and_continue(e, "結果保存")
            return False
    
    def run(self):
        """メイン処理を実行"""
        summary = ProcessingSummary()
        summary.processing_start = datetime.now()
        
        self.logger.info("mediba占い売上データ処理を開始")
        
        # SalesSummaryファイルを検索
        files = self.find_sales_summary_files()
        
        if not files:
            self.logger.warning("SalesSummaryファイルが見つかりませんでした")
            return summary
        
        self.logger.log_file_list(files, "処理")
        
        # 各ファイルを処理
        results = []
        for file_path in files:
            self.logger.log_processing_progress(len(results) + 1, len(files), file_path.name)
            
            result = self.process_sales_data(file_path)
            results.append(result)
            summary.add_result(result)
            
            # 個別結果をログ出力
            if result.success:
                self.logger.log_platform_results("mediba", {
                    'ファイル名': result.file_name,
                    '総レコード数': result.metadata.get('total_records', 0),  
                    '番組ID数': result.metadata.get('unique_programs', 0),
                    '実績合計': result.total_performance,
                    '情報提供料合計': result.total_information_fee
                })
                
                # 上位10件を表示
                if 'merged_result' in result.metadata:
                    merged_df = result.metadata['merged_result']
                    self.logger.info(f"\n=== {result.file_name} ===")
                    self.logger.info(f"総レコード数: {result.metadata['total_records']}")
                    self.logger.info(f"番組ID数: {result.metadata['unique_programs']}")
                    self.logger.info("\n集計結果（上位10件）:")
                    self.logger.info(merged_df.head(10).to_string(index=False))
        
        # 結果をファイルに保存
        if self.save_results(results):
            self.logger.info(f"\n全ての結果が mediba_sales_summary.csv に保存されました")
        
        summary.processing_end = datetime.now()
        self.logger.log_processing_summary(
            summary.total_files,
            summary.successful_files, 
            summary.failed_files,
            summary.processing_duration or 0
        )
        
        return summary


def main():
    """メイン処理"""
    processor = MedibaSalesProcessor()
    processor.run()

if __name__ == "__main__":
    main()