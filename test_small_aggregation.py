#!/usr/bin/env python3
"""
小規模な売上集計テスト（少数ファイルのみ）
"""

import sys
import os
from pathlib import Path
from sales_aggregator import SalesAggregator

def main():
    print("=== 小規模売上集計テスト ===")
    
    base_path = r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書"
    
    try:
        aggregator = SalesAggregator(base_path)
        
        # 個別にファイルを処理（少数のみ）
        test_files = [
            # softbankファイル（動作確認済み）
            Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\2023\202303\SB202303\oidshiharai\OID_PAY_9ATI_202303.PDF"),
            # auファイル（エラーになるが処理される）
            Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\2025\202501\202501cp02お支払い明細書.pdf"),
            # docomoファイル（5行目修正確認）
            Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\2024\202405\bp40000746shiharai20250730143312.csv"),
        ]
        
        results = []
        for file_path in test_files:
            if not file_path.exists():
                print(f"ファイルが存在しません: {file_path}")
                continue
                
            print(f"\n処理中: {file_path.name}")
            
            # ファイルタイプに応じて処理
            if 'oid_pay_9ati' in file_path.name.lower():
                result = aggregator.process_softbank_file(file_path)
                result.platform = 'softbank'
            elif 'cp02お支払い明細書' in file_path.name:
                result = aggregator.process_au_new_file(file_path)
                result.platform = 'au'  
            elif 'bp40000746' in file_path.name:
                result = aggregator.process_docomo_file(file_path)
                result.platform = 'docomo'
            else:
                continue
                
            print(f"  成功: {result.success}")
            if result.errors:
                print(f"  エラー: {result.errors}")
            if result.details:
                print(f"  詳細: {len(result.details)}件")
                for detail in result.details:
                    print(f"    - {detail.content_group}: 実績={detail.performance}, 情報提供料={detail.information_fee}")
            
            if result.success:
                year_month = aggregator._extract_year_month_from_path(file_path)
                results.append({
                    'platform': result.platform,
                    'file_name': result.file_name,
                    'content_details': result.details,
                    '情報提供料合計': result.total_information_fee,
                    '実績合計': result.total_performance,
                    '年月': year_month,
                    '処理日時': '2025-07-30 22:30:00'
                })
        
        print(f"\n=== 結果サマリー ===")
        platform_count = {}
        for result in results:
            platform = result['platform']
            platform_count[platform] = platform_count.get(platform, 0) + 1
            
        for platform, count in platform_count.items():
            print(f"{platform}: {count}件処理成功")
            
        print(f"\n処理成功ファイル総数: {len(results)}")
        
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()