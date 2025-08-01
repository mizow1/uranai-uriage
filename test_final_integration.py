#!/usr/bin/env python3
"""
最終統合テスト - 月別ISP別コンテンツ別売上にau、softbankが表示されるかテスト
"""

import sys
import os
from pathlib import Path
from sales_aggregator import SalesAggregator

def main():
    print("=== 最終統合テスト ===")
    
    base_path = r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書"
    
    try:
        aggregator = SalesAggregator(base_path)
        
        # 特定の年月でテスト（au, softbank, docomoが存在する月）
        test_files = [
            # softbank
            Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\2023\202303\SB202303\oidshiharai\OID_PAY_9ATI_202303.PDF"),
            # au CSV
            Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\2025\202503\au202503\202503cp02お支払い明細書.csv"),
            # docomo
            Path(r"C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\2024\202405\bp40000746shiharai20250730143312.csv"),
        ]
        
        all_results = []
        
        for file_path in test_files:
            if not file_path.exists():
                print(f"ファイルが存在しません: {file_path}")
                continue
                
            print(f"\n処理中: {file_path.name}")
            
            try:
                # ファイルタイプに応じて処理
                if 'oid_pay_9ati' in file_path.name.lower():
                    result = aggregator.process_softbank_file(file_path)
                    platform = 'softbank'
                elif 'cp02お支払い明細書' in file_path.name:
                    result = aggregator.process_au_new_file(file_path)
                    platform = 'au'
                elif 'bp40000746' in file_path.name:
                    result = aggregator.process_docomo_file(file_path)
                    platform = 'docomo'
                else:
                    continue
                    
                if result.success:
                    year_month = aggregator._extract_year_month_from_path(file_path)
                    print(f"  プラットフォーム: {platform}")
                    print(f"  年月: {year_month}")
                    print(f"  情報提供料: {result.total_information_fee:,}円")
                    print(f"  実績合計: {result.total_performance:,}円")
                    
                    # 月別レポート用のデータを準備
                    for detail in result.details:
                        all_results.append({
                            'プラットフォーム': platform,
                            'コンテンツ名': detail.content_group,
                            '年月': year_month,
                            '実績': detail.performance,
                            '情報提供料': detail.information_fee,
                            'ファイル名': result.file_name
                        })
                else:
                    print(f"  処理失敗: {result.errors}")
                    
            except Exception as e:
                print(f"  エラー: {str(e)}")
        
        print(f"\n=== 月別ISP別コンテンツ別売上（統合結果） ===")
        if all_results:
            print("プラットフォーム | コンテンツ名 | 年月 | 実績 | 情報提供料")
            print("-" * 80)
            
            for result in all_results:
                print(f"{result['プラットフォーム']:12} | {result['コンテンツ名']:20} | {result['年月']} | {result['実績']:8,}円 | {result['情報提供料']:8,}円")
            
            # プラットフォーム別集計
            platform_summary = {}
            for result in all_results:
                platform = result['プラットフォーム']
                if platform not in platform_summary:
                    platform_summary[platform] = {
                        '実績合計': 0,
                        '情報提供料': 0,
                        'ファイル数': 0
                    }
                platform_summary[platform]['実績合計'] += result['実績']
                platform_summary[platform]['情報提供料'] += result['情報提供料']
                platform_summary[platform]['ファイル数'] += 1
            
            print(f"\n=== プラットフォーム別集計 ===")
            for platform, summary in platform_summary.items():
                print(f"{platform}: ファイル数={summary['ファイル数']}, 実績合計={summary['実績合計']:,}円, 情報提供料={summary['情報提供料']:,}円")
        else:
            print("処理された結果がありません")
            
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()