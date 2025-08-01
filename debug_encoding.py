#!/usr/bin/env python3
import pandas as pd

# 売上データを読み込み
sales_file = r'C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\月別ISP別コンテンツ別売上.csv'
df = pd.read_csv(sales_file, encoding='utf-8-sig')

# 202501年月のexciteデータを抽出
excite_data = df[(df['年月'] == '202501') & (df['プラットフォーム'] == 'excite')]
print(f"Excite 202501データ件数: {len(excite_data)}")

# シェイプシフター関連のデータを検索
shape_data = excite_data[excite_data['コンテンツ'].str.contains('シェイプシフター', na=False)]
print(f"シェイプシフター関連データ件数: {len(shape_data)}")

if not shape_data.empty:
    for idx, row in shape_data.iterrows():
        print(f"コンテンツ名: '{row['コンテンツ']}'")
        print(f"実績: {row['実績']}")
        print(f"情報提供料: {row['情報提供料']}")
        print(f"コンテンツ名の長さ: {len(row['コンテンツ'])}")
        print(f"コンテンツ名の文字コード: {[ord(c) for c in row['コンテンツ'][:10]]}")
        print("---")

# contents_mapping.csvからshapeのexcite値を確認
mapping_file = r'C:\Users\OW\Dropbox\disk2とローカルの同期\溝口\miz\uriage\contents_mapping.csv'
mapping_df = pd.read_csv(mapping_file, encoding='utf-8-sig')

shape_row = mapping_df[mapping_df.iloc[:, 0] == 'shape']
if not shape_row.empty:
    excite_content = shape_row['excite'].iloc[0]
    print(f"Mapping excite値: '{excite_content}'")
    print(f"Mapping値の長さ: {len(excite_content)}")
    print(f"Mapping値の文字コード: {[ord(c) for c in excite_content[:10]]}")
    
    # 比較
    if not shape_data.empty:
        actual_content = shape_data['コンテンツ'].iloc[0]
        print(f"文字列比較: {actual_content == excite_content}")
        print(f"正規化後比較: {actual_content.strip() == excite_content.strip()}")