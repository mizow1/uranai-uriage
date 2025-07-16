# 売上集計システム

各プラットフォームの課金コンテンツ売上データを処理し、コンテンツ提供元への分配額を算出してCSVファイルに出力するシステムです。

## 対応プラットフォーム

- **ameba占い**: パスワード付きExcelファイル（SATORI実績）
- **楽天占い**: rcmsファイル（Excel/CSV）
- **au占い**: SalesSummaryファイル（Excel/CSV）
- **excite占い**: exciteファイル（Excel/CSV）
- **LINE占い**: LINEファイル（Excel、内訳シート）

## 使用方法

### 1. 依存関係のインストール

```bash
pip install -r requirements.txt
```

### 2. 実行

```bash
python run_sales_aggregator.py
```

または、直接実行する場合：

```bash
python sales_aggregator.py
```

### 3. 設定

実行時に以下の設定を求められます：

- **データフォルダのパス**: 年月フォルダ（yyyymm）が含まれているベースフォルダ
- **出力CSVファイル名**: 結果を保存するCSVファイル名

## 出力形式

CSVファイルには以下の情報が含まれます：

- プラットフォーム名
- ファイル名
- 合計分配額
- 年月
- 処理日時

## ファイル構成

- `sales_aggregator.py`: メインの処理クラス
- `run_sales_aggregator.py`: 実行スクリプト
- `requirements.txt`: 依存関係
- `仕様.md`: 処理仕様

## 処理仕様

各プラットフォームの詳細な処理仕様については `仕様.md` を参照してください。

## 注意事項

- ameba占いファイルはパスワード「cP2eL4T9」でロックされています
- 年月フォルダは6桁の数字（yyyymm）である必要があります
- ファイル名は大文字小文字を区別しません