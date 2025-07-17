# Design Document: 明細書作成機能

## Overview

明細書作成機能は、売上集計システムで生成されたCSVデータ（output.csv）を元に、コンテンツ提供元への支払い明細書を自動生成するシステムです。この機能は既存の明細書テンプレートを参照し、最新の売上データを反映した新しい明細書をExcel形式で作成します。

本機能は以下の主要なコンポーネントで構成されます：
1. データ読み込みモジュール - 売上集計CSVと明細書テンプレートの読み込み
2. データマッチングモジュール - 売上データと明細書テンプレートのマッピング
3. 明細書生成モジュール - 新しい明細書の作成と保存
4. ユーザーインターフェース - 処理状況の表示と入力の受付

## Architecture

システムは以下のアーキテクチャで構築されます：

```
                  +-------------------+
                  |  ユーザーインターフェース  |
                  +-------------------+
                           |
                           v
+-------------+    +-------------------+    +-------------+
| 売上集計CSV  | -> | データ読み込みモジュール | <- | 明細書テンプレート |
+-------------+    +-------------------+    +-------------+
                           |
                           v
                  +-------------------+
                  | データマッチングモジュール |
                  +-------------------+
                           |
                           v
                  +-------------------+
                  |  明細書生成モジュール  |
                  +-------------------+
                           |
                           v
                  +-------------------+
                  |    新規明細書ファイル   |
                  +-------------------+
```

## Components and Interfaces

### 1. データ読み込みモジュール (InvoiceDataLoader)

**責務:**
- 売上集計CSVファイルの読み込み
- 明細書テンプレートフォルダの検索と最新テンプレートの特定
- Excelファイルの読み込みと解析

**インターフェース:**
```python
class InvoiceDataLoader:
    def __init__(self, sales_csv_path, template_base_path):
        # 初期化処理
        
    def load_sales_data(self):
        # 売上集計CSVを読み込み、データフレームとして返す
        
    def find_latest_template(self):
        # 最新の明細書テンプレートを検索して返す
        
    def load_template_data(self, template_path):
        # 明細書テンプレートを読み込み、解析する
```

### 2. データマッチングモジュール (DataMatcher)

**責務:**
- 売上データと明細書テンプレートのコンテンツ名/プラットフォーム名のマッチング
- ファジーマッチングアルゴリズムの実装
- マッチング結果の信頼度評価

**インターフェース:**
```python
class DataMatcher:
    def __init__(self, sales_data, template_data):
        # 初期化処理
        
    def match_content_names(self):
        # コンテンツ名のマッチング処理
        
    def match_platform_names(self):
        # プラットフォーム名のマッチング処理
        
    def get_matching_confidence(self):
        # マッチングの信頼度を評価して返す
        
    def get_matched_data(self):
        # マッチング結果を返す
```

### 3. 明細書生成モジュール (InvoiceGenerator)

**責務:**
- マッチングされたデータを元に新しい明細書を作成
- 適切なフォルダ構造とファイル名での保存
- 既存ファイルの確認と上書き処理

**インターフェース:**
```python
class InvoiceGenerator:
    def __init__(self, matched_data, template_path, output_base_path):
        # 初期化処理
        
    def create_invoice(self):
        # 明細書の作成処理
        
    def determine_output_path(self):
        # 出力先パスとファイル名の決定
        
    def save_invoice(self):
        # 明細書の保存処理
        
    def check_file_exists(self, path):
        # ファイルの存在確認
```

### 4. ユーザーインターフェース (InvoiceGeneratorUI)

**責務:**
- ユーザーからの入力受付
- 処理状況の表示
- エラーメッセージの表示
- 処理結果の表示

**インターフェース:**
```python
class InvoiceGeneratorUI:
    def __init__(self):
        # 初期化処理
        
    def get_input_paths(self):
        # 入力パスの取得
        
    def show_progress(self, message):
        # 進行状況の表示
        
    def show_error(self, error_message, solution=None):
        # エラーメッセージの表示
        
    def show_result(self, result_summary):
        # 処理結果の表示
        
    def confirm_overwrite(self, file_path):
        # ファイル上書きの確認
```

## Data Models

### 1. 売上データモデル (SalesData)

```python
class SalesData:
    def __init__(self):
        self.platform = ""        # プラットフォーム名
        self.content = ""         # コンテンツ名
        self.sales_amount = 0     # 実績額
        self.royalty_amount = 0   # 情報提供料合計
        self.year_month = ""      # 年月（yyyymm）
```

### 2. テンプレートデータモデル (TemplateData)

```python
class TemplateData:
    def __init__(self):
        self.path = ""            # テンプレートファイルパス
        self.year_month = ""      # 年月（yyyymm）
        self.platforms = []       # プラットフォーム名リスト
        self.contents = []        # コンテンツ名リスト
        self.cell_mappings = {}   # セル位置マッピング
```

### 3. マッチング結果モデル (MatchedData)

```python
class MatchedData:
    def __init__(self):
        self.template_path = ""   # 使用するテンプレートパス
        self.output_path = ""     # 出力先パス
        self.mappings = []        # マッピング情報のリスト
        self.confidence = 0.0     # マッチングの信頼度（0.0～1.0）
```

## Error Handling

エラーハンドリングは以下の方針で実装します：

1. **入力ファイル検証エラー**
   - 売上集計CSVファイルが存在しない場合
   - 明細書テンプレートフォルダが存在しない場合
   - 適切な明細書テンプレートが見つからない場合

2. **データ読み込みエラー**
   - CSVファイルの読み込みエラー
   - Excelファイルの読み込みエラー
   - データ形式が予期しないフォーマットの場合

3. **マッチングエラー**
   - マッチング結果の信頼度が低い場合
   - 必要なデータが見つからない場合

4. **ファイル出力エラー**
   - 出力先フォルダへの書き込み権限がない場合
   - ディスク容量不足の場合

各エラーは具体的なメッセージと可能な解決策を提示し、ユーザーが適切に対応できるようにします。

## Testing Strategy

テスト戦略は以下の通りです：

### 1. ユニットテスト

- 各モジュールの個別機能をテスト
- モックデータを使用して外部依存を排除
- エッジケースの検証（空のファイル、異常値など）

### 2. 統合テスト

- モジュール間の連携をテスト
- 実際のファイル形式でのデータ読み込みテスト
- マッチングアルゴリズムの精度検証

### 3. システムテスト

- エンドツーエンドのワークフローテスト
- 実際の売上データと明細書テンプレートを使用したテスト
- パフォーマンステスト（大量データ処理時の動作確認）

### 4. ユーザー受け入れテスト

- 実際の業務シナリオでのテスト
- ユーザーインターフェースの使いやすさ検証
- エラーメッセージの分かりやすさ検証

## Implementation Considerations

### 1. ファジーマッチングアルゴリズム

コンテンツ名とプラットフォーム名のマッチングには、以下のアルゴリズムを検討します：

- レーベンシュタイン距離（編集距離）
- Jaro-Winkler距離
- N-gramベースの類似度
- TF-IDFベクトル類似度

実装では、複数のアルゴリズムを組み合わせて精度を高めます。

### 2. Excelファイル操作

Excelファイルの操作には、以下のライブラリを使用します：

- openpyxl - 新しいExcelファイル形式（.xlsx）の読み書き
- xlrd/xlwt - 古いExcelファイル形式（.xls）の読み書き
- pandas - データフレームとしての操作

### 3. パフォーマンス最適化

- 大量データ処理時のメモリ使用量最適化
- バッチ処理による効率化
- 進捗表示によるユーザー体験向上

### 4. 設定の外部化

- パス設定
- マッチングアルゴリズムのパラメータ
- 出力ファイル形式

これらの設定は設定ファイルに外部化し、コードの変更なしに調整可能にします。
