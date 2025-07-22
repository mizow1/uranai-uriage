# LINE Fortune Email Processor

LINE Fortuneの日次レポートメールを自動処理するシステムです。

## 機能

- LINE Fortuneからの日次レポートメールの自動監視
- CSV添付ファイルの自動抽出と保存
- 日付ベースのファイル名変更
- 年/月のディレクトリ構造での自動整理
- 月次統合CSVファイルの自動生成（line-menu-YYYY-MM.csv）
- **コンテンツ売上集計機能（line-contents-YYYY-MM.csv）**
- 包括的なログ記録とエラー処理

## インストール

1. 必要な依存関係をインストール：
```bash
pip install -r requirements.txt
```

2. 設定ファイルテンプレートを作成：
```bash
python line_fortune_email_processor.py --create-config
```

3. 設定ファイルを編集：
- `line_fortune_config_template.json`を`line_fortune_config.json`にコピー
- メール認証情報などの必要な設定を入力

## 使用方法

### LINE Fortune Email Processor

#### 基本的な実行
```bash
python line_fortune_email_processor.py
```

#### ドライラン（実際の処理は行わない）
```bash
python line_fortune_email_processor.py --dry-run
```

#### カスタム設定ファイルを使用
```bash
python line_fortune_email_processor.py --config my_config.json
```

#### 古いファイルの削除（30日より古いファイルを削除）
```bash
python line_fortune_email_processor.py --cleanup 30
```

#### ログレベルの設定
```bash
python line_fortune_email_processor.py --log-level DEBUG
```

#### CSV統合機能（同フォルダ内のCSVファイルを統合）
```bash
python line_fortune_email_processor.py --merge-csvs
```

実行時の対話フロー：
1. 同フォルダ内のCSVファイルを自動検出・表示（抽出日付も表示）
2. 出力ファイル名を指定（デフォルト: `line-menu-YYYY-MM.csv`）
3. 既存ファイルがある場合は上書き確認
4. 最終実行確認

#### メール検索日付範囲の指定（対話形式）
```bash
# 基本的な実行（対話形式で日付範囲を指定）
python line_fortune_email_processor.py

# 実行例:
# ==================================================
# LINE Fortune Email Processor
# ==================================================
# 処理対象メールの日付範囲を指定してください
# 開始年月日を入力 (デフォルト: 2025-07-22, Enterキーで確定): 2025-07-20
# 終了年月日を入力 (デフォルト: 2025-07-20, Enterキーで確定): 2025-07-22
#
# 処理対象: 2025-07-20 ～ 2025-07-22 のメール
# この設定で実行しますか？ (y/N): y

# ドライランでの実行
python line_fortune_email_processor.py --dry-run
```

### LINE Contents Aggregator（コンテンツ売上集計）

#### 特定年月のコンテンツ集計
```bash
python line_contents_aggregator.py --year 2025 --month 7
```

#### 全ての年月を一括処理
```bash
python line_contents_aggregator.py
```

#### 単一ファイルを処理
```bash
python line_contents_aggregator.py --file "path/to/line-menu-2025-07.csv"
```

#### カスタムベースパスを指定
```bash
python line_contents_aggregator.py --base-path "C:\custom\path" --year 2025
```

## 設定ファイル

`line_fortune_config.json`の設定項目：

- `email.server`: メールサーバー（例: imap.gmail.com）
- `email.port`: ポート番号（例: 993）
- `email.username`: メールアドレス
- `email.password`: アプリパスワード
- `email.use_ssl`: SSL使用フラグ
- `base_path`: ファイル保存先のベースパス
- `log_file`: ログファイル名
- `log_level`: ログレベル（DEBUG/INFO/WARNING/ERROR）
- `sender`: 送信者メールアドレス
- `recipient`: 受信者メールアドレス
- `subject_pattern`: 件名パターン
- `search_days`: メール検索対象日数（日付範囲未指定時のデフォルト: 7日）
- `retry_count`: 再試行回数
- `retry_delay`: 再試行間隔（秒）

## ファイル構造

```
line_fortune_processor/
├── __init__.py
├── config.py              # 設定管理
├── email_processor.py     # メール処理
├── file_processor.py      # ファイル処理
├── consolidation_processor.py  # CSV統合処理
├── logger.py              # ログ記録
└── main_processor.py      # メインコントローラー
```

## 保存されるファイル

### ファイル保存場所
```
C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書\
├── YYYY/           # 年フォルダ
    ├── YYYYMM/     # 月フォルダ
        ├── line/   # LINE関連ファイル専用フォルダ
        │   ├── 元のファイル名_YYYY-MM-DD.csv    # 個別CSVファイル
        │   ├── line-menu-YYYY-MM.csv           # 統合CSVファイル
        │   └── line-contents-YYYY-MM.csv       # コンテンツ集計CSVファイル
        └── logs/
            └── line_fortune_processor.log      # ログファイル
```

### ファイル種類
- **個別CSVファイル**: `元のファイル名_YYYY-MM-DD.csv` - lineサブディレクトリに保存
- **統合CSVファイル**: `line-menu-YYYY-MM.csv` - 月次統合データ
- **コンテンツ集計CSVファイル**: `line-contents-YYYY-MM.csv` - 売上実績集計
- **ログファイル**: `logs/line_fortune_processor.log` - 処理ログ

## LINE Contents Aggregator詳細

### 機能概要

`line_contents_aggregator.py`は、`line-menu-YYYY-MM.csv`ファイルからコンテンツごとの売上実績と情報提供料を計算し、`line-contents-YYYY-MM.csv`ファイルとして集計結果を出力します。

### 処理内容

1. **コンテンツグループ化**: 
   - C列（item_code）の「_」より前の値で同一コンテンツとしてグループ化
   - 例: `kmo2_999`, `kmo2_001` → `kmo2`グループ

2. **実績計算**:
   - D列（ios_paid_amount）× 1.528575 → 1円未満四捨五入 → × 0.366674 ÷ 1.1

3. **情報提供料計算**:
   - 実績 × 0.366674 → 1円未満四捨五入

### 出力フォーマット

`line-contents-YYYY-MM.csv`の列構成：
- `content_group`: コンテンツグループ（例: kmo2, aoki, etc.）
- `content_name`: 代表的なコンテンツ名
- `total_amount`: 合計売上金額
- `performance`: 計算された実績
- `info_fee`: 情報提供料

### 処理例

```
content_group,content_name,total_amount,performance,info_fee
kmo2,ムーンオラクルカードが導く今日の運勢,4200,2140,785
aoki,十大主星が導く今日のあなたの運勢,0,0,0
```

## 新機能・変更点

### v2.0の主な変更点
- **ファイル保存場所の変更**: 添付ファイルが`line`サブディレクトリに保存されるように変更
- **対話形式の日付範囲指定**: コマンドライン実行時に対話形式で日付範囲を指定（従来のオプション指定から変更）
- **デフォルト動作の変更**: デフォルトで今日のメールのみを処理（従来は3日分）
- **バックアップ機能の廃止**: 同名ファイル存在時のバックアップ作成を廃止、上書き保存に変更
- **メール検索制限の緩和**: 検索対象メール数を50件から5000件に拡大

### v2.1の主な変更点
- **対話形式UI**: 実行時に対話形式で日付範囲を指定する方式に変更
- **ユーザビリティ向上**: デフォルト値の提案とEnterキーでの確定機能
- **処理確認機能**: 実行前に処理対象期間の確認プロンプト
- **コマンドライン簡素化**: `--start-date`と`--end-date`オプションを廃止

### v2.2の主な変更点
- **CSV統合機能の追加**: `--merge-csvs`オプションで同フォルダ内の全CSVファイルを統合
- **ファイル名からの日付抽出**: ファイル名から年月日を自動抽出し各行に日付列を追加
- **対話形式の統合処理**: ファイル一覧表示、出力ファイル名カスタマイズ、上書き確認機能
- **統合ファイル命名規則**: `line-menu-YYYY-MM.csv`形式での統合ファイル出力

### reiwaseimeiコンテンツの細分化機能
- `contents_name.xlsx`のマッピングファイルを使用してreiwaseimeiコンテンツをさらに細分化
- amano、chamen、takashima、matsuyama、miyokof の5つのグループに分類

## 対話形式UI詳細

### 実行時の対話フロー
1. **プログラム起動**: `python line_fortune_email_processor.py`
2. **開始日の入力**: デフォルト値（今日）が提案され、Enterキーで確定または任意の日付を入力
3. **終了日の入力**: デフォルト値（開始日と同日）が提案され、Enterキーで確定または任意の日付を入力
4. **処理対象の確認**: 指定した日付範囲が表示され、実行確認を求められる
5. **実行開始**: "y"または"yes"で実行開始、それ以外で中止

### 入力形式と検証
- **日付フォーマット**: YYYY-MM-DD形式（例: 2025-07-22）
- **デフォルト値**: Enterキーのみで確定可能
- **エラーハンドリング**: 無効な日付形式の場合は再入力を促す
- **日付範囲検証**: 開始日が終了日より後の場合はエラーで終了

### 特別な処理モード
- **設定ファイル作成**: `--create-config`実行時は対話処理をスキップ
- **ファイル削除**: `--cleanup`実行時は対話処理をスキップ
- **ドライラン**: `--dry-run`実行時も通常の対話処理を実行

## 注意事項

- メールアプリパスワードの設定が必要です
- 初回実行前に設定ファイルの内容を確認してください
- 処理対象のディレクトリへの書き込み権限が必要です
- **ファイル保存先**: 従来と異なり`line`サブディレクトリ内にファイルが保存されます
- **上書き保存**: 既存ファイルは自動的に上書きされます（バックアップは作成されません）
- **対話形式**: v2.1以降、日付範囲は実行時の対話形式で指定します

## トラブルシューティング

- メール接続エラー：認証情報とサーバー設定を確認
- ファイル保存エラー：ディレクトリの権限を確認
- 詳細なログは`logs/line_fortune_processor.log`を参照