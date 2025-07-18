# LINE Fortune Email Processor

LINE Fortuneの日次レポートメールを自動処理するシステムです。

## 機能

- LINE Fortuneからの日次レポートメールの自動監視
- CSV添付ファイルの自動抽出と保存
- 日付ベースのファイル名変更
- 年/月のディレクトリ構造での自動整理
- 月次統合CSVファイルの自動生成
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

### 基本的な実行
```bash
python line_fortune_email_processor.py
```

### ドライラン（実際の処理は行わない）
```bash
python line_fortune_email_processor.py --dry-run
```

### カスタム設定ファイルを使用
```bash
python line_fortune_email_processor.py --config my_config.json
```

### 古いファイルの削除（30日より古いファイルを削除）
```bash
python line_fortune_email_processor.py --cleanup 30
```

### ログレベルの設定
```bash
python line_fortune_email_processor.py --log-level DEBUG
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

- 個別CSVファイル: `元のファイル名_YYYY-MM-DD.csv`
- 統合CSVファイル: `line-menu-YYYY-MM.csv`
- ログファイル: `logs/line_fortune_processor.log`

## 注意事項

- メールアプリパスワードの設定が必要です
- 初回実行前に設定ファイルの内容を確認してください
- 処理対象のディレクトリへの書き込み権限が必要です

## トラブルシューティング

- メール接続エラー：認証情報とサーバー設定を確認
- ファイル保存エラー：ディレクトリの権限を確認
- 詳細なログは`logs/line_fortune_processor.log`を参照