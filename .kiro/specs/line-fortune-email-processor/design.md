# 設計文書: LINE Fortune メール処理システム（リファクタリング版）

## 概要

LINE Fortune メール処理システムは、LINE Fortuneの日次レポートメールの処理を自動化するためのPythonアプリケーションです。システムはdev-line-fortune@linecorp.comからの特定のメールを監視し、CSV添付ファイルを抽出し、日付ベースの命名規則に従ってリネームし、特定のフォルダ構造に整理します。さらに、システムはデータ分析を容易にするために、月次のCSVファイルを統合します。

このリファクタリング版では、既存の実装を改善し、コードの保守性、テスト可能性、エラー処理を強化します。

## アーキテクチャ

アプリケーションは明確な関心の分離を持つモジュラーアーキテクチャに従います：

1. **設定モジュール**: アプリケーション設定の管理と検証を担当します。
2. **メールモジュール**: メールサーバーへの接続、新規メールの監視、添付ファイルの抽出を担当します。
3. **ファイル処理モジュール**: ファイルのリネーム、ディレクトリの作成、ファイルの保存操作を処理します。
4. **統合モジュール**: 複数のCSVファイルを単一の月次レポートに統合する管理を行います。
5. **ログ記録モジュール**: 監視とトラブルシューティングのための包括的なログ機能を提供します。
6. **メインコントローラー**: 全体的なプロセスフローを調整し、エラー回復を処理します。

## リファクタリングの改善点

現在の実装から特定された主要な改善点：

### 1. コードの複雑性の軽減
- **メール検索ロジックの簡素化**: 複数の検索戦略を統一し、可読性を向上
- **エラーハンドリングの統一**: 一貫したエラー処理パターンの適用
- **設定管理の改善**: 設定の検証と管理の強化

### 2. テスト可能性の向上
- **依存性注入の導入**: モックテストを容易にするための構造改善
- **メソッドの分割**: 大きなメソッドを小さな単位に分割
- **副作用の分離**: 純粋関数とI/O操作の分離

### 3. 保守性の向上
- **定数の外部化**: ハードコードされた値の設定ファイル化
- **ログメッセージの標準化**: 一貫したログフォーマットの適用
- **ドキュメントの充実**: コードコメントとドキュメントの改善

### 4. パフォーマンスの最適化
- **メール検索の効率化**: 不要な検索の削減
- **メモリ使用量の最適化**: 大きなファイル処理時のメモリ効率改善
- **並行処理の検討**: 複数ファイル処理の並列化

## コンポーネントとインターフェース

### メールモジュール

```python
class EmailProcessor:
    def __init__(self, email_config):
        # メールサーバー設定で初期化
        pass
        
    def connect(self):
        # メールサーバーに接続
        pass
        
    def fetch_matching_emails(self, sender, recipient, subject_pattern):
        # 条件に一致するメールを取得
        pass
        
    def extract_attachments(self, email, file_type=".csv"):
        # 指定されたタイプの添付ファイルを抽出
        pass
        
    def extract_date_from_subject(self, subject):
        # メール件名から日付を抽出
        pass
```

### ファイル処理モジュール

```python
class FileProcessor:
    def __init__(self, base_path):
        # ファイル保存のベースパスで初期化
        pass
        
    def create_directory_structure(self, date):
        # 年/月のディレクトリ構造を作成
        pass
        
    def rename_file(self, original_filename, date):
        # 日付を含めるようにファイルをリネーム
        pass
        
    def save_file(self, file_content, filename, directory):
        # 指定されたディレクトリにファイルを保存
        pass
```

### 統合モジュール

```python
class ConsolidationProcessor:
    def __init__(self):
        pass
        
    def consolidate_csv_files(self, directory, output_filename):
        # ディレクトリ内のすべてのCSVファイルを統合
        pass
```

### ログ記録モジュール

```python
class Logger:
    def __init__(self, log_file):
        # ロガーを初期化
        pass
        
    def info(self, message):
        # 情報メッセージをログに記録
        pass
        
    def warning(self, message):
        # 警告メッセージをログに記録
        pass
        
    def error(self, message, exception=None):
        # オプションの例外付きでエラーメッセージをログに記録
        pass
```

### メインコントローラー

```python
class LineFortuneProcessor:
    def __init__(self, config):
        # 設定で初期化
        self.email_processor = EmailProcessor(config['email'])
        self.file_processor = FileProcessor(config['base_path'])
        self.consolidation_processor = ConsolidationProcessor()
        self.logger = Logger(config['log_file'])
        
    def process(self):
        # メイン処理ロジック
        pass
        
    def handle_email(self, email):
        # 単一のメールを処理
        pass
```

## データモデル

### 設定モデル

```python
config = {
    'email': {
        'server': 'imap.example.com',
        'port': 993,
        'username': 'mizoguchi@outward.jp',
        'password': '********',
        'use_ssl': True
    },
    'base_path': r'C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ISP支払通知書',
    'log_file': 'line_fortune_processor.log',
    'sender': 'dev-line-fortune@linecorp.com',
    'recipient': 'mizoguchi@outward.jp',
    'subject_pattern': 'LineFortune Daily Report'
}
```

### メールデータモデル

```python
class Email:
    def __init__(self, id, sender, recipient, subject, date, attachments):
        self.id = id
        self.sender = sender
        self.recipient = recipient
        self.subject = subject
        self.date = date
        self.attachments = attachments  # 添付ファイルオブジェクトのリスト
```

### 添付ファイルデータモデル

```python
class Attachment:
    def __init__(self, filename, content_type, content):
        self.filename = filename
        self.content_type = content_type
        self.content = content  # バイナリコンテンツ
```

## 処理フロー

1. **初期化**:
   - 設定の読み込み
   - コンポーネントの初期化
   - メールサーバーへの接続

2. **メール処理**:
   - 条件に一致するメールの取得
   - 一致する各メールに対して:
     - 件名から日付を抽出
     - CSV添付ファイルを抽出
     - 日付に基づいて対象ディレクトリを決定
     - 必要に応じてディレクトリ構造を作成
     - 添付ファイルのリネームと保存
     - 月次ディレクトリ内のCSVファイルの統合
     - 成功をログに記録

3. **エラー処理**:
   - 一時的なエラーに対する再試行メカニズムの実装
   - すべてのエラーをコンテキスト付きでログに記録
   - 可能な場合は他のメール/ファイルの処理を継続

## エラー処理

アプリケーションは包括的なエラー処理戦略を実装します：

1. **メール接続エラー**: 指数バックオフによる接続の再試行
2. **ファイルシステムエラー**: ファイル操作を最大3回再試行
3. **解析エラー**: 詳細情報をログに記録し、次のメールに進む
4. **認証エラー**: ユーザーに警告し、適切に終了
5. **予期しないエラー**: すべての例外をキャッチし、詳細をログに記録し、可能な場合は継続

## テスト戦略

### ユニットテスト

1. **メールモジュールテスト**:
   - モックサーバーによるメール接続のテスト
   - メールフィルタリングロジックのテスト
   - 添付ファイル抽出のテスト
   - 様々な件名形式からの日付抽出のテスト

2. **ファイル処理テスト**:
   - ディレクトリ作成のテスト
   - ファイルリネームのテスト
   - モックファイルシステムによるファイル保存のテスト

3. **統合テスト**:
   - サンプルファイルによるCSV統合のテスト
   - ヘッダー処理のテスト
   - 空ディレクトリケースのテスト

### 統合テスト

1. モックメールサーバーによるエンドツーエンドフローのテスト
2. 一時ディレクトリによるファイルシステム統合のテスト
3. エラー回復シナリオのテスト

### 手動テスト

1. 実際のメールアカウントによるテスト（テストメールを使用）
2. 正確なファイル配置と命名の検証
3. 統合ファイルの内容の検証

## セキュリティ考慮事項

1. **メール認証**: 認証情報を安全に保存し、環境変数または安全な認証情報ストアの使用を検討
2. **ファイルシステムアクセス**: アプリケーションが対象ディレクトリに適切な権限を持っていることを確認
3. **エラーログ**: 機密情報がログに露出しないようにする

## デプロイメント考慮事項

1. **スケジューリング**: 毎日実行するスケジュールタスクとして設定
2. **監視**: 障害を警告するための監視の実装
3. **更新**: コード変更なしで設定を更新するメカニズムの提供