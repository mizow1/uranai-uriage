﻿# Requirements Document

## Introduction

この機能は、売上集計システムで生成されたCSVデータを元に、コンテンツ提供元への支払い明細書を自動生成するものです。既存の明細書テンプレートを参照し、最新の売上データを反映した新しい明細書を作成します。明細書はExcel形式で出力され、プラットフォーム名、コンテンツ名、売上金額などの情報が含まれます。

## Requirements

### Requirement 1

**User Story:** システム管理者として、売上集計データから明細書を自動生成したい。それにより、手作業での明細書作成の手間と時間を削減できる。

#### Acceptance Criteria

1. WHEN 売上集計CSVファイル（output.csv）が存在する場合 THEN システムはそのデータを読み込むことができる
2. WHEN 明細書テンプレートフォルダ（C:\Users\OW\Dropbox\disk2とローカルの同期\占い\占い売上\履歴\ロイヤリティ\）が存在する場合 THEN システムは最新の明細書テンプレートを特定できる
3. WHEN 明細書テンプレートが見つかった場合 THEN システムはそのフォーマットを新しい明細書作成に使用できる
4. WHEN 明細書テンプレートが見つからない場合 THEN システムは適切なエラーメッセージを表示する

### Requirement 2

**User Story:** システム管理者として、売上データと明細書テンプレートの商品名やプラットフォーム名が完全に一致しなくても、適切にマッチングさせたい。それにより、表記の揺れがあっても正確な明細書を作成できる。

#### Acceptance Criteria

1. WHEN 売上データのコンテンツ名と明細書テンプレートの商品名の表記が異なる場合 THEN システムはファジーマッチングを使用して適切に対応付ける
2. WHEN 売上データのプラットフォーム名と明細書テンプレートのプラットフォーム名の表記が異なる場合 THEN システムはファジーマッチングを使用して適切に対応付ける
3. WHEN マッチングの信頼度が低い場合 THEN システムは警告を表示し、ユーザーに確認を求める

### Requirement 3

**User Story:** システム管理者として、生成された明細書を適切な命名規則で保存したい。それにより、明細書の管理と追跡が容易になる。

#### Acceptance Criteria

1. WHEN 新しい明細書が生成される THEN システムは「yyyymmdd.xls」形式のファイル名で保存する
2. WHEN 明細書が生成される THEN システムは適切なフォルダ構造（年フォルダ内）に保存する
3. WHEN 同名のファイルが既に存在する場合 THEN システムは上書き前に確認を求める

### Requirement 4

**User Story:** システム管理者として、明細書生成プロセスの進行状況と結果を確認したい。それにより、処理の成功または失敗を把握できる。

#### Acceptance Criteria

1. WHEN 明細書生成プロセスが開始される THEN システムは処理開始メッセージを表示する
2. WHEN 明細書生成プロセスが進行中 THEN システムは現在の処理ステップを表示する
3. WHEN 明細書生成プロセスが完了する THEN システムは処理結果の概要（成功または失敗、生成されたファイルのパスなど）を表示する
4. WHEN エラーが発生した場合 THEN システムは具体的なエラーメッセージと可能な解決策を表示する
