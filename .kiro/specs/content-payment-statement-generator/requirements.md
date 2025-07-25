# Requirements Document

## Introduction

コンテンツ関連支払い明細書の自動作成・送信システムは、各プラットフォーム（LINE等）のコンテンツ売上データを基に、指定された年月の支払い明細書を自動生成し、PDF化してメール送信する機能です。このシステムにより、月次の支払い処理業務を効率化し、手作業によるミスを削減します。

## Requirements

### Requirement 1

**User Story:** As a 経理担当者, I want 指定した年月の売上データから支払い明細書を自動生成したい, so that 月次処理の作業時間を短縮し、計算ミスを防げる

#### Acceptance Criteria

1. WHEN ユーザーが処理対象年月（YYYYMM形式）を指定 THEN システム SHALL 指定された年月の売上データを各ソースファイルから取得する
2. WHEN 月別ISP別コンテンツ別売上.csvが存在する THEN システム SHALL 対象年月の各プラットフォームの各コンテンツの「実績」「情報提供料合計」を抽出する
3. WHEN LINE用のline-contents-yyyy-mm.csvが存在する THEN システム SHALL LINEプラットフォームの各コンテンツの「実績」「情報提供料」を抽出する

### Requirement 2

**User Story:** As a 経理担当者, I want コンテンツマッピング情報を使って正しいフォーマットファイルを特定したい, so that 各コンテンツに対応する適切な明細書を作成できる

#### Acceptance Criteria

1. WHEN contents_mapping.csvが読み込まれる THEN システム SHALL 各プラットフォームのコンテンツ名からA列の値（フォーマットファイル名）を取得する
2. WHEN フォーマットファイル名が特定される THEN システム SHALL コンテンツ関連支払明細書フォーマットディレクトリから同名ファイルを検索する
3. WHEN 対象ファイルが見つかる THEN システム SHALL ファイルを指定の出力ディレクトリに「YYYYMM_元ファイル名」として複製する

### Requirement 3

**User Story:** As a 経理担当者, I want Excelファイルに売上データを正しい形式で記入したい, so that 支払い明細書として使用できる

#### Acceptance Criteria

1. WHEN Excelファイルが複製される THEN システム SHALL S3セルに対象年月の翌月5日の日付を記入する
2. WHEN 売上データが処理される THEN システム SHALL 23行目以降に以下の形式で明細を記入する
   - A列：対象年月の翌月5日
   - D列：プラットフォーム名
   - G列：コンテンツ名
   - M列：対象年月
   - S列：「実績」額
   - Y列：「情報提供料」額
   - AC列：rate.csvから取得した料率
3. WHEN rate.csvが参照される THEN システム SHALL 「名称」列がファイル名と一致するレコードのB列値をAC列に設定する

### Requirement 4

**User Story:** As a 経理担当者, I want 作成した明細書をPDF化して自動送信したい, so that 手作業でのメール送信作業を削減できる

#### Acceptance Criteria

1. WHEN Excelファイルの編集が完了する THEN システム SHALL ファイルをPDF形式に変換する
2. WHEN PDF変換が完了する THEN システム SHALL rate.csvから対応するメールアドレス（C列）を取得する
3. WHEN メール送信処理が実行される THEN システム SHALL 以下の設定でGmailを使用してメール送信する
   - 宛先：rate.csvのC列メールアドレス
   - 差出人：mizoguchi@outward.jp
   - CC：ow-fortune@ml.outward.jp
   - BCC：mizoguchi@outward.jp
   - 件名：「今月のお支払額のご連絡」
   - 本文：「今月のお支払額をご連絡いたします。」
   - 添付ファイル：生成されたPDF
4. WHEN メール送信が設定される THEN システム SHALL 対象月の翌月5日に予約送信を設定する

### Requirement 5

**User Story:** As a システム利用者, I want エラーが発生した場合に適切な情報を得たい, so that 問題を特定して対処できる

#### Acceptance Criteria

1. WHEN 必要なファイルが存在しない THEN システム SHALL 具体的なファイルパスとエラー内容を表示する
2. WHEN データ処理中にエラーが発生する THEN システム SHALL エラーの詳細と発生箇所を記録する
3. WHEN メール送信に失敗する THEN システム SHALL 失敗理由と対象ファイル情報をログに記録する