# Salesforce Object Explorer for Google Sheets

Google Apps Script (GAS)を使用して、Salesforce REST APIからオブジェクトとフィールドの一覧を取得し、Google スプレッドシートに出力するツールです。

## 機能

- Salesforceのオブジェクト一覧の取得
- 各オブジェクトのフィールド詳細情報の取得
- Google スプレッドシートへの自動出力
- 設定画面による簡単なセットアップ

## セットアップ手順

### 1. Salesforce側の準備

1. Salesforceにログイン
2. 設定 → アプリケーション → アプリマネージャー
3. 「新規接続アプリケーション」をクリック
4. 以下の情報を入力:
   - 接続アプリケーション名: `GAS Salesforce Explorer`
   - API参照名: `GAS_Salesforce_Explorer`
   - 連絡先メールアドレス: あなたのメールアドレス
   - OAuth設定を有効化: チェック
   - コールバックURL: `https://login.salesforce.com/services/oauth2/callback`
   - 選択したOAuth範囲:
     - データへのアクセスと管理(api)
     - いつでも要求を実行(refresh_token, offline_access)
5. 保存後、Consumer KeyとConsumer Secretをメモ

### 2. セキュリティトークンの取得

1. Salesforceの個人設定を開く
2. 「私のセキュリティトークンのリセット」をクリック
3. メールで送られてきたセキュリティトークンをメモ

### 3. Google Apps Scriptの設定

1. Google スプレッドシートを新規作成
2. 拡張機能 → Apps Script を開く
3. 以下のファイルを作成:
   - `Code.gs` - メインコード
   - `ConfigDialog.html` - 設定画面
   - `appsscript.json` - マニフェストファイル
4. 各ファイルの内容をコピー&ペースト
5. 保存してスプレッドシートをリロード

### 4. 初期設定

1. スプレッドシートのメニューに「Salesforce Explorer」が表示される
2. 「Salesforce Explorer」→「設定を初期化」をクリック
3. 表示されたダイアログに以下を入力:
   - Client ID: SalesforceのConsumer Key
   - Client Secret: SalesforceのConsumer Secret
   - ユーザー名: Salesforceログインユーザー名
   - パスワード: Salesforceログインパスワード
   - セキュリティトークン: メールで取得したトークン
   - ログインURL:
     - 本番環境: `https://login.salesforce.com`
     - Sandbox: `https://test.salesforce.com`

## 使用方法

### オブジェクト一覧の取得

1. メニューから「Salesforce Explorer」→「オブジェクト一覧を取得」
2. 「Objects」シートにオブジェクト一覧が出力される

### フィールド情報の取得

1. メニューから「Salesforce Explorer」→「選択オブジェクトのフィールドを取得」
2. オブジェクト名を入力（例: Account, Contact, CustomObject__c）
3. 「Fields_[オブジェクト名]」シートにフィールド一覧が出力される

## 出力される情報

### オブジェクト一覧
- オブジェクト名
- ラベル
- API名
- カスタムオブジェクトかどうか
- CRUD権限（作成、更新、削除、検索）

### フィールド一覧
- フィールド名
- ラベル
- API名
- データ型
- 長さ
- 必須項目かどうか
- ユニーク制約
- カスタムフィールドかどうか
- CRUD権限
- NULL許可
- 参照先オブジェクト

## トラブルシューティング

### 認証エラーが発生する場合

1. セキュリティトークンが正しいか確認
2. IPアドレス制限を確認（必要に応じて信頼済みIPに追加）
3. ログインURLが環境に合っているか確認

### API制限エラー

- Salesforce組織のAPI呼び出し制限を確認
- 大量のオブジェクトがある場合は、順次実行する

## 注意事項

- API呼び出し回数に制限があるため、頻繁な実行は避けてください
- 大規模な組織の場合、処理に時間がかかる場合があります
- 認証情報は暗号化されてGoogle Apps Scriptのプロパティに保存されます