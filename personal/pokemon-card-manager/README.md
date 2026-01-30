# ポケモンカード管理システム (Google Apps Script版)

## 概要
iPhoneで撮影したトレーディングカードの写真を、Google Photos → Google Drive → AI判定 → Notion DB登録まで自動化するGASスクリプトです。

## 機能
- Google Photosの指定アルバムから新着画像を自動取得
- Google Driveへの画像バックアップ
- Perplexity APIによるカード情報の自動抽出
- 外部APIによる価格情報の補完（オプション）
- Notionデータベースへの自動登録
- エラーログと通知機能

## セットアップ手順

### 1. 必要なAPI・サービスの準備

#### Google Cloud Console
1. [Google Cloud Console](https://console.cloud.google.com/)でプロジェクトを作成
2. Photos Library APIを有効化
3. OAuth 2.0クライアントIDを作成（ウェブアプリケーション用）
4. リダイレクトURIに `https://script.google.com/macros/d/{SCRIPT_ID}/usercallback` を追加

#### Perplexity
1. [Perplexity](https://www.perplexity.ai/)でAPIキーを取得

#### Notion
1. [Notion Integrations](https://www.notion.so/my-integrations)でインテグレーションを作成
2. APIキーを取得
3. データベースを作成し、インテグレーションと共有
4. データベースIDを取得（URLから）

### 2. GASプロジェクトの設定

#### プロジェクト作成
1. [Google Apps Script](https://script.google.com/)で新規プロジェクトを作成
2. 提供されたコードファイルをコピー：
   - `Code.gs` - メイン処理
   - `PhotosAPI.gs` - Google Photos連携
   - `AIAnalysis.gs` - AI判定処理
   - `NotionAPI.gs` - Notion連携
   - `appsscript.json` - マニフェストファイル

#### OAuth2ライブラリの追加
1. ライブラリから追加ボタンをクリック
2. ライブラリID: `1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF`
3. 最新バージョンを選択して追加

#### Script Propertiesの設定
スクリプトエディタで以下を実行：
```javascript
setupScriptProperties()
```

その後、プロジェクト設定 → スクリプトプロパティから以下を設定：
- `PERPLEXITY_API_KEY` - PerplexityのAPIキー
- `NOTION_API_KEY` - NotionのAPIキー
- `NOTION_DATABASE_ID` - NotionデータベースのID
- `PHOTOS_ALBUM_ID` - Google PhotosアルバムのID
- `DRIVE_FOLDER_ID` - 保存先Google DriveフォルダのID
- `PHOTOS_CLIENT_ID` - Google CloudのクライアントID（OAuth2使用時）
- `PHOTOS_CLIENT_SECRET` - Google Cloudのクライアントシークレット（OAuth2使用時）

### 3. Notionデータベースの設定

以下のプロパティを持つデータベースを作成：

| プロパティ名 | タイプ | 説明 |
|------------|-------|------|
| Name | Title | カード名 |
| Game | Select | ゲーム種別（Pokemon/Yu-Gi-Oh!/MTG/Other） |
| Set | Text | セット名 |
| Number | Text | カード番号 |
| Rarity | Text | レアリティ |
| Language | Select | 言語（JP/EN等） |
| Condition | Text | カード状態 |
| Price | Text | 価格情報 |
| Source | URL | Google DriveのURL |
| Status | Select | ステータス（要確認/確定/要再撮影） |
| Notes | Text | メモ |

### 4. トリガーの設定

```javascript
setupTriggers()
```

これにより10分ごとに自動実行されます。

## 使い方

### 基本的なワークフロー
1. iPhoneでカードを撮影
2. Google Photosアプリで指定アルバムに追加
3. GASが自動で処理（10分間隔）
4. Notionデータベースで結果を確認

### テスト実行

#### 接続テスト
```javascript
testConnection()
```
各APIへの接続状態を確認

#### 手動実行
```javascript
main()
```
メイン処理を手動で実行

#### 処理済みIDのリセット
```javascript
resetProcessedIds()
```
再処理が必要な場合に使用

## トラブルシューティング

### よくあるエラー

#### Google Photos API関連
- 「アルバム接続失敗」→ アルバムIDを確認、共有設定を確認
- 「トークン取得失敗」→ OAuth設定を確認、再認証を実行

#### Perplexity API関連
- 「レート制限」→ 処理枚数を減らす、待機時間を増やす
- 「API Key無効」→ APIキーを確認、クレジットを確認

#### Notion API関連
- 「データベース接続失敗」→ DB IDを確認、共有設定を確認
- 「プロパティエラー」→ DBスキーマを確認

### ログの確認
- GASエディタの実行ログ
- Google Spreadsheetsのエラーログ（自動生成）
- メール通知（設定時）

## カスタマイズ

### 処理間隔の変更
`setupTriggers()`関数内の`.everyMinutes(10)`を変更

### 1回の処理枚数
Script Propertiesの`MAX_PHOTOS_PER_RUN`を変更（デフォルト: 50）

### AI判定プロンプト
`AIAnalysis.gs`の`getCardAnalysisPrompt()`関数を編集

### 利用可能なPerplexityモデル
- `llama-3.1-sonar-large-128k-online` - sonar-pro（推奨・デフォルト）
- `llama-3.1-sonar-small-128k-online` - sonar
- `llama-3.1-sonar-huge-128k-online` - sonar-huge

注意: Perplexityのsonarモデルは直接の画像解析はサポートしていませんが、画像URLを基にWeb検索を行い、カード情報を推定します。

### 外部API補完の有効化
Script Propertiesの`USE_EXTERNAL_API`を`true`に設定

## 注意事項
- API利用料金が発生します（特にPerplexity）
- 大量処理時はレート制限に注意
- プライバシー保護のため、Driveフォルダの共有設定に注意

## ライセンス
MIT License

## サポート
問題が発生した場合は、エラーログを確認の上、必要に応じて各APIのドキュメントを参照してください。