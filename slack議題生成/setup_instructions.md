# Slack議題生成システム セットアップ手順

## 必要な準備

### 1. Google Apps Scriptプロジェクトの作成
1. Google Driveで新規スプレッドシートを作成
2. 拡張機能 → Apps Script を開く
3. `IntegratedSlackSystem.gs`と`setup.html`のコードをコピー
4. `appsscript.json`を マニフェストファイルに貼り付け

### 2. 必要なAPIキーの取得

#### Slack Bot Token
1. https://api.slack.com/apps にアクセス
2. 新規アプリを作成
3. OAuth & Permissions で以下のスコープを追加：
   - `channels:history`
   - `channels:read`
   - `chat:write`
   - `users:read`
   - `groups:history`
   - `groups:read`
   - `im:history`
   - `mpim:history`
4. ワークスペースにインストール
5. Bot User OAuth Tokenをコピー

#### OpenAI APIキー
1. https://platform.openai.com にアクセス
2. API Keysページで新規キーを作成
3. キーをコピー（一度しか表示されません）

### 3. スクリプトプロパティの設定

Apps Scriptのプロジェクト設定 → スクリプトプロパティで以下を設定：

| プロパティ名 | 値 | 説明 |
|------------|---|-----|
| SPREADSHEET_ID | スプレッドシートのID | URLの `/d/` と `/edit` の間の文字列 |
| SLACK_BOT_TOKEN | xoxb-で始まるトークン | Slack Bot User OAuth Token |
| OPENAI_API_KEY | sk-で始まるキー | OpenAI APIキー |
| SLACK_SIGNING_SECRET | Slackアプリのsigning secret | Basic Informationから取得 |
| REPORT_EMAIL | 送信先メールアドレス | レポートの送信先 |

### 4. 初回実行

1. スプレッドシートを開く
2. メニューに「🤖 統合議題生成システム」が表示される
3. 「⚙️ 初期設定」を実行
4. 必要な権限を許可

### 5. 必要なシートの自動作成

以下のシートが自動的に作成されます：
- Config：設定管理
- SyncState：同期状態
- Messages：メッセージ一覧
- Categories：カテゴリ定義
- slack_log：Slackログ
- AgendaAnalysis：議題分析結果

### 6. Slackチャンネルの設定

Configシートに監視したいチャンネルを追加：
1. channel_id：チャンネルID（例：C01234567）
2. channel_name：チャンネル名（例：general）
3. enabled：TRUE

### 7. 自動実行の設定

メニューから「⏰ 自動実行タイマー設定」を実行すると：
- 30分ごと：Slack同期
- 1時間ごと：AI分析
- 毎日9時：議題レポート送信

## 動作確認

### 手動テスト
1. 「🔄 Slack全チャンネル同期」：Slackメッセージを取得
2. 「🎯 Slack議題抽出＆メール送信」：議題抽出とメール送信
3. slack_logシートにデータが記録されることを確認

### トラブルシューティング

#### エラー：「OpenAI APIキーが設定されていません」
→ スクリプトプロパティにOPENAI_API_KEYを設定

#### エラー：「Slack APIエラー」
→ Bot TokenとチャンネルIDを確認

#### エラー：「メール送信失敗」
→ REPORT_EMAILが正しいメールアドレスか確認

## セキュリティ注意事項

- APIキーは絶対に公開しない
- スプレッドシートの共有設定に注意
- 定期的にAPIキーを更新する
- ログに個人情報が含まれないよう注意

## サポート

問題が発生した場合：
1. Logsシートでエラーログを確認
2. Apps Scriptのログを確認
3. APIの利用制限を確認