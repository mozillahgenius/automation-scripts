# Slack App 設定ガイド

## Slack Appの作成と必要な権限設定

### 1. Slack Appの作成

1. **Slack API サイトにアクセス**
   - https://api.slack.com/apps
   - 「Create New App」をクリック

2. **From scratchを選択**
   - App Name: `Slack議題生成Bot`（任意の名前）
   - Pick a workspace: 使用するワークスペースを選択

### 2. 必要なBot Token Scopes（権限）

**OAuth & Permissions** ページで以下のスコープを追加：

#### 必須スコープ（Bot Token Scopes）：
```
channels:history     - パブリックチャンネルのメッセージ履歴を読む
channels:read        - パブリックチャンネルの情報を取得
chat:write          - メッセージを投稿する
users:read          - ユーザー情報を取得
groups:history      - プライベートチャンネルのメッセージ履歴を読む
groups:read         - プライベートチャンネルの情報を取得
im:history          - ダイレクトメッセージの履歴を読む
mpim:history        - グループDMの履歴を読む
```

#### スコープ設定のコード内での使用箇所：

```javascript
// IntegratedSlackSystem.gs 内で使用されているSlack API

// channels:read, groups:read が必要
// Line 108-134: conversations.list
const channelsUrl = 'https://slack.com/api/conversations.list';

// channels:read, groups:read が必要  
// Line 206-232: conversations.info
const channelInfoUrl = `https://slack.com/api/conversations.info?channel=${channelId}`;

// channels:history, groups:history, im:history, mpim:history が必要
// Line 234-284: conversations.history
const historyUrl = 'https://slack.com/api/conversations.history';

// channels:history が必要（スレッドの返信取得）
// Line 285-310: conversations.replies
const repliesUrl = `https://slack.com/api/conversations.replies`;

// users:read が必要
// Line 1162-1188: users.info
const userInfoUrl = `https://slack.com/api/users.info?user=${userId}`;

// chat:write が必要
// Line 1419-1540: chat.postMessage
const postMessageUrl = 'https://slack.com/api/chat.postMessage';
```

### 3. Slack Appのインストール

1. **OAuth & Permissions** ページで
   - 「Install to Workspace」ボタンをクリック
   - 権限を確認して「許可する」

2. **Bot User OAuth Token を取得**
   - インストール後に表示される`xoxb-`で始まるトークンをコピー
   - このトークンを`SLACK_BOT_TOKEN`として使用

### 4. Signing Secret の取得

1. **Basic Information** ページで
   - 「App Credentials」セクションを確認
   - 「Signing Secret」の「Show」をクリック
   - 表示されるシークレットをコピー
   - このシークレットを`SLACK_SIGNING_SECRET`として使用

### 5. チャンネルへのBot追加

Botがチャンネルのメッセージを読むには、そのチャンネルにBotを追加する必要があります：

1. **Slackワークスペースで**
   - 監視したいチャンネルを開く
   - チャンネル名をクリック → 「インテグレーション」タブ
   - 「アプリを追加」をクリック
   - 作成したBotを選択して追加

### 6. Google Apps Script での設定

スクリプトプロパティに以下を設定：

```javascript
// プロパティ名: 値
SLACK_BOT_TOKEN: YOUR_SLACK_BOT_TOKEN
SLACK_SIGNING_SECRET: YOUR_SLACK_SIGNING_SECRET
```

### 7. 動作確認チェックリスト

- [ ] Slack Appを作成した
- [ ] 必要なBot Token Scopesをすべて追加した
- [ ] ワークスペースにインストールした
- [ ] Bot User OAuth Tokenを取得した
- [ ] Signing Secretを取得した
- [ ] 監視したいチャンネルにBotを追加した
- [ ] Google Apps ScriptにTokenとSecretを設定した
- [ ] テスト実行でSlackメッセージが取得できることを確認した

### 8. トラブルシューティング

#### エラー: `not_in_channel`
**原因**: BotがチャンネルのメンバーではないZ
**解決**: チャンネルにBotを追加する

#### エラー: `invalid_auth`
**原因**: Bot Tokenが無効
**解決**: トークンが正しくコピーされているか確認

#### エラー: `missing_scope`
**原因**: 必要な権限が不足
**解決**: OAuth & Permissionsで必要なスコープを追加して再インストール

#### エラー: `channel_not_found`
**原因**: チャンネルIDが間違っている
**解決**: Slackでチャンネルを右クリック→「リンクをコピー」でIDを確認

### 9. セキュリティのベストプラクティス

1. **トークンの管理**
   - Bot Tokenは絶対に公開しない
   - Gitにコミットしない
   - 定期的にトークンを再生成する

2. **権限の最小化**
   - 必要最小限のスコープのみを付与
   - 不要になったスコープは削除

3. **アクセス制御**
   - 特定のチャンネルのみを監視対象にする
   - Configシートで有効/無効を管理

### 10. 参考リンク

- [Slack API Documentation](https://api.slack.com/docs)
- [OAuth Scopes](https://api.slack.com/scopes)
- [Bot Users](https://api.slack.com/bot-users)
- [Conversations API](https://api.slack.com/docs/conversations-api)