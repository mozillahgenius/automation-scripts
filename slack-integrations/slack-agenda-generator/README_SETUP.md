# Google Apps Script セットアップ手順

## 重要：ファイルの配置方法

### 1. Google Apps Scriptプロジェクトの構成

Google Apps Scriptエディタで以下のファイルを作成：

#### ファイル構成：
```
📁 Google Apps Scriptプロジェクト
├── 📄 IntegratedSlackSystem.gs （メインコード）
├── 📄 setup.html （設定用HTML）
└── ⚙️ appsscript.json （マニフェストファイル - 特別な設定が必要）
```

### 2. appsscript.jsonの正しい設定方法

**重要：** `appsscript.json`は通常の.gsファイルとは異なり、マニフェストファイルとして設定します。

#### 設定手順：

1. **Google Apps Scriptエディタを開く**
   - スプレッドシート → 拡張機能 → Apps Script

2. **プロジェクト設定を開く**
   - 左側メニューの「プロジェクトの設定」（歯車アイコン）をクリック

3. **「appsscript.json」マニフェストファイルをエディタで表示する」にチェック**
   - 「全般設定」セクションで有効化

4. **エディタに戻る**
   - 左側メニューの「エディタ」をクリック
   - ファイルリストに`appsscript.json`が表示される

5. **appsscript.jsonを編集**
   - クリックして開き、以下の内容に置き換える：

```json
{
  "timeZone": "Asia/Tokyo",
  "dependencies": {
    "enabledAdvancedServices": []
  },
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.labels",
    "https://www.googleapis.com/auth/script.scriptapp",
    "https://www.googleapis.com/auth/userinfo.email"
  ],
  "webapp": {
    "executeAs": "USER_DEPLOYING",
    "access": "ANYONE_ANONYMOUS"
  }
}
```

### 3. ファイルの作成順序

1. **IntegratedSlackSystem.gs**
   - 「ファイルを追加」→「スクリプト」
   - ファイル名: `IntegratedSlackSystem`（.gsは自動付与）
   - コードを貼り付け

2. **setup.html**
   - 「ファイルを追加」→「HTML」
   - ファイル名: `setup`（.htmlは自動付与）
   - HTMLコードを貼り付け

3. **appsscript.json**
   - 上記の手順2で説明した方法で設定

### 4. スクリプトプロパティの設定

プロジェクトの設定 → スクリプトプロパティで以下を追加：

| プロパティ名 | 値の例 |
|------------|-------|
| SPREADSHEET_ID | 1234567890abcdefg |
| SLACK_BOT_TOKEN | xoxb-1234567890-... |
| OPENAI_API_KEY | sk-proj-... |
| REPORT_EMAIL | your-email@example.com |

### 5. 初回実行

1. **保存**
   - Ctrl+S または Cmd+S ですべてのファイルを保存

2. **実行テスト**
   - `testSystem`関数を選択
   - 「実行」ボタンをクリック
   - 初回は権限の承認が必要

3. **スプレッドシートでメニュー確認**
   - スプレッドシートをリロード
   - 「🤖 統合議題生成システム」メニューが表示される

## トラブルシューティング

### エラー：「Unexpected token ':'」
→ appsscript.jsonを通常の.gsファイルとして作成している
→ 解決：上記の手順2に従ってマニフェストファイルとして設定

### エラー：「スクリプトプロパティが見つかりません」
→ プロパティが設定されていない
→ 解決：プロジェクトの設定からスクリプトプロパティを追加

### メニューが表示されない
→ onOpen()関数が実行されていない
→ 解決：スプレッドシートをリロードするか、onOpen()を手動実行

## 必要な外部サービス

1. **Slack App**
   - https://api.slack.com/apps で作成
   - 必要なスコープを設定
   - Bot Tokenを取得

2. **OpenAI API**
   - https://platform.openai.com でアカウント作成
   - APIキーを生成

3. **Gmail**
   - 送信に使用するGoogleアカウント
   - 初回実行時に権限承認が必要