# Slack AI Bot セットアップガイド

## 初期設定手順

### 1. GASプロジェクトの作成

1. [Google Apps Script](https://script.google.com/) にアクセス
2. 新しいプロジェクトを作成
3. プロジェクト名を設定（例：Slack AI Bot）

### 2. コードのアップロード

1. `slack-bot-all-in-one.gs` の内容をコピーしてGASエディタに貼り付け
2. `appsscript.json` を表示するため：
   - プロジェクトの設定 → 「エディタで "appsscript.json" マニフェスト ファイルを表示する」をオン
   - `appsscript.json` の内容をコピーして貼り付け
3. 保存（Ctrl+S または Cmd+S）

### 3. スクリプトプロパティの設定

プロジェクトの設定 → スクリプトプロパティ から以下を設定：

#### 必須プロパティ
- `SLACK_TOKEN`: Slack Bot Token（後で取得）
- `OPEN_AI_TOKEN`: OpenAI API Key

#### オプション
- `GEMINI_TOKEN`: Gemini API Key（Geminiを使用する場合）
- `GOOGLE_NL_API`: Google Natural Language API Key（FAQ機能の高度化に必要）

#### Google Natural Language APIの設定
1. [Google Cloud Console](https://console.cloud.google.com/)にアクセス
2. プロジェクトを作成または選択
3. 「APIとサービス」→「ライブラリ」から「Cloud Natural Language API」を検索
4. APIを有効化
5. 「認証情報」→「APIキー」からAPIキーを作成
6. 取得したAPIキーを`GOOGLE_NL_API`プロパティに設定

### 4. 初期化の実行

GASエディタで以下を実行：

```javascript
initializeBot()
```

これにより：
- スプレッドシート「Slack Bot Data」が自動作成されます
- 必要なシート（log, faq, ドライブ一覧, debug_log）が作成されます
- 設定の確認が行われます
- `SPREADSHEET_ID`プロパティが自動設定されます

コンソールに表示されるスプレッドシートURLをメモしておいてください。

### 5. Webアプリとしてデプロイ

1. デプロイ → 新しいデプロイ
2. 種類：ウェブアプリ
3. 設定：
   - 実行ユーザー：自分
   - アクセスできるユーザー：全員
4. デプロイ
5. **Web App URLをコピー**（重要）

### 6. Slack Appの作成

1. [Slack API](https://api.slack.com/apps) にアクセス
2. 「Create New App」→「From an app manifest」を選択
3. ワークスペースを選択
4. 以下のManifestを貼り付け（`{GAS_WEB_APP_URL}` を実際のURLに置換）：

```yaml
display_information:
  name: ChatGPT
  description: AI Assistant Bot with Document Review
  background_color: "#2eb886"
features:
  bot_user:
    display_name: ChatGPT
    always_online: false
oauth_config:
  scopes:
    bot:
      - app_mentions:read
      - channels:history
      - channels:read
      - chat:write
      - chat:write.public
      - files:read
      - files:write
      - groups:read
      - reactions:read
      - remote_files:read
      - remote_files:share
      - users:read
settings:
  event_subscriptions:
    request_url: {GAS_WEB_APP_URL}
    bot_events:
      - message.channels
      - app_mention
      - file_shared
      - file_change
      - file_deleted
  org_deploy_enabled: false
  socket_mode_enabled: false
  token_rotation_enabled: false
```

5. 「Create」をクリック

### 7. Bot Tokenの取得と設定

1. Slack Appの「OAuth & Permissions」へ
2. 「Bot User OAuth Token」をコピー
3. GASのスクリプトプロパティで `SLACK_TOKEN` に設定

### 8. Slack Appのインストール

1. Slack Appの「Install App」へ
2. 「Install to Workspace」をクリック
3. 権限を確認して許可

### 9. 動作確認

#### GAS側のテスト

GASエディタで以下を実行：

```javascript
// 設定の確認
testSettings()

// Slack接続テスト
testSlackConnection()

// 自然言語処理とスプレッドシート連携テスト
testNLPAndSpreadsheet()
```

`testNLPAndSpreadsheet()`を実行すると：
- Google Natural Language APIの接続テスト
- 形態素解析と名詞抽出テスト
- スプレッドシート全体の分析テスト
- FAQマッチングテスト

#### Slack側のテスト

1. Botを追加したいチャンネルで `/invite @ChatGPT`
2. `@ChatGPT Hello` とメンション
3. Botから返信があれば成功

## トラブルシューティング

### エラーが発生した場合

1. **スプレッドシートの確認**
   - 初期化で作成された「Slack Bot Data」スプレッドシートを開く
   - `debug_log` シートでエラー詳細を確認

2. **よくあるエラー**

   - **「No Slack token」エラー**
     - スクリプトプロパティに `SLACK_TOKEN` が設定されているか確認
   
   - **「URL verification failed」**
     - Web App URLが正しくSlack Appに設定されているか確認
     - GASを再デプロイした場合、新しいURLをSlack Appに更新
   
   - **メンションしても反応しない**
     - Botがチャンネルに追加されているか確認
     - Event Subscriptionsが有効になっているか確認

3. **デバッグログの確認方法**
   ```javascript
   // 最新のログを確認
   function checkLogs() {
     const ss = getActiveSpreadsheet();
     const debugSheet = ss.getSheetByName('debug_log');
     const lastRow = debugSheet.getLastRow();
     const logs = debugSheet.getRange(Math.max(2, lastRow - 10), 1, Math.min(10, lastRow - 1), 4).getValues();
     logs.forEach(log => {
       console.log(log[0], log[1], log[2], log[3]);
     });
   }
   ```

## FAQ機能の設定

### 基本設定
1. スプレッドシートの「faq」シートを開く
2. 以下の列に情報を入力：
   - **A列**: キーワード（質問に含まれる可能性のある単語）
   - **B列**: 回答（そのキーワードに対する回答内容）
   - **C列**: Drive検索フラグ（TRUE/FALSE）
   - **D列**: Drive検索結果（自動入力）

### 自然言語処理の仕組み
Google Natural Language APIが設定されている場合：
1. 質問文から名詞・数値を自動抽出
2. カタカナ→ひらがな、全角→半角に正規化
3. FAQシートのA列・B列と高精度でマッチング

### 例
| A列（キーワード） | B列（回答） | C列 | D列 |
|---|---|---|---|
| 休暇申請 | 休暇申請はシステムから申請してください | FALSE | |
| 経費精算 | 経費精算は月末までに申請書を提出 | TRUE | (自動入力) |

## Drive検索機能の設定

1. スプレッドシートの「ドライブ一覧」シートを開く
2. A列に検索対象のフォルダIDを入力

## 機能一覧

### 基本機能
- **AI対話**: ChatGPT/Geminiを使ったメンション応答
- **スレッド対応**: 会話の文脈を維持
- **チャンネル説明の反映**: チャンネルのpurposeをAIのコンテキストに含める

### ファイル処理
- **PDF読み取り**: OCR対応で日本語も処理可能
- **Word文書処理**: .docx、.doc形式に対応
- **Googleドキュメント**: URLまたは共有リンクから直接アクセス
- **ドキュメントレビュー**: AIによる自動修正・コメント付与

### 自然言語処理機能
- **形態素解析**: Google Natural Language APIで質問文を解析
- **FAQ自動応答**: 高度なマッチングアルゴリズム
- **スプレッドシート分析**: 全シートのテキストを自然言語処理
- **感情分析**: テキストのポジティブ/ネガティブ判定

### Drive連携
- **全文検索**: 指定フォルダ内のファイルを検索
- **スニペット抽出**: キーワードを含む段落・行を自動抽出
- **リッチテキスト出力**: 検索結果をリンク付きでスプレッドシートに記載

## 更新時の注意

GASを更新した場合：
1. 新しいバージョンをデプロイ
2. 新しいWeb App URLをコピー
3. Slack Appの Event Subscriptions で Request URL を更新

## パフォーマンス最適化

### スプレッドシートの最適化
- FAQシートは`SpreadsheetApp.getActive()`を使用して高速アクセス
- 必要最小限の列のみを読み取り（A:B列）
- 文字列正規化処理の最適化

### 自然言語処理の最適化
- テキストを最大5000文字に制限してAPIコストを削減
- 品詞フィルタリングで必要な情報のみ抽出
- キャッシュ機能で重複処理を回避