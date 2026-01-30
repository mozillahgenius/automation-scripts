# Slack議題生成自動化システム

## 概要

SlackのメッセージをAIで分析し、取締役会・監査等委員会・株主総会の議題候補を自動抽出して、議事録案を生成するGoogle Apps Scriptシステムです。

## 主な機能

- 📨 **Slack連携**: 指定チャンネルのメッセージを自動収集
- 🤖 **AI分析**: OpenAI APIで要約・分類・判定
- 📊 **スプレッドシート管理**: 全データを一元管理
- 📢 **自動通知**: Slack/メールで該当案件を通知
- 📝 **ドキュメント生成**: 議事録案をGoogleドキュメントに出力

## セットアップ手順

### 1. 前提条件

- Google Workspace アカウント
- Slack ワークスペースの管理者権限
- OpenAI API アカウント

### 2. Slack App の作成

1. [Slack API](https://api.slack.com/apps) にアクセス
2. 「Create New App」→「From scratch」を選択
3. App Name と Workspace を設定

#### 必要な OAuth Scopes

**Bot Token Scopes**に以下を追加:
- `channels:history` - 公開チャンネルの履歴読み取り
- `groups:history` - プライベートチャンネルの履歴読み取り
- `channels:read` - チャンネル情報の読み取り
- `users:read` - ユーザー情報の読み取り
- `chat:write` - メッセージ投稿

4. 「Install to Workspace」でインストール
5. Bot User OAuth Token をコピー（`xoxb-`で始まる）

### 3. OpenAI API Key の取得

1. [OpenAI Platform](https://platform.openai.com/) にログイン
2. API Keys セクションで新しいキーを作成
3. API Key をコピー（`sk-`で始まる）

### 4. Google スプレッドシートの準備

1. 新しいGoogleスプレッドシートを作成
2. ツール → スクリプトエディタを開く
3. 以下のファイルをコピー:
   - `Code.gs` - メインコード
   - `InitializeSpreadsheet.gs` - 初期化スクリプト
   - `Schemas.gs` - スキーマ定義
   - `setup.html` - 設定画面

### 5. 初期設定

1. スクリプトエディタで実行:
   ```
   InitializeSpreadsheet.gs → initializeSpreadsheet()
   ```
   
2. スプレッドシートのメニューから:
   ```
   🤖 議題生成システム → ⚙️ 初期設定
   ```

3. 設定画面で以下を入力:
   - Slack Bot Token
   - OpenAI API Key
   - 監視対象チャンネルID
   - 通知先設定

### 6. チャンネルIDの確認方法

Slackでチャンネルを右クリック → 「チャンネル詳細を表示」→ 最下部のチャンネルID

### 7. 自動実行の設定

スプレッドシートのメニューから:
```
🤖 議題生成システム → ⏰ 自動実行タイマー設定
```

これにより以下が自動実行されます:
- 5分ごと: Slackメッセージ同期
- 1時間ごと: AI分析
- 毎日9時・15時: 通知送信

## 使い方

### 手動実行

メニューから各機能を手動実行可能:
- 🔄 手動同期（Slack→シート）
- 🤖 AI分析実行
- 📢 通知テスト
- 📝 選択行でドラフト生成

### データ管理

#### Messagesシート
- Slackメッセージと分析結果を管理
- `human_judgement`列で「必要/不要/保留」を選択
- `match_flag`がtrueの行が通知対象

#### Categoriesシート
- 判定カテゴリと基準を定義
- キーワードや判定基準をカスタマイズ可能

#### Configシート
- システム全体の設定を管理
- しきい値やモデル名を調整可能

## トラブルシューティング

### Slack連携エラー

```
Error: invalid_auth
```
→ Bot Tokenを再確認、チャンネルにBotを追加

### OpenAI APIエラー

```
Error: Insufficient quota
```
→ API利用上限を確認、必要に応じてプラン変更

### 通知が届かない

1. Configシートの通知設定を確認
2. Logsシートでエラーを確認
3. トリガーが正しく設定されているか確認

## セキュリティ注意事項

- API KeyやTokenは必ずPropertiesServiceに保存
- コードに直接記載しない
- スプレッドシートの共有範囲を適切に設定
- 機密情報を含むチャンネルは慎重に選択

## カスタマイズ

### AIモデルの変更

Configシートで変更可能:
- `openaiModel`: 要約・分類用（デフォルト: gpt-4o-mini）
- `openaiModelDraft`: ドラフト生成用（デフォルト: gpt-4o）

### 判定しきい値の調整

Configシートの`classificationThreshold`を調整（0.0-1.0）

### プロンプトのカスタマイズ

`Schemas.gs`のPROMPTSオブジェクトを編集

## ライセンス

このプロジェクトは社内利用を前提としています。

## サポート

問題が発生した場合は、Logsシートを確認し、エラーメッセージを元に対処してください。