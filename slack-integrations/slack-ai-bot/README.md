# Slack AI Bot - ファイル読み取り＆ドキュメント修正対応版

## 概要
このSlack BotはChatGPT/Gemini APIと連携し、メンションに応答するだけでなく、添付されたファイルの内容を読み取って処理することができます。
さらに、Word文書やGoogleドキュメントに対して自動的に修正案やコメントを記載し、編集可能なリンクを共有する機能も備えています。

### 主要機能
- **AI対話**: ChatGPT/Geminiを使った自然な会話
- **ファイル処理**: PDF、Word、Googleドキュメントの読み取りと分析
- **ドキュメントレビュー**: AIによる文書の自動修正・コメント付与
- **FAQ自動応答**: Google Natural Language APIを活用した高度な質問マッチング
- **スプレッドシート連携**: FAQデータと自然言語処理の統合
- **Drive検索**: 指定フォルダ内の全文検索と結果の自動抽出

## 対応ファイル形式
- **PDF** - OCR対応で日本語も読み取り可能
- **Word文書** (.docx, .doc)  
- **Googleドキュメント** - URLまたは共有リンクから
- **テキストファイル** (.txt)

## 必要な設定

### 1. Google Apps Script側の設定

#### 必要なAPIの有効化
1. GASエディタで「サービス」→「+」をクリック
2. 以下のAPIを有効化:
   - Drive API (v2)
   - Google Docs API

#### スクリプトプロパティ
以下のプロパティを設定してください：
- `SLACK_TOKEN`: Slack Bot Token
- `OPEN_AI_TOKEN`: OpenAI API Key
- `GEMINI_TOKEN`: Gemini API Key (オプション)
- `GOOGLE_NL_API`: Google Natural Language API Key (FAQ機能の高度化に必要)
- `SPREADSHEET_ID`: 自動生成されるスプレッドシートID（初期化時に自動設定）

### 2. Slack App側の設定

#### OAuth & Permissions
以下のスコープを追加してください：

**Bot Token Scopes:**
- `app_mentions:read`
- `channels:history`
- `channels:read`
- `chat:write`
- `chat:write.public`
- `files:read` ← 新規追加
- `files:write` ← 新規追加
- `groups:read`
- `reactions:read`
- `remote_files:read` ← 新規追加
- `remote_files:share` ← 新規追加
- `users:read`

#### Event Subscriptions
以下のBot Eventsを追加：
- `message.channels`
- `file_shared` ← 新規追加
- `file_change` ← 新規追加
- `file_deleted` ← 新規追加

## 使い方

### 基本的な使い方
1. Slackチャンネルにボットを追加
2. `@ボット名` でメンション
3. ファイルを添付して質問

### FAQ機能の使い方

#### 自然言語処理を使った高度なマッチング
Google Natural Language APIを使用して、質問文から名詞・数値を抽出し、FAQシートと高精度でマッチングします。

```
@ChatGPT 休暇申請の方法は？
→ 「休暇」「申請」「方法」を自動抽出してFAQを検索
```

#### FAQシートの設定
1. スプレッドシートの「faq」シートを開く
2. 以下の列に情報を入力：
   - A列: キーワード
   - B列: 回答
   - C列: Drive検索フラグ（TRUE/FALSE）
   - D列: Drive検索結果（自動入力）

### ファイル処理の例

#### 通常の読み取り
```
@ChatGPT この文書の内容を要約してください
[PDFファイルを添付]
```

```
@ChatGPT このWordファイルの重要なポイントを教えてください
[Word文書を添付]
```

#### ドキュメント修正・レビュー機能
```
@ChatGPT このドキュメントの修正をお願いします
[Word文書またはGoogleドキュメントを添付]
```

```
@ChatGPT この文書をレビューして改善点を指摘してください
[Word文書またはGoogleドキュメントを添付]
```

```
@ChatGPT この提案書の校正とコメントをお願いします
[Word文書またはGoogleドキュメントを添付]
```

### ドキュメント修正機能の詳細

キーワード「修正」「レビュー」「校正」「チェック」「改善」「コメント」を含むメッセージと共にWord/Googleドキュメントを添付すると：

1. **自動的にドキュメントをコピー**
2. **AIが修正案やコメントを直接記載**
   - 赤字取り消し線: 修正が必要な箇所
   - 青字: 修正案
   - 黄色ハイライト: コメント箇所
   - 緑字: 追加提案
3. **共有設定を元のドキュメントから継承**
   - 元のドキュメントと同じメンバーがアクセス可能
   - 元のドキュメントと同じ権限設定を維持
4. **編集可能なリンクを返信**

## ファイル処理の仕組み

1. **ファイル検出**: Slackイベントからファイル情報を取得
2. **ダウンロード**: Slack APIを使用してファイルをダウンロード
3. **変換処理**:
   - PDF/Word: Google Drive APIで一時的にGoogle Docsに変換
   - Google Docs: Document IDを抽出して直接アクセス
4. **テキスト抽出**: DocumentApp.getBody().getText()で内容取得
5. **AI処理**: 抽出したテキストをコンテキストに含めてAIに送信
6. **クリーンアップ**: 一時ファイルを削除

## 自然言語処理の仕組み

### Google Natural Language APIの活用
1. **形態素解析**: `gNL()`関数で質問文を解析
2. **品詞抽出**: `filterGNL()`で名詞・数値を抽出
3. **文字列正規化**:
   - カタカナ→ひらがな変換
   - 全角→半角変換
   - 大文字→小文字変換
4. **FAQマッチング**: 正規化された単語でFAQシートを検索

### スプレッドシート分析機能
`analyzeSpreadsheetWithNLP()`関数により：
- 全シートからテキストを抽出
- 最大5000文字まで自然言語処理
- 名詞・数値の抽出
- 感情分析（ポジティブ/ネガティブスコア）

## 制限事項

- ファイルサイズ: Slackの制限に準拠
- 処理時間: 大きなファイルは処理に時間がかかる場合があります
- アクセス権限: ボットがアクセス可能なファイルのみ処理可能
- OCR精度: PDFの画像テキストは完全でない場合があります

## トラブルシューティング

### ファイルが読み取れない場合
1. Slack Appの権限を確認
2. GASのDrive APIが有効になっているか確認
3. ファイルのアクセス権限を確認

### FAQ機能が動作しない場合
1. Google Natural Language APIキーが設定されているか確認
2. APIが有効になっているか確認（Google Cloud Console）
3. スプレッドシートの「faq」シートにデータが入力されているか確認
4. `testNLPAndSpreadsheet()`関数でテストを実行

### エラーメッセージ
- "ファイル処理中にエラーが発生しました": 権限不足またはファイル形式の問題
- "ファイルタイプはサポートされていません": 対応していない形式

## セキュリティ注意事項

- ファイル内容は一時的にGoogle Driveに保存されます
- 処理後は自動的に削除されますが、ゴミ箱に30日間残ります
- 機密情報を含むファイルの処理には注意してください
- Google Natural Language APIに送信されるテキストは5000文字に制限されます
- スプレッドシートのデータはボットのサービスアカウントがアクセス可能な範囲に限定されます

## 関連ファイル

- `slack-bot-all-in-one.gs`: メインコード（統合版）
- `slack-ai-bot-integrated-original.gs`: 参考実装
- `SETUP.md`: 詳細なセットアップ手順
- `appsscript.json`: GASマニフェストファイル