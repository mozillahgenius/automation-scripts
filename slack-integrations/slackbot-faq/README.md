# slackbot-faq
> Google Apps Script で動作するキーワードマッチ型 FAQ Slack Bot

## 概要
Slack のメッセージイベントを受信し、スプレッドシートに登録されたキーワードと照合して自動回答を返す FAQ Bot です。Google Apps Script の Web アプリとしてデプロイし、Slack Event Subscriptions の Request URL として設定して使用します。

## 主な機能
- Slack メッセージ内のキーワードをスプレッドシートの FAQ データと照合
- 一致するキーワードが見つかった場合、対応する回答を同チャンネルに自動投稿
- Slack API の Challenge リクエストへの自動応答（URL Verification 対応）
- 3秒超過リトライの重複処理防止（CacheService による `client_msg_id` のデデュプリケーション）

## アーキテクチャ
```
Slack (Event Subscriptions)
  |
  v  POST リクエスト
Google Apps Script (doPost)
  |
  +-- Challenge 応答（初回URL検証時）
  +-- CacheService でリトライ重複排除（600秒キャッシュ）
  +-- SpreadsheetApp で FAQ シート（Sheet1）を検索
  +-- キーワード一致時、Slack chat.postMessage API で回答を投稿
```

1. Slack がメッセージイベントを `doPost` に POST 送信
2. `client_msg_id` を CacheService に記録し、3秒タイムアウトによるリトライを防止
3. スプレッドシートの Sheet1 から全行を取得し、A列（キーワード）がメッセージに含まれるか `includes()` で判定
4. 最初にマッチした行の B列（回答）を Slack `chat.postMessage` API で同チャンネルに返信

## ファイル構成
| ファイル | 説明 |
|---|---|
| `Code.js` | メインロジック（doPost ハンドラ、FAQ 検索、Slack 応答） |
| `appsscript.json` | GAS プロジェクト設定（タイムゾーン: Asia/Singapore、V8 ランタイム、Web アプリ設定） |

## 主要関数
| 関数名 | 説明 |
|---|---|
| `doPost(e)` | Slack からの POST リクエストを処理するメインハンドラ。Challenge 応答、リトライ排除、FAQ 検索、Slack 返信を行う |
| `testrun()` | テスト実行用関数。ダミーの Slack イベントデータを作成して `doPost` を呼び出す |

## 使用サービス・API
- **Google Sheets API** (SpreadsheetApp) - FAQ データの読み取り
- **Slack chat.postMessage API** - 回答メッセージの投稿
- **Google Cache Service** (CacheService) - リトライ重複排除
- **Google Content Service** (ContentService) - Challenge レスポンス生成
- **URL Fetch Service** (UrlFetchApp) - Slack API への HTTP リクエスト

## セットアップ手順
1. Google Apps Script プロジェクトを作成し、`Code.js` をコピー
2. スクリプトプロパティに `SLACK_BOT_TOKEN` を設定（Slack Bot User OAuth Token）
3. 紐づけるスプレッドシートに `Sheet1` を作成し、A列にキーワード、B列に回答を入力（1行目はヘッダー）
4. Web アプリとしてデプロイ（実行者: デプロイしたユーザー、アクセス: 全員（匿名含む））
5. デプロイ URL を Slack App の Event Subscriptions → Request URL に設定
6. Slack App の Event Subscriptions で `message.channels`（または `message.im`）イベントを購読

## スクリプトプロパティ
| プロパティ名 | 説明 |
|---|---|
| `SLACK_BOT_TOKEN` | Slack Bot User OAuth Access Token（`xoxb-` で始まるトークン） |

## スプレッドシート構成
| シート名 | 列構成 | 説明 |
|---|---|---|
| `Sheet1` | A列: キーワード, B列: 回答 | FAQ データ（1行目はヘッダー行、2行目以降がデータ） |

## Web アプリ設定（appsscript.json）
| 設定項目 | 値 |
|---|---|
| `executeAs` | `USER_DEPLOYING` |
| `access` | `ANYONE_ANONYMOUS` |
| `runtimeVersion` | `V8` |
| `timeZone` | `Asia/Singapore` |

## 注意事項
- キーワードは部分一致（`text.includes(keyword)`）で判定されるため、短すぎるキーワードは意図しないマッチが発生する可能性がある
- 複数のキーワードがマッチする場合、スプレッドシートの上から順に検索し、最初にマッチした回答のみ返す
- リトライ排除キャッシュの有効期限は 600 秒（10分）
