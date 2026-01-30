# Gmail未処理メール管理システム
> Gmailの未処理メールを優先度判定・経過日数の色分けでレポートし、メール・Slackで通知する自動管理システム

## 概要
Gmail受信トレイの未アーカイブメールを定期的にスキャンし、優先度判定と経過日数に基づいて色分けされたレポートをGoogle Spreadsheetsに自動生成するシステム。高優先度メールや長期間放置されたメールを視覚的に識別し、メール通知およびSlack Webhook通知で対応を促す。

## 主な機能
- Gmail受信トレイの未アーカイブメール自動スキャン（最大50件）
- キーワードベースの3段階優先度判定（高・中・低）
- 経過日数に応じた5段階カラーコーディング（0日 - 14日以上）
- 送信者フィルタリング（noreply等の自動除外）
- 優先度・経過日数でのソート表示
- メール通知（上位10件の詳細 + 概要統計）
- Slack Webhook通知（上位5件 + 概要統計、色分けアタッチメント付き）
- エラー発生時の管理者通知
- 10時-20時の1時間おき自動実行トリガー

## アーキテクチャ
`reportUnarchivedEmails`がメイン処理として、Gmail検索（`in:inbox -is:draft`）でスレッドを取得し、除外送信者をフィルタリング後、各メールの優先度を判定する。優先度は件名・送信者に含まれるキーワード（緊急、至急、重要等）で判定され、結果は優先度降順・経過日数降順にソートされる。スプレッドシートに書き込み後、条件付き書式でカラーコーディングを適用し、メールとSlackへ通知を送信する。

## ファイル構成
| ファイル | 説明 |
|---|---|
| Code.js | 全機能のメインロジック（メール処理、通知、フォーマット） |
| appsscript.json | GASプロジェクト設定（タイムゾーン: Asia/Singapore、V8ランタイム） |

## スプレッドシート列構成（Reportシート）

| 列 | ヘッダー | 内容 |
|---|---|---|
| A | 送信者 | メール送信者 |
| B | 件名 | メールの件名 |
| C | 受信日時 | 最終メッセージの日時 |
| D | メールリンク | Gmailへの直リンク |
| E | 優先度 | 高/中/低 |
| F | 日数経過 | 受信からの経過日数 |

## 主要関数
| 関数名 | 説明 |
|---|---|
| `reportUnarchivedEmails()` | メイン処理。未アーカイブメールのスキャン、レポート生成、通知送信を実行 |
| `prepareSpreadsheet()` | Reportシートの初期化（存在しなければ作成、ヘッダー設定、既存データクリア） |
| `processEmailThreads(threads)` | メールスレッドを処理し、除外フィルタ・優先度判定・ソートを行う |
| `shouldExcludeSender(sender)` | 送信者が除外リストに該当するか判定 |
| `determinePriority(subject, sender)` | 件名・送信者から優先度（高/中/低）を判定 |
| `writeToSpreadsheet(sheet, emailData)` | スプレッドシートへのデータ書き込みとフォーマット適用 |
| `formatSpreadsheet(sheet, dataLength)` | 条件付き書式の設定（優先度カラー、経過日数カラー、クリティカル行ハイライト） |
| `sendNotifications(emailData)` | メール・Slack通知のディスパッチ |
| `sendEmailNotification(...)` | メール通知の送信（概要 + 上位10件の詳細リスト） |
| `sendSlackNotification(...)` | Slack Webhook通知の送信（アタッチメント付き） |
| `sendErrorNotification(error)` | エラー発生時の管理者メール通知 |
| `createEmailDetailsList(emailData)` | メール通知用の詳細リスト文字列を生成（最大10件） |
| `getUrgencyIcon(daysPassed, priority)` | 優先度と経過日数に応じた緊急度ラベルを返す |
| `extractSenderName(sender)` | 送信者文字列から名前部分を抽出 |
| `setupHourlyTriggers()` | 10時-20時の1時間おきトリガーを一括設定 |
| `checkTriggers()` | 設定済みトリガーの確認 |
| `testConfiguration()` | スプレッドシートアクセスとGmail APIの動作テスト |

## 経過日数カラーコーディング

| 経過日数 | 色 | 背景色コード |
|---|---|---|
| 0日（当日） | 薄い緑 | `#ccffcc` |
| 1-2日 | 黄色 | `#ffcc00` |
| 3-6日 | オレンジ | `#ff9900` |
| 7-13日 | 赤 | `#ff4444` |
| 14日以上 | 濃い赤 | `#cc0000` |

## 優先度カラーコーディング

| 優先度 | 背景色 |
|---|---|
| 高 | `#ff9999`（赤） |
| 中 | `#ffff99`（黄） |
| 低 | なし |

## クリティカル行ハイライト
優先度「高」かつ7日以上経過した行は全体が `#ffcccc`（薄い赤）で強調表示される。

## 優先度判定ロジック

| 条件 | 判定結果 |
|---|---|
| 件名・送信者に「緊急」「至急」「重要」「urgent」「important」を含む | 高 |
| 送信者に `@important-client.com` を含む | 高 |
| 件名・送信者に「確認」「承認」を含む | 中 |
| 上記以外 | 低 |

## 使用サービス・API
- Gmail API（GmailApp）- メールスレッド検索・メール送信
- Google Sheets API（SpreadsheetApp）- レポート書き込み・フォーマット
- Slack Incoming Webhook（UrlFetchApp）- Slack通知
- Google Apps Script トリガー（ScriptApp）- 定期実行

## セットアップ手順
1. Google Spreadsheetsで新規スプレッドシートを作成
2. Apps Scriptエディタを開き、Code.jsの内容を貼り付け
3. `CONFIG`オブジェクトの以下を環境に合わせて変更:
   - `EMAIL_RECIPIENT`: 通知先メールアドレス
   - `SLACK_WEBHOOK_URL`: Slack Webhook URL（任意、空文字でスキップ）
   - `EXCLUDE_SENDERS`: 除外したい送信者パターン
   - `PRIORITY_KEYWORDS`: 高優先度と判定するキーワード
   - `MAX_EMAILS_TO_PROCESS`: 処理する最大メール数
4. `setupHourlyTriggers()`を手動実行してトリガーを設定
5. `testConfiguration()`を実行して動作確認

## 設定項目（CONFIGオブジェクト）

| プロパティ名 | デフォルト値 | 説明 |
|---|---|---|
| `SHEET_NAME` | `'Report'` | レポート出力先シート名 |
| `EMAIL_RECIPIENT` | `'admin@example-company.test'` | 通知先メールアドレス |
| `SLACK_WEBHOOK_URL` | `''`（空文字） | Slack Incoming Webhook URL。空の場合Slack通知はスキップ |
| `EXCLUDE_SENDERS` | `['noreply@', 'no-reply@', 'newsletter@']` | 除外する送信者パターンの配列 |
| `PRIORITY_KEYWORDS` | `['緊急', '至急', '重要', 'urgent', 'important']` | 高優先度と判定するキーワードの配列 |
| `MAX_EMAILS_TO_PROCESS` | `50` | 1回の実行で処理する最大メール数 |

## 注意事項
- Slack通知には色分けアタッチメントが使用される（高優先度あり: danger赤、3日以上経過あり: warning黄、それ以外: good緑）
- メール通知は最大10件まで詳細を表示し、残りは件数のみ表示
- 毎回実行時にReportシートのデータはクリアされ、最新状態で再生成される
