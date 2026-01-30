# drive-upload-monitor
> 共有ドライブへのファイルアップロードを監視し、Slack に自動通知＋定期リマインダーを送信するシステム

## 概要
Google 共有ドライブ（ShareDrive）に新しくアップロードされたファイルを定期的に検知し、Slack チャンネルに通知するモニタリングシステムです。さらに、昼（12時）と夕方（17時）にアップロード催促のリマインダーメッセージを自動送信し、チームのファイル管理を促進します。管理用スプレッドシートで複数の共有ドライブの監視設定やファイル一覧を一元管理します。

## 主な機能
- 複数の共有ドライブを対象としたファイルアップロード検知（30分間隔）
- 新規ファイル検出時の Slack 通知（ファイル名、URL、アップロード日時を含む）
- 昼・夕方の定期リマインダーメッセージ送信（ドライブ統計情報付き）
- 監視対象ドライブの有効/無効切り替え（スプレッドシートのチェックボックスで管理）
- 全ドライブのファイル一覧をスプレッドシートに自動記録・更新
- 土日スキップ、簡易祝日スキップ対応
- 夕方のリマインダーには当日アップロードファイル数を表示

## アーキテクチャ
```
[時間トリガー: 30分間隔]
  |
  v
checkAndNotify()
  +-- getActiveDrives(): スプレッドシート「設定」シートから有効なドライブ一覧取得
  +-- checkNewFilesInDrive(): Drive API v2 で過去1時間以内の新規ファイルを検索
  +-- notifyToSlack(): Slack Webhook で新規ファイル情報を通知
  +-- updateAllFileLists(): 全ドライブの全ファイルを「ファイルリスト」シートに記録

[時間トリガー: 毎日12時・17時]
  |
  v
sendReminderMessage()
  +-- 曜日・祝日チェック（土日/祝日はスキップ可能）
  +-- getDriveStatistics(): 監視中ドライブの統計情報取得
  +-- customizeReminderMessage(): 統計情報を含むメッセージ生成
  +-- sendSlackMessage(): Slack Webhook で催促メッセージ送信
```

## ファイル構成
| ファイル | 説明 |
|---|---|
| `Code.js` | メインロジック（ファイル監視、Slack 通知、リマインダー、スプレッドシート管理） |
| `appsscript.json` | GAS プロジェクト設定（タイムゾーン: Asia/Singapore、V8 ランタイム、Drive API v2 有効） |

## 主要関数
| 関数名 | 説明 |
|---|---|
| `setup()` | 初期設定。スプレッドシート初期化、ファイル監視トリガー（30分間隔）とリマインダートリガー（12時・17時）を作成 |
| `checkAndNotify()` | 定期実行メイン処理。有効な共有ドライブをチェックし、新規ファイルがあれば Slack に通知。全ファイルリストも更新 |
| `checkNewFilesInDrive(drive)` | 特定の共有ドライブで過去1時間以内に作成されたファイルを Drive API v2 で検索し、未記録のファイルを返す |
| `notifyToSlack(driveFilesList)` | 新規ファイル情報を Slack Webhook 経由でリッチメッセージとして通知 |
| `updateAllFileLists(activeDrives)` | 全ドライブの全ファイル情報をスプレッドシートのファイルリストシートに書き込み |
| `getAllFilesInDrive(drive)` | Drive API v2 でページネーション付きの全ファイル取得 |
| `sendReminderMessage()` | 時間帯に応じた催促メッセージ送信。土日・祝日スキップ、統計情報付加、当日ファイル数表示 |
| `customizeReminderMessage(template, driveInfo, messageType)` | テンプレートメッセージにドライブ統計情報と当日アップロード数を付加 |
| `getDriveStatistics(activeDrives)` | ファイルリストシートから総ファイル数、ドライブ数、最終更新日を集計 |
| `getTodayUploadedFiles()` | ファイルリストシートから当日アップロードされたファイルを抽出 |
| `isHoliday(date)` | 日本の固定祝日（元日、建国記念の日など10日間）を簡易判定 |
| `sendSlackMessage(messageData, channel)` | Slack Webhook への共通メッセージ送信関数 |
| `initializeSpreadsheet()` | 「設定」シートと「ファイルリスト」シートをヘッダー行付きで作成 |
| `getActiveDrives()` | 設定シートからチェックボックスが TRUE のドライブ一覧を取得 |
| `setupReminderTriggers()` | 既存の催促トリガーを削除し、12時と17時のトリガーを再作成 |
| `sendTestLunchReminder()` | 手動テスト用: 昼の催促メッセージ送信 |
| `sendTestEveningReminder()` | 手動テスト用: 夕方の催促メッセージ送信 |
| `checkReminderSettings()` | 手動実行用: 催促メッセージの現在設定を確認 |
| `resetReminderTriggers()` | 手動実行用: 催促トリガーのリセット |

## 使用サービス・API
- **Google Drive API v2** (Advanced Service) - 共有ドライブ内のファイル検索・一覧取得
- **Google Sheets API** (SpreadsheetApp) - 設定管理、ファイルリスト記録
- **Slack Incoming Webhooks** - ファイル通知・催促メッセージ送信
- **URL Fetch Service** (UrlFetchApp) - Slack Webhook への HTTP リクエスト
- **Script Triggers** (ScriptApp) - 定期実行トリガーの管理

## セットアップ手順
1. Google Apps Script プロジェクトを作成し、`Code.js` をコピー
2. Apps Script エディタで「サービス」から **Drive API (v2)** を追加
3. スクリプトプロパティに `SLACK_WEBHOOK_URL` を設定
4. `CONFIG.SPREADSHEET_ID` を管理用スプレッドシートの ID に変更
5. `CONFIG.SLACK_CHANNEL` を通知先の Slack チャンネル名に変更
6. `setup()` 関数を実行してスプレッドシートの初期化とトリガー設定を行う
7. スプレッドシートの「設定」シートで監視対象の共有ドライブ名と ID を登録し、チェックボックスを有効にする

## スクリプトプロパティ
| プロパティ名 | 説明 |
|---|---|
| `SLACK_WEBHOOK_URL` | Slack Incoming Webhook の URL |

## 設定定数（CONFIG オブジェクト）
| 定数名 | デフォルト値 | 説明 |
|---|---|---|
| `SPREADSHEET_ID` | (ハードコード) | 管理用スプレッドシートの ID |
| `SHEETS.SETTINGS` | `設定` | ドライブ設定シート名 |
| `SHEETS.FILE_LIST` | `ファイルリスト` | ファイル一覧シート名 |
| `SLACK_CHANNEL` | `#clientwork-document` | Slack 通知先チャンネル |
| `CHECK_INTERVAL_MINUTES` | `30` | ファイル監視の実行間隔（分） |
| `REMINDER_SETTINGS.LUNCH_TIME` | `12` | 昼リマインダー送信時刻（24時間制） |
| `REMINDER_SETTINGS.EVENING_TIME` | `17` | 夕方リマインダー送信時刻（24時間制） |
| `REMINDER_SETTINGS.REMINDER_CHANNEL` | `#clientwork-document` | リマインダー送信チャンネル |
| `REMINDER_SETTINGS.SKIP_WEEKENDS` | `true` | 土日のリマインダーをスキップ |
| `REMINDER_SETTINGS.SKIP_HOLIDAYS` | `false` | 祝日のリマインダーをスキップ |

## スプレッドシート構成

### 設定シート
| 列 | ヘッダー | 説明 |
|---|---|---|
| A | 有効/無効 | チェックボックス（TRUE/FALSE） |
| B | 共有ドライブ名 | 表示名 |
| C | 共有ドライブID | Google ドライブの共有ドライブ ID |
| D | 最終チェック日時 | 自動更新 |
| E | ステータス | 正常 / エラーメッセージ |

### ファイルリストシート
| 列 | ヘッダー | 説明 |
|---|---|---|
| A | 共有ドライブ名 | ファイルが所属する共有ドライブ名 |
| B | ファイル名 | ファイル名 |
| C | URL | Google Drive 上の URL |
| D | アップロード日 | ファイル作成日時 |
| E | ファイルID | Google Drive のファイル ID |
| F | 最終更新日 | ファイルの最終更新日時 |

## トリガー構成
| トリガー | 対象関数 | 間隔 | 説明 |
|---|---|---|---|
| 時間ベース | `checkAndNotify` | 30分ごと | ファイルアップロード監視 |
| 時間ベース | `sendReminderMessage` | 毎日12時 | 昼の催促メッセージ |
| 時間ベース | `sendReminderMessage` | 毎日17時 | 夕方の催促メッセージ |
