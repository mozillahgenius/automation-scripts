# Notion ワークスペース変更監視・ユーザー活動分析システム
> Notion APIを使用してワークスペースの変更を定期監視し、ユーザー別の活動分析と日次/週次/月次集計を自動生成するシステム

## 概要
Notion APIと連携してデータベース内のページ変更を定期的に監視し、変更履歴をログとして蓄積するGoogle Apps Scriptシステム。編集日時・編集者・タイトルの変更を検出し、ユーザー別の活動統計（今日・今週・今月・累計の編集回数、平均編集間隔）を自動集計する。活動レポートの生成機能も備え、チームのNotionワークスペース利用状況を可視化する。

## 主な機能
- Notionデータベースの全ページ一括取得（ページネーション対応）
- ページ変更の自動検出（編集日時変更、編集者変更、タイトル変更）
- 変更ログの自動蓄積（変更検出時 + 4時間ごとの定期記録）
- ユーザー別活動統計の自動計算（今日・今週・今月・累計・平均編集間隔）
- ユーザー活動履歴のタイムライン記録
- 活動レベルに応じた色分け表示
- 日付入りのユーザー活動レポート生成
- 4時間ごと / 毎日午前9時の定期実行トリガー

## アーキテクチャ
`scheduledUpdate`関数がトリガーにより定期実行され、PageListシートに登録された各ページのNotion APIから最新情報を取得する。前回の状態と比較して変更（編集日時・編集者・タイトル）を検出し、NotionLogシートに変更ログを記録する。同時にUserActivityシートにユーザー活動履歴を追記し、UserAnalyticsシートでユーザー別の集計統計を更新する。ユーザー情報はNotion Users APIから取得し、ユーザーIDを名前に解決する。

## ファイル構成
| ファイル | 説明 |
|---|---|
| Code.js | 全機能のメインロジック（Notion API連携、変更検出、ユーザー分析） |
| appsscript.json | GASプロジェクト設定（タイムゾーン: Asia/Singapore、V8ランタイム） |

## シート構成

### PageList（ページリスト）
| 列 | ヘッダー | 内容 |
|---|---|---|
| A | Page ID | NotionページID |
| B | Page Title | ページタイトル |
| C | 最終編集日時 | ISO 8601形式の最終編集日時 |
| D | 最終編集者 | 編集者のユーザー名 |

### NotionLog（変更ログ）
| 列 | ヘッダー | 内容 |
|---|---|---|
| A | 記録日時 | ログ記録のタイムスタンプ |
| B | Page ID | 対象ページID |
| C | Page Title | ページタイトル |
| D | 最終編集日時 | 検出時の編集日時 |
| E | 最終編集者 | 検出時の編集者 |
| F | 変更検出 | 変更内容（例: 「編集日時変更 + タイトル変更」）または「定期記録」 |

### UserAnalytics（ユーザー分析）
| 列 | ヘッダー | 内容 |
|---|---|---|
| A | ユーザー名 | Notionユーザー名 |
| B | 編集ページ数 | ユニークな編集ページ数 |
| C | 最終活動日 | 最新の活動日時 |
| D | 今日の編集数 | 当日の編集回数 |
| E | 今週の編集数 | 過去7日間の編集回数 |
| F | 今月の編集数 | 過去30日間の編集回数 |
| G | 総編集回数 | 累計の編集回数 |
| H | 平均編集間隔(日) | 編集間の平均日数 |

### UserActivity（ユーザー活動履歴）
| 列 | ヘッダー | 内容 |
|---|---|---|
| A | 記録日時 | 活動検出のタイムスタンプ |
| B | ユーザー名 | 活動したユーザー |
| C | Page Title | 対象ページタイトル |
| D | 活動種別 | 「ページ編集」または「タイトル変更」 |
| E | Page ID | 対象ページID |

## 主要関数
| 関数名 | 説明 |
|---|---|
| `setupNotionMonitoring()` | 初期セットアップ。全4シートを作成し、ヘッダーを設定 |
| `getAllPagesFromDatabase()` | 指定データベースから全ページを取得し、PageListシートに書き込み |
| `updateAndLogNotionPages()` | PageListの全ページの最新情報を取得し、変更を検出してログに記録 |
| `updateUserAnalytics()` | PageListとUserActivityからユーザー別統計を集計し、UserAnalyticsシートを更新 |
| `generateUserActivityReport()` | 日付入りの活動レポートシートを新規作成（サマリー統計付き） |
| `scheduledUpdate()` | 定期実行用エントリポイント。`updateAndLogNotionPages()`を呼び出す |
| `setupPeriodicTrigger()` | 4時間ごとの定期実行トリガーを設定 |
| `setupDailyTrigger()` | 毎日午前9時の実行トリガーを設定 |
| `fetchAllPagesFromDatabase(databaseId)` | Notion Database Query APIでページネーション付き全ページ取得 |
| `fetchNotionPage(pageId)` | Notion Pages APIで個別ページ情報を取得 |
| `fetchNotionUser(userId)` | Notion Users APIでユーザー情報を取得 |
| `extractPageTitle(page)` | ページオブジェクトからタイトルを抽出（titleプロパティを検索） |
| `isPeriodicLog()` | 現在時刻が4時間ごとの記録タイミングか判定（0, 4, 8, 12, 16, 20時） |

## 変更検出ロジック

以下の3種類の変更を検出する:

| 変更種別 | 比較対象 | 活動種別 |
|---|---|---|
| 編集日時変更 | `last_edited_time`の前回値との比較 | ページ編集 |
| 編集者変更 | `last_edited_by`のユーザー名比較 | ページ編集 |
| タイトル変更 | ページタイトルの前回値との比較 | タイトル変更 |

## ユーザー分析の色分け

### 今日の編集数
| 条件 | 色 |
|---|---|
| 3回以上 | `#d4edda`（緑） |
| 1回以上 | `#fff3cd`（黄） |
| 0回 | 白 |

### 総編集回数（最大値に対する比率）
| 比率 | 色 |
|---|---|
| 80%以上 | `#d4edda`（緑） |
| 50%以上 | `#fff3cd`（黄） |
| 20%以上 | `#f8d7da`（薄赤） |
| 20%未満 | 白 |

## 使用サービス・API
- Notion API v2022-06-28
  - Database Query API（`/v1/databases/{id}/query`）
  - Pages API（`/v1/pages/{id}`）
  - Users API（`/v1/users/{id}`）
- Google Sheets API（SpreadsheetApp）
- Google Apps Script トリガー（ScriptApp）
- UrlFetchApp（HTTP リクエスト）

## セットアップ手順
1. Google Spreadsheetsで新規スプレッドシートを作成
2. Apps Scriptエディタを開き、Code.jsの内容を貼り付け
3. スクリプトプロパティに `NOTION_API_TOKEN` を設定
4. Code.js内の `DATABASE_ID` を監視したいNotionデータベースのIDに変更
5. Notionインテグレーションに対象データベースへのアクセス権を付与
6. `setupNotionMonitoring()` を実行して全シートを初期化
7. `getAllPagesFromDatabase()` を実行してページ一覧を取得
8. `setupPeriodicTrigger()` または `setupDailyTrigger()` でトリガーを設定

## スクリプトプロパティ
| プロパティ名 | 説明 |
|---|---|
| `NOTION_API_TOKEN` | Notionインテグレーションのトークン（`secret_`で始まる） |

## コード内定数
| 定数名 | 値 | 説明 |
|---|---|---|
| `NOTION_VERSION` | `'2022-06-28'` | Notion APIバージョン |
| `LOG_SHEET_NAME` | `'NotionLog'` | ログシート名 |
| `PAGE_LIST_SHEET_NAME` | `'PageList'` | ページリストシート名 |
| `USER_ANALYTICS_SHEET_NAME` | `'UserAnalytics'` | ユーザー分析シート名 |
| `USER_ACTIVITY_SHEET_NAME` | `'UserActivity'` | ユーザー活動履歴シート名 |
| `DATABASE_ID` | （要変更） | 監視対象のNotionデータベースID |

## 注意事項
- Notionインテグレーションに対象データベースへの読み取りアクセス権が必要
- ページネーション対応により、100件以上のページを持つデータベースにも対応
- ユーザー情報の取得に失敗した場合は「Unknown User」と表示される
- 定期ログは4時間ごと（0, 4, 8, 12, 16, 20時）に記録される
- 活動レポート生成時、同日に再実行すると同名シートが上書きされる
- UserAnalyticsシートは更新のたびに全データがクリアされ再生成される
