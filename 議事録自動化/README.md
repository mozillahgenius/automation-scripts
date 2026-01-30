# Google Meet 議事録自動化システム

Google Apps Script (GAS) を使用して、Google Meet の会議録画・トランスクリプトから自動で議事録を生成し、メール送信するシステムです。

## 🎯 主な機能

- ✅ **自動取得**: Google Calendar の会議イベントから Meet トランスクリプトを自動取得
- ✅ **AI生成**: Vertex AI Gemini API を使用して構造化された議事録を生成
- ✅ **自動配信**: 会議参加者と指定した追加宛先に議事録をメール送信
- ✅ **設定管理**: Google Sheets で設定やログを一元管理
- ✅ **自動化**: 時間駆動トリガーで完全自動運用

## 📋 前提条件

- Google Workspace Business Standard 以上
- Google Cloud Platform アカウント
- Google Meet の録画・トランスクリプト機能が有効

## 🚀 クイックスタート

### 1. ファイル構成

```
議事録自動化/
├── Code.gs                    # メインスクリプト
├── Utils.gs                   # ユーティリティ関数集
├── appsscript.json           # マニフェストファイル
├── セットアップ手順書.md        # 詳細なセットアップ手順
└── README.md                 # このファイル
```

### 2. 基本セットアップ

1. **GCP プロジェクト設定**
   - Vertex AI API、Google Meet API、Calendar API、Gmail API を有効化
   - IAM で適切な権限を設定

2. **Apps Script プロジェクト作成**
   - script.google.com で新規プロジェクト作成
   - 提供されたファイルをアップロード

3. **設定スプレッドシート作成**
   - Google Sheets で設定・ログ管理用スプレッドシートを作成

4. **権限設定**
   - 必要なOAuthスコープを許可
   - Admin Console で Meet API の権限を設定

詳細は [セットアップ手順書.md](./セットアップ手順書.md) を参照してください。

## ⚙️ 設定項目

設定スプレッドシートの「設定」シートで以下を管理：

| 設定項目 | 説明 | 例 |
|---------|------|-----|
| `CalendarId` | 監視対象のカレンダーID | `calendar@group.calendar.google.com` |
| `AdditionalRecipients` | 追加送信先（カンマ区切り） | `manager@company.com, team@company.com` |
| `CheckHours` | 何時間前までの会議を対象にするか | `1` |
| `MinutesPrompt` | Gemini API に送信するプロンプト | 詳細なプロンプトテンプレート |

## 📊 実行ログ

「実行ログ」シートで以下を記録：

- **CalendarEventID**: 処理した会議のイベントID
- **ProcessedAt**: 処理実行時刻
- **Status**: SUCCESS または ERROR
- **ErrorMessage**: エラー内容（エラー時のみ）
- **MeetingTitle**: 会議タイトル

## 🔧 主要関数

### メイン処理
- `main()`: 全体のメイン処理（トリガーから実行）

### 設定・データ管理
- `getConfig()`: 設定シートから設定値を読み込み
- `getRecentEvents()`: 直近の会議イベントを取得
- `isProcessed()`: 重複処理チェック
- `logExecution()`: 実行ログの記録

### API連携
- `getTranscriptByEvent()`: Meet API でトランスクリプト取得
- `generateMinutes()`: Vertex AI で議事録生成
- `sendMinutesEmail()`: Gmail でメール送信

### ユーティリティ（Utils.gs）
- `testSystem()`: システム全体のテスト
- `healthCheck()`: システムの健全性チェック
- `getStatistics()`: 実行統計の取得
- `validateConfiguration()`: 設定値の検証
- `processSpecificEvent()`: 特定イベントの手動処理
- `cleanupLogs()`: 古いログのクリーンアップ

## 🔄 自動実行の流れ

1. **トリガー起動** (毎時実行)
2. **設定読み込み** (Sheets から)
3. **イベント取得** (Calendar API)
4. **重複チェック** (実行ログ参照)
5. **トランスクリプト取得** (Meet API、リトライ付き)
6. **議事録生成** (Vertex AI Gemini)
7. **メール送信** (Gmail API)
8. **ログ記録** (Sheets に記録)

## 🛠️ 運用・保守

### 日常監視
```javascript
// システムの健全性チェック
healthCheck()

// 過去7日間の統計取得
getStatistics(7)

// 設定の検証
validateConfiguration()
```

### トラブルシューティング
```javascript
// 特定の会議を手動処理
processSpecificEvent('calendar-event-id')

// システム全体のテスト
testSystem()

// 古いログのクリーンアップ（30日より古いもの）
cleanupLogs(30)
```

## 📝 議事録フォーマット

生成される議事録は以下の構造：

```
## 1. 会議の概要
- **会議名**: [会議タイトル]
- **日時**: YYYY/MM/DD HH:MM
- **参加者**: [参加者リスト]

## 2. 決定事項
- [決定内容の箇条書き]

## 3. アクションアイテム
- **担当者**: [タスク内容] (期限: YYYY/MM/DD)

## 4. 議論内容の要約
- [主要議題・議論の流れ]
```

## 💰 コスト

- **Vertex AI**: 生成トークン数に応じて課金（無料枠あり）
- **Meet API**: 無料（Workspace プランの制限内）
- **その他のGoogle API**: 基本的に無料

## 🔒 セキュリティ

- OAuth 2.0 認証を使用
- API キーは使用せず、トークンベース認証
- 最小権限の原則を適用
- 設定データはスプレッドシートで管理

## ❗ 注意事項

1. **タイムラグ**: Meet トランスクリプトは会議終了後数分〜十数分で利用可能
2. **プライバシー**: 機密会議のトランスクリプト取得には注意
3. **レート制限**: API の利用制限に注意
4. **権限管理**: Admin Console での適切な権限設定が必要

## 📚 参考資料

- [Google Apps Script](https://developers.google.com/apps-script)
- [Google Meet API](https://developers.google.com/meet)
- [Vertex AI API](https://cloud.google.com/vertex-ai/docs)
- [Google Calendar API](https://developers.google.com/calendar)

## 🆘 サポート

問題が発生した場合：

1. 実行ログシートでエラー内容を確認
2. `healthCheck()` でシステム状態を確認
3. `validateConfiguration()` で設定を検証
4. セットアップ手順書の トラブルシューティング を参照

---

**バージョン**: 1.0  
**最終更新**: 2024年12月  
**作成者**: 議事録自動化システム開発チーム
