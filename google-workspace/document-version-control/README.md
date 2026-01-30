# document-version-control / Document バージョン管理台帳

## 概要

Google ドキュメントのバージョン管理を台帳形式で行うシステム。ドキュメントの変更履歴を追跡し、Gmail や Slack を通じて関係者に通知を送信する。メニュー UI から操作可能。

## 技術構成

| 項目 | 詳細 |
|------|------|
| Platform | Google Apps Script |
| APIs | Google Sheets API, Gmail API, Slack Webhook |
| Triggers | スプレッドシート連動（onOpen）/ 手動実行 |

## ファイル構成

- `Code.gs` - メイン処理・ユーティリティ関数
- `DocumentManagementSystem.gs` - ドキュメント管理システムのコアロジック
- `Menu.gs` - スプレッドシートのカスタムメニュー定義
- `SlackNotification.gs` - Slack 通知送信処理
- `GmailIntegration.gs` - Gmail 通知送信処理
- `appsscript.json` - GAS プロジェクト設定

## セットアップ

1. Google スプレッドシートを作成し、スクリプトエディタを開く
2. 各 `.gs` ファイルをプロジェクトにコピー
3. Slack Webhook URL をスクリプトプロパティに設定
4. 台帳用のシートを初期化（メニューから実行可能）
5. 管理対象ドキュメントの情報を台帳に登録

## 使い方

- スプレッドシートを開くとカスタムメニューが表示される
- メニューからドキュメントの登録・更新・バージョン管理を操作
- ドキュメントの更新時に Gmail および Slack で関係者に自動通知
- 台帳シートでバージョン履歴・承認状況を一覧管理
