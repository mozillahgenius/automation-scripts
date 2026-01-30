# calendar-export / カレンダーイベント抽出ツール

## 概要

Google カレンダーのイベント情報を抽出し、Google スプレッドシートに一覧出力するツール。Admin Directory API を使用して組織内ユーザーのカレンダー情報も取得可能。

## 技術構成

| 項目 | 詳細 |
|------|------|
| Platform | Google Apps Script |
| APIs | Google Calendar API, Admin Directory API, Google Sheets API |
| Triggers | 手動実行 |

## ファイル構成

- `Code.gs` - メイン処理（カレンダーイベント取得・スプレッドシート出力）
- `appsscript.json` - GAS プロジェクト設定（OAuth スコープ含む）
- `セットアップ手順.md` - セットアップ手順の詳細ガイド
- `日付修正ガイド.md` - 日付範囲の指定・修正方法ガイド

## セットアップ

1. Google Apps Script プロジェクトを作成
2. `Code.gs` と `appsscript.json` をプロジェクトにコピー
3. `セットアップ手順.md` に従い、必要な API を有効化
4. Admin Directory API を使用する場合は Google Workspace 管理者権限が必要
5. 出力先スプレッドシートの ID をスクリプト内に設定

## 使い方

- スクリプトエディタからメイン関数を実行
- 抽出対象の日付範囲はスクリプト内で指定（詳細は `日付修正ガイド.md` を参照）
- 実行後、指定したスプレッドシートにイベント一覧が出力される
