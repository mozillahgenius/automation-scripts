# offboarding-automation / 退職手続自動化

## 概要

退職者のアカウント処理を自動化するシステム。機密メールの削除、メール転送設定、SSO 連携の検出、デバイス管理など、退職に伴う一連の手続きを Google Workspace 管理 API を通じて実行する。

## 技術構成

| 項目 | 詳細 |
|------|------|
| Platform | Google Apps Script |
| APIs | Gmail API, Admin Directory API, Google Drive API, Google Calendar API |
| Triggers | 手動実行 |

## ファイル構成

- `退職手続.gs` - 全処理（メール削除、転送設定、SSO 検出、デバイス管理）

## セットアップ

1. Google Apps Script プロジェクトを作成し、`退職手続.gs` をコピー
2. Google Workspace 管理者権限を持つアカウントでプロジェクトを作成
3. Admin SDK API を GCP プロジェクトで有効化
4. 必要な OAuth スコープを `appsscript.json` に追加
5. 退職者情報の入力方法（スプレッドシート等）を設定

## 使い方

- 退職者のメールアドレス等の情報を入力
- スクリプトを実行すると以下の処理が自動実行される:
  - 機密メールの検出・削除
  - メール転送設定の構成
  - SSO 連携サービスの検出・レポート
  - デバイス管理（ワイプ・無効化）
  - カレンダー・Drive データの引き継ぎ処理
- 処理結果がログに出力される

> **注意:** 管理者権限が必要な操作を含むため、実行前に対象アカウントと処理内容を十分に確認すること。
