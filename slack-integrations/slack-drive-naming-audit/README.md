# slack-drive-naming-audit / Slack・Google Drive 命名ルール監査システム

## 概要

Slack チャンネルおよび Google Drive のファイル・フォルダに対して、命名規則の準拠状況を自動監査するシステム。違反を検出し、レポートをスプレッドシートに出力する。全階層のフォルダ・ファイルを対象にした包括的な命名監査が可能。

## 技術構成

| 項目 | 詳細 |
|------|------|
| Platform | Google Apps Script |
| APIs | Slack API, Google Drive API, Google Sheets API |
| Triggers | 時間主導型（定期実行） |

## ファイル構成

- `main.gs` - メイン処理・実行エントリポイント
- `slackAPI.gs` - Slack API との通信処理
- `driveAPI.gs` - Google Drive API との通信処理
- `report.gs` - 監査レポート生成・スプレッドシート出力
- `validation.gs` - 命名規則のバリデーションロジック
- `初期設定.gs` - 初期設定・環境構築用スクリプト
- `統合版.gs` - 全機能統合版スクリプト
- `appsscript.json` - GAS プロジェクト設定
- `slack_app_manifest.yaml` - Slack App マニフェスト定義

## セットアップ

1. Google Apps Script プロジェクトを作成し、各 `.gs` ファイルをコピー
2. `slack_app_manifest.yaml` を使用して Slack App を作成
3. Slack Bot Token・Signing Secret をスクリプトプロパティに設定
4. 対象の Google Drive フォルダ ID をスクリプトプロパティに設定
5. `初期設定.gs` を実行して初期データを構築
6. 必要に応じて時間主導型トリガーを設定

## 使い方

- 初回は `初期設定.gs` の関数を実行して環境を構築する
- 定期実行トリガーを設定すると、自動で命名規則監査が実行される
- 監査結果はスプレッドシートにレポートとして出力される
- 違反が検出された場合、対象のファイル・チャンネル情報が一覧表示される
