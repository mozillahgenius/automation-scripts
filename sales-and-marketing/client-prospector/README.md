# client-prospector / 情シス・法務クライアント開拓

## 概要
情報システム部門・法務部門向けの業務委託クライアントを自動的にピックアップするGoogle Apps Scriptツール。PR Times・Indeed・Google Newsなどの複数ソースからシグナルを検出し、Slack通知・CRM連携を行う。

## 技術構成
| 項目 | 詳細 |
|------|------|
| Platform | Google Apps Script |
| Language | JavaScript |
| データ管理 | Google Sheets |
| 情報源 | PR Times RSS, Indeed, Google News |
| 連携 | Zapier, Slack, HubSpot / Salesforce |
| トリガー | 定期実行トリガー |

## ファイル構成
- `ClientPickup.gs` - メインスクリプト（シグナル検出、候補企業ピックアップ、通知・CRM連携）
- `Zapier_Setup_Guide.md` - Zapier連携セットアップガイド
- `情シス・法務 業務委託クライアント自動ピックアップ仕様書 V1.docx` - 仕様書

## セットアップ
1. Google Apps Scriptプロジェクトを作成
2. `ClientPickup.gs` の内容をスクリプトエディタに貼り付け
3. 候補企業の蓄積用スプレッドシートを作成し、シートIDを設定
4. 各情報源の設定:
   - PR Times RSSフィードURL
   - Indeed検索パラメータ
   - Google Newsキーワード
5. 通知・CRM連携の設定:
   - Slack Webhook URL（`Zapier_Setup_Guide.md` 参照）
   - HubSpot / Salesforce API連携（Zapier経由）
6. 定期実行トリガーを設定

## 使い方
- **自動ピックアップ**: トリガーにより定期的に候補企業を自動検出
- **Slack通知**: 新規候補企業が検出されるとSlackに通知
- **CRM連携**: Zapier経由でHubSpot/Salesforceに候補企業を自動登録
- **結果確認**: スプレッドシートで候補企業リストを一覧管理
