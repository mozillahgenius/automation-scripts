# capital-reduction-monitor / 減資公告監視・官報データ収集

## 概要
官報に掲載される減資公告を自動監視し、データを収集・蓄積するGoogle Apps Scriptツール。Google Custom Search APIとWebスクレイピングを併用して官報データを取得する。

## 技術構成
| 項目 | 詳細 |
|------|------|
| Platform | Google Apps Script |
| Language | JavaScript |
| データ管理 | Google Sheets, Google Drive |
| 検索 | Google Custom Search API |
| データ取得 | Webスクレイピング（UrlFetchApp） |
| トリガー | 定期実行トリガー |

## ファイル構成
- `code_v3.js` - メインスクリプト（減資公告監視、データ解析、スプレッドシート出力）
- `gazette_collector.gs` - 官報データ収集スクリプト（官報ページのスクレイピング・データ蓄積）

## セットアップ
1. Google Apps Scriptプロジェクトを作成
2. `code_v3.js` と `gazette_collector.gs` をスクリプトエディタに追加
3. スクリプトプロパティに設定:
   - `CUSTOM_SEARCH_API_KEY`
   - `CUSTOM_SEARCH_ENGINE_ID`
4. データ蓄積用のスプレッドシートを作成し、シートIDを設定
5. Google Driveのデータ保存先フォルダIDを設定
6. 定期実行トリガーを設定（日次推奨）

## 使い方
- **自動監視**: トリガーにより官報の減資公告を定期的にチェック
- **データ収集**: `gazette_collector.gs` で官報ページからデータをスクレイピング
- **データ確認**: 収集結果はスプレッドシートに自動出力される
