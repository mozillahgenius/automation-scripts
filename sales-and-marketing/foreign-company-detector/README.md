# foreign-company-detector / 外資系企業検知

## 概要
日本に進出する外資系企業を4段階のシグナル検知で自動的に発見するGoogle Apps Scriptツール。LinkedIn Jobs・Google News・RSSフィードを横断的に監視し、営業アプローチの早期タイミングを捕捉する。

## 技術構成
| 項目 | 詳細 |
|------|------|
| Platform | Google Apps Script |
| Language | JavaScript |
| データ管理 | Google Sheets |
| 情報源 | LinkedIn Jobs API, Google News, RSSフィード |
| 検知方式 | 4段階シグナル検知 |
| トリガー | 定期実行トリガー |

## ファイル構成
- `ForeignCompanyDetection.gs` - メインスクリプト（4段階シグナル検知、データ収集・分析・出力）
- `ニュース起点×前倒し統合型 日本進出初期外資検知設計.docx` - 設計ドキュメント（ニュース起点統合型）
- `日本進出初期外資系企業 検知・アプローチ設計.docx` - 設計ドキュメント（検知・アプローチ設計）

## セットアップ
1. Google Apps Scriptプロジェクトを作成
2. `ForeignCompanyDetection.gs` の内容をスクリプトエディタに貼り付け
3. 検知結果を蓄積するスプレッドシートを作成し、シートIDを設定
4. 各情報源のAPI設定:
   - LinkedIn Jobs検索パラメータ
   - Google Newsキーワード
   - 監視対象RSSフィードURL
5. 定期実行トリガーを設定

## 使い方
- **自動検知**: トリガーにより定期的に外資系企業の日本進出シグナルを検出
- **4段階検知**: 初期シグナル → 確度判定 → 企業情報収集 → アプローチリスト化
- **結果確認**: スプレッドシートで検知結果・企業情報を一覧確認
- **設計参照**: 詳細なロジックは `.docx` 設計ドキュメントを参照
