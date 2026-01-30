# governance-task-manager / ガバナンスタスク管理

## 概要
メールからガバナンス関連タスクを自動抽出し、開示チェック・外部アドバイザー相談・MECE分類・タイムライン生成などを統合的に管理するGoogle Apps Scriptシステム。OpenAI APIを活用した自然言語処理により、メール本文からタスクを解析・分類し、適切なワークフローへ振り分ける。

## 技術構成
| 項目 | 詳細 |
|------|------|
| Platform | Google Apps Script |
| 言語 | JavaScript (GAS) |
| APIs | Gmail, Google Sheets, Google Drive, OpenAI API |
| 機能 | メール解析、タスク抽出、開示チェック、MECE分類 |
| トリガー | 時間主導型 / メニュー手動実行 |

## ファイル構成
- `integrated_governance_functions.gs` - ガバナンス機能の統合メイン処理
- `data_processor.gs` - データの加工・変換処理
- `config.gs` - 設定値（APIキー、スプレッドシートID等）
- `gmail_outbound.gs` - メール送信処理
- `gmail_inbound.gs` - メール受信・解析処理
- `flow_visualizer.gs` - フロー可視化処理
- `governance_disclosure_checker.gs` - 開示チェック処理
- `external_advisor_consultation.gs` - 外部アドバイザー相談管理
- `parser_and_writer.gs` - メール解析とシート書き込み
- `menu.gs` - カスタムメニュー定義
- `openai_client.gs` - OpenAI API クライアント
- `debug.gs` - デバッグ・テスト用ユーティリティ
- `combined_gas_code.gs` - 統合コード（一括版）
- `appsscript.json` - GASプロジェクト設定
- `仕様書_タスク管理作成.ini` - タスク管理の仕様書

## セットアップ
1. Google Apps Scriptプロジェクトを新規作成
2. `config.gs` にOpenAI APIキーと対象スプレッドシートIDを設定
3. 全 `.gs` ファイルをエディタに追加
4. Gmail、Drive、Sheetsの OAuth スコープを承認
5. `menu.gs` によりスプレッドシートにカスタムメニューが追加される
6. 時間主導型トリガーを設定（メール自動取得用）

## 使い方
- **タスク抽出**: カスタムメニューまたはトリガーでGmailからガバナンス関連メールを取得・解析
- **開示チェック**: `governance_disclosure_checker.gs` により開示要否を自動判定
- **アドバイザー相談**: `external_advisor_consultation.gs` で外部専門家への相談内容を管理
- **フロー可視化**: `flow_visualizer.gs` でタスクフローを視覚的に表示
- **MECE分類**: OpenAI APIにより、タスクを漏れなくダブりなく分類
