# bulk-document-generator / 文書大量作成（テンプレート Docs から一括生成）

## 概要

Google スプレッドシートのデータを基に、Google Docs テンプレートから文書を一括生成するツール。テンプレート内のプレースホルダをスプレッドシートの各行データで置換し、個別の Docs ファイルとして出力する。

## 技術構成

| 項目 | 詳細 |
|------|------|
| Platform | Google Apps Script |
| APIs | Google Docs API, Google Drive API, Google Sheets API |
| Triggers | 手動実行 |

## ファイル構成

- `spreadsheet_to_docs.gs` - メイン処理（テンプレート複製・プレースホルダ置換・一括生成）
- `setup_instructions.md` - セットアップ手順書

## セットアップ

1. Google Apps Script プロジェクトを作成
2. `spreadsheet_to_docs.gs` をプロジェクトにコピー
3. Google Docs でテンプレートを作成（`{{項目名}}` 形式のプレースホルダを配置）
4. スプレッドシートにデータを準備（ヘッダ行がプレースホルダ名に対応）
5. テンプレート Docs ID、スプレッドシート ID、出力先フォルダ ID をスクリプトに設定
6. 詳細は `setup_instructions.md` を参照

## 使い方

- スプレッドシートに生成対象のデータを行ごとに入力
- スクリプトエディタからメイン関数を実行
- スプレッドシートの各行に対してテンプレートが複製され、プレースホルダが置換される
- 生成された Docs ファイルは指定フォルダに一括保存される
