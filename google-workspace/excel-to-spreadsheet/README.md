# excel-to-spreadsheet / Excel から Google スプレッドシートへのインポート

## 概要

Excel ファイル（.xlsx）を Google スプレッドシートにインポートするシンプルなツール。Google Drive 上の Excel ファイルを変換してスプレッドシート形式で保存する。

## 技術構成

| 項目 | 詳細 |
|------|------|
| Platform | Google Apps Script |
| APIs | Google Drive API, Google Sheets API |
| Triggers | 手動実行 |

## ファイル構成

- `gas_simple_excel_import.gs` - Excel インポート処理

## セットアップ

1. Google Apps Script プロジェクトを作成
2. `gas_simple_excel_import.gs` をプロジェクトにコピー
3. インポート対象の Excel ファイルを Google Drive にアップロード
4. スクリプト内にファイル ID または フォルダ ID を設定

## 使い方

- スクリプトエディタからメイン関数を実行
- 指定された Excel ファイルが Google スプレッドシート形式に変換される
- 変換後のスプレッドシートは Google Drive に保存される
