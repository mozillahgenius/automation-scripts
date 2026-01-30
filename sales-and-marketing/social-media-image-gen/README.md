# social-media-image-gen / Instagram・X画像自動生成

## 概要
Google Slides APIとDrive APIを活用し、Instagram・X（旧Twitter）向けの画像を1日3回自動生成するGoogle Apps Scriptツール。テンプレートベースでSNS投稿画像を一括作成する。

## 技術構成
| 項目 | 詳細 |
|------|------|
| Platform | Google Apps Script |
| Language | JavaScript |
| APIs | Google Slides API, Google Drive API |
| 出力 | 画像ファイル（PNG/JPEG） |
| トリガー | 時限トリガー（1日3回） |

## ファイル構成
- `Insta・X自動作成.js` - メインスクリプト（画像生成、テンプレート処理、Drive出力）
- `Drive_API_Setup_Guide.md` - Google Drive API セットアップガイド
- `SlideSpeak_Setup_Guide.md` - SlideSpeak セットアップガイド

## セットアップ
1. Google Apps Scriptプロジェクトを作成
2. `Insta・X自動作成.js` の内容をスクリプトエディタに貼り付け
3. Google Slides APIとDrive APIを有効化（`Drive_API_Setup_Guide.md` 参照）
4. テンプレート用のGoogleスライドを作成し、スライドIDを設定
5. 画像出力先のDriveフォルダIDを設定
6. 時限トリガーを設定（1日3回の実行スケジュール）

## 使い方
- **自動生成**: トリガーにより1日3回、SNS投稿用画像が自動生成される
- **テンプレート編集**: Googleスライドのテンプレートを編集して画像デザインを変更
- **手動実行**: スクリプトエディタから関数を直接実行して即時生成も可能
