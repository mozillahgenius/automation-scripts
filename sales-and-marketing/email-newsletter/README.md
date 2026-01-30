# email-newsletter / メルマガ配信

## 概要
AI関連の最新情報を収集・整理し、メールマガジンとして配信するためのデータベース・プロファイル管理ツール。Gmail・Sheetsと連携したGAS環境で運用する。

## 技術構成
| 項目 | 詳細 |
|------|------|
| Platform | Google Apps Script |
| Language | JavaScript |
| メール送信 | Gmail |
| データ管理 | Google Sheets |
| コンテンツ | AI関連ニュース・記事 |

## ファイル構成
- `AI関連メルマガ.ini` - メルマガ記事データベース（AI関連ニュース・記事の蓄積）
- `profile.ini` - 配信プロファイル設定
- `profile_for_ai.md` - AI向けプロファイル情報（コンテンツ生成用コンテキスト）

## セットアップ
1. Google Apps Scriptプロジェクトを作成
2. メルマガ記事管理用のスプレッドシートを作成
3. `profile.ini` で配信先・配信元の設定を行う
4. `AI関連メルマガ.ini` の記事データを参照し、コンテンツを準備
5. Gmail送信の権限を承認

## 使い方
- **記事管理**: `AI関連メルマガ.ini` にAI関連ニュース・記事を蓄積
- **プロファイル設定**: `profile.ini` で配信設定を管理
- **AI活用**: `profile_for_ai.md` をAIに読み込ませてコンテンツ生成を補助
- **配信**: GASからGmail経由でメルマガを配信
