# corporate-minutes-generator / 会社法議事録作成・承認ワークフロー

## 概要

会社法に準拠した議事録をテンプレートから自動作成し、承認ワークフローで管理するシステム。議案・報告事項の追加、役員管理、テンプレート管理、承認状況の追跡、バージョン管理を HTML ダイアログ UI で操作可能。

## 技術構成

| 項目 | 詳細 |
|------|------|
| Platform | Google Apps Script |
| APIs | Google Sheets API, Google Docs API, Google Drive API |
| Triggers | スプレッドシート連動（onOpen）/ 手動実行 |

## ファイル構成

- `統合Code.gs` - メイン処理（議事録生成、承認ワークフロー、バージョン管理）
- `add_motion.html` - 議案追加ダイアログ
- `add_report.html` - 報告事項追加ダイアログ
- `approval_status.html` - 承認状況確認ダイアログ
- `config.html` - 設定ダイアログ
- `create_draft.html` - 議事録ドラフト作成ダイアログ
- `manage_templates.html` - テンプレート管理ダイアログ
- `officers_management.html` - 役員管理ダイアログ
- `post_meeting_editor.html` - 会議後編集ダイアログ
- `simple_template_manager.html` - 簡易テンプレート管理ダイアログ
- `status_change.html` - ステータス変更ダイアログ
- `test_template.html` - テンプレートテストダイアログ

## セットアップ

1. Google スプレッドシートを作成し、スクリプトエディタを開く
2. `統合Code.gs` と全 HTML ファイルをプロジェクトにコピー
3. 議事録テンプレート用の Google Docs を作成
4. 出力先フォルダを Google Drive に作成
5. テンプレート Docs ID、出力先フォルダ ID をスクリプト内またはスプレッドシートの設定シートに登録
6. 役員情報を初期登録

## 使い方

- スプレッドシートを開くとカスタムメニューが表示される
- メニューから各種操作を選択:
  - **議案追加**: 議案の内容・決議事項を入力
  - **報告事項追加**: 報告事項を入力
  - **ドラフト作成**: テンプレートから議事録ドラフトを生成
  - **承認管理**: 各役員の承認状況を確認・更新
  - **役員管理**: 出席者・署名者の情報を管理
  - **テンプレート管理**: 議事録テンプレートの追加・編集
- 承認完了後、最終版の議事録が Google Docs として保存される
