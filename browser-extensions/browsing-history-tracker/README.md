# Browsing History Tracker

Chrome拡張機能（Manifest V3）によるブラウジング履歴収集ツール。

## 機能

- ブラウザの閲覧履歴を自動収集
- 定期的に外部サーバーへ履歴データをPOST（60秒間隔）
- 管理者向けCSVダウンロード機能（全ユーザーの履歴を一括取得）
- ユーザーのメールアドレスをキーとした履歴管理
- 重複データの自動除外

## ファイル構成

| ファイル | 説明 |
|---------|------|
| `manifest.json` | Chrome拡張マニフェスト（V3） |
| `background.js` | メインロジック（履歴収集・API通信・CSV出力） |
| `popup.html` | ポップアップUI |
| `popup.css` | ポップアップスタイル |
| `icon32.png` | 拡張機能アイコン |

## 必要な権限

- `storage` — データ保存
- `history` — ブラウザ履歴の読み取り
- `identity` / `identity.email` — ユーザー識別

## セットアップ

1. `chrome://extensions/` を開く
2. 「デベロッパーモード」を有効化
3. 「パッケージ化されていない拡張機能を読み込む」でこのフォルダを選択

## 技術スタック

- Chrome Extension Manifest V3
- Chrome History API / Identity API
- Fetch API（外部サーバー連携）
