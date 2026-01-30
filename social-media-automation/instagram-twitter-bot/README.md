# Instagram / Twitter Bot

Selenium + Pythonによる Instagram 自動いいね・Twitter 自動投稿ツール。

## 機能

### Instagram 自動いいね (`favoinsta.py`)
- Instagram にログインし、指定タグの投稿を自動検索
- 人気投稿・最新投稿をランダムに選択していいね
- ランダム待機時間による検知回避
- 実行ログをテキストファイルに記録

### Twitter 自動投稿 (`twitter_tourokuzyun.py`)
- テキストファイルから投稿内容を順番に読み込み
- Chrome プロファイルによるログイン状態維持
- クリップボード経由でのテキスト入力（改行対応）
- 投稿済み行数を記録し、全行投稿後にリセット

## ファイル構成

| ファイル | 説明 |
|---------|------|
| `favoinsta.py` | Instagram 自動いいねスクリプト |
| `twitter_tourokuzyun.py` | Twitter 自動投稿スクリプト |
| `test.py` | テスト用スクリプト |
| `hello.py` | 動作確認用スクリプト |
| `pythonversion.txt` | Python バージョン情報 |

## 前提条件

- Python 3.x
- Selenium（Instagram: v3 / Twitter: v4）
- Chrome / ChromeDriver
- `webdriver-manager`（Twitter用）
- `pyperclip`（Twitter用）

## セットアップ

```bash
pip install selenium==3 webdriver-manager pyperclip
```

## 注意事項

- Instagram / Twitter の利用規約に準拠した運用を推奨
- 認証情報（ユーザー名・パスワード）は別途管理してください
- API 変更やUI変更により動作しなくなる場合があります

## 技術スタック

- Python 3.x
- Selenium WebDriver
- Chrome DevTools Protocol
