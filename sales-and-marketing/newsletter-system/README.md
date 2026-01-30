# newsletter-system / ニュースレター自動配信

## 概要
Perplexity・Grok APIで最新トピックをリサーチし、ニュースレターを自動生成・Gmail SMTP経由で配信するNode.jsアプリケーション。node-cronによるスケジュール実行に対応。

## 技術構成
| 項目 | 詳細 |
|------|------|
| Platform | Node.js |
| Language | TypeScript |
| AI APIs | Perplexity API, Grok API (X AI) |
| メール送信 | Nodemailer (Gmail SMTP) |
| スケジューラ | node-cron |
| HTTP | axios |
| 設定管理 | dotenv |

## ファイル構成
- `package.json` - 依存関係定義
- `tsconfig.json` - TypeScript設定
- `.env.example` - 環境変数テンプレート
- `node_modules/` - 依存パッケージ

## セットアップ
1. リポジトリをクローンまたはファイルをコピー
2. 依存パッケージをインストール:
   ```bash
   npm install
   ```
3. `.env.example` をコピーして `.env` を作成:
   ```bash
   cp .env.example .env
   ```
4. `.env` に各種APIキー・SMTP情報を設定:
   - `PERPLEXITY_API_KEY` - Perplexity APIキー
   - `GROK_API_KEY` - Grok APIキー
   - `SMTP_HOST` / `SMTP_PORT` / `SMTP_USER` / `SMTP_PASS` - Gmail SMTP設定
   - `NEWSLETTER_FROM` / `NEWSLETTER_TO` - 送受信メールアドレス
   - `RESEARCH_TOPICS` - リサーチトピック（カンマ区切り）
5. Gmailのアプリパスワードを発行し `SMTP_PASS` に設定

## 使い方
- **手動実行**: `npx tsx src/index.ts` でニュースレターを即時生成・送信
- **スケジュール実行**: node-cronにより設定した間隔で自動配信
- **トピック変更**: `.env` の `RESEARCH_TOPICS` を編集してリサーチ対象を変更
