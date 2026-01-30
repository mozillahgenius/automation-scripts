# automation-scripts

業務自動化スクリプト集 — Google Apps Script / JavaScript / TypeScript / Python を中心とした55以上の自動化プロジェクト

> **Note:** gas-shared-drive リポジトリから統合されたプロジェクトを含みます。

## カテゴリ一覧

### Slack連携 (`slack-integrations/`)

| プロジェクト | 概要 | 技術 |
|-------------|------|------|
| [slack-ai-bot](slack-integrations/slack-ai-bot/) | ChatGPT/Gemini連携Slackボット（FAQ自動応答・ドキュメント修正） | GAS, Slack API, OpenAI, Gemini |
| [slack-log](slack-integrations/slack-log/) | Slackメッセージ収集・マニュアル/FAQ自動生成 | GAS, Slack API, OpenAI |
| [slack-channel-creator](slack-integrations/slack-channel-creator/) | スプレッドシートからSlackチャンネル自動作成 | GAS, Slack API |
| [slack-agenda-generator](slack-integrations/slack-agenda-generator/) | Slackメッセージから議題候補抽出・議事録案生成 | GAS, Slack API, OpenAI |
| [slack-drive-naming-audit](slack-integrations/slack-drive-naming-audit/) | Slack/Google Drive命名ルール監査 | GAS, Slack API, Drive API |
| [slackbot-faq](slack-integrations/slackbot-faq/) | キーワードマッチ型FAQ Bot（最小構成） | GAS, Slack API |

### Google Workspace自動化 (`google-workspace/`)

| プロジェクト | 概要 | 技術 |
|-------------|------|------|
| [calendar-export](google-workspace/calendar-export/) | カレンダーイベント抽出・グループフィルタリング | GAS, Calendar API, Admin SDK |
| [document-version-control](google-workspace/document-version-control/) | 文書バージョン管理台帳 | GAS, Sheets, Slack |
| [drive-rag-auto-reply](google-workspace/drive-rag-auto-reply/) | Drive内RAG＋OpenAI自動回答生成 | GAS, OpenAI API |
| [drive-structure-visualizer](google-workspace/drive-structure-visualizer/) | Drive階層構造可視化 | GAS, Drive API |
| [excel-to-spreadsheet](google-workspace/excel-to-spreadsheet/) | ExcelからGoogleスプレッドシートへインポート | GAS, Drive API |
| [meeting-background-slides](google-workspace/meeting-background-slides/) | 会議背景スライド自動生成 | GAS, Slides API |
| [meeting-minutes-automation](google-workspace/meeting-minutes-automation/) | Google Meet議事録自動生成（Gemini） | GAS, Gemini API, Calendar API |
| [bulk-document-generator](google-workspace/bulk-document-generator/) | テンプレートDocsから文書一括生成 | GAS, Docs API |
| [offboarding-automation](google-workspace/offboarding-automation/) | 退職手続自動化（メール・SSO・デバイス管理） | GAS, Admin SDK, Gmail API |
| [form-notifications](google-workspace/form-notifications/) | Googleフォーム送信通知 | GAS, Forms API, Gmail API |
| [drive-upload-monitor](google-workspace/drive-upload-monitor/) | 共有Drive新規ファイル監視＋Slack通知 | GAS, Drive API, Slack Webhook |
| [drive-governance](google-workspace/drive-governance/) | エンタープライズDriveガバナンス（命名規則・重複検出・権限管理） | GAS, Drive API |

### AI・ニュース自動化 (`ai-and-news/`)

| プロジェクト | 概要 | 技術 |
|-------------|------|------|
| [ai-newsletter-notion](ai-and-news/ai-newsletter-notion/) | AIニュース定時配信＋Notion自動更新 | GAS, Perplexity, Grok, Notion API |
| [critical-news-monitor](ai-and-news/critical-news-monitor/) | 企業批判記事監視・スコアリング | GAS, Perplexity Sonar API |

### 法務・コンプライアンス (`legal-and-compliance/`)

| プロジェクト | 概要 | 技術 |
|-------------|------|------|
| [corporate-minutes-generator](legal-and-compliance/corporate-minutes-generator/) | 会社法議事録作成・承認ワークフロー | GAS, Sheets, Docs |
| [litigation-management](legal-and-compliance/litigation-management/) | 訴訟案件一元管理・リマインド | GAS, Slack Webhook, Gmail |
| [sanctions-list-processor](legal-and-compliance/sanctions-list-processor/) | 反社チェックリスト自動化（AI活用） | GAS, OpenAI, Perplexity |
| [defamation-monitoring](legal-and-compliance/defamation-monitoring/) | 誹謗中傷対応・名誉毀損監視 | JS, Perplexity, Grok |
| [capital-reduction-monitor](legal-and-compliance/capital-reduction-monitor/) | 減資公告監視・官報データ収集 | GAS/JS, Custom Search API |

### 営業・マーケティング (`sales-and-marketing/`)

| プロジェクト | 概要 | 技術 |
|-------------|------|------|
| [sales-manual](sales-and-marketing/sales-manual/) | 営業マニュアル（Salesforce対応） | Documentation |
| [newsletter-system](sales-and-marketing/newsletter-system/) | ニュースレター自動配信 | Node.js/TS, Perplexity, Grok |
| [social-media-image-gen](sales-and-marketing/social-media-image-gen/) | Instagram・X画像自動生成 | GAS, Slides API |
| [foreign-company-detector](sales-and-marketing/foreign-company-detector/) | 外資系企業検知（日本進出初期） | GAS, LinkedIn, Google News |
| [client-prospector](sales-and-marketing/client-prospector/) | 情シス・法務クライアント開拓 | GAS, Zapier, PR Times RSS |
| [youtube-analyzer](sales-and-marketing/youtube-analyzer/) | YouTube動画分析・管理 | GAS, YouTube Data/Analytics API |
| [email-newsletter](sales-and-marketing/email-newsletter/) | メルマガ配信 | GAS, Gmail |
| [sns-reaction-analyzer](sales-and-marketing/sns-reaction-analyzer/) | SNS反響分析（YouTube/X/Googleトレンド＋AI） | GAS, Gemini, Grok, YouTube API |

### 業務オペレーション (`business-operations/`)

| プロジェクト | 概要 | 技術 |
|-------------|------|------|
| [hubspot-status-reminder](business-operations/hubspot-status-reminder/) | HubSpot未更新案件通知 | GAS, HubSpot API, Gmail |
| [salesforce-field-manager](business-operations/salesforce-field-manager/) | Salesforceオブジェクト・フィールド管理 | GAS, Salesforce REST API |
| [performance-db](business-operations/performance-db/) | 実績DB自動連携 | GAS, Sheets |
| [daily-performance-report](business-operations/daily-performance-report/) | 業績日次レポート自動送信 | GAS, Gmail |
| [governance-task-manager](business-operations/governance-task-manager/) | ガバナンスタスク管理 | GAS, OpenAI, Gmail |
| [license-management](business-operations/license-management/) | ライセンス管理（Google/Microsoft統合） | GAS, Admin SDK |
| [business-flowchart](business-operations/business-flowchart/) | 業務フローチャート自動生成 | GAS, Sheets |
| [automation-methodology](business-operations/automation-methodology/) | 業務自動化の方法論 | Documentation |
| [warehouse-management](business-operations/warehouse-management/) | 倉庫管理システム設計 | System Design |
| [license-management-en](business-operations/license-management-en/) | ライセンスコスト管理（英語版・動的列拡張） | GAS, Sheets |
| [mail-reminder](business-operations/mail-reminder/) | 未処理メール管理＋優先度判定＋Slack通知 | GAS, Gmail, Slack Webhook |
| [notion-activity-log](business-operations/notion-activity-log/) | Notionワークスペース変更監視・アクティビティ分析 | GAS, Notion API |

### セキュリティ (`security/`)

| プロジェクト | 概要 | 技術 |
|-------------|------|------|
| [phishing-training](security/phishing-training/) | 標的型メール訓練 | GAS, Gmail, Perplexity |
| [isms-risk-map](security/isms-risk-map/) | ISMSリスクアセスメント自動化（ヒートマップ・CIA分類） | GAS, Sheets |

### 個人プロジェクト (`personal/`)

| プロジェクト | 概要 | 技術 |
|-------------|------|------|
| [pokemon-card-manager](personal/pokemon-card-manager/) | ポケモンカード管理（AI判定→Notion登録） | GAS, Photos API, Notion, Perplexity |
| [crypto-tracker](personal/crypto-tracker/) | 暗号資産期待値レポート | JS, CoinGecko, Perplexity, Grok |
| [stock-screener](personal/stock-screener/) | 株式銘柄チェック・レポート | GAS, Gmail |
| [crypto-portfolio-manager](personal/crypto-portfolio-manager/) | 暗号資産ポートフォリオ管理（ブレイクアウト検知・リバランス） | GAS, CoinGecko API |

### ブラウザ拡張 (`browser-extensions/`)

| プロジェクト | 概要 | 技術 |
|-------------|------|------|
| [clipboard-history](browser-extensions/clipboard-history/) | クリップボード履歴マネージャー | Chrome Extension (Manifest V3) |
| [browsing-history-tracker](browser-extensions/browsing-history-tracker/) | ブラウジング履歴収集・CSV出力 | Chrome Extension (Manifest V3) |

### SNS自動化 (`social-media-automation/`)

| プロジェクト | 概要 | 技術 |
|-------------|------|------|
| [instagram-twitter-bot](social-media-automation/instagram-twitter-bot/) | Instagram/Twitterの自動操作（いいね・投稿） | Python, Selenium |

## Tech Stack

| 技術 | 用途 |
|------|------|
| **Google Apps Script** | 大半のプロジェクトの実行基盤 |
| **OpenAI API** | テキスト生成・分析 |
| **Perplexity API** | リアルタイム情報検索・AI分析 |
| **Grok API** | AI分析補助 |
| **Google Gemini API** | 議事録生成・テキスト処理 |
| **Slack API** | チャンネル管理・メッセージ処理・通知 |
| **Notion API** | データベース自動更新 |
| **YouTube API** | 動画メタデータ・アナリティクス |
| **Salesforce REST API** | CRM連携 |
| **HubSpot API** | 案件管理連携 |
| **Node.js / TypeScript** | newsletter-system |
| **Chrome Extensions** | clipboard-history, browsing-history-tracker |
| **Python / Selenium** | instagram-twitter-bot |

## GAS共通セットアップ手順

ほとんどのプロジェクトはGoogle Apps Scriptで動作します。

### 1. GASプロジェクト作成
1. [Google Drive](https://drive.google.com) を開く
2. 新規 > その他 > Google Apps Script を選択
3. プロジェクト名を設定

### 2. コードの配置
1. 各プロジェクトの `.gs` ファイルをスクリプトエディタにコピー
2. `.html` ファイルがある場合は「ファイル追加 > HTML」で追加
3. `appsscript.json` がある場合は「プロジェクトの設定 > マニフェストを表示」で内容を反映

### 3. APIキーの設定
多くのプロジェクトで外部APIキーが必要です：
- **スクリプトプロパティ**: プロジェクトの設定 > スクリプトプロパティ でキーを設定
- 必要なAPIキーは各プロジェクトのREADMEを参照

### 4. トリガーの設定
定期実行が必要な場合：
1. スクリプトエディタ > トリガー（時計アイコン）
2. 「トリガーを追加」で関数・実行頻度を設定

### 5. 権限の承認
初回実行時にGoogleアカウントの権限承認が必要です。
「詳細」>「（安全でないページ）に移動」で承認してください。

## ライセンス

社内業務自動化を目的としたプライベートリポジトリです。
