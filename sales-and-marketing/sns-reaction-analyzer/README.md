# ネット反応分析システム
> YouTube/X/Google Trendsのデータ収集とマルチLLM（Gemini/Grok）によるAI分析・コンテンツサジェストを統合したSNS反応分析プラットフォーム

## 概要
複数のSNSプラットフォーム（YouTube、X）とGoogle Trendsからキーワードに基づくデータを収集し、AIによる分析とコンテンツサジェストを自動生成するGoogle Apps Scriptシステム（Version 4.0）。セットアップウィザードによる簡単な初期設定、3段階の動作モード（demo/basic/advanced）、エグゼクティブダッシュボード、スタイリッシュなHTML形式のレポート生成、日次/週次/月次の自動実行に対応する。

## 主な機能
- 簡易セットアップウィザード（5ステップで初期設定完了）
- 3段階の動作モード（demo: APIキー不要、basic: YouTube API、advanced: 全機能）
- YouTube Data API v3によるキーワード検索・動画詳細・チャンネル分析
- X (Twitter) API v2によるツイート検索・センチメント分析・トレンド取得
- Google Ads Keyword Planner APIによる検索ボリューム・競合分析
- Gemini API（gemini-1.5-pro-latest）によるAI分析・コンテンツサジェスト生成
- Grok API（grok-4）によるコンテンツサジェスト生成
- トレンド分析（YouTube投稿傾向、Xセンチメント連動）
- 総合スコア算出（100点満点、YouTube/X/トレンドの重み付け）
- エグゼクティブダッシュボード自動更新
- HTML形式のリッチレポート生成・メール送信
- Google Driveへのレポート自動保存
- コンテンツサジェストのスコアリング・ランキング
- 日次/週次/月次の自動分析トリガー
- エラーログ管理・管理者通知
- カスタム期間レポート（1-365日）
- 組み込みヘルプ（使い方ガイド、APIキー取得ガイド、トラブルシューティング）

## アーキテクチャ
ユーザーが設定シートでキーワード・動作モード・各種ON/OFFを設定すると、分析実行時に`loadSettings()`が設定を読み込む。動作モードに応じて`runDemoAnalysis`/`runBasicAnalysis`/`performFullAnalysis`が実行され、YouTube/X/Trends/Keyword Plannerからデータを収集する。収集データは`calculateScore`で100点満点のスコアに変換され、AI分析（Gemini）とコンテンツサジェスト（Gemini + Grok のデュアルLLM）が生成される。結果はレポートシート、各プラットフォーム別データシート、サジェストシートに保存され、ダッシュボードが更新される。オプションでHTML形式レポートがメール送信・Google Driveに保存される。

## ファイル構成
| ファイル | 説明 |
|---|---|
| Code.js | 全機能のメインロジック（約2300行、Version 4.0） |
| appsscript.json | GASプロジェクト設定（YouTube Advanced Service有効、V8ランタイム） |

## シート構成

### 設定
| 行 | 項目 | 設定値例 | 説明 |
|---|---|---|---|
| 2 | 分析キーワード | AI | カンマ区切りで複数指定可 |
| 3 | レポートタイプ | 短期/中期/長期 | 短期(7日)、中期(30日)、長期(365日) |
| 4 | 通知メール | user@example.com | レポート送信先 |
| 5 | 動作モード | demo/basic/advanced | 利用する機能レベル |
| 6 | YouTube分析 | ON/OFF | YouTube分析の有効/無効 |
| 7 | X分析 | ON/OFF | X分析の有効/無効 |
| 8 | AI要約 | ON/OFF | AI分析の有効/無効 |
| 9 | サジェスト生成 | ON/OFF | コンテンツサジェストの有効/無効 |
| 10 | 監視Xアカウント | user1,user2 | 監視するXアカウント（カンマ区切り） |
| 11 | 監視YouTubeチャンネル | UC... | 監視するチャンネルID（カンマ区切り） |

### ダッシュボード
パフォーマンスサマリー（総合スコア、YouTube再生数、X投稿数、検索トレンド）、オススメコンテンツTOP3、AI分析サマリーを自動表示する。

### レポート
| 列 | ヘッダー | 内容 |
|---|---|---|
| A | 生成日時 | 分析実行のタイムスタンプ |
| B | キーワード | 分析キーワード |
| C | スコア | 総合スコア（0-100） |
| D | YouTube再生数 | YouTube総再生回数 |
| E | X投稿数 | X総ツイート数 |
| F | トレンド | 成長率（%） |
| G | AI分析 | AI分析テキスト |
| H | トップサジェスト | 最上位サジェストのタイトル |

### YouTube_データ
| 列 | 内容 |
|---|---|
| A | タイムスタンプ |
| B | 動画ID |
| C | タイトル |
| D | チャンネル |
| E | 再生回数 |
| F | いいね数 |
| G | コメント数 |
| H | エンゲージメント率 |

### X_データ
| 列 | 内容 |
|---|---|
| A | タイムスタンプ |
| B | ツイートID |
| C | テキスト（先頭100文字） |
| D | いいね数 |
| E | RT数 |
| F | 返信数 |

### サジェスト
| 列 | 内容 |
|---|---|
| A | 生成日時 |
| B | タイトル |
| C | プラットフォーム |
| D | フォーマット |
| E | 予測パフォーマンス |
| F | 成功確率 |
| G | スコア |

### エラーログ
| 列 | 内容 |
|---|---|
| A | タイムスタンプ |
| B | エラータイプ |
| C | メッセージ |
| D | スタック |

## 主要関数

### メニュー・セットアップ
| 関数名 | 説明 |
|---|---|
| `onOpen()` | カスタムメニュー「分析システム」を登録（分析、設定、レポート、テスト、ヘルプのサブメニュー） |
| `runSetupWizard()` | 5ステップのセットアップウィザード実行 |
| `initializeSpreadsheet()` | 全8シートの作成・初期化 |
| `setupBasicSettings()` | キーワード・メール設定のプロンプト入力 |
| `selectOperationMode()` | 動作モード（demo/basic/advanced）の選択 |
| `setupAPIKeys()` | APIキーのプロンプト入力・スクリプトプロパティへの保存 |
| `finalizeSetup()` | セットアップ完了メッセージ表示 |

### 分析実行
| 関数名 | 説明 |
|---|---|
| `runQuickAnalysis(automated)` | クイック分析。動作モードに応じた分析を実行 |
| `runAdvancedAnalysis(automated)` | 詳細分析。全キーワードに対してサジェスト付き分析を実行、前回比較あり |
| `runDemoAnalysis(settings)` | デモモード分析（サンプルデータ） |
| `runBasicAnalysis(settings)` | ベーシックモード分析（YouTube + 基本トレンド） |
| `performFullAnalysis(settings)` | アドバンスモード分析（全プラットフォーム + AI分析 + サジェスト + Keyword Planner） |

### データ収集
| 関数名 | 説明 |
|---|---|
| `fetchYouTubeData(settings)` | YouTube Data API v3でキーワード検索・動画詳細取得 |
| `processYouTubeVideo(video)` | YouTube動画データの加工（再生数、エンゲージメント率等） |
| `fetchChannelData(channelId, settings, apiKey)` | YouTubeチャンネル情報取得 |
| `fetchXData(settings)` | X API v2でツイート検索・トレンド取得・アカウント分析 |
| `processXTweet(tweet)` | Xツイートデータの加工 |
| `fetchKeywordPlannerData(settings)` | Google Ads Keyword Planner APIで検索ボリューム取得 |

### AI分析・サジェスト
| 関数名 | 説明 |
|---|---|
| `performAIAnalysis(data, settings)` | Gemini APIで総合分析を実行 |
| `generateContentSuggestions(data, settings)` | Gemini + Grok + データ分析でコンテンツサジェストを統合生成 |
| `performGeminiSuggestions(data, settings)` | Gemini APIでサジェスト生成 |
| `performGrokSuggestions(data, settings)` | Grok APIでサジェスト生成 |
| `rankSuggestions(suggestions, data)` | サジェストのスコアリング・ランキング |
| `createAnalysisPrompt(data, settings)` | AI分析用プロンプト生成 |
| `createSuggestionsPrompt(data, settings)` | サジェスト生成用プロンプト生成 |
| `parseSuggestions(text)` | LLMの出力テキストをサジェストオブジェクトにパース |

### スコアリング・分析
| 関数名 | 説明 |
|---|---|
| `calculateScore(data)` | 総合スコア算出（ベース50点 + YouTube/X/トレンド加点、0-100） |
| `analyzeSentiment(tweets)` | キーワードベースの簡易センチメント分析（ポジティブ/ネガティブ/ニュートラル） |
| `analyzeTrends(data, settings)` | YouTube投稿傾向・Xセンチメントからトレンド分析 |
| `extractKeywords(text)` | テキストからキーワード抽出（ストップワード除去、頻度分析） |

### レポート・通知
| 関数名 | 説明 |
|---|---|
| `generateHTMLReport(data)` | リッチなHTML形式レポート生成（グラデーションヘッダー、スコアカード、メトリクスグリッド） |
| `generateAndSendHTMLReport(result, settings)` | HTMLレポートをメール送信 + Google Driveに保存 |
| `sendNotification(result, settings)` | テキスト形式メール通知 |
| `saveHTMLReport(html)` | HTMLレポートをGoogle Driveに保存 |
| `generateDailyReport()` | 日次レポート生成（短期） |
| `generateWeeklyReport()` | 週次レポート生成（中期） |
| `generateMonthlyReport()` | 月次レポート生成（長期） |
| `generateCustomReport()` | カスタム期間レポート生成 |

### データ保存
| 関数名 | 説明 |
|---|---|
| `saveAnalysisResult(result)` | 分析結果をレポートシートに保存 |
| `saveYouTubeData(data, timestamp)` | YouTube詳細データをYouTube_データシートに保存 |
| `saveXData(data, timestamp)` | XデータをX_データシートに保存（最大20件） |
| `saveContentSuggestions(suggestions)` | サジェストをサジェストシートに保存 |
| `updateDashboard(result)` | ダッシュボードシートを最新の分析結果で更新 |
| `getPreviousAnalysisResult(keyword)` | 同キーワードの前回分析結果を取得（前回比較用） |

### トリガー・設定
| 関数名 | 説明 |
|---|---|
| `setupTriggers()` | 日次(9時)/週次(月曜10時)/月次(1日11時)トリガーを設定 |
| `loadSettings()` | 設定シートから全設定を読み込み |
| `resetSettings()` | 全設定とAPIキーをリセット |

### テスト
| 関数名 | 説明 |
|---|---|
| `testYouTubeAPI()` | YouTube API接続テスト |
| `testXAPI()` | X API接続テスト |
| `testGeminiAPI()` | Gemini API接続テスト |
| `testSuggestions()` | サジェスト機能テスト |

### ヘルプ
| 関数名 | 説明 |
|---|---|
| `showUserGuide()` | 使い方ガイドをモーダルダイアログで表示 |
| `showAPIKeyGuide()` | APIキー取得方法をモーダルダイアログで表示 |
| `showTroubleshooting()` | トラブルシューティングをモーダルダイアログで表示 |
| `showAdvancedSettings()` | 詳細設定ガイドをモーダルダイアログで表示 |
| `showReleaseNotes()` | 更新情報をモーダルダイアログで表示 |

## スコア算出ロジック

| 要素 | 計算方法 | 最大加点 |
|---|---|---|
| ベーススコア | 固定 | 50 |
| YouTube | 平均エンゲージメント率 x 10 | 25 |
| X | ポジティブ率 x 20 | 20 |
| トレンド | 成長率（-10 - +10で制限） | 10 |

## 動作モード

| モード | 必要なAPIキー | 利用可能機能 |
|---|---|---|
| demo | なし | サンプルデータでの動作確認 |
| basic | YouTube API Key | YouTube分析 + 基本トレンド |
| advanced | YouTube + X + Gemini + Grok + Google Ads | 全機能（AI分析、サジェスト、Keyword Planner含む） |

## 使用サービス・API
- YouTube Data API v3（検索、動画詳細、チャンネル情報）
- X (Twitter) API v2（ツイート検索、トレンド取得）
- Gemini API（gemini-1.5-pro-latest、AI分析・サジェスト）
- Grok API（grok-4、コンテンツサジェスト）
- Google Ads API v17 / Keyword Planner（検索ボリューム・競合分析）
- Google Sheets API（SpreadsheetApp）
- Gmail / MailApp（メール通知）
- Google Drive API（DriveApp、HTMLレポート保存）
- OAuth2 ライブラリ（Google Ads API認証）
- Google Apps Script トリガー（ScriptApp）

## セットアップ手順
1. Google Spreadsheetsで新規スプレッドシートを作成
2. Apps Scriptエディタを開き、Code.jsの内容を貼り付け
3. メニュー「分析システム」>「初回セットアップ」を実行
4. セットアップウィザードに従い設定を完了:
   - Step 1: スプレッドシート初期化（自動）
   - Step 2: キーワード入力
   - Step 3: 動作モード選択（demo/basic/advanced）
   - Step 4: APIキー入力（advancedモードの場合）
   - Step 5: 自動実行設定
5. 「テスト」メニューから各APIの接続テストを実行
6. 「クイック分析」を実行して動作確認

## スクリプトプロパティ
| プロパティ名 | 説明 |
|---|---|
| `OPERATION_MODE` | 動作モード（demo/basic/advanced） |
| `YOUTUBE_API_KEY` | YouTube Data API v3 のAPIキー |
| `X_BEARER_TOKEN` | X (Twitter) API v2 のBearer Token |
| `GEMINI_API_KEY` | Gemini API のAPIキー |
| `GROK_API_KEY` | Grok (xAI) APIキー |
| `GOOGLE_ADS_CLIENT_ID` | Google Ads API OAuth Client ID |
| `GOOGLE_ADS_CLIENT_SECRET` | Google Ads API OAuth Client Secret |
| `GOOGLE_ADS_DEVELOPER_TOKEN` | Google Ads Developer Token |
| `GOOGLE_ADS_CUSTOMER_ID` | Google Ads Customer ID（ハイフンなし） |
| `ADMIN_EMAIL` | エラー通知先の管理者メールアドレス |

## 自動実行スケジュール

| トリガー | 実行タイミング | レポートタイプ |
|---|---|---|
| 日次 | 毎日午前9時 | 短期（7日間） |
| 週次 | 毎週月曜日午前10時 | 中期（30日間） |
| 月次 | 毎月1日午前11時 | 長期（365日間） |

## 注意事項
- X API Bearer Tokenが未設定の場合、モックデータ（空データ）が返される
- YouTube APIは1日10,000クォータの無料枠あり
- Gemini APIは1分間60リクエストの無料枠あり
- Google Ads Keyword Planner APIはDeveloper Token承認に時間がかかる場合がある
- HTMLレポートはGoogle Driveのルートフォルダに保存される
- 重大なAPIエラー発生時はADMIN_EMAIL宛にエラー通知が送信される
- appsscript.jsonでYouTube Advanced Serviceが有効化されている
