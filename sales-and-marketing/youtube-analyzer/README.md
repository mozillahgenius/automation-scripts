# youtube-analyzer / YouTube動画分析・管理

## 概要
YouTube Data API v3・Analytics APIを活用した動画メタデータ管理・分析ツール。ブランドアカウント対応のOAuth2認証を備え、動画情報の一括取得・分析結果のスプレッドシート出力を行う。

## 技術構成
| 項目 | 詳細 |
|------|------|
| Platform | Google Apps Script |
| Language | JavaScript |
| APIs | YouTube Data API v3, YouTube Analytics API |
| データ管理 | Google Sheets |
| 認証 | OAuth2（ブランドアカウント対応） |

## ファイル構成
- `YouTubeAnalyze.js` - メインスクリプト（動画メタデータ取得・分析）
- `YouTubeAnalyze_統合版.js` - 統合版スクリプト（全機能を統合）
- `BrandChannelFix.js` - ブランドチャンネル対応スクリプト
- `OAuth2Auth.js` - OAuth2認証処理
- `DataAPIOnly` - Data API単体利用版

## セットアップ
1. Google Apps Scriptプロジェクトを作成
2. 各 `.js` ファイルをスクリプトエディタに追加
3. Google Cloud ConsoleでAPIを有効化:
   - YouTube Data API v3
   - YouTube Analytics API
4. OAuth2クライアントIDを作成し、`OAuth2Auth.js` に設定
5. ブランドアカウントを使用する場合は `BrandChannelFix.js` の設定を確認
6. 分析結果出力用のスプレッドシートを作成し、シートIDを設定

## 使い方
- **動画情報取得**: チャンネルの全動画メタデータ（タイトル、再生回数、いいね数等）を一括取得
- **分析**: 動画パフォーマンスの分析結果をスプレッドシートに出力
- **ブランドアカウント**: `BrandChannelFix.js` でブランドチャンネルの認証問題を解決
- **統合版**: `YouTubeAnalyze_統合版.js` で全機能をまとめて利用可能
