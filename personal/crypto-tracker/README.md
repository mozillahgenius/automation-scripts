# crypto-tracker / 暗号資産管理

## 概要
暗号資産ポートフォリオの期待値を日次で算出・レポートするGoogle Apps Scriptツール。CoinGecko APIで価格を取得し、Perplexity（sonar-pro）およびGrok-4による市場分析を統合。ステーキング報酬の計算やブレイクアウト検知機能も搭載。

## 技術構成
| 項目 | 詳細 |
|------|------|
| Platform | Google Apps Script |
| 言語 | JavaScript (GAS) |
| APIs | CoinGecko API, Perplexity API (sonar-pro), Grok-4 API |
| データ管理 | Google Sheets |
| 機能 | 期待値算出、ステーキング計算、ブレイクアウト検知 |
| トリガー | 時間主導型（日次） |

## ファイル構成
- `暗号資産管理.js` - メイン処理（価格取得、期待値算出、AI分析、レポート生成）

## セットアップ
1. Google Apps Scriptプロジェクトを新規作成
2. 以下のAPIキーをスクリプトプロパティに設定:
   - CoinGecko APIキー
   - Perplexity APIキー（sonar-proモデル用）
   - Grok-4 APIキー
3. `暗号資産管理.js` のコードをエディタに貼り付け
4. ポートフォリオ情報（保有銘柄、数量、取得単価）をスプレッドシートに入力
5. 時間主導型トリガー（毎日実行）を設定

## 使い方
- **日次レポート**: トリガーにより毎日自動実行され、ポートフォリオの期待値レポートを生成
- **ステーキング管理**: ステーキング中の資産の報酬を自動計算
- **ブレイクアウト検知**: 急激な価格変動を検知し、アラート通知
- **AI分析**: Perplexity / Grok-4 による市場動向の分析コメントを付与
