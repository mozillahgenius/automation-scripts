# license-management / ライセンス管理

## 概要
Google Workspace、Microsoft 365などの各種SaaSライセンスを一元管理するGoogle Apps Scriptツール。Admin SDK Directory APIを活用してライセンス割り当て状況を自動取得し、スプレッドシートで可視化。未使用ライセンスの検出やコスト最適化に活用できる。

## 技術構成
| 項目 | 詳細 |
|------|------|
| Platform | Google Apps Script |
| 言語 | JavaScript (GAS) |
| APIs | Admin SDK Directory API, Google Sheets, Gmail |
| 機能 | ライセンス棚卸し、利用状況可視化、通知 |
| トリガー | 時間主導型（月次/週次） |

## ファイル構成
- `ライセンス管理.gs` - ライセンス管理のメイン処理（取得・集計・通知）

## セットアップ
1. Google Apps Scriptプロジェクトを新規作成
2. Admin SDK（Directory API）の詳細サービスを有効化
3. `ライセンス管理.gs` のコードをエディタに貼り付け
4. Google Workspace管理者権限でOAuthスコープを承認
5. 出力先スプレッドシートのIDをスクリプト内に設定
6. 必要に応じて時間主導型トリガーを設定

## 使い方
- メイン関数を実行すると、各SaaSのライセンス割り当て状況をスプレッドシートに出力
- 未使用ライセンスや重複割り当てを検出し、一覧表示
- 定期実行により、ライセンス利用状況の推移を追跡可能
