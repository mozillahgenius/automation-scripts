# hubspot-status-reminder / HubSpot取引ステータス通知

## 概要
HubSpotの取引（Deal）ステータスを監視し、3日以上更新がない取引を検出してHTML形式のメールで通知するGoogle Apps Scriptツール。緊急度に応じた色分け表示により、対応が必要な案件を一目で把握できる。請求先情報と商品名の整合性チェック機能も含む。

## 技術構成
| 項目 | 詳細 |
|------|------|
| Platform | Google Apps Script |
| 言語 | JavaScript (GAS) |
| APIs | HubSpot API, Gmail, Google Sheets |
| トリガー | 時間主導型（日次） |
| 通知形式 | HTML メール（色分け緊急度表示） |

## ファイル構成
### 取引ステータス/
- `HubspotDealNotifier.gs` - 取引ステータス通知の基本版
- `HubspotDealNotifier_改良版.gs` - 改良版（機能拡張）
- `HubspotDealNotifier_Complete.gs` - 完成版（全機能統合）
- `Huspot Deal Status.xlsx` - 取引ステータスの参考データ
- `セットアップ手順書.md` - 導入手順書

### 請求先情報:商品名/
- `請求先情報チェッカー.gs` - 請求先情報と商品名の整合性チェック
- `appsscript.json` - GASプロジェクト設定

## セットアップ
1. Google Apps Scriptプロジェクトを新規作成
2. HubSpot APIキーを取得し、スクリプトプロパティに設定
3. `HubspotDealNotifier_Complete.gs` のコードをエディタに貼り付け
4. 通知先メールアドレスをスクリプト内で設定
5. 時間主導型トリガー（毎日実行）を設定

## 使い方
- トリガーにより毎日自動実行され、3日以上更新のない取引をメールで通知
- 緊急度が高い案件は赤色、中程度はオレンジ、低い場合は黄色で表示
- 請求先情報チェッカーは手動実行またはメニューから起動
