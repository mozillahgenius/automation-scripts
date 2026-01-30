# warehouse-management / 倉庫管理システム

## 概要
倉庫管理の基幹システム（WAREHOUSE_SYS）を中心とした連携仕様書群。HubSpot、CONTRACT_SVC、BILLING_SERVICE、APPROVAL_SVCなど外部サービスとの統合を設計し、契約から請求までの業務フローにおけるシステム責務を定義する。AS-IS/TO-BEの比較分析を含む。

## 技術構成
| 項目 | 詳細 |
|------|------|
| 種別 | システム設計書 / API仕様書 |
| 基幹システム | WAREHOUSE_SYS（倉庫管理） |
| 連携サービス | HubSpot, CONTRACT_SVC, BILLING_SERVICE, APPROVAL_SVC |
| 設計手法 | AS-IS/TO-BE分析、システム責務定義 |
| ドキュメント形式 | Markdown |

## ファイル構成
- `システム責務定義_核システム中心.md` - WAREHOUSE_SYSを核としたシステム責務の定義書
- `契約から請求_連携仕様_AS-IS_TO-BE.md` - 契約〜請求フローの現状と理想の連携仕様
- `契約から請求_連携仕様_TO-BE_WAREHOUSE_SYS責務版.md` - WAREHOUSE_SYS責務を明確化したTO-BE連携仕様
- `BILLING_SERVICE_請求APIマッピング.md` - BILLING_SERVICE請求APIとのデータマッピング定義
- `作業マスタ_設計.md` - 作業マスタのデータ設計書
- `倉庫管理システム` - 倉庫管理システム関連の追加資料
- `請求業務の改善_WAREHOUSE_SYSシステムへの要望_pptx_-_Google_スライド.png` - 請求業務改善要望のスライド画像

## 使い方
- **全体像の把握**: `システム責務定義_核システム中心.md` からWAREHOUSE_SYSと各外部サービスの責務分担を確認
- **連携仕様の確認**: 契約〜請求の連携仕様書でAS-IS/TO-BEのフローを比較
- **API連携**: `BILLING_SERVICE_請求APIマッピング.md` で具体的なデータマッピングを参照
- **マスタ設計**: `作業マスタ_設計.md` でデータ構造を確認
