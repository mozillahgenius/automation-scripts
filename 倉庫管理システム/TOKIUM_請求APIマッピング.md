タイトル: TOKIUM（MFKessai）請求データ マッピング指針 v0.1

前提: 正式API項目名はベンダ仕様に準拠（ここでは論理項目名で整理）。CSV連携時も同一の論理構造で作成する。

1) 送信トリガー/単位
- トリガー: 案件の月次締め承認後（または案件都度）
- 単位: 顧客×請求先×請求月（1請求書）。案件横断の合算はオプション（顧客の締ルールに従う）。

2) 請求ヘッダ（Invoice）
- customer_id: souco.customer_id（TOKIUMの顧客コード）
- invoice_number: souco採番（YYYYMM-案件ID-枝番など）
- issue_date: 発行日（承認日/当月末 いずれかポリシー）
- billing_period_from / to: 請求対象期間
- due_date: 支払期日（顧客与信/支払サイトから計算）
- bill_to_name/address/email: TOKIUM側マスター同期値（必要時は上書き禁止）
- currency: JPY
- tax_mode / tax_total / total_excl_tax / total_incl_tax: souco計算結果
- remarks: 案件名/倉庫名/締め情報/差異注記など
- department/project_code: 任意（必要なら案件ID）

3) 請求明細（InvoiceLine）
- line_no: 連番
- description: 項目名（作業マスタ.name + 期間/備考）
- quantity: 実績数量（按分後）
- unit: 単位
- unit_price: 上代単価
- amount_excl_tax: 行小計（丸め適用後）
- tax_rate / tax_code: 行別税率（非課税/軽減などに対応）
- note: 伝票番号/輸送内訳/倉庫実績照合ID など

4) 例外パターンの表現
- 手数料のみ: descriptionに「手数料（元金額:XXX）」を付記。amountはPERCENT計算。
- 輸送のみ: descriptionに便種/距離/重量等のサマリー。
- 倉庫実績ベース: descriptionに外部請求書のID/日付を明記。

5) マッピングルール
- ヘッダ顧客/請求先: 原則TOKIUMマスターを真実源。soucoはcustomer_idのみ送る。
- 税/端数: 行→請求書の順に丸め。TOKIUM側の自動税計算と二重計算にならないよう、金額を明示送信。
- 同一顧客の複数案件: ポリシー選択（A: 合算1通, B: 案件別複数通）。運用開始はBを推奨。
- 添付: 詳細明細CSVや差異報告PDFを添付（可能な場合）。

6) エラー/バリデーション
- customer_id未登録 → エラーファイル（ワークキュー）
- 行金額が負 → 訂正伝票扱い（マイナス行を許容するか要確認）
- 税率不一致 → 税コードマッピングテーブルで変換（設定漏れはブロック）

7) 受信（入金/ステータス）
- invoice_id/token: TOKIUMの発行IDをsoucoに保存
- status: 発行/送付/入金/遅延 など
- payment: 入金日/入金額/手数料をsoucoへ同期→消込

8) CSVフォールバック（TOKIUMのCSV仕様を前提に雛形）
ヘッダ:
- 顧客コード, 請求書番号, 発行日, 請求期間自, 請求期間至, 支払期日, 通貨, 税区分, 備考
明細:
- 行番号, 品目名, 数量, 単位, 単価, 金額(税抜), 税率, 備考

9) フィールドソース一覧（souco→TOKIUM）
- 顧客コード: customer_id
- 請求書番号: invoice.number（採番規則管理）
- 発行日/請求期間/期日: invoiceヘッダ
- 品目名/数量/単位/単価/金額/税率: invoice_lines（作業マスタ×実績計算結果）
- 備考: 案件名/倉庫名/外部実績IDなどを可読化

10) 未決定事項（要ベンダ確認）
- マイナス行/訂正伝票の取り扱い
- 請求書の部門/プロジェクト付与項目
- 添付ファイルAPIの可用性と制限
- 税計算の優先（送信金額優先 or 先方自動計算）

