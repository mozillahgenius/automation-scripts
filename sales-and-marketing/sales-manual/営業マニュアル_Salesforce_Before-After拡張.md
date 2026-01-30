文書名：営業マニュアル 拡張（Salesforce導入時のBefore/After比較付き）
版数：v1.0
作成日：2025-07-08

0. 目的と使い方
- 本書は既存の「営業マニュアル」の手順（SOP）を踏襲しつつ、Salesforce導入時に何が変わるかをBefore/Afterで示す拡張版です。
- 対象：営業、SE、SalesOps、管理部門。実運用の差分を素早く把握するために、各フェーズのSOP直後にBefore/Afterを併記します。
- 本書は参照資料であり、原本（営業マニュアル）には変更を加えません。

1. Day0：商談発生・案件登録・初動（SOP＋Before/After）
SOP（原本準拠、要点）
- [ ] CRMに案件登録（顧客/案件名/金額概算/確度/担当）
- [ ] 案件番号の採番と専用フォルダ作成確認
- [ ] 高難易度は社内キックオフ招集・SE工数仮押さえ
- [ ] 顧客へお礼・次回候補・ヒアリングシート送付

Before（現行想定）
- 案件登録：Excelや別CRMで手動、命名/採番ルールは担当者運用
- 初動SLA：個人タスク管理（メール/ToDo）で遅延が埋もれやすい
- 証跡：メール/メモが分散、活動履歴の網羅性にばらつき
- フォルダ：共有ドライブで手動作成、命名ルールが乱れやすい

After（Salesforce導入時）
- オブジェクト：Lead/Account/Contact/Opportunityを標準化。案件名は自動命名規則（例：{顧客}_{案件}_{YYYYMM}）をFlowで付与
- 必須項目/検証：必須・検証ルールで登録の完全性を担保、レコードタイプで商材別必須を出し分け
- 初動SLA：タスク自動生成＋エスカレーション（48h未活動で上長通知）。List View/ダッシュボードで見える化
- 活動記録：メール連携（Inbox/Einstein Activity Capture）で自動ログ、面談はEventとして登録
- ファイル：提案関連はFilesへ集約。外部ストレージ連携（SharePoint/Google）使用時はリンク項目で関連付け

2. Day1‑3：顧客理解・議事録・提案骨子（SOP＋Before/After）
SOP（要点）
- [ ] BANT‑CCで要件深掘り、MoMは1時間以内に共有
- [ ] 競合マトリクス、提案骨子（5枚）、概算見積（±30%）
- [ ] 勝率自己評価と社内共有

Before
- MoM：Word/Google Docで作成しメール配布、URLは各自管理
- 競合/調査：個人PC/フォルダに散在、参照/再利用が困難
- ToDo：個人リスト管理で漏れが発生

After（Salesforce）
- MoM：Notes/Filesに保管しOpportunityに関連付け。共有リンクで権限管理
- ToDo：次回アクションはTask化、期限と担当で追跡。未完了はダッシュボードで把握
- 競合：カスタム「Competition」オブジェクト（案件関連）で論点・差別化・勝敗要因を構造化
- 研究資料：Files/Quipを関連付け、バージョン管理を活用

3. Day4‑6：提案作成・REDレビュー・見積（SOP＋Before/After）
SOP（要点）
- [ ] REDレビュー（30分、想定問答10問）
- [ ] 提案ブラッシュアップ（ストーリー/根拠/体裁チェック）
- [ ] 粗利シミュレーション、承認ルール確認

Before
- RED：メール配布→会議、修正指示がメール散在
- 見積/粗利：Excelで作成、式崩れや版ズレリスク
- 提出前チェック：個人チェックで抜け漏れ

After（Salesforce）
- RED：Event＋Chatter/Quipでレビュー履歴を一元化。チェックリストはTaskテンプレで一括割当
- 見積：Quote/Quote Line Items＋Price Bookを使用（CPQは任意）。ディスカウント閾値でApproval Processを起動
- 粗利：カスタム項目で粗利率算出、25%未満は自動で経理レビュー必須にルーティング
- 提出前チェック：Validation/Checklist（カスタム）で必須ファイルや除外項目の完了を担保

4. Day7‑9：稟議・承認（SOP＋Before/After）
SOP（要点）
- [ ] 稟議起案（概要/利益率/リスク/添付）
- [ ] ルート自動判定、停滞フォロー、社外提出前に社内承認

Before
- 稟議：別システム/紙決裁。進捗トラッキングや証跡の突合が手作業

After（Salesforce）
- 承認：Approval Processで金額/粗利/割引に応じて多段承認。合議/リマインド/代理承認を設定
- 証跡：承認履歴・コメントを案件に集約、変更理由は必須入力
- 通知：Slack/メール連携で停滞を自動通知

5. Day10+：顧客提案・交渉・フォロー（SOP＋Before/After）
SOP（要点）
- [ ] プレゼン、Q&A、条件交渉、宿題整理、1週間以内に回答

Before
- 提出履歴・宿題管理がメール/Excelに分散

After（Salesforce）
- 提出：Filesに提出版を格納。提出日/有効期限/改訂履歴を項目管理
- 宿題：Task/ケースで期限/担当を可視化。次回ステップはOpportunityのNext Stepに記載
- クロージング：Close Plan（カスタム）で意思決定プロセスを可視化

6. 受注/契約・引継ぎ（SOP＋Before/After）
SOP（要点）
- [ ] 受注登録、契約書差分レビュー、社内速報、計上/請求連携、引継ぎKO設定

Before
- 契約PDFがメール/フォルダで錯綜、版ズレリスク
- 引継ぎ：Excelチェックリストで手渡し

After（Salesforce）
- 契約：Contract（必要に応じOrder）を使用、電子署名（DocuSign/Adobe Sign）連携で締結履歴を保持
- 引継ぎ：Closed Wonをトリガに「Handoff」カスタムレコードとChecklistを自動生成、PM/TSにオーナー割当
- 会計連携：受注/請求データを連携（中間はIntegration/連携項目で受領確認）

7. 知識化・ナレッジ（SOP＋Before/After）
SOP（要点）
- [ ] 成功事例、改善点、横展開の登録

Before
- 成功事例が共有ドライブや社内SNSに散在、検索性が低い

After（Salesforce）
- ナレッジ：Salesforce Knowledge（または「Success Story」カスタム）に構造化登録、タグで検索性を確保
- レビュー：上長承認後に全社公開、改訂は版管理

8. レポート/ダッシュボード（導入効果の可視化例）
- 初動SLA遵守率（活動なし48h）
- REDレビュー実施率と要改善件数
- 粗利レビュー実施率・値引き閾値逸脱件数
- 稟議リードタイム（承認所要日数）
- Handoff完了率（チェックリスト100%）
- ナレッジ登録率（成約30日以内）

9. 最小構成：設定/自動化（Salesforce）
- 標準：Lead, Account, Contact, Opportunity, Activity, Notes/Files, Quote, Product/Price Book, Contract, Order（任意）, Case（任意）, Knowledge（任意）
- カスタム例：Competition, RED Review, Handoff Checklist, Close Plan, 例外申請/是正計画（必要時）
- 自動化：
  - Flow（命名規則、SLAタスク、Closed Won→Handoff生成）
  - Validation（必須/除外チェック）
  - Approval Process（値引き/粗利/金額）
  - Assignment/Escalation Rules（リード/タスク）
- 権限：プロファイル＋権限セットでSoDを実現（起案/承認/編集の分離）

10. 項目マッピング（抜粋）
- 顧客名 → Account.Name
- 担当者 → Owner / Opportunity Owner
- 案件名 → Opportunity.Name（命名規則適用）
- 金額概算 → Opportunity.Amount（精度区分：概算/正式）
- 確度 → Opportunity.Stage / Probability
- 次アクション → Task.Subject/Due Date / Opportunity.Next Step
- 議事録URL → Notes/Files（関連）
- 競合情報 → Competition（カスタム）
- 粗利率 → Opportunity（カスタム項目）/Quote（カスタム）

11. 運用メモ（現場Tips）
- メールは必ずSFから送るかBcc to Salesforceで自動ログ化
- 面談はEventで登録し、参加者をContact/Leadで紐付け
- 提案の版管理はFilesで行い、提出版にタグ「外部提出」を付与
- 「未活動一覧」リストビューをホームに固定しSLA漏れをゼロにする

（終）
