# 批判的記事監視システム - 使用方法

## 概要
Perplexity Sonar APIを使用して、指定した会社に関する批判的な記事を検索し、スプレッドシートに記録してメールレポートを送信するGoogle Apps Script (GAS)です。

## セットアップ

### 1. Perplexity APIキーの取得
1. [Perplexity AI](https://www.perplexity.ai/)でアカウントを作成
2. APIキーを取得
3. GASのスクリプトプロパティに設定（推奨）または、コード内の`CONFIG.PERPLEXITY_API_KEY`に直接入力

### 2. スクリプトプロパティの設定（推奨）
```javascript
// GASエディタから設定
// プロジェクトの設定 > スクリプトプロパティ
// キー: PERPLEXITY_API_KEY
// 値: your-api-key-here
```

### 3. スプレッドシートの準備
- 新規または既存のGoogleスプレッドシートを開く
- 拡張機能 > Apps Script を選択
- `CriticalNewsMonitor.gs`のコードを貼り付けて保存

## 主要機能

### 1. システム全体のセットアップ
システムを使用する前に、まず全体をセットアップします：

```javascript
// 完全なシステムセットアップ（記事保存＋検索設定シート）
setupCompleteSystem();

// 既存設定をリセットして新規セットアップ
setupCompleteSystem(true);

// 検索設定シートのみセットアップ
setupConfigSheet();

// 記事保存シートのみセットアップ
setupSpreadsheet();

// システム全体の診断
diagnoseCompleteSetup();
```

**セットアップ内容:**
- **記事保存シート**: 美しいヘッダー書式、条件付き書式、フィルター機能
- **検索設定シート**: 企業監視設定の管理、データ検証、サンプル設定

### 2. 検索設定の管理
スプレッドシートで企業監視設定を簡単に管理できます：

**検索設定シートの項目:**
- **企業名**: 監視対象の企業名
- **キーワード**: 検索キーワード（カンマ区切り）
- **検索期間(日)**: 過去何日分を検索するか
- **批判的記事のみ**: TRUE/FALSE
- **メール送信**: TRUE/FALSE  
- **送信先メール**: レポート送信先
- **実行頻度**: 毎日/週次/月次/手動
- **有効/無効**: 設定の有効化/無効化
- **最終実行日時**: 自動更新
- **備考**: メモ欄

**使用例:**
```javascript
// 設定に基づく検索実行
executeSearchByConfig('毎日');     // 毎日設定のみ実行
executeSearchByConfig('週次');     // 週次設定のみ実行  
executeSearchByConfig('すべて');   // 全有効設定を実行

// 設定の読み取り
const configs = getSearchConfigs();        // 有効な設定のみ
const allConfigs = getSearchConfigs(false); // 全設定

// 設定の検証
validateSearchConfigs();
```

### 3. 批判的記事の検索（基本）
```javascript
// 基本的な使用方法
searchCriticalArticles('企業名', 7);  // 過去7日間の批判的記事を検索

// 詳細な使用方法
searchCriticalArticles(
  '企業名',           // 必須: 検索対象の会社名
  30,                 // 過去何日分を検索（デフォルト: 7日）
  ['不祥事', '訴訟'], // キーワード指定（オプション）
  'report@example.com', // メール送信先（オプション）
  true                // メール送信するか（デフォルト: true）
);
```

### 4. キーワードベースの検索
```javascript
// キーワードで記事を検索（批判的でない記事も含む）
searchArticlesByKeywords(
  '企業名',                  // 必須: 会社名
  ['新製品', 'イノベーション'], // 必須: 検索キーワード
  14,                        // 過去何日分（デフォルト: 7日）
  false,                     // 批判的記事のみか（デフォルト: false）
  'report@example.com',      // メール送信先（オプション）
  true                       // メール送信するか（デフォルト: true）
);
```

### 5. 定期実行の設定

#### 毎日実行（トリガー設定）
```javascript
// dailyCheck関数を毎日実行するよう設定
// GASエディタ: 編集 > 現在のプロジェクトのトリガー > トリガーを追加
// - 実行する関数: dailyCheck
// - イベントのソース: 時間主導型
// - 時間ベースのトリガータイプ: 日付ベースのタイマー
// - 時刻: 任意の時刻（例: 午前9-10時）
```

**重要**: 現在は設定シートベースの実行に変更されています。監視対象企業は「検索設定」シートで管理してください。

新しい定期実行の設定方法:
```javascript
// 設定シートの「実行頻度」列を「毎日」に設定した企業が自動実行されます
function dailyCheck() {
  executeSearchByConfig('毎日');
}

// 週次実行の場合
function weeklyReport() {
  executeSearchByConfig('週次');  
}

// 月次実行の場合  
function monthlyReport() {
  executeSearchByConfig('月次');
}
```

#### 週次レポート
```javascript
function weeklyReport() {
  const companies = ['企業A', '企業B'];
  const daysBack = 7;
  
  for (const company of companies) {
    searchCriticalArticles(company, daysBack, [], 'weekly@example.com');
    Utilities.sleep(2000);
  }
}
```

## メール送信機能

### メールレポートの内容
- HTML形式の見やすい表形式レポート
- 以下の情報を含む:
  - 日付
  - 件名（記事タイトル）
  - 内容の要約
  - 記者名/執筆者名
  - 記事URL（クリック可能なリンク）
  - 批判度スコア（1-10）

### メール送信先の設定方法

#### 方法1: 関数の引数で指定
```javascript
searchCriticalArticles('企業名', 7, [], 'custom@example.com');
```

#### 方法2: デフォルト設定を変更
```javascript
// 関数内でデフォルトのメールアドレスを設定
const toEmail = recipient || 'default@example.com';
```

## スプレッドシートのデータ構造

| 列 | 内容 |
|---|---|
| A | 検索日時 |
| B | 対象企業 |
| C | 記事日付 |
| D | 件名 |
| E | 内容要約 |
| F | 記者名/執筆者 |
| G | URL |
| H | 媒体名 |
| I | 批判度スコア |

## カスタマイズ

### 検索パラメータの調整
```javascript
const CONFIG = {
  DEFAULT_DAYS_BACK: 7,        // デフォルトの検索期間
  MAX_RESULTS_PER_SEARCH: 10,  // 1回の検索で取得する最大記事数
  SHEET_NAME: '批判的記事'     // シート名
};
```

### 批判度スコアの閾値変更
```javascript
// filterCriticalArticles関数内
if (article.criticalScore && article.criticalScore < 5) {
  // 5以上を批判的とする（変更可能）
  return false;
}
```

## トラブルシューティング

### よくあるエラーと対処法

1. **APIキーエラー**
   - エラー: `Perplexity APIキーが設定されていません`
   - 対処: スクリプトプロパティまたはCONFIGにAPIキーを設定

2. **API制限エラー**
   - エラー: `API Error: Rate limit exceeded`
   - 対処: `Utilities.sleep(2000)`で待機時間を増やす

3. **メール送信エラー**
   - エラー: `Service invoked too many times`
   - 対処: 1日のメール送信数を減らすか、バッチ処理に変更

4. **スプレッドシートエラー**
   - エラー: `Cannot read property 'getRange' of null`
   - 対処: シート名が正しいか確認

## テスト方法

### 基本的なテスト
```javascript
// API接続テスト
testApiConnection();

// 検索機能テスト
testSearch();
```

### カスタムテスト
```javascript
function customTest() {
  // 特定の企業で少ない日数でテスト
  const results = searchCriticalArticles('テスト企業', 3, ['テスト'], null, false);
  console.log('検索結果:', results);
}
```

## 注意事項

1. **API利用制限**
   - Perplexity APIの利用制限に注意
   - 大量の検索を行う場合は適切な待機時間を設定

2. **プライバシー**
   - 検索結果には機密情報が含まれる可能性があるため、適切なアクセス制限を設定

3. **コスト**
   - Perplexity APIの使用料金を確認
   - 不要な検索を避けるため、トリガーの設定は慎重に

## サポート

問題が発生した場合は、以下を確認してください:
1. GASの実行ログ（表示 > ログ）
2. Perplexity APIのドキュメント
3. Google Apps Scriptのドキュメント