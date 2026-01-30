# ポケモンカード管理システム セットアップガイド

## エラー解決方法

### Photos Library APIエラーの解決

現在、Google Photos Library APIが無効になっているため、以下の2つの方法から選択してください。

---

## 方法1: Google Photos APIを有効化する（推奨）

### 手順

1. **Google Cloud Consoleにアクセス**
   - エラーメッセージに表示されているURLをクリック:
   - https://console.developers.google.com/apis/api/photoslibrary.googleapis.com/overview?project=138255947511

2. **APIを有効化**
   - 「有効にする」ボタンをクリック
   - 数分待機（APIの有効化には時間がかかることがあります）

3. **OAuth同意画面の設定（必要な場合）**
   - 左メニューから「OAuth同意画面」を選択
   - アプリ名: ポケモンカード管理
   - ユーザーサポートメール: あなたのメールアドレス
   - 承認済みドメイン: 不要
   - 保存

4. **スクリプトを再実行**
   ```javascript
   // GASエディタで実行
   main()
   ```

---

## 方法2: Google Drive直接アップロード版を使用（簡単）

Google Photos APIを使わず、Google Driveに直接画像をアップロードする方式です。

### セットアップ手順

1. **初期セットアップを実行**
   ```javascript
   // GASエディタで実行
   initialDriveSetup()
   ```

2. **作成されたフォルダに画像をアップロード**
   - コンソールに表示されるフォルダURLにアクセス
   - カード画像をドラッグ&ドロップでアップロード

3. **処理を実行**
   ```javascript
   // 手動実行
   processImagesFromDrive()

   // または1時間ごとに自動実行されます
   ```

### Drive版のメリット
- API設定不要
- すぐに使い始められる
- PCからの直接アップロードが簡単

### Drive版のデメリット
- iPhoneから直接アップロードする場合は、Googleドライブアプリが必要
- Google Photosの自動バックアップ機能は使えない

---

## 通知メールアドレスの設定

エラー通知を受け取るには、以下を実行してください：

```javascript
// GASエディタで実行
PropertiesService.getScriptProperties().setProperty('NOTIFICATION_EMAIL', 'your-email@example.com');
```

---

## トラブルシューティング

### 権限エラーが出る場合

1. GASエディタで任意の関数を手動実行
2. 権限確認ダイアログで「許可」
3. 必要な権限をすべて承認

### APIキーが未設定の場合

```javascript
// 必要なAPIキーを設定
PropertiesService.getScriptProperties().setProperty('PERPLEXITY_API_KEY', 'your-api-key');
PropertiesService.getScriptProperties().setProperty('NOTION_API_KEY', 'your-notion-key');
PropertiesService.getScriptProperties().setProperty('NOTION_DATABASE_ID', 'your-database-id');
```

### 接続テスト

```javascript
// 各サービスへの接続をテスト
testConnection()
```

---

## どちらの方法を選ぶべきか？

### Google Photos APIを有効化すべき場合
- iPhoneから自動バックアップしたい
- 既にGoogle Photosを使っている
- 完全自動化したい

### Drive版を使うべき場合
- すぐに使い始めたい
- API設定が面倒
- PCから画像をアップロードすることが多い
- テスト的に使ってみたい

---

## サポート

問題が解決しない場合は、以下の情報を確認してください：

1. エラーログ（スプレッドシートの「エラーログ」シート）
2. 実行ログ（GASエディタの「実行」メニュー）
3. スクリプトプロパティの設定値

```javascript
// 設定値の確認
function checkSettings() {
  const props = PropertiesService.getScriptProperties().getProperties();
  console.log('現在の設定:', props);
}
```