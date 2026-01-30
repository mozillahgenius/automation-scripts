# SlideSpeak API セットアップガイド

## 1. SlideSpeak APIキーの取得

### ステップ1: SlideSpeakアカウントの作成
1. [SlideSpeak公式サイト](https://slidespeak.co/)にアクセス
2. 「Sign Up」または「Get Started」をクリック
3. メールアドレスとパスワードでアカウントを作成

### ステップ2: APIキーの取得
1. ダッシュボードにログイン
2. 「Settings」または「API」セクションに移動
3. 「Generate API Key」または「Create New API Key」をクリック
4. 生成されたAPIキーをコピー

### ステップ3: Google Apps ScriptにAPIキーを設定
1. Google Apps Scriptエディタを開く
2. `Insta・X自動作成.js`ファイルの243行目を編集：
```javascript
SLIDESPEAK_API_KEY: 'YOUR_ACTUAL_API_KEY_HERE', // ここに取得したAPIキーを貼り付け
```

## 2. APIプランの確認

SlideSpeak APIには通常以下のプランがあります：

- **Free Plan**: 月間10-50枚の画像生成
- **Starter Plan**: 月間500枚の画像生成
- **Pro Plan**: 月間5000枚の画像生成
- **Enterprise Plan**: 無制限

1日3回×4枚 = 12枚/日 = 約360枚/月の使用量となるため、最低でもStarter Plan以上が必要です。

## 3. API制限事項

- **レート制限**: 通常1分間に10-60リクエスト
- **画像サイズ**: 最大解像度は通常1920x1920px
- **処理時間**: 1枚あたり2-10秒

## 4. トラブルシューティング

### エラー: "Unauthorized" または "Invalid API Key"
- APIキーが正しくコピーされているか確認
- APIキーの前後に余分なスペースがないか確認
- APIキーがアクティブか確認（ダッシュボードで確認）

### エラー: "Rate limit exceeded"
- API呼び出しの間隔を空ける（1秒以上）
- プランのアップグレードを検討

### エラー: "HTML response instead of JSON"
- APIエンドポイントURLが正しいか確認
- APIのバージョンが最新か確認
- SlideSpeak側のメンテナンス状況を確認

## 5. 代替案

SlideSpeakが利用できない場合の代替サービス：

### 1. Bannerbear API
- URL: https://www.bannerbear.com/
- 特徴: テンプレートベースの画像生成
- 価格: 月$49から

### 2. Placid API
- URL: https://placid.app/
- 特徴: 動的画像生成、テンプレート管理
- 価格: 月$19から

### 3. Canva API (Connect API)
- URL: https://www.canva.com/developers/
- 特徴: Canvaのデザインツールと連携
- 価格: 要問い合わせ

### 4. HTML/CSS to Image API
- URL: https://htmlcsstoimage.com/
- 特徴: HTMLを画像に変換
- 価格: 月$19から

## 6. Google Apps Script設定の推奨事項

### PropertiesServiceを使用したAPIキーの保護

```javascript
// APIキーを安全に保存
function setApiKey() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('SLIDESPEAK_API_KEY', 'your-actual-api-key');
}

// APIキーを取得
function getApiKey() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty('SLIDESPEAK_API_KEY');
}
```

### 使用方法
1. `setApiKey()`関数を一度実行してAPIキーを保存
2. コード内でAPIキーを使用する際は`getApiKey()`を呼び出す

```javascript
const CONFIG = {
  SLIDESPEAK_API_KEY: getApiKey(), // PropertiesServiceから取得
  // ...
};
```

## 7. テスト手順

1. **単一画像生成テスト**
   - メニューから「🧪 テスト生成」→「テスト画像生成（1枚）」を実行

2. **API接続テスト**
   - `testSlideSpeakConnection()`関数を実行

3. **フル生成テスト**
   - メニューから「🧪 テスト生成」→「朝の生成テスト」を実行

## 8. お問い合わせ

SlideSpeak APIに関する質問：
- サポート: support@slidespeak.co
- ドキュメント: https://docs.slidespeak.co/api

このシステムに関する質問：
- 設定ファイル内のコメントを参照
- エラーログを確認（スプレッドシートの「生成ログ」シート）