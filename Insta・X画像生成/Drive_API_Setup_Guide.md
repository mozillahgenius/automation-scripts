# Google Drive API 有効化ガイド

## Google Apps ScriptでDrive APIを有効にする手順

### 1. Google Apps Scriptエディタでプロジェクトを開く
1. Google Apps Scriptエディタで「Insta・X自動作成.js」のプロジェクトを開く

### 2. サービスを追加
1. 左側のメニューから「サービス」（＋マークがある項目）をクリック
2. 「サービスを追加」ダイアログが開く

### 3. Drive APIを追加
1. サービス一覧から「Drive API」を探す
2. 「Drive API」を選択
3. バージョンは「v2」を選択（推奨）
4. 識別子は「Drive」のまま（デフォルト）
5. 「追加」ボタンをクリック

### 4. 追加を確認
1. 左側のメニューの「サービス」セクションに「Drive」が追加されていることを確認
2. これでDrive APIが使用可能になります

## トラブルシューティング

### エラー: ReferenceError: Drive is not defined
このエラーが出る場合は、Drive APIが有効になっていません。上記の手順に従ってAPIを有効にしてください。

### エラー: TypeError: Drive.Files is undefined
Drive APIは有効になっているが、正しく初期化されていない可能性があります。
1. プロジェクトを保存
2. ページをリロード
3. 再度実行

### 代替方法（Drive APIが使えない場合）
Drive APIが使用できない場合、以下の制限付き機能で動作します：
- PPTXファイルの直接変換はできません
- 代わりにHTML→PDF→画像の変換を使用します
- 画像品質が若干低下する可能性があります

## 権限について

Drive APIを使用する際、以下の権限が必要になります：
- Google Driveへのアクセス権
- ファイルの作成・編集・削除権限
- Google Slidesへのアクセス権

初回実行時に権限の許可を求められるので、「許可」をクリックしてください。

## 注意事項

1. **プロジェクトごとの設定**
   - Drive APIの有効化はプロジェクトごとに必要です
   - 他のプロジェクトにコピーした場合は再度有効化が必要

2. **バージョンについて**
   - v2を推奨（安定性が高い）
   - v3も使用可能ですが、APIの呼び出し方法が異なります

3. **クォータ制限**
   - Google Drive APIには使用制限があります
   - 通常の使用では問題ありませんが、大量の変換を行う場合は注意

## 確認方法

Drive APIが正しく有効になっているか確認するには：

```javascript
function testDriveAPI() {
  try {
    if (typeof Drive !== 'undefined') {
      console.log('Drive API is enabled ✓');
      
      // テスト: ルートフォルダの情報を取得
      const rootFolder = Drive.Files.get('root');
      console.log('Root folder name:', rootFolder.title);
      
      return true;
    } else {
      console.log('Drive API is NOT enabled ✗');
      return false;
    }
  } catch (error) {
    console.error('Drive API test error:', error);
    return false;
  }
}
```

このテスト関数を実行して、「Drive API is enabled ✓」と表示されれば設定完了です。