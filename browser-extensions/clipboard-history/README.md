# クリップボード履歴マネージャー

macOS用Chrome拡張機能で、複数のコピー履歴を保持して管理できます。

## 機能

- ✂️ コピーしたテキストを自動的に履歴として保存
- 📋 最大200件の履歴を保持
- 🔍 履歴から選択して再度コピー可能
- 🗑️ 個別削除、または全削除が可能
- ⏰ 相対時刻表示（「5分前」など）
- 🎨 macOS風のモダンなUI

## インストール方法

1. Chrome拡張機能の管理画面を開く
   - Chromeのアドレスバーに `chrome://extensions/` と入力
   - または、メニュー > その他のツール > 拡張機能

2. 右上の「デベロッパーモード」をONにする

3. 「パッケージ化されていない拡張機能を読み込む」をクリック

4. このプロジェクトのフォルダ（`clipboard-history-extension`）を選択

## アイコンの作成（オプション）

拡張機能のアイコンは `icons` フォルダに以下のサイズで配置してください：
- icon16.png (16x16px)
- icon48.png (48x48px)
- icon128.png (128x128px)

アイコンがない場合、Chrome標準のアイコンが表示されますが、機能には影響ありません。

### 簡単なアイコン作成方法

macOSのプレビューアプリやオンラインツールを使って、以下のようなアイコンを作成できます：
- クリップボードのイラスト
- テキストの重なったイメージ
- シンプルな記号（📋など）

## 使い方

### 基本的な使い方

1. **コピー**: ウェブページ上で通常通りテキストをコピー（Cmd + C）すると、自動的に履歴に保存されます

2. **履歴を確認**: ツールバーの拡張機能アイコンをクリックしてポップアップを開きます

3. **再コピー**: 履歴リストからアイテムをクリックすると、そのテキストがクリップボードにコピーされます

4. **削除**: 各アイテムにマウスを乗せると表示される削除ボタンで個別削除が可能です

5. **全削除**: 右上のゴミ箱アイコンで履歴を全てクリアできます

## ファイル構成

```
clipboard-history-extension/
├── manifest.json         # 拡張機能の設定ファイル
├── background.js         # バックグラウンドで動作するスクリプト
├── content.js           # ウェブページに注入されるスクリプト
├── popup.html           # ポップアップUI
├── popup.css            # ポップアップのスタイル
├── popup.js             # ポップアップのロジック
├── icons/               # アイコンフォルダ（要作成）
│   ├── icon16.png
│   ├── icon48.png
│   └── icon128.png
└── README.md            # このファイル
```

## 技術仕様

- Manifest V3
- Chrome Storage API
- Content Scripts
- Background Service Worker

## 実装方法の詳細

この拡張機能は3つの主要コンポーネントで構成されています。

### 1. Background Script (background.js)

**役割**: クリップボード履歴の管理とメッセージハンドリング

**主な機能**:
- `addToHistory()`: コピーされたテキストを履歴に追加（重複は削除して先頭に配置）
- `getHistory()`: 保存された履歴を取得
- `clearHistory()`: 全履歴をクリア
- `removeItem()`: 指定したインデックスのアイテムを削除

**データ構造**:
```javascript
clipboardHistory: [
  {
    text: "コピーされたテキスト",
    timestamp: 1697158800000  // Date.now()
  },
  // ...最大200件
]
```

**ストレージ**: Chrome Storage API (`chrome.storage.local`) を使用してブラウザのローカルストレージに保存

### 2. Content Script (content.js)

**役割**: ウェブページ上のコピーイベントを検知

**実装詳細**:
- `copy` イベントリスナー: ユーザーがCmd+CまたはCtrl+Cでコピーした時に発火
- `cut` イベントリスナー: カット操作も履歴に保存
- `window.getSelection().toString()`: 選択されたテキストを取得
- `chrome.runtime.sendMessage()`: Background Scriptにテキストを送信

**注入スコープ**: manifest.jsonの `content_scripts` で `<all_urls>` に指定されているため、全てのウェブページで動作

### 3. Popup UI (popup.html/css/js)

**役割**: 履歴の表示とユーザーインタラクション

**popup.js の主な機能**:

#### 履歴の表示
```javascript
loadHistory()
  → chrome.runtime.sendMessage({ action: 'getHistory' })
  → 各アイテムを createHistoryItem() でDOM要素に変換
  → historyList に追加
```

#### クリップボードへの再コピー
```javascript
copyToClipboard(text)
  → navigator.clipboard.writeText(text)  // Clipboard API使用
  → 視覚的フィードバック表示
  → ポップアップを自動的に閉じる
```

#### 時刻のフォーマット
`formatTime()` 関数が相対時刻を計算：
- 1分未満: "たった今"
- 1時間未満: "N分前"
- 24時間未満: "N時間前"
- 7日未満: "N日前"
- それ以上: "月/日"

### アーキテクチャ図

```
┌─────────────────────┐
│   Web Page          │
│                     │
│  ┌───────────────┐  │
│  │ Content.js    │  │ ← コピーイベント検知
│  └───────┬───────┘  │
└──────────┼──────────┘
           │ sendMessage
           ▼
┌─────────────────────┐
│ Background.js       │
│ (Service Worker)    │
│                     │
│ • 履歴管理          │
│ • データ永続化      │
│ • メッセージ処理    │
└─────────┬───────────┘
          │ storage
          ▼
┌─────────────────────┐
│ Chrome Storage API  │
│ (ローカルストレージ)│
└─────────────────────┘
          ▲
          │ getHistory
┌─────────┴───────────┐
│ Popup UI            │
│                     │
│ • 履歴表示          │
│ • 再コピー機能      │
│ • 削除機能          │
└─────────────────────┘
```

### 通信フロー

#### コピー時のフロー
1. ユーザーがウェブページでテキストを選択してCmd+C
2. content.js の `copy` イベントリスナーが発火
3. `window.getSelection().toString()` で選択テキストを取得
4. `chrome.runtime.sendMessage()` で background.js に送信
5. background.js の `addToHistory()` が履歴に追加
6. Chrome Storage API でデータを永続化

#### 再コピー時のフロー
1. ユーザーが拡張機能アイコンをクリック
2. popup.js の `loadHistory()` が履歴を取得
3. ユーザーが履歴アイテムをクリック
4. `navigator.clipboard.writeText()` でクリップボードに書き込み
5. 視覚的フィードバック表示後、ポップアップを閉じる

### セキュリティ考慮事項

- **ローカル保存**: 履歴は全てローカルに保存され、外部サーバーへの送信は一切なし
- **権限の最小化**: 必要最小限の権限のみを manifest.json で要求
  - `storage`: ローカルストレージへのアクセス
  - `clipboardRead`: クリップボードからの読み取り
  - `clipboardWrite`: クリップボードへの書き込み
- **XSS対策**: popup.js では `textContent` を使用して XSS を防止

### パフォーマンス最適化

- **重複削除**: 同じテキストが再度コピーされた場合、古いものを削除して先頭に移動
- **上限管理**: 200件を超えた古い履歴は自動的に削除 (`slice(0, MAX_HISTORY_SIZE)`)
- **非同期処理**: Storage API の操作は全て async/await で非同期処理
- **効率的なDOM操作**: 履歴表示時は一度に全ての要素を生成してから追加

## カスタマイズ

### 履歴の保存数を変更

`background.js` の `MAX_HISTORY_SIZE` を変更してください：

```javascript
const MAX_HISTORY_SIZE = 200; // お好みの数に変更（デフォルト: 200）
```

### UIの色やスタイルを変更

`popup.css` を編集してお好みのスタイルにカスタマイズできます。

## 注意事項

- この拡張機能はテキストのみをサポートしています（画像やファイルは非対応）
- プライベートブラウジングモードでは動作が制限される場合があります
- 履歴はローカルストレージに保存され、外部に送信されることはありません

## ライセンス

MIT License

## 貢献

バグ報告や機能追加のリクエストは歓迎します！
