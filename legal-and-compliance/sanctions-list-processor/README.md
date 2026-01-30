# sanctions-list-processor / 反社チェックリスト自動化

## 概要
制裁リスト・反社会的勢力リストのバッチ処理をAIで自動化するGoogle Apps Scriptツール。OpenAI・Perplexity・Grokを活用し、リスト照合・名寄せ・ローマ字変換を一括処理する。

## 技術構成
| 項目 | 詳細 |
|------|------|
| Platform | Google Apps Script |
| Language | TypeScript / JavaScript |
| AI APIs | OpenAI API, Perplexity API, Grok API |
| データ管理 | Google Sheets, Google Drive |
| 検索 | Google Custom Search API |
| UI | HTML Service（ファイル・フォルダピッカー） |

## ファイル構成
- `Code.gs` - メインスクリプト（API連携、リスト処理基盤）
- `Code2.gs` - 補助スクリプト（追加処理ロジック）
- `processAntiSocialList.gs` - 反社リスト処理（基本版）
- `processAntiSocialListWithAI.gs` - AI活用版リスト処理
- `processAntiSocialListModular.gs` - モジュラー版リスト処理
- `processAntiSocialListDynamic.gs` - 動的処理版
- `processAntiSocialListComplete.gs` - 完全版リスト処理
- `file-picker.html` - ファイル選択UI
- `folder-picker.html` - フォルダ選択UI
- `recent-files.html` - 最近のファイル選択UI
- `setup_instructions.md` - セットアップ手順
- `openai_setup_instructions.md` - OpenAI API設定手順
- `データ仕様まとめ（制裁リスト／暴力団排除リスト）.ini` - データ仕様書
- `ローマ字変換/katakana_to_romaji.bas` - カタカナ→ローマ字変換（VBA版）
- `ローマ字変換/katakana_to_romaji.ts` - カタカナ→ローマ字変換（TypeScript版）

## セットアップ
1. Google Apps Scriptプロジェクトを作成
2. 各 `.gs` ファイルをスクリプトエディタに追加
3. HTMLファイルをプロジェクトに追加
4. スクリプトプロパティにAPIキーを設定:
   - `OPENAI_API_KEY`
   - `PERPLEXITY_API_KEY`
   - `GROK_API_KEY`
   - `CUSTOM_SEARCH_API_KEY` / `CUSTOM_SEARCH_ENGINE_ID`
5. Google Drive・Sheetsへのアクセス権限を承認
6. 詳細は `setup_instructions.md` および `openai_setup_instructions.md` を参照

## 使い方
- スプレッドシートのメニューまたはサイドバーからリスト処理を実行
- ファイルピッカーで対象リストを選択し、AIによる自動照合・名寄せを実行
- 処理結果はスプレッドシートに出力される
