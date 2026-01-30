# drive-rag-auto-reply / Drive 参照 RAG 自動回答システム

## 概要

Google Drive 内の `inbox_questionnaires/` フォルダを監視し、新規の質問票を検出すると、過去の QA データから類似質問を Jaccard 類似度で検索し、OpenAI API を使用して回答を自動生成するシステム。生成された回答は Gmail の下書きとして保存される。

## 技術構成

| 項目 | 詳細 |
|------|------|
| Platform | Google Apps Script + OpenAI API |
| APIs | Google Drive API, Gmail API, OpenAI Chat Completions API |
| Triggers | 時間主導型（定期実行）/ 手動実行 |

## ファイル構成

- `Code.gs` - 全処理（フォルダ監視、類似検索、OpenAI 連携、下書き生成）

## セットアップ

1. Google Apps Script プロジェクトを作成し、`Code.gs` をコピー
2. OpenAI API キーをスクリプトプロパティに設定
3. Google Drive に `inbox_questionnaires/` フォルダを作成
4. 過去の QA データを所定の形式で Drive に保存
5. 時間主導型トリガーを設定して定期監視を有効化

## 使い方

- `inbox_questionnaires/` フォルダに質問票ファイルを配置
- トリガーにより自動検出され、過去の QA データから類似質問を検索
- Jaccard 類似度に基づくマッチング結果をコンテキストとして OpenAI に送信
- 生成された回答が Gmail の下書きとして保存される
- 下書きを確認・編集後、手動で送信する
