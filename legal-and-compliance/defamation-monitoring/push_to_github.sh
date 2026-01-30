#!/bin/bash

echo "GitHubへのプッシュを開始します..."
echo ""
echo "Personal Access Token (ghp_で始まる)を入力してください："
read -s TOKEN
echo ""

# リモートリポジトリを追加
git remote add origin https://github.com/mozillahgenius/defamation-flow.git 2>/dev/null || git remote set-url origin https://github.com/mozillahgenius/defamation-flow.git

# トークンを使ってpush
echo "プッシュ中..."
if git push https://mozillahgenius:${TOKEN}@github.com/mozillahgenius/defamation-flow.git main; then
    echo ""
    echo "✅ 成功！リポジトリはこちら："
    echo "https://github.com/mozillahgenius/defamation-flow"
    
    # 今後の認証を簡単にする
    git remote set-url origin https://mozillahgenius:${TOKEN}@github.com/mozillahgenius/defamation-flow.git
    echo ""
    echo "今後は 'git push' だけでプッシュできます"
else
    echo ""
    echo "❌ エラーが発生しました"
    echo "1. GitHubで 'defamation-flow' リポジトリを作成しましたか？"
    echo "2. Personal Access Tokenは正しいですか？"
    echo "3. リポジトリ名は正しいですか？"
fi