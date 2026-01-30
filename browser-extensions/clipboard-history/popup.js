// DOM要素
const historyList = document.getElementById('historyList');
const clearAllBtn = document.getElementById('clearAll');

// 初期化
document.addEventListener('DOMContentLoaded', () => {
  loadHistory();

  // クリアボタンのイベントリスナー
  clearAllBtn.addEventListener('click', clearAllHistory);
});

// 履歴を読み込んで表示
async function loadHistory() {
  const response = await chrome.runtime.sendMessage({ action: 'getHistory' });
  const history = response.history || [];

  if (history.length === 0) {
    historyList.innerHTML = `
      <div class="empty-state">
        <p>まだクリップボード履歴がありません</p>
        <p class="hint">テキストをコピーすると、ここに表示されます</p>
      </div>
    `;
    return;
  }

  historyList.innerHTML = '';

  history.forEach((item, index) => {
    const itemElement = createHistoryItem(item, index);
    historyList.appendChild(itemElement);
  });
}

// 履歴アイテムのHTML要素を作成
function createHistoryItem(item, index) {
  const div = document.createElement('div');
  div.className = 'history-item';

  const contentDiv = document.createElement('div');
  contentDiv.className = 'item-content';

  const textDiv = document.createElement('div');
  textDiv.className = 'item-text';
  textDiv.textContent = item.text;

  const timeDiv = document.createElement('div');
  timeDiv.className = 'item-time';
  timeDiv.textContent = formatTime(item.timestamp);

  contentDiv.appendChild(textDiv);
  contentDiv.appendChild(timeDiv);

  const deleteBtn = document.createElement('button');
  deleteBtn.className = 'btn-delete';
  deleteBtn.title = '削除';
  deleteBtn.innerHTML = `
    <svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor">
      <path d="M5.5 5.5A.5.5 0 0 1 6 6v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5zm2.5 0a.5.5 0 0 1 .5.5v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5zm3 .5a.5.5 0 0 0-1 0v6a.5.5 0 0 0 1 0V6z"/>
      <path fill-rule="evenodd" d="M14.5 3a1 1 0 0 1-1 1H13v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V4h-.5a1 1 0 0 1-1-1V2a1 1 0 0 1 1-1H6a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1h3.5a1 1 0 0 1 1 1v1zM4.118 4 4 4.059V13a1 1 0 0 0 1 1h6a1 1 0 0 0 1-1V4.059L11.882 4H4.118zM2.5 3V2h11v1h-11z"/>
    </svg>
  `;

  // クリックでコピー
  contentDiv.addEventListener('click', () => {
    copyToClipboard(item.text);
  });

  // 削除ボタンのイベント
  deleteBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    removeHistoryItem(index);
  });

  div.appendChild(contentDiv);
  div.appendChild(deleteBtn);

  return div;
}

// クリップボードにコピーして貼り付け
async function copyToClipboard(text) {
  try {
    // background scriptに貼り付けメッセージを送信
    await chrome.runtime.sendMessage({
      action: 'pasteToActiveElement',
      text: text
    });

    // 視覚的なフィードバック
    showCopyFeedback();

    // 少し待ってからポップアップを閉じる
    setTimeout(() => {
      window.close();
    }, 300);
  } catch (error) {
    console.error('貼り付けに失敗しました:', error);
  }
}

// コピー成功のフィードバック
function showCopyFeedback() {
  const items = document.querySelectorAll('.history-item');
  items.forEach(item => {
    item.style.background = '#e8f5e9';
  });
}

// 特定の履歴アイテムを削除
async function removeHistoryItem(index) {
  await chrome.runtime.sendMessage({ action: 'removeItem', index });
  loadHistory();
}

// 履歴を全てクリア
async function clearAllHistory() {
  if (confirm('クリップボード履歴を全て削除しますか？')) {
    await chrome.runtime.sendMessage({ action: 'clearHistory' });
    loadHistory();
  }
}

// タイムスタンプをフォーマット
function formatTime(timestamp) {
  const now = Date.now();
  const diff = now - timestamp;

  const seconds = Math.floor(diff / 1000);
  const minutes = Math.floor(seconds / 60);
  const hours = Math.floor(minutes / 60);
  const days = Math.floor(hours / 24);

  if (seconds < 60) {
    return 'たった今';
  } else if (minutes < 60) {
    return `${minutes}分前`;
  } else if (hours < 24) {
    return `${hours}時間前`;
  } else if (days < 7) {
    return `${days}日前`;
  } else {
    const date = new Date(timestamp);
    return `${date.getMonth() + 1}/${date.getDate()}`;
  }
}
