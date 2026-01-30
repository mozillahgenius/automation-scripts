// クリップボード履歴の最大保存数
const MAX_HISTORY_SIZE = 200;

// 拡張機能のインストール時の初期化
chrome.runtime.onInstalled.addListener(() => {
  chrome.storage.local.set({ clipboardHistory: [] });
});

// 拡張機能アイコンクリック時の処理
chrome.action.onClicked.addListener(() => {
  // オプションページを開く
  chrome.runtime.openOptionsPage();
});

// キーボードショートカットの処理
chrome.commands.onCommand.addListener((command) => {
  if (command === 'open-clipboard-history') {
    // 履歴を小さなウィンドウで開く
    chrome.windows.create({
      url: 'popup.html',
      type: 'popup',
      width: 420,
      height: 600,
      top: 100,
      left: 100
    });
  }
});

// content scriptからのメッセージを受信
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'addToHistory') {
    addToHistory(request.text);
  } else if (request.action === 'getHistory') {
    getHistory().then(history => {
      sendResponse({ history });
    });
    return true; // 非同期レスポンスを有効化
  } else if (request.action === 'clearHistory') {
    clearHistory();
    sendResponse({ success: true });
  } else if (request.action === 'removeItem') {
    removeItem(request.index);
    sendResponse({ success: true });
  } else if (request.action === 'copyToClipboard') {
    copyToClipboard(request.text);
    sendResponse({ success: true });
  } else if (request.action === 'pasteToActiveElement') {
    pasteToActiveElement(request.text);
    sendResponse({ success: true });
  } else if (request.action === 'updateMaxHistorySize') {
    updateMaxHistorySize(request.maxHistorySize);
    sendResponse({ success: true });
  }
});

// 履歴に追加
async function addToHistory(text) {
  if (!text || text.trim() === '') return;

  const result = await chrome.storage.local.get(['clipboardHistory', 'maxHistorySize']);
  const clipboardHistory = result.clipboardHistory || [];
  const maxHistorySize = result.maxHistorySize || MAX_HISTORY_SIZE;

  // 重複を削除（既存のものを削除して先頭に追加）
  const filteredHistory = clipboardHistory.filter(item => item.text !== text);

  // 新しいアイテムを先頭に追加
  const newHistory = [
    {
      text,
      timestamp: Date.now()
    },
    ...filteredHistory
  ].slice(0, maxHistorySize);

  await chrome.storage.local.set({ clipboardHistory: newHistory });
}

// 履歴を取得
async function getHistory() {
  const { clipboardHistory = [] } = await chrome.storage.local.get('clipboardHistory');
  return clipboardHistory;
}

// 履歴をクリア
async function clearHistory() {
  await chrome.storage.local.set({ clipboardHistory: [] });
}

// 特定のアイテムを削除
async function removeItem(index) {
  const { clipboardHistory = [] } = await chrome.storage.local.get('clipboardHistory');
  clipboardHistory.splice(index, 1);
  await chrome.storage.local.set({ clipboardHistory });
}

// クリップボードにコピー
async function copyToClipboard(text) {
  // アクティブなタブにメッセージを送信してコピーを実行
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (tab) {
    chrome.tabs.sendMessage(tab.id, { action: 'copyText', text });
  }
}

// アクティブな要素にテキストを貼り付け
async function pasteToActiveElement(text) {
  // アクティブなタブにメッセージを送信して貼り付けを実行
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (tab) {
    chrome.tabs.sendMessage(tab.id, { action: 'pasteText', text });
  }
}

// 最大履歴サイズを更新
async function updateMaxHistorySize(newSize) {
  const { clipboardHistory = [] } = await chrome.storage.local.get('clipboardHistory');

  // 新しいサイズに合わせて履歴を切り詰める
  if (clipboardHistory.length > newSize) {
    const trimmedHistory = clipboardHistory.slice(0, newSize);
    await chrome.storage.local.set({ clipboardHistory: trimmedHistory });
  }
}
