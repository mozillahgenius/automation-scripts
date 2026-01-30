// DOM要素
const openShortcutsBtn = document.getElementById('openShortcuts');
const saveSettingsBtn = document.getElementById('saveSettings');
const clearAllHistoryBtn = document.getElementById('clearAllHistory');
const maxHistorySizeInput = document.getElementById('maxHistorySize');
const saveStatus = document.getElementById('saveStatus');
const historyCountSpan = document.getElementById('historyCount');

// 初期化
document.addEventListener('DOMContentLoaded', () => {
  loadSettings();
  loadHistoryCount();

  // イベントリスナー
  openShortcutsBtn.addEventListener('click', openShortcutsPage);
  saveSettingsBtn.addEventListener('click', saveSettings);
  clearAllHistoryBtn.addEventListener('click', clearAllHistory);
});

// 設定を読み込み
async function loadSettings() {
  const { maxHistorySize = 200 } = await chrome.storage.local.get('maxHistorySize');
  maxHistorySizeInput.value = maxHistorySize;
}

// 履歴件数を読み込み
async function loadHistoryCount() {
  const response = await chrome.runtime.sendMessage({ action: 'getHistory' });
  const count = response.history ? response.history.length : 0;
  historyCountSpan.textContent = `${count}件`;
}

// ショートカット設定ページを開く
function openShortcutsPage() {
  chrome.tabs.create({
    url: 'chrome://extensions/shortcuts'
  });
}

// 設定を保存
async function saveSettings() {
  const maxHistorySize = parseInt(maxHistorySizeInput.value);

  if (maxHistorySize < 10 || maxHistorySize > 500) {
    showSaveStatus('保存件数は10〜500の範囲で設定してください', false);
    return;
  }

  await chrome.storage.local.set({ maxHistorySize });

  // background scriptに設定変更を通知
  chrome.runtime.sendMessage({
    action: 'updateMaxHistorySize',
    maxHistorySize
  });

  showSaveStatus('設定を保存しました', true);
}

// 保存ステータスを表示
function showSaveStatus(message, isSuccess) {
  saveStatus.textContent = message;
  saveStatus.style.color = isSuccess ? '#34a853' : '#e53935';

  setTimeout(() => {
    saveStatus.textContent = '';
  }, 3000);
}

// 全履歴を削除
async function clearAllHistory() {
  if (!confirm('本当に全ての履歴を削除しますか？この操作は取り消せません。')) {
    return;
  }

  await chrome.runtime.sendMessage({ action: 'clearHistory' });
  loadHistoryCount();
  alert('全ての履歴を削除しました');
}
