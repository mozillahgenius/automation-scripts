// コピーイベントをリスン（より確実な検知）
document.addEventListener('copy', handleCopy);
document.addEventListener('cut', handleCut);

// コピー処理
function handleCopy(e) {
  try {
    // まずは選択テキスト（入力欄含む）を取得
    const selectedText = getSelectedTextFallback();
    if (selectedText) {
      chrome.runtime.sendMessage({ action: 'addToHistory', text: selectedText }).catch(() => {});
      return;
    }

    // 取得できない場合は少し待ってからクリップボードを読む
    setTimeout(() => {
      if (navigator.clipboard && navigator.clipboard.readText) {
        navigator.clipboard.readText().then(text => {
          const t = (text || '').trim();
          if (t) {
            chrome.runtime.sendMessage({ action: 'addToHistory', text: t }).catch(() => {});
          }
        }).catch(() => {});
      }
    }, 50);
  } catch (error) {
    console.error('Copy handler error:', error);
  }
}

// カット処理
function handleCut(e) {
  try {
    const selectedText = getSelectedTextFallback();
    if (selectedText) {
      chrome.runtime.sendMessage({ action: 'addToHistory', text: selectedText }).catch(() => {});
      return;
    }
    setTimeout(() => {
      if (navigator.clipboard && navigator.clipboard.readText) {
        navigator.clipboard.readText().then(text => {
          const t = (text || '').trim();
          if (t) {
            chrome.runtime.sendMessage({ action: 'addToHistory', text: t }).catch(() => {});
          }
        }).catch(() => {});
      }
    }, 50);
  } catch (error) {
    console.error('Cut handler error:', error);
  }
}

// 入力欄やコンテンツから選択テキストを安全に取得
function getSelectedTextFallback() {
  const active = document.activeElement;
  if (active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA')) {
    try {
      const start = active.selectionStart;
      const end = active.selectionEnd;
      if (typeof start === 'number' && typeof end === 'number' && end > start) {
        return String(active.value).slice(start, end).trim();
      }
    } catch (_) {}
  }
  const sel = window.getSelection && window.getSelection();
  if (sel) {
    const text = String(sel.toString() || '').trim();
    if (text) return text;
  }
  return '';
}

// background scriptからのメッセージをリスン
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'copyText') {
    copyTextToClipboard(request.text);
    sendResponse({ success: true });
  } else if (request.action === 'pasteText') {
    // アクティブな要素にテキストを貼り付け
    pasteTextToActiveElement(request.text);
    sendResponse({ success: true });
  }
});

// クリップボードにテキストをコピー
function copyTextToClipboard(text) {
  // Clipboard APIを優先的に使用
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).catch(err => {
      // フォールバック
      fallbackCopyToClipboard(text);
    });
  } else {
    fallbackCopyToClipboard(text);
  }
}

// フォールバック用のコピー処理
function fallbackCopyToClipboard(text) {
  const textarea = document.createElement('textarea');
  textarea.value = text;
  textarea.style.position = 'fixed';
  textarea.style.opacity = '0';
  document.body.appendChild(textarea);
  textarea.select();

  try {
    document.execCommand('copy');
  } catch (error) {
    console.error('コピーに失敗しました:', error);
  }

  document.body.removeChild(textarea);
}

// アクティブな要素にテキストを貼り付け
function pasteTextToActiveElement(text) {
  const activeElement = document.activeElement;

  // 入力可能な要素かチェック
  if (!activeElement) {
    return;
  }

  const tagName = activeElement.tagName.toLowerCase();
  const isEditable = activeElement.isContentEditable;
  const isInput = tagName === 'input' || tagName === 'textarea';

  if (isInput) {
    // input, textarea の場合
    const start = activeElement.selectionStart;
    const end = activeElement.selectionEnd;
    const currentValue = activeElement.value;

    // カーソル位置にテキストを挿入
    const newValue = currentValue.substring(0, start) + text + currentValue.substring(end);
    activeElement.value = newValue;

    // カーソル位置を調整
    const newPosition = start + text.length;
    activeElement.selectionStart = newPosition;
    activeElement.selectionEnd = newPosition;

    // inputイベントを発火（Reactなどのフレームワーク対応）
    activeElement.dispatchEvent(new Event('input', { bubbles: true }));
    activeElement.dispatchEvent(new Event('change', { bubbles: true }));
  } else if (isEditable) {
    // contentEditable の場合
    document.execCommand('insertText', false, text);
  } else {
    // 編集可能な要素でない場合はクリップボードにコピーのみ
    copyTextToClipboard(text);
  }
}
