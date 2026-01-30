/**
 * Slack チャンネル自動作成 Google Apps Script
 *
 * スプレッドシートのE列に値が追加されたときに、
 * Slackチャンネルを自動作成します。
 * 日本語チャンネル名にも対応しています。
 */

// ========================================
// 設定
// ========================================

// Slack Bot Token（xoxb-で始まるトークン）
// スクリプトプロパティから取得
function getSlackToken() {
  return PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
}

// 対象のシート名（必要に応じて変更）
const TARGET_SHEET_NAME = 'チャンネル名';

// 対象の列（E列 = 5）
const TARGET_COLUMN = 5;

// ステータス列（F列 = 6）- 作成結果を記録
const STATUS_COLUMN = 6;

// ========================================
// メイン処理
// ========================================

/**
 * スプレッドシート変更時のトリガー関数（onChange用）
 * スクリプトによる変更も検知可能
 * @param {Object} e - イベントオブジェクト
 */
function onSpreadsheetChange(e) {
  // 変更タイプがEDITの場合のみ処理
  if (e.changeType !== 'EDIT') {
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) {
    console.log('対象シートが見つかりません: ' + TARGET_SHEET_NAME);
    return;
  }

  // E列のデータを取得して処理
  processNewChannels(sheet);
}

/**
 * E列の新規チャンネル名を処理
 * F列が空の行のみ処理対象
 * @param {Sheet} sheet - シートオブジェクト
 */
function processNewChannels(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return; // ヘッダー行のみの場合はスキップ
  }

  // E列とF列のデータを取得（2行目から）
  const dataRange = sheet.getRange(2, TARGET_COLUMN, lastRow - 1, 2);
  const data = dataRange.getValues();

  for (let i = 0; i < data.length; i++) {
    const channelName = data[i][0];
    const status = data[i][1];
    const row = i + 2; // 実際の行番号（1始まり、ヘッダー除く）

    // チャンネル名があり、ステータスが空の場合のみ処理
    if (channelName && String(channelName).trim() !== '' && !status) {
      console.log('新規データ検出: 行=' + row + ', チャンネル名=' + channelName);
      createSlackChannel(String(channelName).trim(), sheet, row);
    }
  }
}

/**
 * Slackチャンネルを作成
 * @param {string} originalName - 元のチャンネル名（日本語可）
 * @param {Sheet} sheet - シートオブジェクト
 * @param {number} row - 行番号
 */
function createSlackChannel(originalName, sheet, row) {
  const token = getSlackToken();

  if (!token) {
    sheet.getRange(row, STATUS_COLUMN).setValue('エラー: トークン未設定');
    console.error('トークンが設定されていません');
    return;
  }

  // チャンネル名を正規化（Slack仕様に合わせる）
  const normalizedName = normalizeChannelName(originalName);
  console.log('チャンネル名正規化: ' + originalName + ' -> ' + normalizedName);

  if (!normalizedName) {
    sheet.getRange(row, STATUS_COLUMN).setValue('エラー: 無効なチャンネル名');
    return;
  }

  // Slack API呼び出し
  const url = 'https://slack.com/api/conversations.create';
  const payload = {
    name: normalizedName,
    is_private: false
  };

  const options = {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json; charset=utf-8'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    console.log('Slack API呼び出し開始: ' + normalizedName);
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    console.log('Slack API応答: ' + responseText);

    const result = JSON.parse(responseText);

    if (result.ok) {
      const channelId = result.channel.id;
      const createdName = result.channel.name;
      sheet.getRange(row, STATUS_COLUMN).setValue('✓ 作成完了: ' + createdName);
      console.log('チャンネル作成成功: ' + createdName + ' (ID: ' + channelId + ')');
    } else {
      const errorMessage = getErrorMessage(result.error);
      sheet.getRange(row, STATUS_COLUMN).setValue('✗ エラー: ' + errorMessage);
      console.error('チャンネル作成失敗: ' + result.error + ' / ' + errorMessage);
    }
  } catch (error) {
    sheet.getRange(row, STATUS_COLUMN).setValue('✗ 通信エラー: ' + error.message);
    console.error('通信エラー: ' + error);
  }
}

/**
 * チャンネル名を正規化（Slack仕様に合わせる）
 * - 小文字に変換
 * - スペースをハイフンに変換
 * - 日本語はそのまま保持（Slackは日本語チャンネル名に対応）
 * - 使用できない文字を除去
 * - 最大80文字に制限
 *
 * @param {string} name - 元の名前
 * @returns {string} 正規化された名前
 */
function normalizeChannelName(name) {
  if (!name) return '';

  let normalized = name
    // 小文字に変換（英字のみ影響）
    .toLowerCase()
    // 全角スペースを半角に
    .replace(/　/g, ' ')
    // スペースをハイフンに
    .replace(/\s+/g, '-')
    // 連続するハイフンを1つに
    .replace(/-+/g, '-')
    // 先頭と末尾のハイフンを除去
    .replace(/^-+|-+$/g, '')
    // Slackで使用できない文字を除去（英数字、ハイフン、アンダースコア、日本語は許可）
    .replace(/[^\w\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FFF\u3400-\u4DBF-]/g, '');

  // 最大80文字に制限
  if (normalized.length > 80) {
    normalized = normalized.substring(0, 80);
  }

  return normalized;
}

/**
 * Slackエラーコードを日本語メッセージに変換
 * @param {string} error - エラーコード
 * @returns {string} 日本語エラーメッセージ
 */
function getErrorMessage(error) {
  const errorMessages = {
    'name_taken': 'このチャンネル名は既に使用されています',
    'invalid_name_required': 'チャンネル名が必要です',
    'invalid_name_punctuation': '使用できない文字が含まれています',
    'invalid_name_maxlength': 'チャンネル名が長すぎます（最大80文字）',
    'invalid_name_specials': '特殊文字は使用できません',
    'invalid_name': '無効なチャンネル名です',
    'no_channel': 'チャンネルが見つかりません',
    'restricted_action': 'この操作は制限されています',
    'method_not_supported_for_channel_type': 'このチャンネルタイプではサポートされていません',
    'missing_scope': '必要な権限がありません（channels:write）',
    'not_authed': '認証されていません',
    'invalid_auth': '無効な認証情報です',
    'token_revoked': 'トークンが無効化されています',
    'ratelimited': 'API制限に達しました。しばらく待ってから再試行してください'
  };

  return errorMessages[error] || error;
}

// ========================================
// セットアップ・テスト用関数
// ========================================

/**
 * Slack Bot Tokenを設定
 * 初回セットアップ時にスクリプトエディタから実行
 */
function setSlackToken() {
  const token = Browser.inputBox('Slack Bot Tokenを入力してください（xoxb-で始まるもの）');
  if (token && token !== 'cancel') {
    PropertiesService.getScriptProperties().setProperty('SLACK_BOT_TOKEN', token);
    Browser.msgBox('トークンを保存しました');
  }
}

/**
 * トークン設定確認
 */
function checkToken() {
  const token = getSlackToken();
  if (token) {
    Browser.msgBox('トークンは設定されています（先頭: ' + token.substring(0, 10) + '...）');
  } else {
    Browser.msgBox('トークンが設定されていません。setSlackToken()を実行してください。');
  }
}

/**
 * 手動でチャンネル作成をテスト
 */
function testCreateChannel() {
  const channelName = Browser.inputBox('作成するチャンネル名を入力してください');
  if (!channelName || channelName === 'cancel') {
    return;
  }

  const token = getSlackToken();
  if (!token) {
    Browser.msgBox('トークンが設定されていません');
    return;
  }

  const normalizedName = normalizeChannelName(channelName);
  Browser.msgBox('正規化後の名前: ' + normalizedName);

  const url = 'https://slack.com/api/conversations.create';
  const payload = {
    name: normalizedName,
    is_private: false
  };

  const options = {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json; charset=utf-8'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const result = JSON.parse(response.getContentText());

  if (result.ok) {
    Browser.msgBox('成功！チャンネル「' + result.channel.name + '」を作成しました');
  } else {
    Browser.msgBox('エラー: ' + getErrorMessage(result.error));
  }
}

/**
 * onSpreadsheetChangeトリガーを設定（重複防止付き）
 */
function createChangeTrigger() {
  // 既存のonSpreadsheetChangeトリガーを全て削除
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onSpreadsheetChange') {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    }
  });

  // 新しいトリガーを1つだけ作成
  ScriptApp.newTrigger('onSpreadsheetChange')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onChange()
    .create();

  if (deletedCount > 0) {
    Browser.msgBox('既存のトリガー' + deletedCount + '個を削除し、新しいトリガーを1つ設定しました');
  } else {
    Browser.msgBox('変更トリガーを設定しました');
  }
}

/**
 * 全てのトリガーを削除
 */
function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;

  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
    deletedCount++;
  });

  Browser.msgBox(deletedCount + '個のトリガーを削除しました');
}

/**
 * 現在のトリガー一覧を表示
 */
function listAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();

  if (triggers.length === 0) {
    Browser.msgBox('トリガーはありません');
    return;
  }

  let message = '現在のトリガー一覧:\n\n';
  triggers.forEach((trigger, index) => {
    message += (index + 1) + '. ' + trigger.getHandlerFunction() +
               ' (' + trigger.getEventType() + ')\n';
  });

  Browser.msgBox(message);
}

/**
 * 初期セットアップを一括実行
 * トークン設定とトリガー設定を順番に行う
 */
function initialSetup() {
  // 1. トークン設定
  const token = Browser.inputBox('【1/2】Slack Bot Tokenを入力してください（xoxb-で始まるもの）');
  if (!token || token === 'cancel') {
    Browser.msgBox('セットアップをキャンセルしました');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('SLACK_BOT_TOKEN', token);

  // 2. トリガー設定（重複削除込み）
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onSpreadsheetChange') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('onSpreadsheetChange')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onChange()
    .create();

  Browser.msgBox('【2/2】セットアップ完了！\n\nトークンを保存し、トリガーを設定しました。\nシート「' + TARGET_SHEET_NAME + '」のE列にチャンネル名が追加されると自動でチャンネルを作成します。');
}

/**
 * 手動で全ての未処理行を処理
 * テストやリカバリー用
 */
function processAllPending() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) {
    Browser.msgBox('シート「' + TARGET_SHEET_NAME + '」が見つかりません');
    return;
  }

  processNewChannels(sheet);
  Browser.msgBox('処理完了');
}
