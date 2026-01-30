/**
 * Slack チャンネル自動作成 Google Apps Script（v3）
 *
 * 機能:
 * 「フォーム形式の回答」シートに新規データ追加時、
 * 「チャンネル名」シートにコピー＆採番 → 5秒待機 → Slackチャンネル自動作成
 * を一括で実行
 *
 * シート構造:
 * A列: 用途
 * B列: 表示名（日本語可）
 * C列: 用途コード
 * D列: チャンネル番号
 * E列: Slack作成用チャンネル名
 * F列: バリデーション
 * G列: zapier_action
 * H列: Slack作成ステータス（結果出力）
 *
 * ============================================
 * セットアップ手順:
 * 1. このファイルの内容を全てコピー
 * 2. Google Apps Script に貼り付け
 * 3. 既存のファイルは全て削除
 * 4. deleteAllTriggers() を実行して不要なトリガーを削除
 * 5. initialSetup() を実行（トークン設定＋トリガー設定）
 * ============================================
 *
 * 重要: onChangeトリガーのみを使用
 *       採番完了後5秒待機してからSlackチャンネルを作成
 */

// ========================================
// 設定
// ========================================

// Slack Bot Token（xoxb-で始まるトークン）
function getSlackToken() {
  return PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
}

// シート名
const SOURCE_SHEET_NAME = 'フォーム形式の回答';
const TARGET_SHEET_NAME = 'チャンネル名';

// 列番号の設定
const HEADER_ROWS = 1;
const COL_PURPOSE = 1;      // A列: 用途
const COL_NAME = 2;         // B列: 表示名（日本語可）
const COL_CODE = 3;         // C列: 用途コード
const COL_NUMBER = 4;       // D列: チャンネル番号
const COL_SLACK_NAME = 5;   // E列: Slack作成用チャンネル名
const COL_VALIDATION = 6;   // F列: バリデーション
const COL_ZAPIER = 7;       // G列: zapier_action
const COL_STATUS = 8;       // H列: Slack作成ステータス（結果出力）

// ========================================
// メイン処理: 変更時トリガー（onChangeのみ）
// ========================================

/**
 * スプレッドシート変更時のトリガー関数
 * @param {Object} e - イベントオブジェクト
 */
function onSpreadsheetChange(e) {
  // 変更タイプをログ出力
  Logger.log('変更タイプ: ' + e.changeType);

  // EDIT または INSERT_ROW の場合のみ処理
  if (e.changeType !== 'EDIT' && e.changeType !== 'INSERT_ROW') {
    Logger.log('対象外の変更タイプのためスキップ');
    return;
  }

  // 新規エントリを処理（採番 → 5秒待機 → Slack作成を一括実行）
  processNewEntryAndCreateChannel();
}

// ========================================
// 一括処理: 採番 → 5秒待機 → Slackチャンネル作成
// ========================================

/**
 * 新規エントリの処理とSlackチャンネル作成を一括実行
 * 1. フォーム回答シートからデータをコピー
 * 2. 採番処理
 * 3. 5秒待機
 * 4. Slackチャンネル作成
 */
function processNewEntryAndCreateChannel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  const targetSheet = ss.getSheetByName(TARGET_SHEET_NAME);

  if (!sourceSheet || !targetSheet) {
    Logger.log('エラー: シートが見つかりません');
    return;
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  let targetRow = -1;
  let channelName = '';

  try {
    const sourceLastRow = sourceSheet.getLastRow();
    if (sourceLastRow <= HEADER_ROWS) {
      Logger.log('フォーム形式の回答シートにデータがありません');
      return;
    }

    // フォーム形式の回答シートの最新行からデータを取得
    const colA = sourceSheet.getRange(sourceLastRow, 1).getValue(); // 用途
    const colB = sourceSheet.getRange(sourceLastRow, 2).getValue(); // チャンネル名

    Logger.log('新規データ検出: 用途=' + colA + ', チャンネル名=' + colB);

    if (!colA) {
      Logger.log('用途が空のためスキップ');
      return;
    }

    // 既に処理済みかチェック
    const targetLastRow = targetSheet.getLastRow();
    if (targetLastRow > HEADER_ROWS) {
      const existingData = targetSheet
        .getRange(HEADER_ROWS + 1, 1, targetLastRow - HEADER_ROWS, 2)
        .getValues();

      for (const row of existingData) {
        if (row[0] === colA && row[1] === colB) {
          Logger.log('既に処理済みのデータです');
          return;
        }
      }
    }

    // 書き込み先行を決定
    if (targetLastRow <= HEADER_ROWS) {
      targetRow = HEADER_ROWS + 1;
    } else {
      // A列が空欄の行を探す
      const aColValues = targetSheet
        .getRange(HEADER_ROWS + 1, COL_PURPOSE, targetLastRow - HEADER_ROWS, 1)
        .getValues()
        .flat();

      targetRow = -1;
      for (let i = 0; i < aColValues.length; i++) {
        if (aColValues[i] === '' || aColValues[i] === null) {
          targetRow = HEADER_ROWS + 1 + i;
          break;
        }
      }

      if (targetRow === -1) {
        targetRow = targetLastRow + 1;
      }
    }

    Logger.log('書き込み先行: ' + targetRow);

    // A列とB列にデータをコピー
    targetSheet.getRange(targetRow, COL_PURPOSE).setValue(colA);
    targetSheet.getRange(targetRow, COL_NAME).setValue(colB);

    // 採番処理
    const purpose = String(colA).trim();
    const numCell = targetSheet.getRange(targetRow, COL_NUMBER);
    const [minN, maxN] = rangeByPurpose_(purpose);

    Logger.log('番号範囲: ' + minN + '〜' + maxN);

    // 既存番号から最大を探す
    let currentMax = minN - 1;
    const currentLastRow = targetSheet.getLastRow();
    if (currentLastRow > HEADER_ROWS) {
      const allValues = targetSheet
        .getRange(HEADER_ROWS + 1, COL_NUMBER, currentLastRow - HEADER_ROWS, 1)
        .getValues()
        .flat();

      for (const v of allValues) {
        if (typeof v === 'number' && v >= minN && v <= maxN) {
          if (v > currentMax) currentMax = v;
        }
      }
    }

    const nextN = currentMax + 1;
    Logger.log('次の番号: ' + nextN);

    if (nextN > maxN) {
      numCell.setValue('OVER_RANGE');
      Logger.log('警告: 番号範囲を超えました');
    } else {
      numCell.setValue(nextN);
      Logger.log('採番完了: ' + nextN);
    }

    // スプレッドシートの変更を即座に反映
    SpreadsheetApp.flush();
    Logger.log('採番処理完了。5秒待機します...');

    // E列のチャンネル名を取得（数式の場合は計算結果を取得）
    channelName = targetSheet.getRange(targetRow, COL_SLACK_NAME).getDisplayValue();
    Logger.log('E列のチャンネル名: ' + channelName);

  } finally {
    lock.releaseLock();
  }

  // 採番処理が完了した場合のみ、5秒待機後にSlackチャンネル作成
  if (targetRow > 0 && channelName && String(channelName).trim() !== '') {
    // 5秒待機
    Utilities.sleep(5000);
    Logger.log('5秒待機完了。Slackチャンネル作成を開始します。');

    // Slackチャンネル作成
    createSlackChannel(String(channelName).trim(), targetSheet, targetRow);
  } else if (targetRow > 0) {
    Logger.log('E列にチャンネル名がないため、Slack作成はスキップします');
  }
}

/**
 * 用途に応じた番号範囲を返す
 */
function rangeByPurpose_(purpose) {
  switch (purpose) {
    case 'デフォルト':
    case 'デフォルトチャンネル':
      return [0, 99];
    case '組織':
      return [100, 199];
    case 'プロジェクト':
      return [200, 899];
    case 'プライベート':
      return [900, 999];
    default:
      Logger.log('警告: 未知の用途「' + purpose + '」- デフォルト範囲を使用');
      return [0, 99];
  }
}

// ========================================
// Slack チャンネル作成処理
// ========================================

/**
 * Slackチャンネルを作成
 * @param {string} originalName - 元のチャンネル名（日本語可）
 * @param {Sheet} sheet - シートオブジェクト
 * @param {number} row - 行番号
 */
function createSlackChannel(originalName, sheet, row) {
  const token = getSlackToken();

  if (!token) {
    sheet.getRange(row, COL_STATUS).setValue('エラー: トークン未設定');
    return;
  }

  // チャンネル名を正規化（Slack仕様に合わせる）
  const normalizedName = normalizeChannelName(originalName);

  if (!normalizedName) {
    sheet.getRange(row, COL_STATUS).setValue('エラー: 無効なチャンネル名');
    return;
  }

  Logger.log('Slackチャンネル作成: ' + normalizedName);

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
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    if (result.ok) {
      const channelId = result.channel.id;
      const createdName = result.channel.name;
      sheet.getRange(row, COL_STATUS).setValue('✓ 作成完了: ' + createdName);
      Logger.log('チャンネル作成成功: ' + createdName + ' (ID: ' + channelId + ')');
    } else {
      const errorMessage = getErrorMessage(result.error);
      sheet.getRange(row, COL_STATUS).setValue('✗ エラー: ' + errorMessage);
      Logger.log('チャンネル作成失敗: ' + result.error);
    }
  } catch (error) {
    sheet.getRange(row, COL_STATUS).setValue('✗ 通信エラー: ' + error.message);
    Logger.log('通信エラー: ' + error);
  }
}

/**
 * チャンネル名を正規化（Slack仕様に合わせる）
 */
function normalizeChannelName(name) {
  if (!name) return '';

  let normalized = name
    .toLowerCase()
    .replace(/　/g, ' ')
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-+|-+$/g, '')
    .replace(/[^\w\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FFF\u3400-\u4DBF-]/g, '');

  if (normalized.length > 80) {
    normalized = normalized.substring(0, 80);
  }

  return normalized;
}

/**
 * Slackエラーコードを日本語メッセージに変換
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
 * 一括処理の手動テスト
 */
function testProcessNewEntryAndCreateChannel() {
  processNewEntryAndCreateChannel();
  Logger.log('テスト完了');
}

/**
 * 特定行でSlack作成をテスト（デバッグ用）
 */
function testSlackCreationForRow() {
  const row = 2; // テストしたい行番号を指定
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TARGET_SHEET_NAME);

  if (!sheet) {
    Browser.msgBox('シートが見つかりません: ' + TARGET_SHEET_NAME);
    return;
  }

  const channelName = sheet.getRange(row, COL_SLACK_NAME).getDisplayValue();
  Browser.msgBox('行' + row + 'のE列の値: ' + channelName);

  if (channelName && String(channelName).trim() !== '') {
    createSlackChannel(String(channelName).trim(), sheet, row);
    Browser.msgBox('処理完了。H列を確認してください。');
  } else {
    Browser.msgBox('E列が空です');
  }
}

// ========================================
// トリガー管理関数
// ========================================

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

  Browser.msgBox(deletedCount + '個のトリガーを全て削除しました');
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
 * 必要なトリガーを設定（onChangeのみ）
 */
function createTriggers() {
  // 既存のトリガーを全て削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });

  const ss = SpreadsheetApp.getActive();

  // onSpreadsheetChange トリガー（変更時のみ）
  ScriptApp.newTrigger('onSpreadsheetChange')
    .forSpreadsheet(ss)
    .onChange()
    .create();

  Browser.msgBox('トリガーを設定しました:\n・onSpreadsheetChange（変更時）');
}

/**
 * 初期セットアップを一括実行
 */
function initialSetup() {
  // 1. トークン設定
  const token = Browser.inputBox('【1/2】Slack Bot Tokenを入力してください（xoxb-で始まるもの）');
  if (!token || token === 'cancel') {
    Browser.msgBox('セットアップをキャンセルしました');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('SLACK_BOT_TOKEN', token);

  // 2. トリガー設定（onChangeのみ）
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });

  const ss = SpreadsheetApp.getActive();

  // onChangeトリガーのみ設定
  ScriptApp.newTrigger('onSpreadsheetChange')
    .forSpreadsheet(ss)
    .onChange()
    .create();

  Browser.msgBox('【2/2】セットアップ完了！\n\n' +
    'トークンを保存し、トリガーを設定しました。\n\n' +
    '動作:\n' +
    '・「フォーム形式の回答」シートにデータ追加\n' +
    '　→「チャンネル名」シートにコピー＆採番\n' +
    '　→ 5秒待機\n' +
    '　→ Slackチャンネル自動作成（結果はH列）');
}
