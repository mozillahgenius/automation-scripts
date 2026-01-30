/**
 * Slack チャンネル自動作成 Google Apps Script（v4）
 *
 * 機能:
 * 「フォーム形式の回答」シートに新規データ追加時、
 * 「チャンネル名」シートにコピー＆採番 → 5秒待機 → Slackチャンネル自動作成
 * → メンバー招待 → マネージャー権限付与
 * を一括で実行
 *
 * シート構造（フォーム形式の回答）:
 * A列: 用途
 * B列: 表示名（日本語可）
 * C列: マネージャー（ユーザーID）
 * D列: メンバー（ユーザーID、カンマ区切り可）
 *
 * シート構造（チャンネル名）:
 * A列: 用途
 * B列: 表示名（日本語可）
 * C列: マネージャー（ユーザーID）
 * D列: メンバー（ユーザーID、カンマ区切り可）
 * E列: 用途コード
 * F列: チャンネル番号
 * G列: Slack作成用チャンネル名
 * H列: バリデーション
 * I列: zapier_action
 * J列: Slack作成ステータス（結果出力）
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

// 列番号の設定（フォーム形式の回答シート）
const SRC_COL_PURPOSE = 1;    // A列: 用途
const SRC_COL_NAME = 2;       // B列: 表示名
const SRC_COL_MANAGER = 3;    // C列: マネージャー
const SRC_COL_MEMBERS = 4;    // D列: メンバー

// 列番号の設定（チャンネル名シート）
const HEADER_ROWS = 1;
const COL_PURPOSE = 1;        // A列: 用途
const COL_NAME = 2;           // B列: 表示名（日本語可）
const COL_MANAGER = 3;        // C列: マネージャー（ユーザーID）
const COL_MEMBERS = 4;        // D列: メンバー（ユーザーID）
const COL_CODE = 5;           // E列: 用途コード
const COL_NUMBER = 6;         // F列: チャンネル番号
const COL_SLACK_NAME = 7;     // G列: Slack作成用チャンネル名
const COL_VALIDATION = 8;     // H列: バリデーション
const COL_ZAPIER = 9;         // I列: zapier_action
const COL_STATUS = 10;        // J列: Slack作成ステータス（結果出力）

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
// 一括処理: 採番 → 5秒待機 → Slackチャンネル作成 → メンバー招待
// ========================================

/**
 * 新規エントリの処理とSlackチャンネル作成を一括実行
 * 1. フォーム回答シートからデータをコピー
 * 2. 採番処理（プライベート以外）
 * 3. 5秒待機
 * 4. Slackチャンネル作成（プライベート以外）
 * 5. メンバー招待
 * 6. マネージャー権限付与
 *
 * 注意: プライベートチャンネルは自動作成・採番しない（データコピーのみ）
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
  let managerIds = '';
  let memberIds = '';
  let isPrivate = false;

  try {
    const sourceLastRow = sourceSheet.getLastRow();
    if (sourceLastRow <= HEADER_ROWS) {
      Logger.log('フォーム形式の回答シートにデータがありません');
      return;
    }

    // フォーム形式の回答シートの最新行からデータを取得
    const colA = sourceSheet.getRange(sourceLastRow, SRC_COL_PURPOSE).getValue();   // 用途
    const colB = sourceSheet.getRange(sourceLastRow, SRC_COL_NAME).getValue();      // チャンネル名
    const colC = sourceSheet.getRange(sourceLastRow, SRC_COL_MANAGER).getValue();   // マネージャー
    const colD = sourceSheet.getRange(sourceLastRow, SRC_COL_MEMBERS).getValue();   // メンバー

    Logger.log('新規データ検出: 用途=' + colA + ', チャンネル名=' + colB + ', マネージャー=' + colC + ', メンバー=' + colD);

    if (!colA) {
      Logger.log('用途が空のためスキップ');
      return;
    }

    // プライベートチャンネルかどうかをチェック
    const purpose = String(colA).trim();
    isPrivate = (purpose === 'プライベート');

    // デバッグ用ログ
    Logger.log('用途の値: [' + purpose + '], isPrivate: ' + isPrivate);
    Logger.log('文字コード: ' + purpose.split('').map(c => c.charCodeAt(0)).join(','));

    if (isPrivate) {
      Logger.log('プライベートチャンネルのため、自動作成・採番はスキップします');
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

    // A列〜D列にデータをコピー
    targetSheet.getRange(targetRow, COL_PURPOSE).setValue(colA);   // 用途
    targetSheet.getRange(targetRow, COL_NAME).setValue(colB);      // 表示名
    targetSheet.getRange(targetRow, COL_MANAGER).setValue(colC);   // マネージャー
    targetSheet.getRange(targetRow, COL_MEMBERS).setValue(colD);   // メンバー

    // プライベートチャンネルの場合はデータコピーのみで終了
    if (isPrivate) {
      // ステータスに手動作成が必要であることを記載
      targetSheet.getRange(targetRow, COL_STATUS).setValue('手動作成が必要（プライベートチャンネル）');
      SpreadsheetApp.flush();
      Logger.log('プライベートチャンネル: データコピー完了。手動作成が必要です。');
      return;
    }

    // マネージャーとメンバーのIDを保持
    managerIds = String(colC || '').trim();
    memberIds = String(colD || '').trim();

    // 採番処理（プライベート以外）
    const numCell = targetSheet.getRange(targetRow, COL_NUMBER);
    const [minN, maxN] = rangeByPurpose_(purpose);

    Logger.log('番号範囲: ' + minN + '〜' + maxN);

    // 既存番号を全て収集（手動追加分も含む）
    // 文字列・数値どちらも対応
    const usedNumbers = new Set();
    const currentLastRow = targetSheet.getLastRow();
    if (currentLastRow > HEADER_ROWS) {
      const allValues = targetSheet
        .getRange(HEADER_ROWS + 1, COL_NUMBER, currentLastRow - HEADER_ROWS, 1)
        .getValues()
        .flat();

      for (const v of allValues) {
        // 数値または数値に変換可能な文字列を処理
        const num = typeof v === 'number' ? v : parseInt(String(v), 10);
        if (!isNaN(num) && num >= minN && num <= maxN) {
          usedNumbers.add(num);
        }
      }
    }

    Logger.log('使用済み番号: ' + Array.from(usedNumbers).sort((a, b) => a - b).join(', '));

    // 範囲内で使用されていない最小の番号を探す
    let nextN = -1;
    for (let n = minN; n <= maxN; n++) {
      if (!usedNumbers.has(n)) {
        nextN = n;
        break;
      }
    }

    Logger.log('次の番号: ' + nextN);

    if (nextN === -1) {
      numCell.setValue('OVER_RANGE');
      Logger.log('警告: 番号範囲内に空きがありません');
    } else {
      numCell.setValue(nextN);
      Logger.log('採番完了: ' + nextN);

      // E列に用途コードを書き込み
      const purposeCode = getPurposeCode_(purpose);
      targetSheet.getRange(targetRow, COL_CODE).setValue(purposeCode);
      Logger.log('用途コード: ' + (purposeCode || '(なし)'));

      // G列にチャンネル名を書き込み
      channelName = generateChannelName_(purpose, nextN, colB);
      targetSheet.getRange(targetRow, COL_SLACK_NAME).setValue(channelName);
      Logger.log('生成されたチャンネル名: ' + channelName);
    }

    // スプレッドシートの変更を即座に反映
    SpreadsheetApp.flush();
    Logger.log('採番処理完了。5秒待機します...');

  } finally {
    lock.releaseLock();
  }

  // 採番処理が完了した場合のみ、5秒待機後にSlackチャンネル作成（プライベート以外）
  if (!isPrivate && targetRow > 0 && channelName && String(channelName).trim() !== '') {
    // 5秒待機
    Utilities.sleep(5000);
    Logger.log('5秒待機完了。Slackチャンネル作成を開始します。');

    // Slackチャンネル作成（メンバー招待・マネージャー権限付与も含む）
    createSlackChannelWithMembers(
      String(channelName).trim(),
      targetSheet,
      targetRow,
      managerIds,
      memberIds
    );
  } else if (targetRow > 0 && !isPrivate) {
    Logger.log('G列にチャンネル名がないため、Slack作成はスキップします');
  }
}

/**
 * 用途に応じた番号範囲を返す
 * 注意: プライベートは自動採番しないため含まない
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
    default:
      Logger.log('警告: 未知の用途「' + purpose + '」- デフォルト範囲を使用');
      return [0, 99];
  }
}

/**
 * 用途に応じた用途コードを返す
 * 注意: プライベートは自動作成しないため含まない
 */
function getPurposeCode_(purpose) {
  switch (purpose) {
    case 'デフォルト':
    case 'デフォルトチャンネル':
      return ''; // デフォルトは用途コードなし
    case '組織':
      return 'org';
    case 'プロジェクト':
      return 'pj';
    default:
      return '';
  }
}

/**
 * チャンネル名を生成
 * デフォルト: 番号_チャンネル名 (例: 01_test)
 * その他: 用途コード_番号_チャンネル名 (例: org_101_test)
 * @param {string} purpose - 用途
 * @param {number} number - チャンネル番号
 * @param {string} displayName - 表示名
 * @returns {string} - 生成されたチャンネル名
 */
function generateChannelName_(purpose, number, displayName) {
  const code = getPurposeCode_(purpose);
  const paddedNumber = String(number).padStart(2, '0');
  const name = String(displayName || '').trim();

  if (code === '') {
    // デフォルトの場合: 番号_チャンネル名
    return paddedNumber + '_' + name;
  } else {
    // その他の場合: 用途コード_番号_チャンネル名
    return code + '_' + paddedNumber + '_' + name;
  }
}

// ========================================
// Slack チャンネル作成処理（メンバー招待・権限付与含む）
// ========================================

/**
 * Slackチャンネルを作成し、メンバー招待とマネージャー権限付与を行う
 * @param {string} originalName - 元のチャンネル名（日本語可）
 * @param {Sheet} sheet - シートオブジェクト
 * @param {number} row - 行番号
 * @param {string} managerIds - マネージャーのユーザーID（カンマ区切り可）
 * @param {string} memberIds - メンバーのユーザーID（カンマ区切り可）
 */
function createSlackChannelWithMembers(originalName, sheet, row, managerIds, memberIds) {
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

  // ステータスメッセージを格納
  let statusMessages = [];

  // 1. Slackチャンネル作成
  const channelId = createChannel(token, normalizedName);

  if (!channelId) {
    sheet.getRange(row, COL_STATUS).setValue('✗ チャンネル作成失敗');
    return;
  }

  statusMessages.push('✓ チャンネル作成: ' + normalizedName);
  Logger.log('チャンネル作成成功: ' + normalizedName + ' (ID: ' + channelId + ')');

  // 2. メンバーを招待（マネージャーも含む）
  const allUserIds = combineUserIds(managerIds, memberIds);

  if (allUserIds.length > 0) {
    const inviteResult = inviteUsersToChannel(token, channelId, allUserIds);
    if (inviteResult.success) {
      statusMessages.push('✓ メンバー招待: ' + inviteResult.invited + '人');
    } else {
      statusMessages.push('△ メンバー招待: ' + inviteResult.message);
    }
  }

  // 3. マネージャーに管理者権限を付与
  const managerIdList = parseUserIds(managerIds);

  if (managerIdList.length > 0) {
    const managerResult = setChannelManagers(token, channelId, managerIdList);
    if (managerResult.success) {
      statusMessages.push('✓ 管理者設定: ' + managerResult.set + '人');
    } else {
      statusMessages.push('△ 管理者設定: ' + managerResult.message);
    }
  }

  // ステータスを更新
  sheet.getRange(row, COL_STATUS).setValue(statusMessages.join(' / '));
}

/**
 * Slackチャンネルを作成
 * @param {string} token - Slack Bot Token
 * @param {string} channelName - チャンネル名
 * @returns {string|null} - チャンネルID（失敗時はnull）
 */
function createChannel(token, channelName) {
  const url = 'https://slack.com/api/conversations.create';
  const payload = {
    name: channelName,
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
      return result.channel.id;
    } else {
      Logger.log('チャンネル作成失敗: ' + result.error);
      return null;
    }
  } catch (error) {
    Logger.log('通信エラー: ' + error);
    return null;
  }
}

/**
 * ユーザーをチャンネルに招待
 * @param {string} token - Slack Bot Token
 * @param {string} channelId - チャンネルID
 * @param {string[]} userIds - ユーザーIDの配列
 * @returns {Object} - 結果オブジェクト
 */
function inviteUsersToChannel(token, channelId, userIds) {
  if (userIds.length === 0) {
    return { success: true, invited: 0, message: '招待するユーザーなし' };
  }

  const url = 'https://slack.com/api/conversations.invite';
  let invitedCount = 0;
  let errors = [];

  for (const userId of userIds) {
    const payload = {
      channel: channelId,
      users: userId
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
        invitedCount++;
        Logger.log('ユーザー招待成功: ' + userId);
      } else if (result.error === 'already_in_channel') {
        invitedCount++;
        Logger.log('ユーザーは既にチャンネルに参加: ' + userId);
      } else {
        errors.push(userId + ':' + result.error);
        Logger.log('ユーザー招待失敗: ' + userId + ' - ' + result.error);
      }
    } catch (error) {
      errors.push(userId + ':通信エラー');
      Logger.log('通信エラー: ' + error);
    }

    // API制限対策で少し待機
    Utilities.sleep(200);
  }

  return {
    success: errors.length === 0,
    invited: invitedCount,
    message: errors.length > 0 ? errors.join(', ') : ''
  };
}

/**
 * チャンネルの管理者権限を設定
 * conversations.setPurpose や admin.conversations.setTeams などは
 * Enterprise Grid のみなので、ここでは代替手段を使用
 *
 * 注意: Slack APIでチャンネル管理者を設定するには
 * admin.conversations.* スコープが必要（Enterprise Gridのみ）
 * 通常のワークスペースでは、チャンネル作成者が自動的に管理者になる
 *
 * @param {string} token - Slack Bot Token
 * @param {string} channelId - チャンネルID
 * @param {string[]} managerIds - マネージャーのユーザーID配列
 * @returns {Object} - 結果オブジェクト
 */
function setChannelManagers(token, channelId, managerIds) {
  if (managerIds.length === 0) {
    return { success: true, set: 0, message: '設定するマネージャーなし' };
  }

  // Enterprise Grid以外では admin.conversations.* APIは使用できない
  // 代替手段として、チャンネルの説明にマネージャー情報を記載する方法を使用
  // または、Botがチャンネル管理者権限を持っている場合のみ設定可能

  // admin.conversations.setConversationPrefs を試行
  // このAPIはEnterprise Gridでのみ利用可能
  const url = 'https://slack.com/api/admin.conversations.setConversationPrefs';

  let setCount = 0;
  let errors = [];

  // まず、現在のBotの権限でマネージャー設定が可能か確認
  // 多くの場合、この機能はEnterprise Gridでのみ利用可能

  for (const managerId of managerIds) {
    const payload = {
      channel_id: channelId,
      prefs: {
        who_can_post: {
          type: ['admin', 'owner', 'org_admin'],
          user: [managerId]
        }
      }
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
        setCount++;
        Logger.log('マネージャー権限設定成功: ' + managerId);
      } else {
        // Enterprise Grid以外では 'missing_scope' や 'not_allowed' エラーになる
        if (result.error === 'missing_scope' || result.error === 'not_allowed' ||
            result.error === 'feature_not_enabled') {
          Logger.log('マネージャー権限設定: Enterprise Grid機能のため利用不可');
          return {
            success: false,
            set: 0,
            message: 'Enterprise Grid機能（利用不可）'
          };
        }
        errors.push(managerId + ':' + result.error);
        Logger.log('マネージャー権限設定失敗: ' + managerId + ' - ' + result.error);
      }
    } catch (error) {
      errors.push(managerId + ':通信エラー');
      Logger.log('通信エラー: ' + error);
    }

    Utilities.sleep(200);
  }

  return {
    success: errors.length === 0,
    set: setCount,
    message: errors.length > 0 ? errors.join(', ') : ''
  };
}

/**
 * ユーザーID文字列をパース（カンマ区切り対応）
 * @param {string} userIdString - ユーザーID文字列
 * @returns {string[]} - ユーザーIDの配列
 */
function parseUserIds(userIdString) {
  if (!userIdString || userIdString.trim() === '') {
    return [];
  }

  return userIdString
    .split(/[,、\s]+/)
    .map(id => id.trim())
    .filter(id => id !== '' && id.startsWith('U'));
}

/**
 * マネージャーとメンバーのIDを結合（重複排除）
 * @param {string} managerIds - マネージャーID文字列
 * @param {string} memberIds - メンバーID文字列
 * @returns {string[]} - 結合されたユーザーIDの配列
 */
function combineUserIds(managerIds, memberIds) {
  const managers = parseUserIds(managerIds);
  const members = parseUserIds(memberIds);

  const combined = [...new Set([...managers, ...members])];
  return combined;
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
    'missing_scope': '必要な権限がありません',
    'not_authed': '認証されていません',
    'invalid_auth': '無効な認証情報です',
    'token_revoked': 'トークンが無効化されています',
    'ratelimited': 'API制限に達しました。しばらく待ってから再試行してください',
    'user_not_found': 'ユーザーが見つかりません',
    'channel_not_found': 'チャンネルが見つかりません',
    'already_in_channel': 'ユーザーは既にチャンネルに参加しています',
    'cant_invite_self': '自分自身を招待できません',
    'not_in_channel': 'ユーザーはチャンネルに参加していません'
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

  const channelId = createChannel(token, normalizedName);

  if (channelId) {
    Browser.msgBox('成功！チャンネルを作成しました\nID: ' + channelId);
  } else {
    Browser.msgBox('チャンネル作成に失敗しました');
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
  const managerIds = sheet.getRange(row, COL_MANAGER).getValue();
  const memberIds = sheet.getRange(row, COL_MEMBERS).getValue();

  Browser.msgBox(
    '行' + row + 'のデータ:\n' +
    'チャンネル名(G列): ' + channelName + '\n' +
    'マネージャー(C列): ' + managerIds + '\n' +
    'メンバー(D列): ' + memberIds
  );

  if (channelName && String(channelName).trim() !== '') {
    createSlackChannelWithMembers(
      String(channelName).trim(),
      sheet,
      row,
      String(managerIds || ''),
      String(memberIds || '')
    );
    Browser.msgBox('処理完了。J列を確認してください。');
  } else {
    Browser.msgBox('G列が空です');
  }
}

/**
 * ユーザー招待のテスト
 */
function testInviteUser() {
  const channelId = Browser.inputBox('チャンネルIDを入力してください');
  if (!channelId || channelId === 'cancel') return;

  const userId = Browser.inputBox('招待するユーザーIDを入力してください（U...）');
  if (!userId || userId === 'cancel') return;

  const token = getSlackToken();
  if (!token) {
    Browser.msgBox('トークンが設定されていません');
    return;
  }

  const result = inviteUsersToChannel(token, channelId, [userId]);
  Browser.msgBox('結果: ' + JSON.stringify(result));
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
    '　→ Slackチャンネル自動作成\n' +
    '　→ メンバー招待\n' +
    '　→ マネージャー権限設定（Enterprise Gridのみ）\n' +
    '　→ 結果はJ列に出力');
}

/**
 * 必要なSlack Bot権限の一覧を表示
 */
function showRequiredScopes() {
  Browser.msgBox(
    '必要なSlack Bot Token Scopes:\n\n' +
    '【必須】\n' +
    '・channels:write - チャンネル作成\n' +
    '・channels:manage - チャンネル管理\n' +
    '・groups:write - プライベートチャンネル作成\n\n' +
    '【メンバー招待用】\n' +
    '・channels:read - チャンネル情報取得\n' +
    '・users:read - ユーザー情報取得\n\n' +
    '【マネージャー権限設定用（Enterprise Gridのみ）】\n' +
    '・admin.conversations:write\n\n' +
    'Slack App設定画面の「OAuth & Permissions」で設定してください'
  );
}

// ========================================
// スプレッドシート初期セットアップ
// ========================================

/**
 * スプレッドシートの初期セットアップ
 * 必要なシートとヘッダーを自動作成
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. 「フォーム形式の回答」シートを作成
  let sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) {
    sourceSheet = ss.insertSheet(SOURCE_SHEET_NAME);
    Logger.log('シートを作成しました: ' + SOURCE_SHEET_NAME);
  } else {
    Logger.log('シートは既に存在します: ' + SOURCE_SHEET_NAME);
  }

  // フォーム形式の回答シートのヘッダーを設定
  const sourceHeaders = ['用途', '表示名', 'マネージャー', 'メンバー'];
  const sourceHeaderRange = sourceSheet.getRange(1, 1, 1, sourceHeaders.length);

  // 既存のヘッダーをチェック
  const existingSourceHeaders = sourceHeaderRange.getValues()[0];
  const sourceNeedsHeader = existingSourceHeaders.every(cell => cell === '');

  if (sourceNeedsHeader) {
    sourceHeaderRange.setValues([sourceHeaders]);
    sourceHeaderRange.setFontWeight('bold');
    sourceHeaderRange.setBackground('#4a86e8');
    sourceHeaderRange.setFontColor('#ffffff');
    Logger.log('ヘッダーを設定しました: ' + SOURCE_SHEET_NAME);
  }

  // 列幅を調整
  sourceSheet.setColumnWidth(1, 120); // 用途
  sourceSheet.setColumnWidth(2, 200); // 表示名
  sourceSheet.setColumnWidth(3, 150); // マネージャー
  sourceSheet.setColumnWidth(4, 200); // メンバー

  // 用途のドロップダウンを設定（A列）
  const sourceDataValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['デフォルト', '組織', 'プロジェクト', 'プライベート'], true)
    .setAllowInvalid(false)
    .build();
  sourceSheet.getRange(2, 1, 100, 1).setDataValidation(sourceDataValidation);

  // 2. 「チャンネル名」シートを作成
  let targetSheet = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!targetSheet) {
    targetSheet = ss.insertSheet(TARGET_SHEET_NAME);
    Logger.log('シートを作成しました: ' + TARGET_SHEET_NAME);
  } else {
    Logger.log('シートは既に存在します: ' + TARGET_SHEET_NAME);
  }

  // チャンネル名シートのヘッダーを設定
  const targetHeaders = [
    '用途',           // A列
    '表示名',         // B列
    'マネージャー',   // C列
    'メンバー',       // D列
    '用途コード',     // E列
    'チャンネル番号', // F列
    'Slackチャンネル名', // G列
    'バリデーション', // H列
    'zapier_action',  // I列
    'ステータス'      // J列
  ];
  const targetHeaderRange = targetSheet.getRange(1, 1, 1, targetHeaders.length);

  // 既存のヘッダーをチェック
  const existingTargetHeaders = targetHeaderRange.getValues()[0];
  const targetNeedsHeader = existingTargetHeaders.every(cell => cell === '');

  if (targetNeedsHeader) {
    targetHeaderRange.setValues([targetHeaders]);
    targetHeaderRange.setFontWeight('bold');
    targetHeaderRange.setBackground('#34a853');
    targetHeaderRange.setFontColor('#ffffff');
    Logger.log('ヘッダーを設定しました: ' + TARGET_SHEET_NAME);
  }

  // 列幅を調整
  targetSheet.setColumnWidth(1, 100);  // 用途
  targetSheet.setColumnWidth(2, 180);  // 表示名
  targetSheet.setColumnWidth(3, 130);  // マネージャー
  targetSheet.setColumnWidth(4, 180);  // メンバー
  targetSheet.setColumnWidth(5, 80);   // 用途コード
  targetSheet.setColumnWidth(6, 100);  // チャンネル番号
  targetSheet.setColumnWidth(7, 220);  // Slackチャンネル名
  targetSheet.setColumnWidth(8, 100);  // バリデーション
  targetSheet.setColumnWidth(9, 100);  // zapier_action
  targetSheet.setColumnWidth(10, 300); // ステータス

  // ヘッダー行を固定
  sourceSheet.setFrozenRows(1);
  targetSheet.setFrozenRows(1);

  Browser.msgBox(
    'スプレッドシートのセットアップが完了しました！\n\n' +
    '作成されたシート:\n' +
    '・' + SOURCE_SHEET_NAME + '（データ入力用）\n' +
    '・' + TARGET_SHEET_NAME + '（処理結果）\n\n' +
    '次のステップ:\n' +
    '1. initialSetup() を実行してトークンとトリガーを設定\n' +
    '2. 「' + SOURCE_SHEET_NAME + '」シートにデータを入力'
  );
}

/**
 * 完全な初期セットアップ（スプレッドシート作成 + トークン + トリガー）
 */
function fullSetup() {
  // 1. スプレッドシートのセットアップ
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 「フォーム形式の回答」シートを作成
  let sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) {
    sourceSheet = ss.insertSheet(SOURCE_SHEET_NAME);
  }

  // フォーム形式の回答シートのヘッダーを設定
  const sourceHeaders = ['用途', '表示名', 'マネージャー', 'メンバー'];
  const sourceHeaderRange = sourceSheet.getRange(1, 1, 1, sourceHeaders.length);
  const existingSourceHeaders = sourceHeaderRange.getValues()[0];

  if (existingSourceHeaders.every(cell => cell === '')) {
    sourceHeaderRange.setValues([sourceHeaders]);
    sourceHeaderRange.setFontWeight('bold');
    sourceHeaderRange.setBackground('#4a86e8');
    sourceHeaderRange.setFontColor('#ffffff');
  }

  sourceSheet.setColumnWidth(1, 120);
  sourceSheet.setColumnWidth(2, 200);
  sourceSheet.setColumnWidth(3, 150);
  sourceSheet.setColumnWidth(4, 200);

  // 用途のドロップダウンを設定
  const sourceDataValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['デフォルト', '組織', 'プロジェクト', 'プライベート'], true)
    .setAllowInvalid(false)
    .build();
  sourceSheet.getRange(2, 1, 100, 1).setDataValidation(sourceDataValidation);

  // 「チャンネル名」シートを作成
  let targetSheet = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!targetSheet) {
    targetSheet = ss.insertSheet(TARGET_SHEET_NAME);
  }

  const targetHeaders = [
    '用途', '表示名', 'マネージャー', 'メンバー', '用途コード',
    'チャンネル番号', 'Slackチャンネル名', 'バリデーション', 'zapier_action', 'ステータス'
  ];
  const targetHeaderRange = targetSheet.getRange(1, 1, 1, targetHeaders.length);
  const existingTargetHeaders = targetHeaderRange.getValues()[0];

  if (existingTargetHeaders.every(cell => cell === '')) {
    targetHeaderRange.setValues([targetHeaders]);
    targetHeaderRange.setFontWeight('bold');
    targetHeaderRange.setBackground('#34a853');
    targetHeaderRange.setFontColor('#ffffff');
  }

  targetSheet.setColumnWidth(1, 100);
  targetSheet.setColumnWidth(2, 180);
  targetSheet.setColumnWidth(3, 130);
  targetSheet.setColumnWidth(4, 180);
  targetSheet.setColumnWidth(5, 80);
  targetSheet.setColumnWidth(6, 100);
  targetSheet.setColumnWidth(7, 220);
  targetSheet.setColumnWidth(8, 100);
  targetSheet.setColumnWidth(9, 100);
  targetSheet.setColumnWidth(10, 300);

  sourceSheet.setFrozenRows(1);
  targetSheet.setFrozenRows(1);

  // 2. トークン設定
  const token = Browser.inputBox(
    '【1/2】Slack Bot Tokenを入力してください\n（xoxb-で始まるもの）'
  );
  if (!token || token === 'cancel') {
    Browser.msgBox(
      'セットアップを中断しました。\n\n' +
      'スプレッドシートは作成されました。\n' +
      'トークンとトリガーの設定は initialSetup() を実行してください。'
    );
    return;
  }
  PropertiesService.getScriptProperties().setProperty('SLACK_BOT_TOKEN', token);

  // 3. トリガー設定
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });

  ScriptApp.newTrigger('onSpreadsheetChange')
    .forSpreadsheet(ss)
    .onChange()
    .create();

  Browser.msgBox(
    '【2/2】セットアップが完了しました！\n\n' +
    '設定内容:\n' +
    '✓ シート作成（' + SOURCE_SHEET_NAME + ', ' + TARGET_SHEET_NAME + '）\n' +
    '✓ ヘッダー設定\n' +
    '✓ Slack Bot Token保存\n' +
    '✓ トリガー設定\n\n' +
    '使い方:\n' +
    '「' + SOURCE_SHEET_NAME + '」シートに\n' +
    '用途・表示名・マネージャー・メンバーを入力すると\n' +
    '自動でSlackチャンネルが作成されます。'
  );
}
