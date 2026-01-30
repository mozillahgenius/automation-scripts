/**
 * 命名ルール判定ロジック
 * 共通の違反判定と違反ログ管理
 */

// =================================================================
// 汎用名前検証関数
// =================================================================
function validateName(name, rules, whitelist, targetType) {
  // ホワイトリストチェック
  if (isWhitelisted(name, whitelist)) {
    return {
      violation: false,
      violationType: null,
      violationMessage: null,
      matchedRule: 'Whitelisted'
    };
  }
  
  // ルールチェック（優先度順）
  for (const rule of rules) {
    try {
      const regex = new RegExp(rule.regex);
      if (regex.test(name)) {
        // ルールにマッチ（正常）
        return {
          violation: false,
          violationType: null,
          violationMessage: null,
          matchedRule: rule.ruleName
        };
      }
    } catch (error) {
      console.error(`不正な正規表現: ${rule.ruleName} - ${rule.regex}`, error);
    }
  }
  
  // どのルールにもマッチしない（違反）
  const defaultMessage = getDefaultViolationMessage(targetType);
  return {
    violation: true,
    violationType: 'RULE_VIOLATION',
    violationMessage: defaultMessage,
    matchedRule: 'None'
  };
}

// =================================================================
// ホワイトリストチェック
// =================================================================
function isWhitelisted(name, whitelist) {
  for (const item of whitelist) {
    if (item.isRegex) {
      // 正規表現パターン
      try {
        const regex = new RegExp(item.pattern);
        if (regex.test(name)) {
          return true;
        }
      } catch (error) {
        console.error(`ホワイトリストの正規表現エラー: ${item.pattern}`, error);
      }
    } else {
      // 完全一致
      if (name === item.pattern) {
        return true;
      }
    }
  }
  return false;
}

// =================================================================
// デフォルト違反メッセージ
// =================================================================
function getDefaultViolationMessage(targetType) {
  const messages = {
    'Slack': 'チャンネル名が命名ルールに準拠していません',
    'Drive': '共有ドライブ名が命名ルールに準拠していません',
    'DriveFolder': 'フォルダ名が命名ルールに準拠していません'
  };
  return messages[targetType] || '命名ルール違反';
}

// =================================================================
// 違反ログ更新
// =================================================================
function updateViolationsLog() {
  const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
  const violationsSheet = spreadsheet.getSheetByName('Violations_Log');
  
  if (!violationsSheet) {
    throw new Error('Violations_Log シートが見つかりません');
  }
  
  // 既存の違反ログを取得
  const existingViolations = getExistingViolations(violationsSheet);
  
  // 現在の違反を収集
  const currentViolations = collectCurrentViolations(spreadsheet);
  
  // 違反ステータスを更新
  const updatedViolations = mergeViolations(existingViolations, currentViolations);
  
  // シートを更新
  saveViolationsToSheet(violationsSheet, updatedViolations);
  
  return updatedViolations;
}

// =================================================================
// 既存違反ログ取得
// =================================================================
function getExistingViolations(violationsSheet) {
  const lastRow = violationsSheet.getLastRow();
  if (lastRow <= 1) return new Map();
  
  const data = violationsSheet.getRange(2, 1, lastRow - 1, 14).getValues();
  const violations = new Map();
  
  for (const row of data) {
    const key = `${row[1]}_${row[2]}`; // Type_ResourceID
    violations.set(key, {
      violationId: row[0],
      type: row[1],
      resourceId: row[2],
      resourceName: row[3],
      fullPath: row[4],
      violationType: row[5],
      violationMessage: row[6],
      matchedRule: row[7],
      severity: row[8],
      firstDetected: row[9],
      lastConfirmed: row[10],
      status: row[11],
      resolvedDate: row[12],
      notes: row[13]
    });
  }
  
  return violations;
}

// =================================================================
// 現在の違反収集
// =================================================================
function collectCurrentViolations(spreadsheet) {
  const violations = [];
  const now = new Date();
  
  // Slackチャンネルの違反
  const slackSheet = spreadsheet.getSheetByName('Slack_Channels');
  if (slackSheet && slackSheet.getLastRow() > 1) {
    const slackData = slackSheet.getRange(2, 1, slackSheet.getLastRow() - 1, 11).getValues();
    for (const row of slackData) {
      if (row[7] === 'TRUE') { // Violation列
        violations.push({
          type: 'Slack',
          resourceId: row[0],
          resourceName: row[1],
          fullPath: row[1],
          violationType: row[8],
          violationMessage: row[9],
          matchedRule: row[10],
          severity: determineSeverity(row[8]),
          timestamp: now
        });
      }
    }
  }
  
  // 共有ドライブの違反
  const drivesSheet = spreadsheet.getSheetByName('Drive_SharedDrives');
  if (drivesSheet && drivesSheet.getLastRow() > 1) {
    const drivesData = drivesSheet.getRange(2, 1, drivesSheet.getLastRow() - 1, 8).getValues();
    for (const row of drivesData) {
      if (row[4] === 'TRUE') { // Violation列
        violations.push({
          type: 'Drive',
          resourceId: row[0],
          resourceName: row[1],
          fullPath: row[1],
          violationType: row[5],
          violationMessage: row[6],
          matchedRule: row[7],
          severity: determineSeverity(row[5]),
          timestamp: now
        });
      }
    }
  }
  
  // フォルダの違反
  const foldersSheet = spreadsheet.getSheetByName('Drive_Folders');
  if (foldersSheet && foldersSheet.getLastRow() > 1) {
    const foldersData = foldersSheet.getRange(2, 1, foldersSheet.getLastRow() - 1, 13).getValues();
    for (const row of foldersData) {
      if (row[9] === 'TRUE') { // Violation列
        violations.push({
          type: 'DriveFolder',
          resourceId: row[0],
          resourceName: row[1],
          fullPath: row[4],
          violationType: row[10],
          violationMessage: row[11],
          matchedRule: row[12],
          severity: determineSeverity(row[10]),
          timestamp: now
        });
      }
    }
  }
  
  return violations;
}

// =================================================================
// 違反情報マージ
// =================================================================
function mergeViolations(existingViolations, currentViolations) {
  const mergedViolations = [];
  const currentViolationKeys = new Set();
  const now = new Date();
  
  // 現在の違反を処理
  for (const violation of currentViolations) {
    const key = `${violation.type}_${violation.resourceId}`;
    currentViolationKeys.add(key);
    
    if (existingViolations.has(key)) {
      // 既存の違反（ONGOING）
      const existing = existingViolations.get(key);
      mergedViolations.push({
        ...existing,
        lastConfirmed: now,
        status: 'ONGOING',
        violationType: violation.violationType,
        violationMessage: violation.violationMessage,
        severity: violation.severity
      });
    } else {
      // 新規違反（NEW）
      mergedViolations.push({
        violationId: generateViolationId(),
        ...violation,
        firstDetected: now,
        lastConfirmed: now,
        status: 'NEW',
        resolvedDate: '',
        notes: ''
      });
    }
  }
  
  // 解決された違反を処理
  for (const [key, existing] of existingViolations) {
    if (!currentViolationKeys.has(key) && existing.status !== 'RESOLVED') {
      mergedViolations.push({
        ...existing,
        status: 'RESOLVED',
        resolvedDate: now,
        notes: '自動解決確認'
      });
    } else if (existing.status === 'RESOLVED') {
      // 解決済みの違反はそのまま保持
      mergedViolations.push(existing);
    }
  }
  
  return mergedViolations;
}

// =================================================================
// 違反情報のシート保存
// =================================================================
function saveViolationsToSheet(violationsSheet, violations) {
  // 既存データをクリア
  if (violationsSheet.getLastRow() > 1) {
    violationsSheet.getRange(2, 1, violationsSheet.getLastRow() - 1, violationsSheet.getLastColumn()).clear();
  }
  
  if (violations.length === 0) return;
  
  // データを整形
  const data = violations.map(v => [
    v.violationId,
    v.type,
    v.resourceId,
    v.resourceName,
    v.fullPath,
    v.violationType,
    v.violationMessage,
    v.matchedRule,
    v.severity,
    v.firstDetected,
    v.lastConfirmed,
    v.status,
    v.resolvedDate || '',
    v.notes || ''
  ]);
  
  // シートに書き込み
  violationsSheet.getRange(2, 1, data.length, data[0].length).setValues(data);
  
  // ステータスで色分け
  for (let i = 0; i < data.length; i++) {
    const statusCell = violationsSheet.getRange(i + 2, 12);
    const status = data[i][11];
    
    if (status === 'NEW') {
      statusCell.setBackground('#FFE0E0'); // 赤系
    } else if (status === 'ONGOING') {
      statusCell.setBackground('#FFF0E0'); // オレンジ系
    } else if (status === 'RESOLVED') {
      statusCell.setBackground('#E0FFE0'); // 緑系
    }
  }
}

// =================================================================
// ユーティリティ関数
// =================================================================
function generateViolationId() {
  return `VIO-${new Date().getTime()}-${Math.random().toString(36).substr(2, 5).toUpperCase()}`;
}

function determineSeverity(violationType) {
  // ルールからSeverityを判定（将来的に拡張可能）
  return violationType === 'RULE_VIOLATION' ? 'ERROR' : 'WARN';
}

// =================================================================
// 違反サマリ取得
// =================================================================
function getViolationsSummary() {
  const spreadsheet = SpreadsheetApp.openById(getSpreadsheetId());
  const violationsSheet = spreadsheet.getSheetByName('Violations_Log');
  
  if (!violationsSheet || violationsSheet.getLastRow() <= 1) {
    return {
      total: 0,
      new: 0,
      ongoing: 0,
      resolved: 0,
      byType: {},
      bySeverity: {}
    };
  }
  
  const data = violationsSheet.getRange(2, 1, violationsSheet.getLastRow() - 1, 14).getValues();
  const summary = {
    total: 0,
    new: 0,
    ongoing: 0,
    resolved: 0,
    byType: {
      Slack: 0,
      Drive: 0,
      DriveFolder: 0
    },
    bySeverity: {
      ERROR: 0,
      WARN: 0
    }
  };
  
  for (const row of data) {
    const status = row[11];
    const type = row[1];
    const severity = row[8];
    
    if (status !== 'RESOLVED') {
      summary.total++;
      
      if (status === 'NEW') summary.new++;
      if (status === 'ONGOING') summary.ongoing++;
      
      if (type) summary.byType[type] = (summary.byType[type] || 0) + 1;
      if (severity) summary.bySeverity[severity] = (summary.bySeverity[severity] || 0) + 1;
    } else {
      summary.resolved++;
    }
  }
  
  return summary;
}