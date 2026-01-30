/**
 * Slack AI Bot - 完全統合版 (All-in-One)
 * Version: 3.0
 * 
 * 統合された機能:
 * - 初期セットアップ機能（ステップバイステップ）
 * - Slack Bot本体（メッセージ処理、スレッド対応）
 * - ファイル処理（PDF、Word、Googleドキュメント）
 * - ドキュメント編集・レビュー機能
 * - FAQ検索・Drive検索
 * - Natural Language API連携
 * - Mermaidグラフ対応
 * - デバッグ・ログ機能
 */

// =====================================
// URL Verification Challenge処理
// =====================================

/**
 * Slackチャンネル一覧を取得
 * Botが参加しているチャンネルを表示
 */
function testListChannels() {
  console.log('========================================');
  console.log('Botが参加しているチャンネル一覧');
  console.log('========================================\n');
  
  const config = Settings();
  if (!config?.SLACK_TOKEN) {
    console.log('❌ SLACK_TOKENが設定されていません');
    return;
  }
  
  const url = 'https://slack.com/api/conversations.list';
  const payload = {
    token: config.SLACK_TOKEN,
    types: 'public_channel,private_channel',
    limit: 100
  };
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      payload: payload,
      muteHttpExceptions: true
    });
    
    const data = JSON.parse(response.getContentText());
    
    if (!data.ok) {
      console.log('❌ エラー:', data.error);
      if (data.error === 'missing_scope') {
        console.log('\n必要なスコープ: channels:read');
        console.log('Slack App設定でスコープを追加してください');
      }
      return;
    }
    
    console.log('Botが参加しているチャンネル:\n');
    
    let botChannels = [];
    data.channels.forEach(channel => {
      if (channel.is_member) {
        botChannels.push(channel);
        console.log(`✅ ${channel.name} (ID: ${channel.id})`);
      }
    });
    
    if (botChannels.length === 0) {
      console.log('❌ Botはどのチャンネルにも参加していません\n');
      console.log('解決方法:');
      console.log('1. Slackでチャンネルを開く');
      console.log('2. /invite @YourBotName を実行');
      console.log('3. このテストを再実行');
    } else {
      console.log(`\n合計 ${botChannels.length} チャンネルに参加中`);
      console.log('\n上記のIDをtestDirectPost()のTEST_CHANNELに設定してください');
    }
    
  } catch (e) {
    console.log('❌ エラー:', e.toString());
  }
}

/**
 * Slack Challenge テスト関数
 * GASエディタから実行してChallenge処理をテスト
 */
function testSlackChallenge() {
  console.log('========================================');
  console.log('Slack URL Verification Challenge テスト');
  console.log('========================================\n');
  
  // テスト用のChallengeリクエストを作成
  const testChallenge = 'test_challenge_' + Date.now();
  const testRequest = {
    postData: {
      contents: JSON.stringify({
        token: 'test_token',
        challenge: testChallenge,
        type: 'url_verification'
      })
    }
  };
  
  console.log('送信するchallenge: ' + testChallenge);
  
  // doPost関数をテスト
  try {
    const response = doPost(testRequest);
    const responseText = response.getContent();
    
    console.log('受信したレスポンス: ' + responseText);
    
    if (responseText === testChallenge) {
      console.log('\n✅ テスト成功！');
      console.log('SlackのEvent Subscriptionsで URL を設定できます。');
    } else {
      console.log('\n❌ テスト失敗');
      console.log('期待値: ' + testChallenge);
      console.log('実際値: ' + responseText);
    }
  } catch (error) {
    console.log('\n❌ エラー発生: ' + error.toString());
  }
  
  console.log('\n========================================');
  console.log('トラブルシューティング:');
  console.log('1. GASを新しくデプロイしましたか？');
  console.log('2. デプロイの設定は「全員」にアクセス可能ですか？');
  console.log('3. 新しいWeb App URLを使用していますか？');
  console.log('========================================');
}

// =====================================
// セクション1: 初期セットアップ関数
// =====================================

/**
 * 【最初に実行】ステップ1: スプレッドシートIDを手動で設定
 * 
 * 1. Google Driveで新しいスプレッドシートを作成
 * 2. URLからIDをコピー（/d/と/editの間の文字列）
 * 3. 下記のSPREADSHEET_IDに貼り付け
 * 4. この関数を実行
 */
function setupStep1_SetSpreadsheetId() {
  // ★★★ ここにスプレッドシートIDを入力してください ★★★
  const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';
  
  if (SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE') {
    console.log('❌ エラー: SPREADSHEET_IDを設定してください');
    console.log('\n手順:');
    console.log('1. Google Driveで新しいスプレッドシートを作成');
    console.log('2. スプレッドシートを開く');
    console.log('3. URLから以下の部分をコピー:');
    console.log('   https://docs.google.com/spreadsheets/d/【ここの部分】/edit');
    console.log('4. コピーしたIDを上記のSPREADSHEET_IDに貼り付け');
    console.log('5. この関数を再度実行');
    return;
  }
  
  // スクリプトプロパティに保存
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', SPREADSHEET_ID);
  console.log('✅ スプレッドシートIDを保存しました: ' + SPREADSHEET_ID);
  
  // スプレッドシートを開いてみる
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log('✅ スプレッドシートに接続成功');
    console.log('   名前: ' + ss.getName());
    console.log('   URL: ' + ss.getUrl());
    console.log('\n次のステップ: setupStep2_InitializeSheets() を実行してください');
  } catch (e) {
    console.log('❌ スプレッドシートを開けません: ' + e.toString());
    console.log('IDが正しいか、アクセス権限があるか確認してください');
  }
}

/**
 * ステップ2: シートを初期化
 */
function setupStep2_InitializeSheets() {
  const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  
  if (!SPREADSHEET_ID) {
    console.log('❌ エラー: 先にsetupStep1_SetSpreadsheetId()を実行してください');
    return;
  }
  
  try {
    console.log('スプレッドシートを開いています...');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    console.log('✅ スプレッドシート接続成功: ' + ss.getName());
    
    // 既存のシートを確認
    const sheets = ss.getSheets();
    console.log('\n現在のシート数: ' + sheets.length);
    
    // logシートの作成または確認
    console.log('\n1. logシートを設定中...');
    let logSheet = ss.getSheetByName('log');
    if (!logSheet) {
      if (sheets.length > 0 && sheets[0].getName() === 'シート1') {
        sheets[0].setName('log');
        logSheet = sheets[0];
        console.log('   デフォルトシートをlogに変更');
      } else {
        logSheet = ss.insertSheet('log');
        console.log('   logシートを作成');
      }
    }
    
    if (logSheet.getLastRow() === 0) {
      logSheet.appendRow(['Timestamp', 'Message']);
      logSheet.getRange('1:1').setFontWeight('bold');
      logSheet.setFrozenRows(1);
    }
    
    // faqシートの作成
    console.log('\n2. faqシートを設定中...');
    let faqSheet = ss.getSheetByName('faq');
    if (!faqSheet) {
      faqSheet = ss.insertSheet('faq');
      faqSheet.appendRow(['キーワード', '回答', '検索フラグ', 'Drive検索結果']);
      faqSheet.getRange('1:1').setFontWeight('bold');
      faqSheet.setFrozenRows(1);
      faqSheet.setColumnWidth(1, 150);
      faqSheet.setColumnWidth(2, 400);
      faqSheet.setColumnWidth(3, 100);
      faqSheet.setColumnWidth(4, 400);
      
      // サンプルFAQデータを追加
      const sampleFAQs = [
        ['休暇申請', '休暇申請はシステムから申請してください。上長の承認が必要です。', false, ''],
        ['経費精算', '経費精算は月末までに申請書を提出してください。領収書の添付が必要です。', false, ''],
        ['会議室予約', '会議室の予約はGoogleカレンダーから行えます。', false, ''],
        ['VPN接続', 'VPNの設定方法は社内Wikiの「ITサポート」ページを参照してください。', false, ''],
        ['パスワード変更', 'パスワードは90日ごとに変更が必要です。システム設定から変更できます。', false, '']
      ];
      
      sampleFAQs.forEach((faq, index) => {
        faqSheet.getRange(index + 2, 1, 1, 4).setValues([faq]);
      });
      
      console.log('   faqシートを作成し、サンプルデータを追加');
    } else {
      console.log('   faqシートは既に存在します');
    }
    
    // ドライブ一覧シートの作成
    console.log('\n3. ドライブ一覧シートを設定中...');
    let driveSheet = ss.getSheetByName('ドライブ一覧');
    if (!driveSheet) {
      driveSheet = ss.insertSheet('ドライブ一覧');
      driveSheet.appendRow(['フォルダID', 'フォルダ名', '説明']);
      driveSheet.getRange('1:1').setFontWeight('bold');
      driveSheet.setFrozenRows(1);
      driveSheet.setColumnWidth(1, 300);
      driveSheet.setColumnWidth(2, 200);
      driveSheet.setColumnWidth(3, 300);
      console.log('   ドライブ一覧シートを作成');
    }
    
    // debug_logシートの作成
    console.log('\n4. debug_logシートを設定中...');
    let debugSheet = ss.getSheetByName('debug_log');
    if (!debugSheet) {
      debugSheet = ss.insertSheet('debug_log');
      debugSheet.appendRow(['Timestamp', 'Category', 'Message', 'Data']);
      debugSheet.getRange('1:1').setFontWeight('bold');
      debugSheet.setFrozenRows(1);
      debugSheet.setColumnWidth(1, 150);
      debugSheet.setColumnWidth(2, 100);
      debugSheet.setColumnWidth(3, 300);
      debugSheet.setColumnWidth(4, 400);
      console.log('   debug_logシートを作成');
    }
    
    console.log('\n✅ スプレッドシートの初期化完了！');
    console.log('URL: ' + ss.getUrl());
    console.log('\n次のステップ: setupStep3_SetAPIKeys() を実行してください');
    
  } catch (e) {
    console.log('❌ エラー: ' + e.toString());
  }
}

/**
 * ステップ3: APIキーを設定確認
 */
function setupStep3_SetAPIKeys() {
  console.log('========================================');
  console.log('APIキーの設定確認');
  console.log('========================================');
  
  const props = PropertiesService.getScriptProperties().getProperties();
  
  console.log('\n現在の設定:');
  console.log('✅ SPREADSHEET_ID: ' + (props.SPREADSHEET_ID ? '設定済み' : '未設定'));
  console.log((props.SLACK_TOKEN ? '✅' : '❌') + ' SLACK_TOKEN: ' + (props.SLACK_TOKEN ? '設定済み' : '未設定'));
  console.log((props.OPEN_AI_TOKEN ? '✅' : '❌') + ' OPEN_AI_TOKEN: ' + (props.OPEN_AI_TOKEN ? '設定済み' : '未設定'));
  console.log('   OPEN_AI_MODEL: ' + (props.OPEN_AI_MODEL ? props.OPEN_AI_MODEL : '未設定（既定: gpt-5）'));
  console.log('   GEMINI_TOKEN: ' + (props.GEMINI_TOKEN ? '設定済み（オプション）' : '未設定（オプション）'));
  console.log('   GOOGLE_NL_API: ' + (props.GOOGLE_NL_API ? '設定済み（オプション）' : '未設定（オプション）'));
  
  if (!props.SLACK_TOKEN || !props.OPEN_AI_TOKEN) {
    console.log('\n⚠️ 必要なAPIキーが設定されていません');
    console.log('\n設定方法:');
    console.log('1. プロジェクトの設定 → スクリプト プロパティ');
    console.log('2. 「プロパティを追加」をクリック');
    console.log('3. 以下のプロパティを追加:');
    console.log('   - SLACK_TOKEN: Slack Bot Token');
    console.log('   - OPEN_AI_TOKEN: OpenAI API Key');
    console.log('   - （任意）OPEN_AI_MODEL: 使用モデル名 例) gpt-5, gpt-4o');
  } else {
    console.log('\n✅ 必要なAPIキーはすべて設定されています');
    console.log('\n次のステップ: setupStep4_TestConnection() を実行してください');
  }
}

/**
 * FAQ機能のテスト
 */
function testFAQ() {
  console.log('========================================');
  console.log('FAQ機能テスト');
  console.log('========================================\n');
  
  const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!ssId) {
    console.log('❌ SPREADSHEET_IDが設定されていません');
    console.log('setupStep2_CreateSpreadsheet() を実行してください');
    return;
  }
  
  try {
    const ss = SpreadsheetApp.openById(ssId);
    console.log('✅ スプレッドシートにアクセスできました');
    
    let faqSheet = ss.getSheetByName('faq');
    if (!faqSheet) {
      console.log('⚠️ FAQシートが存在しません。作成します...');
      faqSheet = ss.insertSheet('faq');
      
      // サンプルFAQデータを追加
      const sampleData = [
        ['キーワード', '回答'],
        ['休暇', '休暇申請はシステムから申請してください。上長の承認が必要です。'],
        ['経費', '経費精算は月末までに申請書を提出してください。領収書の添付が必要です。'],
        ['会議室', '会議室の予約はGoogleカレンダーから行えます。'],
        ['VPN', 'VPNの設定方法は社内Wikiの「ITサポート」ページを参照してください。']
      ];
      
      faqSheet.getRange(1, 1, sampleData.length, 2).setValues(sampleData);
      console.log('✅ FAQシートを作成し、サンプルデータを追加しました');
    }
    
    // FAQデータを表示
    const faqs = faqSheet.getRange('A:B').getValues()
      .filter(row => !row.every(cell => cell.toString().trim() === ''));
    
    console.log('\n現在のFAQデータ:');
    console.log('================');
    faqs.forEach((row, i) => {
      if (i === 0) {
        console.log(`[ヘッダー] ${row[0]} | ${row[1]}`);
      } else if (row[0] && row[1]) {
        console.log(`${i}. [${row[0]}] => ${row[1].substring(0, 50)}...`);
      }
    });
    
    // テスト質問でFAQ検索をテスト
    console.log('\nテスト質問でFAQ検索をテスト:');
    console.log('================');
    
    // 実際のFAQデータに合わせたテスト質問
    const testQuestions = [
      'IR情報の開示手順',
      '説明会資料のアドレス',
      'HPに開示',
      '休暇を取りたい'  // FAQにないケース
    ];
    
    testQuestions.forEach(question => {
      console.log(`\n質問: "${question}"`);
      const faqRole = getFaqRole(question);
      if (faqRole) {
        console.log('✅ FAQマッチあり');
        console.log('システムプロンプト:');
        console.log(faqRole.content.substring(0, 200) + '...');
      } else {
        console.log('❌ FAQマッチなし');
      }
    });
    
    console.log('\nスプレッドシートURL:');
    console.log(ss.getUrl());
    console.log('\n✨ FAQシートにキーワードと回答を追加してください');
    
  } catch (e) {
    console.log('❌ エラー:', e.toString());
  }
}

// ===========================
// Drive検索機能
// ===========================

/**
 * FAQシートの A列キーワード、C列チェックを元に、
 * ドライブ一覧シート A列のすべてのフォルダIDを検索対象として
 * 指定キーワードの含まれるファイル本文／行を抜き出し、
 * D列に結果をリッチテキストで書き出します。
 */
function updateFaqDriveResults() {
  const ss             = SpreadsheetApp.getActive();
  const faqSheet       = ss.getSheetByName('faq');
  const driveListSheet = ss.getSheetByName('ドライブ一覧');
  const lastFaqRow     = faqSheet.getLastRow();
  const lastDriveRow   = driveListSheet.getLastRow();
  if (lastFaqRow < 2 || lastDriveRow < 2) return;

  // ドライブ一覧シート A2:A に書かれたフォルダIDを取得
  const folderIds = driveListSheet
    .getRange(`A2:A${lastDriveRow}`)
    .getValues()
    .flat()
    .filter(id => id);

  // FAQシートの A列キーワード、C列検索フラグを取得
  const faqData = faqSheet.getRange(`A2:C${lastFaqRow}`).getValues();

  faqData.forEach((row, i) => {
    const [ keyword, /*manualAnswer*/, doSearch ] = row;
    const rowNum = i + 2;
    const resultCell = faqSheet.getRange(rowNum, 4); // D列

    if (doSearch === true) {
      // C列が TRUE の場合のみ、全フォルダを検索して結果をまとめる
      let allResults = [];
      folderIds.forEach(folderId => {
        const res = searchDriveLinkReturn(keyword, folderId);
        allResults = allResults.concat(res);
      });
      // 抜き出した結果を D列に書き込む
      if (allResults.length) {
        cellSetLink(resultCell, allResults);
      } else {
        resultCell.setValue('該当ファイルが見つかりませんでした');
      }
    } else {
      // C列が FALSE の場合は D列をクリア
      resultCell.clearContent();
    }
  });
}

/**
 * Drive API で指定フォルダ内を全文検索し、
 * ファイルごとに段落／行を抜き出して配列で返す
 */
function searchDriveLinkReturn(keyword, folderId) {
  const ret = [];
  const baseUrl = 'https://www.googleapis.com/drive/v3/files';
  const params = {
    q: `'${folderId}' in parents and trashed = false and fullText contains '${keyword}'`,
    corpora: 'allDrives',
    includeItemsFromAllDrives: true,
    supportsAllDrives: true,
    fields: 'files(id,name,mimeType,webViewLink)'
  };
  const query = Object.entries(params)
    .map(([k,v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
    .join('&');
  const url = `${baseUrl}?${query}`;
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    const files = JSON.parse(response.getContentText()).files || [];
    
    files.forEach(file => {
      let snippets = [];
      try {
        if (file.mimeType === 'application/vnd.google-apps.document') {
          snippets = extractSnippetFromDoc(file.id, keyword);
        } else if (file.mimeType === 'application/vnd.google-apps.spreadsheet') {
          snippets = extractSnippetFromSheet(file.id, keyword);
        } else if (file.mimeType === 'application/pdf') {
          // PDF を Docs に変換して抜き出す
          const blob = DriveApp.getFileById(file.id).getBlob();
          const tmpFile = DriveApp.createFile(blob).setName('temp');
          const resource = { title: 'temp-doc', mimeType: MimeType.GOOGLE_DOCS };
          const converted = Drive.Files.insert(resource, tmpFile.getBlob());
          snippets = extractSnippetFromDoc(converted.id, keyword);
          DriveApp.getFileById(converted.id).setTrashed(true);
          tmpFile.setTrashed(true);
        }
      } catch (e) {
        debugLog('Drive', `処理エラー (${file.name})`, e.toString());
      }
      if (snippets.length > 0) {
        ret.push({ file: file, snippets: snippets });
      }
    });
  } catch (e) {
    debugLog('Drive', 'Search error', e.toString());
  }
  
  return ret;
}

/**
 * Googleドキュメントからキーワードを含む段落を抽出
 */
function extractSnippetFromDoc(docId, keyword) {
  const paras = DocumentApp.openById(docId).getBody().getParagraphs();
  return paras
    .map(p => p.getText().trim())
    .filter(t => t.includes(keyword));
}

/**
 * スプレッドシートからキーワードを含む行を抽出 (タブ区切り)
 */
function extractSnippetFromSheet(sheetId, keyword) {
  const rows = SpreadsheetApp.openById(sheetId)
    .getSheets()
    .flatMap(sh => sh.getDataRange().getValues());
  return rows
    .filter(r => r.some(c => c.toString().includes(keyword)))
    .map(r => r.join('\t'));
}

/**
 * 結果をリッチテキスト (リンク付き) でセルに書き込む
 */
function cellSetLink(range, data) {
  const maxLen = 5000;
  let text = '';
  const links = [];
  
  data.forEach(item => {
    const nameBlock = item.file.name + '\n';
    const snippetBlock = item.snippets
      .slice(0, 2)
      .map(s => s.replace(/\t/g,' ').replace(/\n/g,' ').trim())
      .join('\n') + '\n';
    
    let block = nameBlock + snippetBlock;
    if (text.length + block.length > maxLen) {
      block = block.substring(0, maxLen - text.length);
    }
    
    const start = text.length;
    text += block;
    const end = start + nameBlock.length;
    if (end <= maxLen) {
      links.push({ start, end, url: item.file.webViewLink });
    }
  });
  
  const builder = SpreadsheetApp.newRichTextValue().setText(text);
  links.forEach(l => builder.setLinkUrl(l.start, l.end, l.url));
  range.setRichTextValue(builder.build());
}

/**
 * ステップ4: 接続テスト
 */
function setupStep4_TestConnection() {
  console.log('========================================');
  console.log('接続テスト');
  console.log('========================================');
  
  // スプレッドシート接続テスト
  console.log('\n1. スプレッドシート接続テスト...');
  const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  try {
    const ss = SpreadsheetApp.openById(ssId);
    console.log('✅ スプレッドシート接続成功');
    const debugSheet = ss.getSheetByName('debug_log');
    if (debugSheet) {
      debugSheet.appendRow([new Date(), 'Test', 'Connection test', 'Success']);
      console.log('✅ テストデータ書き込み成功');
    }
  } catch (e) {
    console.log('❌ スプレッドシート接続失敗: ' + e.toString());
    return;
  }
  
  // Slack接続テスト
  console.log('\n2. Slack API接続テスト...');
  const slackToken = PropertiesService.getScriptProperties().getProperty('SLACK_TOKEN');
  if (slackToken) {
    try {
      const url = 'https://slack.com/api/auth.test';
      const response = UrlFetchApp.fetch(url, {
        method: 'post',
        payload: { token: slackToken },
        muteHttpExceptions: true
      });
      const data = JSON.parse(response.getContentText());
      
      if (data.ok) {
        console.log('✅ Slack接続成功');
        console.log('   ユーザー: ' + data.user);
        console.log('   チーム: ' + data.team);
      } else {
        console.log('❌ Slack接続失敗: ' + data.error);
      }
    } catch (e) {
      console.log('❌ Slackテストエラー: ' + e.toString());
    }
  }
  
  console.log('\n✅ セットアップ完了！');
  console.log('\n次のステップ:');
  console.log('1. デプロイ → 新しいデプロイ → ウェブアプリ');
  console.log('2. Web App URLをコピー');
  console.log('3. Slack Appの設定でURLを登録');
}

// ===========================
// グローバル設定
// ===========================

const DEBUG_MODE = true;
const DEBUG_SHEET_NAME = 'debug_log';

// ===========================
// デバッグ機能
// ===========================

/**
 * アクティブなスプレッドシートを取得
 */
function getActiveSpreadsheet() {
  const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!ssId) {
    throw new Error('SPREADSHEET_ID not configured. Run setupStep1_SetSpreadsheetId() first');
  }
  
  try {
    return SpreadsheetApp.openById(ssId);
  } catch (e) {
    console.log('Error opening spreadsheet:', e.toString());
    throw new Error('Cannot open spreadsheet: ' + e.toString());
  }
}

function debugLog(category, message, data = null) {
  // 引数チェック
  if (!category || category === undefined) {
    console.log('debugLog: category is undefined');
    return;
  }
  if (!message || message === undefined) {
    console.log('debugLog: message is undefined');
    return;
  }
  
  if (!DEBUG_MODE) return;
  
  // コンソールに出力
  if (data !== null && data !== undefined) {
    console.log(`[${category}] ${message}`, data);
  } else {
    console.log(`[${category}] ${message}`);
  }
  
  Logger.log(`[${category}] ${message} ${data ? JSON.stringify(data) : ''}`);
  
  // スプレッドシートへの記録
  try {
    const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!ssId) {
      console.log('debugLog: SPREADSHEET_ID not found');
      return;
    }
    
    const ss = SpreadsheetApp.openById(ssId);
    let debugSheet = ss.getSheetByName(DEBUG_SHEET_NAME);
    
    if (debugSheet) {
      debugSheet.appendRow([
        new Date(),
        String(category),
        String(message),
        data ? JSON.stringify(data) : ''
      ]);
    }
  } catch (e) {
    Logger.log('Debug log error: ' + e.toString());
  }
}

// ===========================
// 設定管理
// ===========================

function Settings() {
  try {
    const env = PropertiesService.getScriptProperties().getProperties();
    const required = ['SLACK_TOKEN', 'OPEN_AI_TOKEN', 'SLACK_BOT_USER_ID'];
    const missing = required.filter(key => !env[key]);

    if (missing.length > 0) {
      debugLog('Settings', 'Missing required properties', missing);
      // SLACK_BOT_USER_IDが未設定の場合のヘルプメッセージ
      if (missing.includes('SLACK_BOT_USER_ID')) {
        console.log('SLACK_BOT_USER_ID取得方法:');
        console.log('1. Slack Appの設定ページでOAuth & Permissionsセクションを開く');
        console.log('2. Bot User OAuth Tokenの下にある「Bot User ID」をコピー');
        console.log('3. GASのプロジェクト設定 > スクリプトプロパティに追加');
      }
      throw new Error(`Missing required properties: ${missing.join(', ')}`);
    }

    debugLog('Settings', 'Properties loaded successfully', Object.keys(env));
    return env;
  } catch (e) {
    debugLog('Settings', 'Error loading properties', e.toString());
    throw e;
  }
}

// ===========================
// ユーティリティ関数
// ===========================

function katakanaToHiragana(text) {
  return text.replace(/[\u30a1-\u30f6]/g, function(match) {
    var chr = match.charCodeAt(0) - 0x60;
    return String.fromCharCode(chr);
  });
}

function toHalfWidth(str) {
  str = str.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
  return str;
}

// ===========================
// Slack Bot クラス
// ===========================

class SlackBot {
  constructor(e) {
    debugLog('SlackBot', 'Constructor called', e);
    this.requestEvent = e;
    this.postData = null;
    this.slackEvent = null;
    this.responseData = this.init();
    this.verification();
  }

  responseJsonData(json) {
    debugLog('SlackBot', 'Response JSON', json);
    return ContentService.createTextOutput(JSON.stringify(json)).setMimeType(ContentService.MimeType.JSON);
  }

  init() {
    const e = this.requestEvent;
    debugLog('SlackBot', 'Init started', { hasPostData: !!e?.postData });
    
    if (!e?.postData) {
      const error = { error: 'postData is missing or undefined.' };
      debugLog('SlackBot', 'No postData', error);
      return error;
    }
    
    try {
      const contents = e.postData.contents;
      debugLog('SlackBot', 'PostData contents', contents);
      
      this.postData = JSON.parse(contents);
      debugLog('SlackBot', 'Parsed postData', this.postData);
      
      if (this.postData.type === 'url_verification') {
        debugLog('SlackBot', 'URL verification detected');
        return null;
      }
      
      if (this.postData.type === 'event_callback') {
        this.slackEvent = this.postData;
        debugLog('SlackBot', 'Event callback detected', this.slackEvent);
        return null;
      }
      
      const error = { error: 'Unknown event type', type: this.postData.type };
      debugLog('SlackBot', 'Unknown event type', error);
      return error;
      
    } catch (error) {
      const err = { error: 'Invalid JSON format in postData contents.', details: error.toString() };
      debugLog('SlackBot', 'JSON parse error', err);
      return err;
    }
  }

  verification() {
    if (!this.postData || this.responseData) return null;
    
    if (this.postData.type === 'url_verification') {
      this.responseData = { "challenge": this.postData.challenge };
      debugLog('SlackBot', 'URL verification response', this.responseData);
      return this.responseData;
    }
    return null;
  }

  hasCache(key) {
    if (!key) return true;
    const cache = CacheService.getScriptCache();
    const cached = cache.get(key);
    if (cached) {
      debugLog('SlackBot', 'Cache hit', key);
      return true;
    }
    cache.put(key, 'true', 30 * 60);
    debugLog('SlackBot', 'Cache miss, stored', key);
    return false;
  }

  handleEvent(type, callback = () => { }) {
    debugLog('SlackBot', 'HandleEvent', { type, hasEvent: !!this.slackEvent });
    
    if (!this.slackEvent || this.responseData) return null;
    
    const event = this.slackEvent?.event;
    if (!event || event.type !== type) {
      debugLog('SlackBot', 'Event type mismatch', { expected: type, actual: event?.type });
      return null;
    }
    
    const callbackResponse = callback({ event });
    if (callbackResponse) {
      this.responseData = callbackResponse;
      debugLog('SlackBot', 'Callback response set', callbackResponse);
    }
    return callbackResponse;
  }

  handleBase(type, targetType, callback = () => {}) {
    return this.handleEvent(type, ({ event }) => {
      debugLog('SlackBot', 'HandleBase', { type, targetType, event });
      
      const { text: message, channel, thread_ts: threadTs, ts, client_msg_id, bot_id, app_id } = event;
      
      if (bot_id || app_id) {
        debugLog('SlackBot', 'Bot message ignored', { bot_id, app_id });
        return null;
      }
      
      if (event.type !== targetType) {
        debugLog('SlackBot', 'Type mismatch in handleBase', { expected: targetType, actual: event.type });
        return null;
      }
      
      const cacheKey = `${channel}:${client_msg_id}`;
      if (this.hasCache(cacheKey)) {
        debugLog('SlackBot', 'Duplicate message ignored', cacheKey);
        return null;
      }
      
      return callback ? callback({ message, channel, threadTs: threadTs ?? ts, event }) : null;
    });
  }

  handleMentionEventBase(callback) {
    debugLog('SlackBot', 'HandleMentionEventBase called');
    // messageイベントとapp_mentionイベントの両方を処理
    const mentionResult = this.handleBase("app_mention", "app_mention", (args) => {
      // Bot自身のIDを取得
      const botUserId = Settings().SLACK_BOT_USER_ID;
      if (!botUserId) {
        debugLog('SlackBot', 'Bot User ID not set - skipping app_mention');
        return null;
      }

      // Bot自身へのメンションが含まれているかチェック
      const botMentionPattern = `<@${botUserId}>`;
      if (args.message && args.message.includes(botMentionPattern)) {
        debugLog('SlackBot', 'app_mention contains bot mention', {
          botUserId: botUserId,
          message: args.message.substring(0, 50)
        });
        return callback(args);
      }

      // 他のユーザーへのメンションの場合は無視
      debugLog('SlackBot', 'app_mention for different user - ignoring', {
        message: args.message ? args.message.substring(0, 50) : 'no message'
      });
      return null;
    });
    if (mentionResult) return mentionResult;

    // messageイベントでメンションが含まれている場合も処理
    return this.handleBase("message", "message", (args) => {
      // Bot自身のIDを取得
      const botUserId = Settings().SLACK_BOT_USER_ID;
      if (!botUserId) {
        debugLog('SlackBot', 'Bot User ID not set - skipping message');
        return null;
      }

      // Bot自身へのメンションが含まれているかチェック
      const botMentionPattern = `<@${botUserId}>`;
      if (args.message && args.message.includes(botMentionPattern)) {
        debugLog('SlackBot', 'Message contains bot mention', {
          botUserId: botUserId,
          message: args.message.substring(0, 50)
        });
        return callback(args);
      }

      // Slack Connectの場合、チーム間で異なるBot IDの可能性があるため、
      // app_mentionイベント以外は無視
      debugLog('SlackBot', 'Message without bot mention - ignoring', {
        message: args.message ? args.message.substring(0, 50) : 'no message'
      });
      return null;
    });
  }

  response() {
    debugLog('SlackBot', 'Final response', this.responseData);
    
    if (this.responseData) {
      return this.responseJsonData(this.responseData);
    }
    
    return ContentService.createTextOutput('');
  }
}

// ===========================
// 自然言語処理 (Google Natural Language API)
// ===========================

/**
 * Google Natural Language API analyzeSyntax
 * テキストを形態素解析して品詞情報を取得
 */
function gNL(textdata) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GOOGLE_NL_API');
  if (!apiKey) {
    debugLog('NLP', 'No Google NL API key found');
    return null;
  }

  const url = "https://language.googleapis.com/v1/documents:analyzeSyntax?key=" + apiKey;
  const payload = {
    document: {
      type: "PLAIN_TEXT",
      content: textdata,
      language: "ja" // 日本語を明示的に指定
    },
    encodingType: "UTF8"
  };
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    if (result.error) {
      debugLog('NLP', 'API error', result.error.message);
      return null;
    }

    debugLog('NLP', 'Analysis success', { tokensCount: result.tokens?.length });
    return result;
  } catch(e) {
    debugLog('NLP', 'Exception in gNL', e.toString());
    return null;
  }
}

/**
 * Google Natural Language API の戻り値から指定品詞の単語を抽出
 * @param {Object} gNLobj - gNL関数の戻り値
 * @param {Array} tags - 抽出したい品詞タグの配列 ['NOUN','NUM','NUMBER']
 * https://cloud.google.com/natural-language/docs/morphology?hl=ja
 */
function filterGNL(gNLobj, tags) {
  if (!gNLobj || !gNLobj.tokens) return [];

  const words = gNLobj.tokens
    .filter(token => tags.includes(token.partOfSpeech.tag))
    .map(token => token.text.content);

  debugLog('NLP', 'Filtered words', { tags: tags, words: words });
  return words;
}

/**
 * スプレッドシート全体からテキストを抽出して自然言語処理
 * @param {string} spreadsheetId - スプレッドシートのID
 * @param {number} maxRows - 各シートから取得する最大行数（デフォルト: 100）
 */
function analyzeSpreadsheetWithNLP(spreadsheetId, maxRows = 100) {
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheets = ss.getSheets();
    let allText = [];
    let sheetContents = {};

    // 各シートからテキストを抽出
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const lastRow = Math.min(sheet.getLastRow(), maxRows);

      if (lastRow > 0) {
        const range = sheet.getRange(1, 1, lastRow, sheet.getLastColumn());
        const values = range.getValues();

        // シートごとのテキストを保存
        sheetContents[sheetName] = [];

        values.forEach((row, rowIndex) => {
          const rowText = row
            .filter(cell => cell !== null && cell !== '')
            .map(cell => String(cell).trim())
            .join(' ');

          if (rowText) {
            sheetContents[sheetName].push({
              row: rowIndex + 1,
              text: rowText
            });
            allText.push(rowText);
          }
        });
      }
    });

    // 全テキストを結合（最大5000文字に制限）
    const combinedText = allText.join('\n').substring(0, 5000);

    // 自然言語処理を実行
    const nlpResult = gNL(combinedText);

    if (!nlpResult) {
      return {
        success: false,
        error: 'Natural Language API failed',
        sheetContents: sheetContents
      };
    }

    // 名詞と数値を抽出
    const nouns = filterGNL(nlpResult, ['NOUN']);
    const numbers = filterGNL(nlpResult, ['NUM', 'NUMBER']);

    // 感情分析も実行
    const sentiment = analyzeSentiment(combinedText);

    return {
      success: true,
      sheetContents: sheetContents,
      analysis: {
        nouns: nouns,
        numbers: numbers,
        sentiment: sentiment,
        totalTokens: nlpResult.tokens?.length || 0
      }
    };

  } catch (e) {
    debugLog('NLP', 'Error in analyzeSpreadsheetWithNLP', e.toString());
    return {
      success: false,
      error: e.toString()
    };
  }
}

/**
 * テキストの感情分析
 */
function analyzeSentiment(text) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GOOGLE_NL_API');
  if (!apiKey) return null;

  const url = `https://language.googleapis.com/v1/documents:analyzeSentiment?key=${apiKey}`;

  const payload = {
    document: {
      type: 'PLAIN_TEXT',
      content: text,
      language: 'ja'
    },
    encodingType: 'UTF8'
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(response.getContentText());

    if (result.error) {
      debugLog('NLP', 'Sentiment analysis error', result.error.message);
      return null;
    }

    return {
      score: result.documentSentiment?.score,
      magnitude: result.documentSentiment?.magnitude
    };
  } catch (e) {
    debugLog('NLP', 'Sentiment analysis exception', e.toString());
    return null;
  }
}

// ===========================
// Slack API 関数
// ===========================

function getChannelInfo(channelId) {
  debugLog('API', 'Getting channel info', channelId);
  
  const url = 'https://slack.com/api/conversations.info';
  const config = Settings();
  
  if (!config?.SLACK_TOKEN) {
    debugLog('API', 'No Slack token');
    return null;
  }
  
  const payload = {
    token: config.SLACK_TOKEN,
    channel: channelId,
  };
  
  const options = {
    method: 'post',
    payload,
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (!data.ok) {
      debugLog('API', 'Channel info error', data.error);
      return null;
    }
    
    debugLog('API', 'Channel info success', data.channel?.name);
    return data.channel;
  } catch (e) {
    debugLog('API', 'Channel info exception', e.toString());
    return null;
  }
}

function getThreadMessages(channelId, threadTs) {
  debugLog('API', 'Getting thread messages', { channelId, threadTs });
  
  const url = 'https://slack.com/api/conversations.replies';
  const config = Settings();
  
  if (!config?.SLACK_TOKEN) return [];
  
  const payload = {
    token: config.SLACK_TOKEN,
    channel: channelId,
    ts: threadTs,
  };
  
  const options = {
    method: 'get',
    payload,
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (!data.ok) {
      debugLog('API', 'Thread messages error', data.error);
      return [];
    }
    
    debugLog('API', 'Thread messages success', data.messages?.length);
    return data.messages || [];
  } catch (e) {
    debugLog('API', 'Thread messages exception', e.toString());
    return [];
  }
}

/**
 * Wordドキュメントのレビュー結果をGoogleドキュメントに保存（簡略化版）
 */
function saveReviewToGoogleDoc(wordDocumentContext, reviewResult, userRequest) {
  try {
    debugLog('Review', 'Starting to save review to Google Doc');
    
    // 新しいGoogleドキュメントを作成
    const doc = DocumentApp.create('ドキュメントレビュー_' + new Date().getTime());
    const body = doc.getBody();
    
    // タイトルを追加
    const title = body.appendParagraph('📝 ドキュメントレビュー結果');
    title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    title.setBold(true);
    
    // 基本情報
    body.appendParagraph('レビュー日時: ' + new Date().toLocaleString('ja-JP'));
    if (wordDocumentContext.files && wordDocumentContext.files[0]) {
      body.appendParagraph('ファイル名: ' + wordDocumentContext.files[0].name);
    }
    
    body.appendHorizontalRule();
    
    // 比較表を作成
    const table = body.appendTable();
    
    // ヘッダー行
    const headerRow = table.appendTableRow();
    const header1 = headerRow.appendTableCell('元の文書内容');
    header1.setBackgroundColor('#f0f0f0');
    header1.setBold(true);
    const header2 = headerRow.appendTableCell('AIレビュー・改善提案');
    header2.setBackgroundColor('#e3f2fd');
    header2.setBold(true);
    
    // コンテンツ行
    const contentRow = table.appendTableRow();
    const originalCell = contentRow.appendTableCell();
    originalCell.appendParagraph(wordDocumentContext.originalContent || '内容が取得できませんでした');
    
    const reviewCell = contentRow.appendTableCell();
    reviewCell.appendParagraph(reviewResult || 'レビュー結果がありません');
    reviewCell.setBackgroundColor('#f5f5f5');
    
    // ドキュメントを保存
    doc.saveAndClose();
    
    // ドキュメントのIDを取得
    const docId = doc.getId();
    
    // 共有設定を適用
    try {
      const file = DriveApp.getFileById(docId);
      // 組織内でリンクを知っている人が閲覧可能に設定
      file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      debugLog('Review', 'Document sharing set to domain with link view');
    } catch (sharingError) {
      try {
        // フォールバック: リンクを知っている人が閲覧可能
        const file = DriveApp.getFileById(docId);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        debugLog('Review', 'Document sharing set to anyone with link view');
      } catch (fallbackError) {
        debugLog('Review', 'Failed to set document sharing', fallbackError.toString());
      }
    }
    
    const docUrl = `https://docs.google.com/document/d/${docId}/edit`;
    
    debugLog('Review', 'Review saved successfully', { docId: docId, url: docUrl });
    return docUrl;
    
  } catch (e) {
    debugLog('Review', 'Error in saveReviewToGoogleDoc', {
      error: e.toString(),
      stack: e.stack
    });
    return null;
  }
}

/**
 * ドキュメント作成と共有設定のテスト関数
 */
function testDocumentCreationAndSharing() {
  console.log('========================================');
  console.log('ドキュメント作成・共有設定テスト');
  console.log('========================================\n');
  
  try {
    // テスト用ドキュメントを作成
    console.log('1. テスト用ドキュメントを作成中...');
    const doc = DocumentApp.create('テスト_ドキュメント_' + new Date().getTime());
    const body = doc.getBody();
    
    // テスト内容を追加
    body.appendParagraph('これはドキュメント作成・共有設定のテストです。');
    body.appendParagraph('作成日時: ' + new Date().toLocaleString('ja-JP'));
    
    doc.saveAndClose();
    const docId = doc.getId();
    console.log('✅ ドキュメント作成成功: ' + docId);
    
    // 共有設定を適用
    console.log('\n2. 共有設定を適用中...');
    const file = DriveApp.getFileById(docId);
    
    try {
      // 組織内でリンクを知っている人が閲覧可能に設定
      file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      console.log('✅ 組織内リンク共有（閲覧）を設定しました');
    } catch (sharingError) {
      console.log('⚠️ 組織内共有に失敗、フォールバックを試行中...');
      try {
        // フォールバック: リンクを知っている人が閲覧可能
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        console.log('✅ 一般リンク共有（閲覧）を設定しました');
      } catch (fallbackError) {
        console.log('❌ 共有設定に失敗: ' + fallbackError.toString());
      }
    }
    
    // 結果を表示
    const docUrl = `https://docs.google.com/document/d/${docId}/edit`;
    console.log('\n3. 結果:');
    console.log('ドキュメントURL: ' + docUrl);
    console.log('ファイル名: ' + file.getName());
    
    // 現在の共有設定を確認
    try {
      const access = file.getSharingAccess();
      const permission = file.getSharingPermission();
      console.log('現在の共有設定:');
      console.log('  アクセス: ' + access);
      console.log('  権限: ' + permission);
    } catch (e) {
      console.log('共有設定の確認に失敗: ' + e.toString());
    }
    
    console.log('\n✅ テスト完了！上記URLにアクセスして確認してください。');
    
  } catch (e) {
    console.log('❌ テスト失敗: ' + e.toString());
    console.log('スタック: ' + e.stack);
  }
}

/**
 * Markdown記法をプレーンテキストに変換
 */
function convertMarkdownToSlack(message) {
  if (!message) return message;
  
  // **bold** を bold に変換（アスタリスクを完全に削除）
  let converted = message.replace(/\*\*(.+?)\*\*/g, '$1');
  
  // *italic* を italic に変換（シングルアスタリスクも削除）
  converted = converted.replace(/(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)/g, '$1');
  
  // _italic_ を italic に変換（アンダースコアも削除）
  converted = converted.replace(/\_(.+?)\_/g, '$1');
  
  // # ヘッダーを削除
  converted = converted.replace(/^#{1,6}\s+/gm, '');
  
  // - や * のリストマーカーを・に変換
  converted = converted.replace(/^[\*\-]\s+/gm, '・');
  
  // バックティック3つのコードブロックを保持（mermaid以外）
  // ```言語 を ``` に統一
  converted = converted.replace(/```(?!mermaid)\w*\n/g, '```\n');
  
  return converted;
}

/**
 * Mermaidコードを含むメッセージを処理してURLを追加
 */
function processMermaidInMessage(message) {
  if (!message) return message;
  
  // Mermaidコードブロックを検出
  const mermaidPattern = /```mermaid\n([\s\S]*?)```/g;
  let processedMessage = message;
  let matches = [];
  let match;
  
  while ((match = mermaidPattern.exec(message)) !== null) {
    matches.push({
      fullMatch: match[0],
      code: match[1].trim()
    });
  }
  
  // 各Mermaidコードブロックに対してURLを追加
  matches.forEach(mermaidMatch => {
    const mermaidCode = mermaidMatch.code;
    
    // Base64エンコード
    const base64Code = Utilities.base64Encode(mermaidCode, Utilities.Charset.UTF_8);
    
    // Mermaid.inkを使用して画像として表示
    const mermaidImageUrl = `https://mermaid.ink/img/${base64Code}`;
    
    // Mermaid LiveエディタのURLを生成（簡略化版）
    const mermaidLiveUrl = `https://mermaid-js.github.io/mermaid-live-editor/edit#base64:${base64Code}`;
    
    // メッセージにURLを追加
    const urlSection = `\n\n📊 Mermaidグラフを表示:\n• <${mermaidImageUrl}|画像として表示>\n• <${mermaidLiveUrl}|エディタで開く>`;
    
    // 元のMermaidコードブロックの後にURLを追加
    processedMessage = processedMessage.replace(
      mermaidMatch.fullMatch,
      mermaidMatch.fullMatch + urlSection
    );
  });
  
  return processedMessage;
}

function postMessage(message, channel, threadTs = null) {
  debugLog('API', 'Posting message START', { 
    channel: channel, 
    threadTs: threadTs, 
    messageLength: message?.length,
    channelType: typeof channel
  });
  
  const url = 'https://slack.com/api/chat.postMessage';
  const config = Settings();
  
  if (!config?.SLACK_TOKEN) {
    console.log('❌ SLACK_TOKEN not found in Script Properties');
    debugLog('API', 'No Slack token for posting');
    return false;
  }
  
  if (!channel) {
    console.log('❌ Channel is undefined or null');
    debugLog('API', 'No channel specified');
    return false;
  }
  
  // MarkdownをSlack形式に変換してからMermaidコードを処理
  const markdownConverted = convertMarkdownToSlack(message);
  const processedMessage = processMermaidInMessage(markdownConverted);
  
  const payload = {
    token: config.SLACK_TOKEN,
    channel: channel,
    text: processedMessage,
    unfurl_links: true,
    ...(threadTs ? { thread_ts: threadTs } : {}),
  };
  
  const options = {
    method: 'post',
    payload,
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseText = response.getContentText();
    const data = JSON.parse(responseText);
    
    if (!data.ok) {
      console.log('❌ Slack API Error:', data.error);
      console.log('Full response:', responseText);
      debugLog('API', 'Message post error', { 
        error: data.error, 
        response: data,
        channel: channel,
        needed_scope: data.needed,
        provided_scope: data.provided
      });
      return false;
    }
    
    console.log('✅ Message posted successfully to channel:', channel);
    debugLog('API', 'Message posted successfully', {
      ts: data.ts,
      channel: data.channel
    });
    return true;
  } catch (e) {
    console.log('❌ Exception during message post:', e.toString());
    debugLog('API', 'Message post exception', {
      error: e.toString(),
      channel: channel
    });
    return false;
  }
}

// ===========================
// ファイル処理機能
// ===========================

/**
 * Slackファイル情報を取得
 */
function getSlackFileInfo(fileId) {
  const config = Settings();
  if (!config?.SLACK_TOKEN) {
    debugLog('File', 'No Slack token');
    return null;
  }
  
  const url = 'https://slack.com/api/files.info';
  const payload = {
    token: config.SLACK_TOKEN,
    file: fileId
  };
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      payload: payload,
      muteHttpExceptions: true
    });
    
    const data = JSON.parse(response.getContentText());
    
    if (!data.ok) {
      debugLog('File', 'Failed to get file info', data.error);
      return null;
    }
    
    return data.file;
  } catch (e) {
    debugLog('File', 'Error getting file info', e.toString());
    return null;
  }
}

/**
 * Slackファイルをダウンロード
 */
function downloadSlackFile(url) {
  const config = Settings();
  if (!config?.SLACK_TOKEN) return null;
  
  try {
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + config.SLACK_TOKEN
      },
      muteHttpExceptions: true
    });
    
    return response.getBlob();
  } catch (e) {
    debugLog('File', 'Download error', e.toString());
    return null;
  }
}

/**
 * Word文書を処理してテキストを抽出（GoogleドキュメントIDも返す）
 */
function processWordDocument(blob, keepFile = false) {
  try {
    debugLog('File', 'Processing Word document', { keepFile });

    // Google Driveにアップロードして変換
    const file = Drive.Files.insert({
      title: 'Document_Review_' + new Date().getTime(),
      mimeType: 'application/vnd.google-apps.document'
    }, blob, {
      convert: true,
      ocr: false
    });

    // Google Documentとして開く
    const doc = DocumentApp.openById(file.id);
    const text = doc.getBody().getText();

    // keepFileがfalseの場合のみ削除
    if (!keepFile) {
      Drive.Files.remove(file.id);
    }

    debugLog('File', 'Word document processed', { textLength: text.length, fileId: file.id });
    return { text, fileId: keepFile ? file.id : null };

  } catch (e) {
    debugLog('File', 'Word processing error', e.toString());
    throw new Error('Word文書の処理に失敗しました: ' + e.toString());
  }
}

/**
 * PDFを処理してテキストを抽出
 */
function processPDFDocument(blob, keepFile = false) {
  try {
    debugLog('File', 'Processing PDF document', { keepFile });
    
    // Google DriveのOCR機能を使用
    const file = Drive.Files.insert({
      title: 'temp_pdf_' + new Date().getTime(),
      mimeType: 'application/pdf'
    }, blob, {
      ocr: true,
      ocrLanguage: 'ja'
    });
    
    // テキストを抽出
    const doc = DocumentApp.openById(file.id);
    const text = doc.getBody().getText();
    
    // keepFileがfalseの場合のみ削除
    if (!keepFile) {
      Drive.Files.remove(file.id);
    }
    
    debugLog('File', 'PDF processed', { textLength: text.length, fileId: file.id });
    return { text, fileId: keepFile ? file.id : null };
    
  } catch (e) {
    debugLog('File', 'PDF processing error', e.toString());
    throw new Error('PDFの処理に失敗しました: ' + e.toString());
  }
}

/**
 * GoogleドキュメントのURLから内容を取得
 */
function getGoogleDocumentContent(url) {
  try {
    debugLog('GoogleDoc', 'Processing Google Document URL', { url });
    
    // Google ドキュメントのIDを抽出
    let docId = null;
    
    // 様々なGoogleドキュメントのURL形式に対応
    const patterns = [
      /\/document\/d\/([a-zA-Z0-9-_]+)/,  // https://docs.google.com/document/d/DOC_ID/...
      /\/open\?id=([a-zA-Z0-9-_]+)/,      // https://docs.google.com/open?id=DOC_ID
      /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/, // スプレッドシート
      /\/presentation\/d\/([a-zA-Z0-9-_]+)/, // プレゼンテーション
    ];
    
    for (const pattern of patterns) {
      const match = url.match(pattern);
      if (match) {
        docId = match[1];
        break;
      }
    }
    
    if (!docId) {
      debugLog('GoogleDoc', 'Could not extract document ID from URL');
      return null;
    }
    
    // ドキュメントのタイプを判定
    let content = '';
    
    if (url.includes('/spreadsheets/')) {
      // Googleスプレッドシート
      try {
        const sheet = SpreadsheetApp.openById(docId);
        const sheets = sheet.getSheets();
        content = 'Googleスプレッドシート: ' + sheet.getName() + '\n\n';
        
        sheets.forEach((s, index) => {
          if (index < 3) { // 最初の3シートのみ
            content += `シート: ${s.getName()}\n`;
            const range = s.getDataRange();
            const values = range.getValues();
            const maxRows = Math.min(values.length, 50); // 最大50行
            
            for (let i = 0; i < maxRows; i++) {
              content += values[i].join('\t') + '\n';
            }
            content += '\n';
          }
        });
      } catch (e) {
        debugLog('GoogleDoc', 'Error reading spreadsheet', e.toString());
        return null;
      }
    } else if (url.includes('/presentation/')) {
      // Googleスライド（プレゼンテーション）
      try {
        const presentation = SlidesApp.openById(docId);
        const slides = presentation.getSlides();
        content = 'Googleプレゼンテーション: ' + presentation.getName() + '\n\n';
        
        slides.forEach((slide, index) => {
          if (index < 10) { // 最初の10スライドのみ
            content += `スライド ${index + 1}:\n`;
            const shapes = slide.getShapes();
            shapes.forEach(shape => {
              const text = shape.getText();
              if (text) {
                content += text.asString() + '\n';
              }
            });
            content += '\n';
          }
        });
      } catch (e) {
        debugLog('GoogleDoc', 'Error reading presentation', e.toString());
        return null;
      }
    } else {
      // Googleドキュメント（デフォルト）
      try {
        const doc = DocumentApp.openById(docId);
        const body = doc.getBody();
        content = body.getText();
        
        if (!content) {
          debugLog('GoogleDoc', 'Document is empty');
          return null;
        }
      } catch (e) {
        debugLog('GoogleDoc', 'Error reading document', e.toString());
        return null;
      }
    }
    
    debugLog('GoogleDoc', 'Content extracted', { length: content.length });
    return content;
    
  } catch (e) {
    debugLog('GoogleDoc', 'Error processing Google Document', e.toString());
    return null;
  }
}

/**
 * メッセージからGoogleドキュメントのURLを検出
 */
function extractGoogleDocUrls(message) {
  if (!message) return [];
  
  // GoogleドキュメントのURLパターン
  const urlPattern = /https?:\/\/docs\.google\.com\/[^\s<>]+/gi;
  const matches = message.match(urlPattern);
  
  if (!matches) return [];
  
  // 重複を除去
  return [...new Set(matches)];
}

/**
 * ファイル添付イベントを処理
 */
function processFileAttachment(event) {
  if (!event.files || event.files.length === 0) {
    return null;
  }
  
  debugLog('File', 'Processing file attachments', { count: event.files.length });
  
  const results = [];
  
  for (const file of event.files) {
    try {
      debugLog('File', 'Processing file', { 
        name: file.name, 
        type: file.mimetype,
        id: file.id 
      });
      
      // ファイル情報を取得
      const fileInfo = getSlackFileInfo(file.id);
      if (!fileInfo) {
        results.push({
          name: file.name || 'Unknown',
          error: 'ファイル情報を取得できませんでした'
        });
        continue;
      }

      // fileInfoから必要な情報を取得（file.mimetypeがundefinedの場合のため）
      const fileName = file.name || fileInfo.name || 'Unknown';
      const mimeType = file.mimetype || fileInfo.mimetype;

      // ファイルをダウンロード
      const blob = downloadSlackFile(fileInfo.url_private);
      if (!blob) {
        results.push({
          name: fileName,
          error: 'ファイルをダウンロードできませんでした'
        });
        continue;
      }

      let content = '';

      // ファイルタイプに応じて処理
      let googleDocId = null;

      if (mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
          mimeType === 'application/msword') {
        // Word文書 - Googleドキュメントとして保存
        const result = processWordDocument(blob, true); // keepFile = true
        content = result.text;
        googleDocId = result.fileId;
      } else if (mimeType === 'application/pdf') {
        // PDF
        const result = processPDFDocument(blob, false); // keepFile = false for PDF
        content = result.text;
      } else if (mimeType && mimeType.startsWith('text/')) {
        // テキストファイル
        content = blob.getDataAsString();
      } else {
        results.push({
          name: fileName,
          error: `サポートされていないファイル形式です (${mimeType || 'unknown'})`
        });
        continue;
      }

      results.push({
        name: fileName,
        type: mimeType,
        content: content,
        googleDocId: googleDocId // GoogleドキュメントIDを保存
      });
      
    } catch (e) {
      debugLog('File', 'Error processing file', e.toString());
      results.push({
        name: file.name || 'unknown',
        error: e.toString()
      });
    }
  }
  
  return results;
}

// ===========================
// Google Natural Language API
// ===========================

/**
 * Google Natural Language API analyzeSyntax
 */
function gNL(textdata) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GOOGLE_NL_API');
  if (!apiKey) {
    debugLog('NL', 'No Google NL API key');
    return null;
  }
  
  const url = "https://language.googleapis.com/v1/documents:analyzeSyntax?key=" + apiKey;
  const payload = {
    document: {
      type: "PLAIN_TEXT",
      content: textdata
    },
    encodingType: "UTF8"
  };
  
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    return JSON.parse(response.getContentText());
  } catch(e) {
    debugLog('NL', 'Error calling NL API', e.toString());
    return null;
  }
}

/**
 * Google Natural Language API の戻り値より必要なものを抽出する
 * 品詞の場合は tagsの欄に ['NOUN','NUM','NUMBER']
 */
function filterGNL(gNLobj, tags) {
  if (!gNLobj || !gNLobj.tokens) return [];
  const words = gNLobj.tokens
    .filter(token => tags.includes(token.partOfSpeech.tag))
    .map(token => token.text.content);
  return words;
}

/**
 * カタカナをひらがなに変換
 */
function katakanaToHiragana(text) {
  return text.replace(/[ァ-ヶ]/g, function(match) {
    const chr = match.charCodeAt(0) - 0x60;
    return String.fromCharCode(chr);
  });
}

/**
 * 全角を半角に変換
 */
function toHalfWidth(str) {
  str = str.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
  return str;
}

// ===========================
// 文字列正規化関数
// ===========================

/**
 * カタカナをひらがなに変換
 */
function katakanaToHiragana(str) {
  return str.replace(/[\u30a1-\u30f6]/g, function(match) {
    const chr = match.charCodeAt(0) - 0x60;
    return String.fromCharCode(chr);
  });
}

/**
 * 全角文字を半角に変換
 */
function toHalfWidth(str) {
  // 全角英数字を半角に変換
  str = str.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
  // 全角スペースを半角に
  str = str.replace(/　/g, ' ');
  return str;
}

// ===========================
// FAQ検索機能
// ===========================

/**
 * FAQロールを取得（参考コードに忠実に再現）
 */
function getFaqRole(question) {
  try {
    // 参考コードのロジックを忠実に再現
    const morpths = filterGNL(gNL(question), ['NOUN', 'NUM', 'NUMBER']);
    let words = [];

    for (let i = 0; i < morpths.length; i++) {
      let d = katakanaToHiragana(
        toHalfWidth(morpths[i]).toLowerCase().replace(',', '')
      );
      if (d.indexOf('-')) {
        const arr = morpths[i].split('-');
        for (let n = 0; n < arr.length; n++) {
          words.push(
            katakanaToHiragana(
              toHalfWidth(arr[n]).toLowerCase().replace(',', '')
            )
          );
        }
        continue;
      }
      words.push(d);
    }

    // SpreadsheetApp.getActive()を使用（参考コードと同じ）
    const faqs = SpreadsheetApp.getActive()
      .getSheetByName('faq')
      .getRange('A:B')
      .getValues()
      .filter((row) => !row.every((cell) => cell.toString().trim() === ''));

    let sfaqs = [], result = [];

    // FAQデータを正規化
    for (let i = 1; i < faqs.length; i++) {
      sfaqs[i] = faqs[i].map((cell) =>
        katakanaToHiragana(toHalfWidth(cell.toString()).toLowerCase().replace(',', ''))
      );
    }

    // マッチング処理
    for (let i = 1; i < sfaqs.length; i++) {
      if (sfaqs[i].some((faq) => words.some((w) => faq.includes(w)))) {
        if (result.length === 0) result.push(faqs[0]);
        result.push(faqs[i]);
      }
    }

    if (!result.length) return null;

    return {
      role: 'system',
      content:
        '今から記載するJSON形式のFAQを踏まえて回答を望む(FAQの回答とは言わない)' +
        JSON.stringify(result),
    };
  } catch (e) {
    return null;
  }
}


function mergeRoleAndThread(optionRole, threadMessages) {
  debugLog('Thread', 'Merging thread messages', { count: threadMessages.length });
  
  for (let i = 0; i < threadMessages.length; i++) {
    const msg = threadMessages[i];
    
    // メッセージからメンションを除去
    let cleanText = msg.text || '';
    cleanText = cleanText.replace(/<@[A-Z0-9]+>/g, '').trim();
    
    // 空のメッセージはスキップ
    if (!cleanText) continue;
    
    // ボットのメッセージかユーザーのメッセージかを判定
    // app_id、bot_id、またはuserがボットのIDと一致する場合はassistant
    const isBot = msg.hasOwnProperty('app_id') || msg.hasOwnProperty('bot_id') || 
                  (msg.user && Settings().SLACK_BOT_USER_ID && msg.user === Settings().SLACK_BOT_USER_ID);
    
    const role = isBot ? 'assistant' : 'user';
    
    optionRole.push({
      role: role,
      content: cleanText
    });
    
    debugLog('Thread', `Added ${role} message`, { 
      text: cleanText.substring(0, 50),
      isBot: isBot,
      user: msg.user
    });
  }
  
  debugLog('Thread', 'Thread merge complete', { totalRoles: optionRole.length });
}

// ===========================
// AI レスポンス機能
// ===========================

function chatGPTResponse(message, { optionRole = [], temperature }) {
  debugLog('AI', 'ChatGPT request', { messageLength: message?.length, roles: optionRole.length });
  
  const config = Settings();
  if (!config?.OPEN_AI_TOKEN) {
    debugLog('AI', 'No OpenAI token');
    return 'OpenAI APIトークンが設定されていません。';
  }
  
  const apiKey = config.OPEN_AI_TOKEN;
  const url = 'https://api.openai.com/v1/chat/completions';
  // モデルはスクリプトプロパティ OPEN_AI_MODEL から取得（未設定時は既定の gpt-5）
  const modelName = config.OPEN_AI_MODEL || 'gpt-5';
  debugLog('AI', 'Using OpenAI model', modelName);
  const payload = {
    model: modelName,
    messages: [...optionRole, { role: 'user', content: message }],
    temperature: temperature ?? 1,
  };
  
  const options = {
    method: 'post',
    muteHttpExceptions: true,
    headers: {
      'Authorization': 'Bearer ' + apiKey,
      'Content-Type': 'application/json',
    },
    payload: JSON.stringify(payload),
  };
  
  try {
    const res = UrlFetchApp.fetch(url, options);
    const result = JSON.parse(res.getContentText());
    
    if (result.error) {
      debugLog('AI', 'ChatGPT error', result.error);
      return `AIエラー: ${result.error.message}`;
    }
    
    const content = result?.choices?.[0]?.message?.content || '';
    debugLog('AI', 'ChatGPT success', { responseLength: content.length });
    return content;
  } catch (e) {
    debugLog('AI', 'ChatGPT exception', e.toString());
    return `AI処理エラー: ${e.toString()}`;
  }
}

// ===========================
// メインエントリポイント
// ===========================

function doGet(e) {
  debugLog('Main', 'GET request received', e.parameter);
  
  return ContentService
    .createTextOutput(JSON.stringify({
      status: 'ok',
      message: 'Slack Bot is running',
      timestamp: new Date().toISOString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  // Slack URL Verification Challenge処理（最優先）
  try {
    if (e && e.postData && e.postData.contents) {
      const params = JSON.parse(e.postData.contents);
      
      // URL Verificationの場合は即座にchallengeを返す
      if (params.type === 'url_verification') {
        console.log('Challenge received: ' + params.challenge);
        // challengeを平文テキストで返す（重要）
        return ContentService.createTextOutput(params.challenge);
      }
    }
  } catch (error) {
    console.log('Challenge error: ' + error.toString());
  }
  
  // 通常のイベント処理
  debugLog('Main', 'POST request received');
  
  try {
    if (!e.postData) {
      debugLog('Main', 'No postData');
      return ContentService.createTextOutput('OK');
    }
    
    const params = JSON.parse(e.postData.contents);
    debugLog('Main', 'Parsed params', { type: params.type, event_type: params.event?.type });
    
    // イベントコールバックの処理
    if (params.type !== 'event_callback') {
      debugLog('Main', 'Not event_callback', params.type);
      return ContentService.createTextOutput('OK');
    }
    
    // 重複受信防止
    const cache = CacheService.getScriptCache();
    if (params.event && params.event.client_msg_id) {
      if (cache.get(params.event.client_msg_id) === 'done') {
        debugLog('Main', 'Duplicate message');
        return ContentService.createTextOutput('');
      }
      cache.put(params.event.client_msg_id, 'done', 600);
    }
    
    const Bot = new SlackBot(e);
    
    // メンションイベント用ハンドラ
    const run = ({ event, message, channel, threadTs }) => {
      debugLog('Main', 'Run handler', { 
        channel, 
        message: message?.substring(0, 50), 
        threadTs,
        hasFiles: event.files ? event.files.length : 0
      });
      
      try {
        // ファイル処理とGoogleドキュメントURL処理
        let fileContext = null;
        let googleDocContext = null;
        
        // メッセージからGoogleドキュメントのURLを検出
        const googleDocUrls = extractGoogleDocUrls(message);
        if (googleDocUrls.length > 0) {
          debugLog('Main', 'Google Doc URLs found', { count: googleDocUrls.length, urls: googleDocUrls });
          
          const docContents = [];
          for (const url of googleDocUrls) {
            const content = getGoogleDocumentContent(url);
            if (content) {
              docContents.push({
                url: url,
                content: content
              });
            }
          }
          
          if (docContents.length > 0) {
            googleDocContext = {
              docs: docContents,
              content: docContents.map(doc => {
                const truncatedContent = doc.content.substring(0, 5000);
                return `【GoogleドキュメントURL: ${doc.url}】\n【内容】\n${truncatedContent}${doc.content.length > 5000 ? '\n...(以下省略)' : ''}`;
              }).join('\n\n========================================\n\n')
            };
            
            debugLog('Main', 'Google Docs processed', { 
              success: docContents.length,
              totalChars: googleDocContext.content.length
            });
          }
        }
        
        // ファイルがイベントに直接添付されている場合
        if (event.files && event.files.length > 0) {
          debugLog('Main', 'Files found in event', { 
            count: event.files.length,
            files: event.files.map(f => ({ name: f.name, type: f.mimetype, id: f.id }))
          });
          
          const fileResults = processFileAttachment(event);
          if (fileResults && fileResults.length > 0) {
            // ファイル内容をコンテキストとして準備
            const successfulFiles = fileResults.filter(r => r.content);
            const failedFiles = fileResults.filter(r => r.error);
            
            if (successfulFiles.length > 0) {
              fileContext = {
                files: successfulFiles,
                content: successfulFiles.map(f => {
                  const truncatedContent = f.content.substring(0, 5000);
                  return `【ファイル名: ${f.name}】\n【ファイルタイプ: ${f.type}】\n【内容】\n${truncatedContent}${f.content.length > 5000 ? '\n...(以下省略)' : ''}`;
                }).join('\n\n========================================\n\n')
              };
              
              debugLog('Main', 'Files processed successfully', { 
                success: successfulFiles.length,
                failed: failedFiles.length,
                totalChars: fileContext.content.length
              });
            }
            
            // エラーがあった場合は通知
            if (failedFiles.length > 0 && successfulFiles.length === 0) {
              const errorMessage = '⚠️ ファイル処理エラー:\n' + failedFiles.map(f => 
                `・${f.name}: ${f.error}`
              ).join('\n');
              postMessage(errorMessage, channel, threadTs || event.ts);
              return; // エラーのみの場合は処理を中断
            }
          }
        }
        
        // file_sharedイベントの場合の処理
        if (event.type === 'file_shared' && event.file_id) {
          debugLog('Main', 'File shared event detected', { file_id: event.file_id });
          
          try {
            const fileInfo = getSlackFileInfo(event.file_id);
            if (fileInfo) {
              const tempEvent = {
                files: [{
                  id: fileInfo.id,
                  name: fileInfo.name,
                  mimetype: fileInfo.mimetype,
                  url_private: fileInfo.url_private
                }]
              };
              
              const fileResults = processFileAttachment(tempEvent);
              if (fileResults && fileResults.length > 0 && fileResults[0].content) {
                fileContext = {
                  files: [fileResults[0]],
                  content: `【ファイル名: ${fileResults[0].name}】\n【内容】\n${fileResults[0].content.substring(0, 5000)}`
                };
                
                debugLog('Main', 'File shared processed', { name: fileInfo.name });
              }
            }
          } catch (e) {
            debugLog('Main', 'File shared processing error', e.toString());
          }
        }
        
        // チャンネル説明を取得
        const channelInfo = getChannelInfo(channel);
        const channelDescription = channelInfo?.purpose?.value || '';
        
        // ベースロール定義
        const baseRole = [
          { role: 'system', content: 'You are a helpful assistant.' },
          { role: 'system', content: 'Please answer in Japanese.' },
          { role: 'system', content: 'Be concise and clear.' },
          { role: 'system', content: 'When creating graphs, diagrams, flowcharts, or visualizations, always use Mermaid syntax wrapped in ```mermaid code blocks. Include the mermaid code blocks for any visual representation of data, processes, or relationships.' }
        ];
        
        // ファイルコンテキストがある場合は追加しない（後でメッセージに直接含める）
        // この部分は削除して、メッセージ本文に直接ファイル内容を含める
        
        // メンション部分を除去
        let text = message || '';
        const mentionMatch = text.match(/<@[A-Z0-9]+>/);
        if (mentionMatch) {
          text = text.replace(mentionMatch[0], '').trim();
        }
        
        // ファイルが添付されていて、テキストが空の場合のデフォルトメッセージ
        if (fileContext && !text) {
          text = '添付されたファイルの内容を確認し、要約や重要なポイントを教えてください。';
        }
        
        debugLog('Main', 'Processed text', text);
        debugLog('Main', 'Context status', {
          hasFileContext: !!fileContext,
          fileCount: fileContext ? fileContext.files.length : 0,
          messageLength: text.length
        });
        
        // 簡易コマンド対応
        if (text === 'Hello' || text === 'hello') {
          postMessage('こんにちは！ご用件をお聞かせください。', channel, threadTs || event.ts);
          return;
        }
        
        if (text === 'help') {
          postMessage('このボットはAIアシスタントです。質問や依頼をメンションと共に送信してください。', channel, threadTs || event.ts);
          return;
        }
        
        // AI レスポンス取得
        const optionRole = [...baseRole];
        
        // FAQから追加ロールを取得
        const faq = getFaqRole(text);
        if (faq) {
          optionRole.push(faq);
          debugLog('Main', 'FAQ role added to context');
        } else {
          debugLog('Main', 'No FAQ matches found for this query');
        }
        
        // スレッド履歴を取得してマージ
        if (event.thread_ts) {
          debugLog('Main', 'Processing thread context', { threadTs: event.thread_ts });
          const threadMessages = getThreadMessages(channel, event.thread_ts);
          
          if (threadMessages && threadMessages.length > 0) {
            // 最新のメッセージが現在のメッセージの場合は除外
            const filteredMessages = threadMessages.filter(msg => 
              msg.ts !== event.ts && // 現在のメッセージを除外
              msg.ts !== event.thread_ts // スレッドの最初のメッセージが重複しないように
            );
            
            debugLog('Main', 'Thread messages filtered', { 
              original: threadMessages.length, 
              filtered: filteredMessages.length 
            });
            
            mergeRoleAndThread(optionRole, filteredMessages);
          }
        } else {
          debugLog('Main', 'No thread context (new thread)');
        }
        
        // ファイルコンテキストまたはGoogleドキュメントコンテキストがある場合は、内容を直接メッセージに含める
        let finalMessage = text;
        let wordDocumentContext = null; // Wordドキュメントのコンテキスト情報
        
        // Googleドキュメントとファイルのコンテキストをマージ
        let allContent = [];
        let contentDescription = [];
        
        if (googleDocContext && googleDocContext.content) {
          allContent.push(googleDocContext.content);
          contentDescription.push('Googleドキュメント');
        }
        
        if (fileContext && fileContext.content) {
          allContent.push(fileContext.content);
          const fileNames = fileContext.files.map(f => f.name).join(', ');
          contentDescription.push(`添付ファイル（${fileNames}）`);
          
          // Wordドキュメントの情報を保存（シンプルに修正）
          // Wordファイルがあるかチェック
          const wordFiles = fileContext.files.filter(f => 
            f.type && (f.type.includes('word') || f.type.includes('msword') || 
            f.type.includes('openxmlformats-officedocument.wordprocessingml'))
          );
          
          debugLog('Main', 'Checking for Word files', {
            totalFiles: fileContext.files.length,
            wordFilesFound: wordFiles.length,
            fileTypes: fileContext.files.map(f => f.type)
          });
          
          if (wordFiles.length > 0) {
            wordDocumentContext = {
              files: wordFiles,
              originalContent: wordFiles[0].content // 元のコンテンツを取得
            };
            debugLog('Main', 'Word document context prepared', {
              hasContext: true,
              filesCount: wordFiles.length,
              contentLength: wordFiles[0].content ? wordFiles[0].content.length : 0
            });
          }
        }
        
        if (allContent.length > 0) {
          // コンテンツを直接プロンプトに含める
          const contentLabel = contentDescription.join('と');
          
          // ユーザーのメッセージとコンテンツを組み合わせる
          finalMessage = `以下の${contentLabel}の内容について、${text || 'レビューや要約をお願いします'}：

========================================
内容：
========================================
${allContent.join('\n\n========================================\n\n')}
========================================

上記の内容に基づいて回答してください。`;
          
          debugLog('Main', 'Message with content', { 
            messageLength: finalMessage.length,
            hasGoogleDocs: !!googleDocContext,
            hasFiles: !!fileContext,
            totalContent: allContent.length
          });
        }
        
        let responseText = chatGPTResponse(finalMessage, { optionRole }) || '申し訳ありません。応答の生成に失敗しました。';
        
        // ChatGPTのレスポンスからMarkdown記法を削除
        responseText = convertMarkdownToSlack(responseText);
        
        // Wordドキュメントのレビュー結果をGoogleドキュメントに保存
        let googleDocUrl = null;
        if (wordDocumentContext && responseText && responseText.length > 0) {
          debugLog('Main', 'Attempting to save review to Google Doc', {
            hasWordContext: !!wordDocumentContext,
            responseLength: responseText.length
          });
          
          googleDocUrl = saveReviewToGoogleDoc(
            wordDocumentContext,
            responseText,
            text
          );
          
          debugLog('Main', 'Save review result', {
            success: !!googleDocUrl,
            url: googleDocUrl
          });
        }
        
        // GoogleドキュメントURLを追加
        let finalResponse = responseText;
        if (googleDocUrl) {
          finalResponse += `\n\n📝 レビュー結果をGoogleドキュメントに保存しました:\n${googleDocUrl}`;
          debugLog('Main', 'Added Google Doc URL to response');
        } else if (wordDocumentContext) {
          debugLog('Main', 'Google Doc URL not created despite Word context existing');
        }
        
        // メッセージ送信
        const success = postMessage(finalResponse, channel, threadTs || event.ts);
        debugLog('Main', 'Message post result', success);
        
      } catch (error) {
        debugLog('Main', 'Handler error', error.toString());
        postMessage('エラーが発生しました: ' + error.toString(), channel, event.ts);
      }
    };
    
    // メンションされたときのみ処理
    Bot.handleMentionEventBase(run);
    
    return Bot.response();
    
  } catch (error) {
    debugLog('Main', 'Fatal error', error.toString());
    return ContentService.createTextOutput('');
  }
}

// ===========================
// テスト関数
// ===========================

/**
 * 最小限のChallenge検証テスト
 * SlackのURL検証が失敗する場合はこれを試す
 */
function minimalChallengeTest() {
  // 最小限のdoPost関数の動作確認
  const testData = {
    postData: {
      contents: JSON.stringify({
        type: 'url_verification',
        challenge: 'test_123'
      })
    }
  };
  
  const result = doPost(testData);
  console.log('Result:', result.getContent());
  console.log('Expected: test_123');
  console.log('Match:', result.getContent() === 'test_123');
}

/**
 * メンションイベントのテスト
 * 実際のイベント処理をシミュレート
 */
function testMentionEvent() {
  console.log('========================================');
  console.log('メンションイベントのテスト');
  console.log('========================================\n');
  
  // 実際のSlackイベントをシミュレート
  const testEvent = {
    postData: {
      contents: JSON.stringify({
        type: 'event_callback',
        event: {
          type: 'app_mention',
          text: '<@U123456> こんにちは、テストです',
          channel: 'C09BW2EEVAR',  // testチャンネルを使用
          ts: '1234567890.123456',
          user: 'U987654321',
          client_msg_id: 'test_' + Date.now()
        },
        team_id: 'T123456',
        event_id: 'Ev123456'
      })
    }
  };
  
  console.log('テストイベント:', JSON.stringify(testEvent.postData.contents, null, 2));
  
  try {
    const response = doPost(testEvent);
    console.log('\n処理完了');
    console.log('レスポンス:', response.getContent());
    
    // debug_logシートを確認
    const ss = getActiveSpreadsheet();
    const debugSheet = ss.getSheetByName('debug_log');
    if (debugSheet && debugSheet.getLastRow() > 1) {
      const lastLogs = debugSheet.getRange(Math.max(2, debugSheet.getLastRow() - 4), 1, 5, 4).getValues();
      console.log('\n最新のデバッグログ:');
      lastLogs.forEach(log => {
        if (log[0]) {
          console.log(`[${log[1]}] ${log[2]}`);
          if (log[3]) console.log('  Data:', log[3]);
        }
      });
    }
  } catch (error) {
    console.log('エラー:', error.toString());
    console.log(error.stack);
  }
}

/**
 * Slack投稿テスト（直接投稿）
 * チャンネルIDを指定して直接メッセージを送信
 */
function testDirectPost() {
  console.log('========================================');
  console.log('Slack直接投稿テスト');
  console.log('========================================\n');
  
  // testチャンネルを使用（変更可能）
  const TEST_CHANNEL = 'C09BW2EEVAR';  // testチャンネル
  // const TEST_CHANNEL = 'CR81GRMGS';  // generalチャンネル（代替）
  
  console.log('チャンネルIDの見つけ方:');
  console.log('1. Slackでチャンネルを右クリック');
  console.log('2. "リンクをコピー"を選択');
  console.log('3. URLの最後の部分がチャンネルID (Cから始まる)');
  console.log('   例: https://xxx.slack.com/archives/C05XXXXXX\n');
  
  if (TEST_CHANNEL === 'C1234567890') {
    console.log('❌ TEST_CHANNELを実際のチャンネルIDに変更してください');
    console.log('\n【手順】');
    console.log('1. Slackでテストしたいチャンネルを開く');
    console.log('2. チャンネル名をクリック');
    console.log('3. "About"タブの一番下にあるChannel IDをコピー');
    console.log('4. このコードのTEST_CHANNELの値を置き換える');
    console.log('5. 保存して再実行\n');
    
    // チャンネル一覧を取得して表示
    testListChannels();
    return;
  }
  
  try {
    const success = postMessage('テストメッセージ from GAS: ' + new Date().toLocaleString('ja-JP'), TEST_CHANNEL);
    
    if (success) {
      console.log('✅ メッセージ送信成功！');
      console.log('Slackでチャンネルを確認してください');
    } else {
      console.log('❌ メッセージ送信失敗');
      console.log('SLACK_TOKENとチャンネルIDを確認してください');
    }
  } catch (error) {
    console.log('❌ エラー:', error.toString());
  }
}


/**
 * debugLog関数のテスト
 */
function testDebugLog() {
  console.log('========================================');
  console.log('debugLogテスト開始');
  console.log('========================================\n');
  
  // 様々なパターンでテスト
  debugLog('Test', 'Simple message');
  debugLog('Test', 'Message with data', { key: 'value', number: 123 });
  debugLog('Test', 'Message with null', null);
  debugLog('Test', 'Message with undefined', undefined);
  debugLog('Test', 'Message with array', [1, 2, 3]);
  
  console.log('\nテスト完了。debug_logシートを確認してください。');
}

/**
 * 実際のSlackイベントをログに記録
 * デバッグ用にイベントの内容を確認
 */
function logSlackEvent() {
  console.log('========================================');
  console.log('最新のSlackイベントログを確認');
  console.log('========================================\n');
  
  try {
    const ss = getActiveSpreadsheet();
    const debugSheet = ss.getSheetByName('debug_log');
    
    if (!debugSheet || debugSheet.getLastRow() <= 1) {
      console.log('デバッグログがありません');
      console.log('Slackでボットにメンションして、イベントを発生させてください');
      return;
    }
    
    const lastRow = debugSheet.getLastRow();
    const numRows = Math.min(20, lastRow - 1);
    const startRow = Math.max(2, lastRow - numRows + 1);
    
    const logs = debugSheet.getRange(startRow, 1, numRows, 4).getValues();
    
    console.log('最新のデバッグログ:\n');
    logs.forEach(log => {
      if (log[0]) {
        const timestamp = new Date(log[0]).toLocaleString('ja-JP');
        console.log(`[${timestamp}] ${log[1]}: ${log[2]}`);
        if (log[3]) {
          try {
            const data = JSON.parse(log[3]);
            console.log('  Data:', JSON.stringify(data, null, 2));
          } catch(e) {
            console.log('  Data:', log[3]);
          }
        }
      }
    });
    
    console.log('\n========================================');
    console.log('トラブルシューティング:');
    console.log('1. イベントが記録されていない場合:');
    console.log('   - URL Verificationが完了しているか確認');
    console.log('   - Event Subscriptionsが有効か確認');
    console.log('2. イベントは来ているが処理されない場合:');
    console.log('   - event.typeがapp_mentionか確認');
    console.log('   - channelとtextが正しく取得できているか確認');
    console.log('========================================');
    
  } catch (e) {
    console.log('エラー:', e.toString());
  }
}

/**
 * 最近のファイルイベントを確認
 */
function checkFileEvents() {
  console.log('========================================');
  console.log('ファイル関連のイベントを確認');
  console.log('========================================\n');
  
  try {
    const ss = getActiveSpreadsheet();
    const debugSheet = ss.getSheetByName('debug_log');
    
    if (!debugSheet || debugSheet.getLastRow() <= 1) {
      console.log('デバッグログがありません');
      return;
    }
    
    const lastRow = debugSheet.getLastRow();
    const numRows = Math.min(50, lastRow - 1);
    const startRow = Math.max(2, lastRow - numRows + 1);
    
    const logs = debugSheet.getRange(startRow, 1, numRows, 4).getValues();
    
    console.log('ファイル関連のログ:\n');
    let fileLogFound = false;
    
    logs.forEach(log => {
      if (log[1] && (log[1].includes('File') || log[2].includes('file') || log[2].includes('File'))) {
        fileLogFound = true;
        const timestamp = new Date(log[0]).toLocaleString('ja-JP');
        console.log(`[${timestamp}] ${log[1]}: ${log[2]}`);
        if (log[3]) {
          try {
            const data = JSON.parse(log[3]);
            console.log('  Data:', JSON.stringify(data, null, 2));
          } catch(e) {
            console.log('  Data:', log[3]);
          }
        }
      }
    });
    
    if (!fileLogFound) {
      console.log('ファイル関連のログが見つかりません。');
      console.log('\n最新のMainイベント:');
      
      logs.forEach(log => {
        if (log[1] === 'Main' && log[2].includes('Run handler')) {
          const timestamp = new Date(log[0]).toLocaleString('ja-JP');
          console.log(`[${timestamp}] ${log[2]}`);
          if (log[3]) {
            try {
              const data = JSON.parse(log[3]);
              console.log('  Data:', JSON.stringify(data, null, 2));
            } catch(e) {
              console.log('  Data:', log[3]);
            }
          }
        }
      });
    }
    
  } catch (e) {
    console.log('エラー:', e.toString());
  }
}

/**
 * ファイル処理テスト
 * Word文書の処理をシミュレート
 */
function testFileProcessing() {
  console.log('========================================');
  console.log('ファイル処理テスト');
  console.log('========================================\n');
  
  // テスト用のイベントを作成（ファイル付き）
  const testEvent = {
    type: 'message',
    text: '<@U123456> このファイルをレビューしてください',
    channel: 'C09BW2EEVAR',
    ts: '1234567890.123456',
    files: [
      {
        id: 'F123456',
        name: 'test_document.docx',
        mimetype: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        url_private: 'https://files.slack.com/files-pri/T123456/F123456/test.docx'
      }
    ]
  };
  
  console.log('テストイベント:', JSON.stringify(testEvent, null, 2));
  
  // ファイル処理を実行
  try {
    const results = processFileAttachment(testEvent);
    
    if (results && results.length > 0) {
      console.log('\n処理結果:');
      results.forEach(result => {
        console.log(`\nファイル: ${result.name}`);
        if (result.content) {
          console.log('内容（最初の200文字）:');
          console.log(result.content.substring(0, 200));
        } else if (result.error) {
          console.log('エラー:', result.error);
        }
      });
    } else {
      console.log('ファイルが処理されませんでした');
    }
  } catch (e) {
    console.log('エラー:', e.toString());
  }
  
  console.log('\n========================================');
  console.log('注意事項:');
  console.log('1. Drive APIを有効にする必要があります');
  console.log('2. サービス → Drive API v2 を追加');
  console.log('3. appsscript.jsonに必要なスコープを追加');
  console.log('========================================');
}

/**
 * 簡単な投稿テスト（チャンネル選択付き）
 */
function quickPostTest() {
  // まずチャンネル一覧を取得
  console.log('利用可能なチャンネルを確認中...\n');
  
  const config = Settings();
  if (!config?.SLACK_TOKEN) {
    console.log('❌ SLACK_TOKENが設定されていません');
    return;
  }
  
  const url = 'https://slack.com/api/conversations.list';
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    payload: {
      token: config.SLACK_TOKEN,
      types: 'public_channel',
      limit: 10
    },
    muteHttpExceptions: true
  });
  
  const data = JSON.parse(response.getContentText());
  
  if (!data.ok || !data.channels || data.channels.length === 0) {
    console.log('❌ チャンネルが見つかりません');
    console.log('Botをチャンネルに招待してください: /invite @BotName');
    return;
  }
  
  // Botが参加している最初のチャンネルを使用
  const testChannel = data.channels.find(ch => ch.is_member);
  
  if (!testChannel) {
    console.log('❌ Botが参加しているチャンネルがありません');
    console.log('\n利用可能なチャンネル:');
    data.channels.forEach(ch => {
      console.log(`- #${ch.name} (ID: ${ch.id})`);
    });
    console.log('\nSlackで上記のチャンネルに /invite @BotName を実行してください');
    return;
  }
  
  console.log(`✅ テストチャンネル: #${testChannel.name} (${testChannel.id})`);
  console.log('メッセージを送信中...\n');
  
  const message = `テスト送信 [${new Date().toLocaleString('ja-JP')}]`;
  const success = postMessage(message, testChannel.id);
  
  if (success) {
    console.log('✅ 送信成功！');
    console.log(`Slackの #${testChannel.name} を確認してください`);
  } else {
    console.log('❌ 送信失敗');
    console.log('debug_logシートでエラー詳細を確認してください');
  }
}

function testSettings() {
  try {
    const settings = Settings();
    console.log('Settings test passed:', settings);
    return true;
  } catch (e) {
    console.log('Settings test failed:', e.toString());
    return false;
  }
}

/**
 * Google自然言語処理APIとスプレッドシート連携のテスト
 */
function testNLPAndSpreadsheet() {
  console.log('========================================');
  console.log('Google NLP & スプレッドシート連携テスト');
  console.log('========================================\n');

  // 1. Google NLP APIのテスト
  console.log('1. Google Natural Language APIテスト');
  console.log('-----------------------------------------');

  const testText = '休暇申請の方法を教えてください。経費精算の締切はいつですか？';
  console.log('テストテキスト: "' + testText + '"\n');

  const nlResult = gNL(testText);
  if (nlResult) {
    console.log('✅ NLP API接続成功');

    const nouns = filterGNL(nlResult, ['NOUN']);
    console.log('\n抽出された名詞:');
    nouns.forEach(noun => console.log('  - ' + noun));

    const numbers = filterGNL(nlResult, ['NUM', 'NUMBER']);
    if (numbers.length > 0) {
      console.log('\n抽出された数値:');
      numbers.forEach(num => console.log('  - ' + num));
    }
  } else {
    console.log('❌ NLP API接続失敗');
    console.log('ℹ️ GOOGLE_NL_APIキーが設定されているか確認してください');
  }

  // 2. スプレッドシート分析テスト
  console.log('\n2. スプレッドシート分析テスト');
  console.log('-----------------------------------------');

  const ssId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (ssId) {
    console.log('スプレッドシートID: ' + ssId);

    const analysisResult = analyzeSpreadsheetWithNLP(ssId, 50);

    if (analysisResult.success) {
      console.log('✅ スプレッドシート分析成功\n');

      console.log('分析結果:');
      console.log('  - トークン数: ' + analysisResult.analysis.totalTokens);
      console.log('  - 抽出された名詞数: ' + analysisResult.analysis.nouns.length);

      if (analysisResult.analysis.nouns.length > 0) {
        console.log('\n主要な名詞 (最初の10個):');
        analysisResult.analysis.nouns.slice(0, 10).forEach(noun => {
          console.log('    - ' + noun);
        });
      }

      if (analysisResult.analysis.sentiment) {
        console.log('\n感情分析:');
        console.log('  - スコア: ' + analysisResult.analysis.sentiment.score);
        console.log('  - 強度: ' + analysisResult.analysis.sentiment.magnitude);
      }

      console.log('\nシートごとのテキスト抽出結果:');
      Object.keys(analysisResult.sheetContents).forEach(sheetName => {
        const contents = analysisResult.sheetContents[sheetName];
        if (contents.length > 0) {
          console.log('  シート: ' + sheetName + ' (' + contents.length + '行のテキスト)');
        }
      });
    } else {
      console.log('❌ スプレッドシート分析失敗');
      console.log('エラー: ' + analysisResult.error);
    }
  } else {
    console.log('❌ SPREADSHEET_IDが設定されていません');
  }

  // 3. FAQ検索テスト
  console.log('\n3. FAQ検索機能テスト');
  console.log('-----------------------------------------');

  const testQuestions = [
    '休暇の申請方法を教えて',
    '経費精算について',
    '会議室の予約'
  ];

  testQuestions.forEach(question => {
    console.log('\n質問: "' + question + '"');
    const faqResult = getFaqRole(question);

    if (faqResult && faqResult.content) {
      console.log('✅ FAQマッチあり');
      try {
        const content = faqResult.content.replace('JSON形式のFAQを踏まえて回答を望む(FAQの回答とは言わない)', '');
        const faqData = JSON.parse(content);
        if (Array.isArray(faqData) && faqData.length > 1) {
          console.log('  マッチしたFAQ:');
          for (let i = 1; i < faqData.length; i++) {
            console.log('    - キーワード: ' + faqData[i][0]);
          }
        }
      } catch (e) {
        console.log('  FAQデータのパースに失敗');
      }
    } else {
      console.log('ℹ️ FAQマッチなし');
    }
  });

  console.log('\n========================================');
  console.log('テスト完了');
  console.log('========================================');
}

function testSlackConnection() {
  try {
    const config = Settings();
    if (!config.SLACK_TOKEN) {
      console.log('SLACK_TOKENが設定されていません');
      return false;
    }
    
    const url = 'https://slack.com/api/auth.test';
    const options = {
      method: 'post',
      payload: { token: config.SLACK_TOKEN },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const data = JSON.parse(response.getContentText());
    
    if (data.ok) {
      console.log('Slack接続成功');
      console.log('  ユーザー: ' + data.user);
      console.log('  チーム: ' + data.team);
      return true;
    } else {
      console.log('Slack接続失敗: ' + data.error);
      return false;
    }
  } catch (e) {
    console.log('テストエラー: ' + e.toString());
    return false;
  }
}

function testPostMessage() {
  const testChannel = 'C1234567890'; // ★テスト用チャンネルIDに変更してください
  const testMessage = 'テストメッセージ: ' + new Date().toISOString();
  
  const result = postMessage(testMessage, testChannel);
  console.log('Test message result:', result);
  return result;
}

// ===========================
// Slack Manifest (JSON)
// ===========================

const SLACK_MANIFEST = {
  "display_information": {
    "name": "ChatGPT",
    "description": "AI Assistant Bot",
    "background_color": "#2eb886"
  },
  "features": {
    "bot_user": {
      "display_name": "ChatGPT",
      "always_online": false
    }
  },
  "oauth_config": {
    "scopes": {
      "bot": [
        "app_mentions:read",
        "channels:history",
        "channels:read",
        "chat:write",
        "chat:write.public",
        "files:read",
        "groups:read",
        "reactions:read",
        "users:read"
      ]
    }
  },
  "settings": {
    "event_subscriptions": {
      "request_url": "{YOUR_WEB_APP_URL}",  // ★ここにWeb App URLを設定
      "bot_events": [
        "app_mention",
        "message.channels"
      ]
    },
    "org_deploy_enabled": false,
    "socket_mode_enabled": false,
    "token_rotation_enabled": false
  }
};

/**
 * Slack Manifestを取得（URLを設定済み）
 */
function getSlackManifest(webAppUrl) {
  const manifest = JSON.parse(JSON.stringify(SLACK_MANIFEST));
  manifest.settings.event_subscriptions.request_url = webAppUrl;
  return manifest;
}

/**
 * セットアップ手順を表示
 */
function showSetupInstructions() {
  console.log('========================================');
  console.log('Slack Bot セットアップ手順');
  console.log('========================================');
  console.log('\n【初期設定】');
  console.log('1. setupStep1_SetSpreadsheetId() - スプレッドシートIDを設定');
  console.log('2. setupStep2_InitializeSheets() - シートを初期化');
  console.log('3. setupStep3_SetAPIKeys() - APIキーを確認');
  console.log('4. setupStep4_TestConnection() - 接続テスト');
  console.log('\n【デプロイ】');
  console.log('5. デプロイ → 新しいデプロイ → ウェブアプリ');
  console.log('6. Web App URLをコピー');
  console.log('\n【Slack設定】');
  console.log('7. https://api.slack.com/apps で新しいアプリを作成');
  console.log('8. From an app manifest を選択');
  console.log('9. getSlackManifest("YOUR_WEB_APP_URL") の結果を貼り付け');
  console.log('10. OAuth & Permissions から Bot Token をコピー');
  console.log('11. GASのスクリプトプロパティに SLACK_TOKEN として設定');
  console.log('12. Install to Workspace でインストール');
  console.log('\n【動作確認】');
  console.log('13. Slackチャンネルで /invite @ChatGPT');
  console.log('14. @ChatGPT Hello でテスト');
}
