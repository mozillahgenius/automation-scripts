/**
 * Google Cloud Project の設定と確認
 */

/**
 * 現在のプロジェクト設定を確認
 */
function checkProjectSettings() {
  console.log('=== プロジェクト設定確認 ===');
  
  // 1. OAuth トークンの確認
  try {
    const token = ScriptApp.getOAuthToken();
    console.log('✓ OAuth トークン取得成功');
  } catch (error) {
    console.error('✗ OAuth トークン取得失敗:', error);
  }
  
  // 2. プロジェクト番号の確認
  try {
    // GASのプロジェクト情報を取得する方法は限定的
    console.log('GCP_PROJECT_ID (コード内定数):', GCP_PROJECT_ID);
    
    // スクリプトのプロジェクト番号を推測
    const scriptId = ScriptApp.getScriptId();
    console.log('Script ID:', scriptId);
  } catch (error) {
    console.error('プロジェクト情報取得エラー:', error);
  }
  
  // 3. 利用可能なスコープの確認
  console.log('\n=== 設定されているスコープ ===');
  console.log('appsscript.json で以下のスコープが設定されています:');
  console.log('- calendar.readonly');
  console.log('- gmail.send');
  console.log('- meetings.space.readonly');
  console.log('- drive.readonly');
  console.log('- script.external_request');
  console.log('- spreadsheets');
  console.log('- cloud-platform');
  console.log('- userinfo.email');
  console.log('- script.scriptapp');
  
  console.log('\n=== 推奨される対処法 ===');
  console.log('1. GASエディタで「プロジェクトの設定」を開く');
  console.log('2. 「Google Cloud Platform（GCP）プロジェクト」セクションを確認');
  console.log('3. プロジェクト番号が「company-gas」と関連付けられているか確認');
  console.log('4. 異なる場合は「プロジェクトを変更」をクリック');
  console.log('5. プロジェクト番号に「company-gas」のプロジェクト番号を入力');
}

/**
 * GCPプロジェクトとの連携を設定する手順を表示
 */
function showGCPSetupInstructions() {
  console.log('=== GCPプロジェクトとの連携設定手順 ===\n');
  
  console.log('【ステップ1】GCPプロジェクト番号の確認');
  console.log('1. https://console.cloud.google.com/home/dashboard?project=company-gas にアクセス');
  console.log('2. ダッシュボードで「プロジェクト番号」をコピー（数字のみ）');
  console.log('');
  
  console.log('【ステップ2】GASプロジェクトの設定');
  console.log('1. GASエディタ左側の歯車アイコン「プロジェクトの設定」をクリック');
  console.log('2. 「Google Cloud Platform（GCP）プロジェクト」セクションを探す');
  console.log('3. 「プロジェクトを変更」をクリック');
  console.log('4. プロジェクト番号を入力（ステップ1でコピーした番号）');
  console.log('5. 「プロジェクトを設定」をクリック');
  console.log('');
  
  console.log('【ステップ3】権限の再認証');
  console.log('1. 任意の関数を実行');
  console.log('2. 新しい認証画面が表示される');
  console.log('3. 権限を承認');
  console.log('');
  
  console.log('【ステップ4】Meet APIの動作確認');
  console.log('1. testMeetAPIConnection() を実行');
  console.log('2. エラーが出なければ成功');
}

/**
 * Meet API接続テスト（簡易版）
 */
function testMeetAPIConnection() {
  console.log('=== Meet API 接続テスト ===');
  
  const baseUrl = 'https://meet.googleapis.com/v2';
  const accessToken = ScriptApp.getOAuthToken();
  
  const options = {
    method: 'GET',
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Accept': 'application/json'
    },
    muteHttpExceptions: true
  };
  
  try {
    // シンプルなAPIコール（会議レコードのリスト取得）
    const url = `${baseUrl}/conferenceRecords?pageSize=1`;
    console.log('テストURL:', url);
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log('レスポンスコード:', responseCode);
    
    if (responseCode === 200) {
      console.log('✅ Meet API接続成功！');
      const data = JSON.parse(responseText);
      if (data.conferenceRecords) {
        console.log(`${data.conferenceRecords.length}件の会議レコードが見つかりました`);
      } else {
        console.log('会議レコードはありませんが、API接続は成功しています');
      }
    } else if (responseCode === 403) {
      console.error('❌ 権限エラー');
      const errorData = JSON.parse(responseText);
      
      if (errorData.error && errorData.error.message.includes('project')) {
        const projectMatch = errorData.error.message.match(/project (\d+)/);
        if (projectMatch) {
          console.error(`現在のプロジェクト番号: ${projectMatch[1]}`);
          console.error('このプロジェクトでMeet APIが有効化されていません');
          console.error('');
          console.error('👉 showGCPSetupInstructions() を実行して設定手順を確認してください');
        }
      } else {
        console.error('詳細:', responseText);
      }
    } else {
      console.error(`エラー (${responseCode}):`, responseText);
    }
  } catch (error) {
    console.error('接続エラー:', error);
  }
}

/**
 * 代替案: Drive APIを使用して録画ファイルにアクセス
 */
function findMeetRecordings() {
  console.log('=== Google Drive から Meet 録画を検索 ===');
  
  try {
    // Meet録画は通常「Meet Recordings」フォルダに保存される
    const folders = DriveApp.getFoldersByName('Meet Recordings');
    
    if (folders.hasNext()) {
      const folder = folders.next();
      console.log('Meet Recordings フォルダが見つかりました');
      
      const files = folder.getFiles();
      let count = 0;
      
      while (files.hasNext() && count < 5) {
        const file = files.next();
        console.log(`- ${file.getName()}`);
        console.log(`  作成日: ${file.getDateCreated()}`);
        console.log(`  URL: ${file.getUrl()}`);
        count++;
      }
      
      if (count === 0) {
        console.log('録画ファイルが見つかりません');
      }
    } else {
      console.log('Meet Recordings フォルダが見つかりません');
      console.log('録画が有効になっていない可能性があります');
    }
  } catch (error) {
    console.error('Drive検索エラー:', error);
  }
}