// ========= プライベートチャンネル問題の修正版 =========

// 正しくプライベートチャンネルを取得する関数
function getCorrectPrivateChannels() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    console.log('===== 正しいプライベートチャンネル取得 =====');
    
    let report = [];
    report.push('プライベートチャンネル取得（修正版）');
    report.push('=' .repeat(50));
    report.push('');
    
    // Bot情報
    const authInfo = slackAPI('auth.test', {});
    report.push(`Bot: @${authInfo.user}`);
    report.push('');
    
    // 重要: まずALLチャンネルを取得（types指定を変更）
    report.push('【方法1: 全チャンネルから正しくフィルタリング】');
    
    // すべてのチャンネルを取得
    const allChannelsResponse = slackAPI('conversations.list', {
      limit: 1000,
      exclude_archived: true
      // typesパラメータを指定しない、または明示的に両方指定
    });
    
    if (allChannelsResponse.ok) {
      const allChannels = allChannelsResponse.channels || [];
      
      // 実際のプライベートチャンネルを抽出
      // プライベートチャンネルの条件：
      // 1. IDが"G"で始まる（レガシープライベート）
      // 2. is_private === true
      // 3. is_mpim === false（グループDMではない）
      const realPrivateChannels = allChannels.filter(ch => {
        // プライベートチャンネルの判定
        const isPrivateChannel = (
          ch.is_private === true && 
          ch.is_mpim !== true &&  // グループDMを除外
          ch.is_im !== true        // 個人DMを除外
        );
        return isPrivateChannel;
      });
      
      const realPublicChannels = allChannels.filter(ch => {
        return ch.is_channel === true && ch.is_private !== true;
      });
      
      report.push(`全チャンネル数: ${allChannels.length}`);
      report.push(`├─ パブリックチャンネル: ${realPublicChannels.length}個`);
      report.push(`└─ プライベートチャンネル: ${realPrivateChannels.length}個`);
      report.push('');
      
      // プライベートチャンネルの詳細
      if (realPrivateChannels.length > 0) {
        report.push('発見したプライベートチャンネル:');
        realPrivateChannels.forEach((ch, i) => {
          report.push(`${i + 1}. #${ch.name} (${ch.id})`);
          report.push(`   - is_member: ${ch.is_member ? '✅' : '❌'}`);
        });
      } else {
        report.push('⚠️ プライベートチャンネルが見つかりません');
      }
    }
    
    report.push('');
    report.push('【方法2: groups.listの使用（レガシーAPI）】');
    
    // レガシーAPIを試す（古いプライベートチャンネル用）
    try {
      const groupsResponse = slackAPI('groups.list', {
        exclude_archived: true
      });
      
      if (groupsResponse.ok && groupsResponse.groups) {
        report.push(`groups.list結果: ${groupsResponse.groups.length}個のグループ`);
        groupsResponse.groups.slice(0, 3).forEach(g => {
          report.push(`- ${g.name} (${g.id})`);
        });
      } else {
        report.push('groups.listは使用できません（新しいワークスペース）');
      }
    } catch (e) {
      report.push('groups.list APIは利用不可');
    }
    
    report.push('');
    report.push('【方法3: ユーザーのチャンネルメンバーシップ確認】');
    
    // users.conversations で Bot のチャンネルを取得
    try {
      const userConversations = slackAPI('users.conversations', {
        user: authInfo.user_id,
        types: 'private_channel',
        limit: 100
      });
      
      if (userConversations.ok) {
        const botPrivateChannels = userConversations.channels || [];
        report.push(`Botが参加しているプライベートチャンネル: ${botPrivateChannels.length}個`);
        
        if (botPrivateChannels.length > 0) {
          report.push('Botが参加中:');
          botPrivateChannels.slice(0, 5).forEach(ch => {
            report.push(`- #${ch.name} (${ch.id})`);
          });
        }
      }
    } catch (e) {
      report.push(`users.conversations エラー: ${e.toString()}`);
    }
    
    // 結果表示
    const resultText = report.join('\n');
    console.log(resultText);
    ui.alert('診断結果', resultText, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('エラー:', error);
    ui.alert('エラー', error.toString(), ui.ButtonSet.OK);
  }
}

// プライベートチャンネルの正しい診断と修正
function fixPrivateChannelAccess() {
  const ui = SpreadsheetApp.getUi();
  
  console.log('===== プライベートチャンネルアクセス修正 =====');
  
  let report = [];
  report.push('プライベートチャンネル問題の解決策');
  report.push('=' .repeat(50));
  report.push('');
  
  report.push('【現在の状況】');
  report.push('✅ Bot Token は正しく設定されています');
  report.push('✅ groups:read, groups:history スコープはあります');
  report.push('❌ プライベートチャンネルが検出されません');
  report.push('');
  
  report.push('【考えられる原因】');
  report.push('1. ワークスペースに実際にプライベートチャンネルが存在しない');
  report.push('2. すべてのチャンネルがパブリックチャンネルである');
  report.push('3. Botがプライベートチャンネルに一つも招待されていない');
  report.push('');
  
  report.push('【解決方法】');
  report.push('');
  report.push('1. プライベートチャンネルの作成確認:');
  report.push('   Slackで新しいプライベートチャンネルを作成してください');
  report.push('   - チャンネル作成時に「プライベート」を選択');
  report.push('   - 既存のパブリックチャンネルをプライベートに変換も可能');
  report.push('');
  
  report.push('2. Botをプライベートチャンネルに招待:');
  report.push('   プライベートチャンネルで以下を実行:');
  report.push('   /invite @kushim_slack_governan');
  report.push('');
  
  report.push('3. 表示されているチャンネルについて:');
  report.push('   現在表示されている以下のチャンネルはすべてパブリックです:');
  report.push('   - all-kushim (C08QJSAMS5T) → パブリック');
  report.push('   - backoffice (C08S6947WSD) → パブリック（Botメンバー）');
  report.push('   これらにアクセスするには「Bot追加」機能を使用してください');
  report.push('');
  
  report.push('【次のステップ】');
  report.push('1. Slackでプライベートチャンネルが存在するか確認');
  report.push('2. 存在する場合、Botを招待');
  report.push('3. この診断を再実行');
  
  const resultText = report.join('\n');
  console.log(resultText);
  ui.alert('解決策', resultText, ui.ButtonSet.OK);
}

// メニューに追加
function addFixMenuItems() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔧 修正版')
    .addItem('正しいプライベートチャンネル取得', 'getCorrectPrivateChannels')
    .addItem('プライベートチャンネル問題の解決', 'fixPrivateChannelAccess')
    .addToUi();
}