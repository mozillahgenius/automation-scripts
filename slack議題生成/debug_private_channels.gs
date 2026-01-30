// ========= プライベートチャンネル完全デバッグ診断 =========
// この関数はプライベートチャンネルアクセス問題を根本的に診断します

function debugPrivateChannelsComplete() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    console.log('===== プライベートチャンネル完全診断開始 =====');
    
    let report = [];
    report.push('プライベートチャンネル完全診断レポート');
    report.push('=' .repeat(50));
    report.push('');
    
    // 1. Bot認証情報の確認
    report.push('【1. Bot認証情報】');
    const authInfo = slackAPI('auth.test', {});
    report.push(`Bot名: @${authInfo.user || 'unknown'}`);
    report.push(`Bot ID: ${authInfo.user_id || 'unknown'}`);
    report.push(`Team: ${authInfo.team || 'unknown'}`);
    report.push(`Token Type: ${SLACK_BOT_TOKEN.startsWith('xoxb-') ? 'Bot Token ✅' : 'User Token ⚠️'}`);
    report.push('');
    
    // 2. 必要なスコープの確認
    report.push('【2. 必要なスコープの確認】');
    report.push('プライベートチャンネルに必要なスコープ:');
    report.push('- groups:read（プライベートチャンネル一覧取得）');
    report.push('- groups:history（プライベートチャンネル履歴取得）');
    report.push('');
    
    // 3. conversations.listでプライベートチャンネルのみを取得
    report.push('【3. プライベートチャンネル取得テスト】');
    
    // 3-1. プライベートチャンネルのみを指定して取得
    console.log('プライベートチャンネルのみを取得...');
    const privateResponse = slackAPI('conversations.list', {
      types: 'private_channel',
      limit: 100,
      exclude_archived: true
    });
    
    if (!privateResponse.ok) {
      report.push(`❌ エラー: ${privateResponse.error}`);
      if (privateResponse.error === 'missing_scope') {
        report.push('→ groups:read スコープが不足しています');
      }
    } else {
      const privateChannels = privateResponse.channels || [];
      report.push(`プライベートチャンネル数: ${privateChannels.length}個`);
      
      if (privateChannels.length === 0) {
        report.push('⚠️ プライベートチャンネルが0個です');
        report.push('考えられる原因:');
        report.push('1. Botがどのプライベートチャンネルにも招待されていない');
        report.push('2. ワークスペースにプライベートチャンネルが存在しない');
      } else {
        report.push('');
        report.push('取得したプライベートチャンネル:');
        privateChannels.slice(0, 5).forEach((ch, i) => {
          report.push(`${i + 1}. #${ch.name} (${ch.id})`);
          report.push(`   - is_private: ${ch.is_private}`);
          report.push(`   - is_member: ${ch.is_member}`);
          report.push(`   - is_channel: ${ch.is_channel}`);
          report.push(`   - is_group: ${ch.is_group}`);
        });
        if (privateChannels.length > 5) {
          report.push(`... 他 ${privateChannels.length - 5} チャンネル`);
        }
      }
    }
    
    report.push('');
    
    // 4. パブリックチャンネルとプライベートチャンネルを両方取得して比較
    report.push('【4. 全チャンネル取得テスト（パブリック＋プライベート）】');
    
    const allResponse = slackAPI('conversations.list', {
      types: 'public_channel,private_channel',
      limit: 1000,
      exclude_archived: true
    });
    
    if (allResponse.ok) {
      const allChannels = allResponse.channels || [];
      
      // チャンネルIDのプレフィックスで分類
      const cChannels = allChannels.filter(ch => ch.id && ch.id.startsWith('C'));
      const gChannels = allChannels.filter(ch => ch.id && ch.id.startsWith('G'));
      const otherChannels = allChannels.filter(ch => ch.id && !ch.id.startsWith('C') && !ch.id.startsWith('G'));
      
      // is_privateフラグで分類
      const privateByFlag = allChannels.filter(ch => ch.is_private === true);
      const publicByFlag = allChannels.filter(ch => ch.is_private === false || ch.is_private === undefined);
      
      report.push(`全チャンネル数: ${allChannels.length}個`);
      report.push('');
      report.push('IDプレフィックスによる分類:');
      report.push(`- Cで始まる（通常パブリック）: ${cChannels.length}個`);
      report.push(`- Gで始まる（通常プライベート）: ${gChannels.length}個`);
      report.push(`- その他: ${otherChannels.length}個`);
      report.push('');
      report.push('is_privateフラグによる分類:');
      report.push(`- is_private=true: ${privateByFlag.length}個`);
      report.push(`- is_private=false/undefined: ${publicByFlag.length}個`);
      
      // 不一致の検出
      report.push('');
      report.push('【ID と is_private フラグの不一致チェック】');
      const mismatches = [];
      
      allChannels.forEach(ch => {
        const expectedPrivate = ch.id && ch.id.startsWith('G');
        const actualPrivate = ch.is_private === true;
        
        if (expectedPrivate !== actualPrivate) {
          mismatches.push({
            name: ch.name,
            id: ch.id,
            expectedPrivate: expectedPrivate,
            actualPrivate: actualPrivate
          });
        }
      });
      
      if (mismatches.length > 0) {
        report.push(`⚠️ ${mismatches.length}個のチャンネルで不一致を検出:`);
        mismatches.slice(0, 5).forEach(m => {
          report.push(`- #${m.name} (${m.id}): ID判定=${m.expectedPrivate}, フラグ=${m.actualPrivate}`);
        });
      } else {
        report.push('✅ すべてのチャンネルでIDとフラグが一致');
      }
    }
    
    report.push('');
    
    // 5. 特定のプライベートチャンネルへのアクセステスト
    report.push('【5. プライベートチャンネルアクセステスト】');
    
    if (privateResponse.ok && privateResponse.channels && privateResponse.channels.length > 0) {
      const testChannel = privateResponse.channels[0];
      report.push(`テスト対象: #${testChannel.name} (${testChannel.id})`);
      
      // conversations.history でアクセステスト
      try {
        const historyResponse = slackAPI('conversations.history', {
          channel: testChannel.id,
          limit: 1
        });
        
        if (historyResponse.ok) {
          report.push('✅ メッセージ履歴にアクセス可能');
        } else {
          report.push(`❌ アクセス不可: ${historyResponse.error}`);
          if (historyResponse.error === 'not_in_channel') {
            report.push('→ Botがチャンネルメンバーではありません');
          }
        }
      } catch (e) {
        report.push(`❌ エラー: ${e.toString()}`);
      }
    } else {
      report.push('テスト対象のプライベートチャンネルがありません');
    }
    
    report.push('');
    report.push('【6. 推奨アクション】');
    
    // プライベートチャンネルが0の場合の対処法
    if (!privateResponse.channels || privateResponse.channels.length === 0) {
      report.push('プライベートチャンネルにアクセスするには:');
      report.push('');
      report.push('1. Slack App の設定を確認:');
      report.push('   - https://api.slack.com/apps でアプリを選択');
      report.push('   - OAuth & Permissions → Scopes で以下を確認:');
      report.push('     ✓ groups:read');
      report.push('     ✓ groups:history');
      report.push('');
      report.push('2. アプリを再インストール:');
      report.push('   - スコープ追加後、"Reinstall to Workspace" をクリック');
      report.push('');
      report.push('3. プライベートチャンネルにBotを招待:');
      report.push('   - 各プライベートチャンネルで: /invite @' + (authInfo.user || 'bot-name'));
      report.push('   - または: チャンネル設定 → Integrations → Add apps');
      report.push('');
      report.push('4. Bot Token の確認:');
      report.push('   - xoxb- で始まるBot Tokenを使用しているか確認');
      report.push('   - User Token (xoxp-) では制限がある場合があります');
    }
    
    // 結果を表示
    const resultText = report.join('\n');
    console.log(resultText);
    
    // UIに表示（長すぎる場合は最初の部分のみ）
    const displayText = resultText.length > 3000 ? 
      resultText.substring(0, 2900) + '\n\n... (詳細はログを確認してください)' :
      resultText;
    
    ui.alert('診断結果', displayText, ui.ButtonSet.OK);
    
  } catch (error) {
    console.error('診断中にエラー:', error);
    ui.alert('エラー', `診断中にエラーが発生しました:\n${error.toString()}`, ui.ButtonSet.OK);
  }
}

// メニューに追加するための関数
function addDebugMenuItems() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔧 デバッグ')
    .addItem('プライベートチャンネル完全診断', 'debugPrivateChannelsComplete')
    .addToUi();
}