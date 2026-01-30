// ========= UI安全関数 =========
// トリガーやWeb Appから実行してもエラーにならない関数

/**
 * UIが利用可能かチェック
 * @returns {boolean} UIが利用可能な場合true
 */
function isUiAvailable() {
  try {
    SpreadsheetApp.getUi();
    return true;
  } catch (e) {
    return false;
  }
}

/**
 * 安全にUIアラートを表示（UIが使えない場合はログ出力）
 * @param {string} title - アラートのタイトル
 * @param {string} message - アラートのメッセージ
 * @param {ButtonSet} buttons - ボタンセット（省略可）
 */
function showAlertSafely(title, message, buttons) {
  if (isUiAvailable()) {
    const ui = SpreadsheetApp.getUi();
    if (buttons) {
      ui.alert(title, message, buttons);
    } else {
      ui.alert(title, message, ui.ButtonSet.OK);
    }
  } else {
    // UIが使えない場合はログに出力
    console.log(`[ALERT] ${title}: ${message}`);
  }
}

/**
 * 安全にプロンプトを表示（UIが使えない場合はnull返却）
 * @param {string} title - プロンプトのタイトル
 * @param {string} prompt - プロンプトのメッセージ
 * @param {ButtonSet} buttons - ボタンセット
 * @returns {PromptResponse|null} レスポンスまたはnull
 */
function showPromptSafely(title, prompt, buttons) {
  if (isUiAvailable()) {
    const ui = SpreadsheetApp.getUi();
    return ui.prompt(title, prompt, buttons);
  } else {
    console.log(`[PROMPT] ${title}: ${prompt}`);
    return null;
  }
}

/**
 * プライベートチャンネル対応版のチャンネル同期（UI安全版）
 * トリガーから実行可能
 */
function syncAllChannelsSafe() {
  try {
    console.log('===== 全チャンネル同期開始（UI安全版） =====');
    
    // Bot情報取得
    const authInfo = slackAPI('auth.test', {});
    console.log(`Bot: @${authInfo.user}, Team: ${authInfo.team}`);
    
    // 全チャンネル取得（パブリック＋プライベート）
    const allChannels = [];
    let cursor = '';
    
    do {
      const params = {
        types: 'public_channel,private_channel',
        limit: 1000,
        exclude_archived: true
      };
      if (cursor) params.cursor = cursor;
      
      const response = slackAPI('conversations.list', params);
      if (response.ok && response.channels) {
        allChannels.push(...response.channels);
        cursor = response.response_metadata?.next_cursor || '';
      } else {
        break;
      }
    } while (cursor);
    
    console.log(`取得したチャンネル総数: ${allChannels.length}`);
    
    // チャンネルを分類
    const publicChannels = allChannels.filter(ch => !ch.is_private);
    const privateChannels = allChannels.filter(ch => ch.is_private === true);
    const memberChannels = allChannels.filter(ch => ch.is_member === true);
    
    console.log(`内訳:`);
    console.log(`- パブリック: ${publicChannels.length}個`);
    console.log(`- プライベート: ${privateChannels.length}個`);
    console.log(`- Botがメンバー: ${memberChannels.length}個`);
    
    // スプレッドシートの準備
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let messagesSheet = ss.getSheetByName(SHEETS.MESSAGES);
    if (!messagesSheet) {
      messagesSheet = createMessagesSheet(ss);
    }
    
    let slackLogSheet = ss.getSheetByName(SHEETS.SLACK_LOG);
    if (!slackLogSheet) {
      slackLogSheet = createSlackLogSheetInSpreadsheet(ss);
    }
    
    // 各チャンネルのメッセージを取得
    let totalMessages = 0;
    let processedChannels = 0;
    let errors = [];
    
    for (const channel of memberChannels) {
      try {
        console.log(`処理中: #${channel.name} (${channel.id}) - ${channel.is_private ? 'プライベート' : 'パブリック'}`);
        
        // メッセージ履歴取得
        const messages = fetchChannelHistory(channel.id, 7); // 過去7日分
        
        if (messages && messages.length > 0) {
          // メッセージを保存
          const messageRows = [];
          const logRows = [];
          
          for (const message of messages) {
            // Messagesシート用
            messageRows.push(prepareMessageRow(channel.id, message));
            
            // slack_logシート用
            logRows.push(prepareSlackLogRow(channel.id, channel.name, message));
          }
          
          // バッチ保存
          if (messageRows.length > 0) {
            saveMessagesBatch(messagesSheet, messageRows);
          }
          if (logRows.length > 0) {
            saveSlackLogBatch(slackLogSheet, logRows);
          }
          
          totalMessages += messages.length;
          console.log(`  → ${messages.length}件のメッセージを保存`);
        } else {
          console.log(`  → 新しいメッセージなし`);
        }
        
        processedChannels++;
        
      } catch (error) {
        const errorMsg = `チャンネル ${channel.name} でエラー: ${error.toString()}`;
        console.error(errorMsg);
        errors.push(errorMsg);
        
        // not_in_channelエラーの場合
        if (error.toString().includes('not_in_channel')) {
          console.log(`  → Botを招待してください: /invite @${authInfo.user}`);
        }
      }
    }
    
    // 結果のサマリー
    const summary = [
      '===== 同期完了 =====',
      `処理チャンネル数: ${processedChannels}/${memberChannels.length}`,
      `取得メッセージ数: ${totalMessages}`,
      `パブリック未参加: ${publicChannels.filter(ch => !ch.is_member).length}個`,
      `プライベート未参加: ${privateChannels.filter(ch => !ch.is_member).length}個`
    ];
    
    if (errors.length > 0) {
      summary.push('');
      summary.push('エラー:');
      errors.forEach(e => summary.push(`- ${e}`));
    }
    
    const summaryText = summary.join('\n');
    console.log(summaryText);
    
    // UI表示（可能な場合のみ）
    showAlertSafely('同期完了', summaryText);
    
    return {
      success: true,
      totalMessages: totalMessages,
      processedChannels: processedChannels,
      errors: errors
    };
    
  } catch (error) {
    console.error('全チャンネル同期エラー:', error);
    showAlertSafely('エラー', `同期中にエラーが発生しました:\n${error.toString()}`);
    
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * プライベートチャンネルの存在確認（デバッグ用）
 */
function checkPrivateChannelsExist() {
  console.log('===== プライベートチャンネル存在確認 =====');
  
  // 方法1: conversations.list with types
  console.log('\n方法1: types=private_channel指定');
  const privateOnly = slackAPI('conversations.list', {
    types: 'private_channel',
    limit: 100
  });
  console.log(`結果: ${privateOnly.channels ? privateOnly.channels.length : 0}個`);
  
  // 方法2: 全チャンネル取得してフィルタ
  console.log('\n方法2: 全チャンネルからフィルタ');
  const all = slackAPI('conversations.list', {
    types: 'public_channel,private_channel',
    limit: 1000
  });
  
  if (all.channels) {
    const realPrivate = all.channels.filter(ch => ch.is_private === true);
    console.log(`is_private=true: ${realPrivate.length}個`);
    
    if (realPrivate.length > 0) {
      console.log('プライベートチャンネル詳細:');
      realPrivate.forEach(ch => {
        console.log(`- #${ch.name} (${ch.id})`);
        console.log(`  is_member: ${ch.is_member}`);
        console.log(`  is_private: ${ch.is_private}`);
        console.log(`  is_channel: ${ch.is_channel}`);
        console.log(`  is_group: ${ch.is_group}`);
        console.log(`  is_mpim: ${ch.is_mpim}`);
      });
    }
  }
  
  // 方法3: users.conversations
  console.log('\n方法3: users.conversations (Botが参加中のプライベート)');
  try {
    const authInfo = slackAPI('auth.test', {});
    const userConv = slackAPI('users.conversations', {
      user: authInfo.user_id,
      types: 'private_channel',
      limit: 100
    });
    console.log(`結果: ${userConv.channels ? userConv.channels.length : 0}個`);
    
    if (userConv.channels && userConv.channels.length > 0) {
      userConv.channels.forEach(ch => {
        console.log(`- #${ch.name} (${ch.id})`);
      });
    }
  } catch (e) {
    console.log(`エラー: ${e.toString()}`);
  }
  
  // 方法4: チャンネル別メッセージ取得テスト
  console.log('\n方法4: 特定チャンネルのメッセージ取得テスト');
  const testChannelIds = [
    'SLACK_CHANNEL_ID_1', // general
    'SLACK_CHANNEL_ID_2'  // channel-2
  ];
  
  testChannelIds.forEach(id => {
    try {
      const info = slackAPI('conversations.info', { channel: id });
      if (info.channel) {
        console.log(`\nチャンネル: #${info.channel.name} (${id})`);
        console.log(`- is_private: ${info.channel.is_private}`);
        console.log(`- is_member: ${info.channel.is_member}`);
        
        // メッセージ取得テスト
        try {
          const history = slackAPI('conversations.history', {
            channel: id,
            limit: 1
          });
          console.log(`- メッセージ取得: ${history.ok ? '成功' : '失敗'}`);
        } catch (e) {
          console.log(`- メッセージ取得: エラー (${e.toString().split(':')[0]})`);
        }
      }
    } catch (e) {
      console.log(`チャンネル ${id}: 情報取得失敗`);
    }
  });
}