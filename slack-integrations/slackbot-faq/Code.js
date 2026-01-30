function testrun(){
  var e = {postData:{contents:JSON.stringify({event:{text:'test'},challenge:'test'})}};
  doPost(e);
}

function doPost(e) {
  var params = JSON.parse(e.postData.contents);  

  // Slackからのチャレンジリクエストへの応答
  if (params.challenge) {
    return ContentService.createTextOutput(params.challenge);
  }

  // 3秒超過リトライ対処
  let cache = CacheService.getScriptCache();
  if (cache.get(params.event.client_msg_id) == 'done') {
    return ContentService.createTextOutput();
  } else {
    cache.put(params.event.client_msg_id, 'done', 600);
  }

  // 以降は通常のイベント処理
  var text = params.event.text; // Slackから送信されたテキストを取得
  var response = ""; // 返信用のテキストを格納する変数

  // スプレッドシートから質問と回答を取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  var range = sheet.getDataRange();
  var values = range.getValues();

  for (var i = 1; i < values.length; i++) {
    var questionKeyword = values[i][0];
    var answer = values[i][1];
    if (text.includes(questionKeyword)) {
      response = answer;
      break;
    }
  }
  
  // Slackにメッセージを返信
  if (response !== "") {
    var url = 'https://slack.com/api/chat.postMessage';
    var token = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN'); // SlackのBot User OAuth Access Tokenをここに設定
    var channel = params.event.channel;
    var payload = {
      "channel": channel,
      "text": response,
      "as_user": true
    };
    var options = {
      "method" : "post",
      "contentType" : "application/json",
      "headers": {"Authorization": "Bearer " + token},
      "payload" : JSON.stringify(payload)
    };

    UrlFetchApp.fetch(url, options);
  }
  return HtmlService.createHtmlOutput("OK");
}