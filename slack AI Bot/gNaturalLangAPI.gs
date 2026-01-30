/**
 * Google Natural Language API alnalyzeSyntax
 */
function gNL(textdata) {
  var apiKey = ScriptProperties.getProperty('GOOGLE_NL_API');  // ここに取得したAPIキーを入れる
    //形態素解析（品詞取得） = analyzeSyntax
  var url = "https://language.googleapis.com/v1/documents:analyzeSyntax?key=" + apiKey;
  var payload = {
    document: {
      type: "PLAIN_TEXT",
      content: textdata
    },
    encodingType: "UTF8"
  };  
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };
  
  var response = UrlFetchApp.fetch(url, options);
  //SheetLog.log('NL:' + response.getContentText());
  //Logger.log(response.getContentText());
  try{
    return JSON.parse(response.getContentText());  
  }catch(e){
    Logger.log(response.getContentText());
    Logger.log(e);
    return null;
  }
}

/**
 * Google Natural Language API の戻り値より必要なものを抽出する
 * 品詞の場合は tagsの欄に ['NOUN','NUM','NUMBER']
 * https://cloud.google.com/natural-language/docs/morphology?hl=ja
 */
function filterGNL(gNLobj,tags){
  if (!gNLobj)return [];
  var words = gNLobj.tokens
    .filter(token => tags.includes(token.partOfSpeech.tag)) // 配列内の品詞と一致するものを抽出
    .map(token => token.text.content); 
  return words;
}

/**
{
  "sentences": [
    {
      "text": {
        "content": "入力された文のテキスト",
        "beginOffset": 0
      }
    }
  ],
  "tokens": [
    {
      "text": {
        "content": "トークンのテキスト",
        "beginOffset": 0
      },
      "partOfSpeech": {
        "tag": "品詞タグ",
        "aspect": "体",
        "case": "格",
        "form": "形式",
        "gender": "性",
        "mood": "法",
        "number": "数",
        "person": "人称",
        "proper": "固有名詞かどうか",
        "reciprocity": "相互関係",
        "tense": "時制",
        "voice": "態"
      },
      "dependencyEdge": {
        "headTokenIndex": 1,
        "label": "依存関係ラベル"
      },
      "lemma": "レンマ（原形）"
    }
  ],
  "language": "ja"
}


https://cloud.google.com/natural-language/docs/morphology?hl=ja

 * 
 * 
 */