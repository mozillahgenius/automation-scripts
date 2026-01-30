function katakanaToHiragana(text)
{
  return text.replace(/[\u30a1-\u30f6]/g, function(match) {
    // カタカナの文字コードからひらがなの文字コードへ変換
    var chr = match.charCodeAt(0) - 0x60;
    return String.fromCharCode(chr);
  });
}

function toHalfWidth(str) {
  // 全角英数字を半角に変換
  str = str.replace(/[Ａ-Ｚａ-ｚ０-９]/g, function(s) {
    return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
  return str;
}


function test_katakanaToHiragana(){
  Logger.log(katakanaToHiragana('アイウエオ'));
}