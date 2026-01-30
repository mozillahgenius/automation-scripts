function main(workbook: ExcelScript.Workbook) {
  try {
    // アクティブなワークシートを取得
    const worksheet = workbook.getActiveWorksheet();
    
    if (!worksheet) {
      console.log("ワークシートが見つかりませんでした。");
      return;
    }
    
    // C列の使用されている範囲を取得
    const usedRange = worksheet.getUsedRange();
    
    // データが存在しない場合は終了
    if (!usedRange) {
      console.log("シートにデータが見つかりませんでした。");
      return;
    }
    
    // 使用範囲の情報を取得
    const rowCount = usedRange.getRowCount();
    const firstRow = usedRange.getRowIndex();
    
    console.log(`処理開始: ${rowCount}行のデータを確認します`);
    
    // C列のデータを1行ずつ処理
    for (let i = 0; i < rowCount; i++) {
      // C列の各セルを取得 (行インデックス, 列インデックス2=C列, 行数1, 列数1)
      const cell = worksheet.getRangeByIndexes(firstRow + i, 2, 1, 1);
      
      if (!cell) {
        continue;
      }
      
      // セルの値を取得
      const cellValues = cell.getValues();
      
      if (!cellValues || !cellValues[0]) {
        continue;
      }
      
      const cellValue = cellValues[0][0];
      
      // 空のセルまたは値がない場合はスキップ
      if (!cellValue || cellValue === "") {
        continue;
      }
      
      // カタカナをローマ字に変換
      const romajiText = convertToRomaji(String(cellValue));
      
      // 変換した値をセルに書き戻す
      cell.setValue(romajiText);
    }
    
    console.log("変換が完了しました。");
    
  } catch (error) {
    console.log("エラーが発生しました: " + error);
  }
}

function convertToRomaji(katakana: string): string {
  let result = "";
  let i = 0;
  
  while (i < katakana.length) {
    const char = katakana[i];
    let nextChar = "";
    
    // 次の文字を確認
    if (i < katakana.length - 1) {
      nextChar = katakana[i + 1];
    }
    
    // 2文字の組み合わせをチェック
    if (nextChar) {
      const twoChars = char + nextChar;
      const twoCharRomaji = getTwoCharRomaji(twoChars);
      
      if (twoCharRomaji) {
        result += twoCharRomaji;
        i += 2;
        continue;
      }
    }
    
    // 促音の処理
    if (char === "ッ" && i < katakana.length - 1) {
      const nextCharRomaji = getSingleCharRomaji(katakana[i + 1]);
      if (nextCharRomaji && nextCharRomaji[0]) {
        result += nextCharRomaji[0];
      }
      i++;
      continue;
    }
    
    // 1文字の変換
    const singleCharRomaji = getSingleCharRomaji(char);
    if (singleCharRomaji) {
      result += singleCharRomaji;
    } else {
      // 変換できない文字はそのまま
      result += char;
    }
    
    i++;
  }
  
  return result;
}

function getTwoCharRomaji(twoChars: string): string | null {
  const twoCharMap: { [key: string]: string } = {
    // 拗音
    "キャ": "kya", "キュ": "kyu", "キョ": "kyo",
    "シャ": "sha", "シュ": "shu", "ショ": "sho",
    "チャ": "cha", "チュ": "chu", "チョ": "cho",
    "ニャ": "nya", "ニュ": "nyu", "ニョ": "nyo",
    "ヒャ": "hya", "ヒュ": "hyu", "ヒョ": "hyo",
    "ミャ": "mya", "ミュ": "myu", "ミョ": "myo",
    "リャ": "rya", "リュ": "ryu", "リョ": "ryo",
    "ギャ": "gya", "ギュ": "gyu", "ギョ": "gyo",
    "ジャ": "ja", "ジュ": "ju", "ジョ": "jo",
    "ビャ": "bya", "ビュ": "byu", "ビョ": "byo",
    "ピャ": "pya", "ピュ": "pyu", "ピョ": "pyo",
    "ヂャ": "dya", "ヂュ": "dyu", "ヂョ": "dyo",
    
    // 外来語表記
    "ファ": "fa", "フィ": "fi", "フェ": "fe", "フォ": "fo",
    "ヴァ": "va", "ヴィ": "vi", "ヴェ": "ve", "ヴォ": "vo",
    "ウィ": "wi", "ウェ": "we", "ウォ": "wo",
    "ティ": "ti", "ディ": "di",
    "トゥ": "tu", "ドゥ": "du",
    "ツァ": "tsa", "ツィ": "tsi", "ツェ": "tse", "ツォ": "tso",
    "チェ": "che",
    "ジェ": "je",
    "シェ": "she"
  };
  
  return twoCharMap[twoChars] || null;
}

function getSingleCharRomaji(char: string): string | null {
  const singleCharMap: { [key: string]: string } = {
    // 基本の五十音
    "ア": "a", "イ": "i", "ウ": "u", "エ": "e", "オ": "o",
    "カ": "ka", "キ": "ki", "ク": "ku", "ケ": "ke", "コ": "ko",
    "サ": "sa", "シ": "shi", "ス": "su", "セ": "se", "ソ": "so",
    "タ": "ta", "チ": "chi", "ツ": "tsu", "テ": "te", "ト": "to",
    "ナ": "na", "ニ": "ni", "ヌ": "nu", "ネ": "ne", "ノ": "no",
    "ハ": "ha", "ヒ": "hi", "フ": "fu", "ヘ": "he", "ホ": "ho",
    "マ": "ma", "ミ": "mi", "ム": "mu", "メ": "me", "モ": "mo",
    "ヤ": "ya", "ユ": "yu", "ヨ": "yo",
    "ラ": "ra", "リ": "ri", "ル": "ru", "レ": "re", "ロ": "ro",
    "ワ": "wa", "ヲ": "wo", "ン": "n",
    
    // 濁音
    "ガ": "ga", "ギ": "gi", "グ": "gu", "ゲ": "ge", "ゴ": "go",
    "ザ": "za", "ジ": "ji", "ズ": "zu", "ゼ": "ze", "ゾ": "zo",
    "ダ": "da", "ヂ": "ji", "ヅ": "zu", "デ": "de", "ド": "do",
    "バ": "ba", "ビ": "bi", "ブ": "bu", "ベ": "be", "ボ": "bo",
    
    // 半濁音
    "パ": "pa", "ピ": "pi", "プ": "pu", "ペ": "pe", "ポ": "po",
    
    // 小文字
    "ァ": "a", "ィ": "i", "ゥ": "u", "ェ": "e", "ォ": "o",
    "ャ": "ya", "ュ": "yu", "ョ": "yo",
    "ヮ": "wa",
    
    // 特殊文字
    "ー": "-",
    "ヴ": "vu"
  };
  
  return singleCharMap[char] || null;
}