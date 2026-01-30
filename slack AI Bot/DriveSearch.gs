/**
 * FAQシートの A列キーワード、C列チェックを元に、
 * ドライブ一覧シート A列のすべてのフォルダIDを検索対象として
 * 指定キーワードの含まれるファイル本文／行を抜き出し、
 * D列に結果をリッチテキストで書き出します。
 *
 * B列の手動回答はそのまま残し、FAQシートの構造は変更しません。
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
 * @returns Array<{file: Object, snippets: string[]}>
 */
function searchDriveLinkReturn(keyword, folderId) {
  const ret = [];
  const baseUrl = 'https://www.googleapis.com/drive/v3/files';
  const params = {
    q:                          `\'${folderId}\' in parents and trashed = false and fullText contains '${keyword}'`,
    corpora:                    'allDrives',
    includeItemsFromAllDrives:  true,
    supportsAllDrives:          true,
    fields:                     'files(id,name,mimeType,webViewLink)'
  };
  const query = Object.entries(params)
    .map(([k,v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
    .join('&');
  const url = `${baseUrl}?${query}`;

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
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
        const blob      = DriveApp.getFileById(file.id).getBlob();
        const tmpFile   = DriveApp.createFile(blob).setName('temp');
        const resource  = { title: 'temp-doc', mimeType: MimeType.GOOGLE_DOCS };
        const converted = Drive.Files.insert(resource, tmpFile.getBlob());
        snippets        = extractSnippetFromDoc(converted.id, keyword);
        DriveApp.getFileById(converted.id).setTrashed(true);
        tmpFile.setTrashed(true);
      }
    } catch (e) {
      Logger.log(`処理エラー (${file.name}): ${e}`);
    }
    if (snippets.length > 0) {
      ret.push({ file: file, snippets: snippets });
    }
  });

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
  let text     = '';
  const links  = [];

  data.forEach(item => {
    const nameBlock    = item.file.name + '\n';
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
    const end   = start + nameBlock.length;
    if (end <= maxLen) {
      links.push({ start, end, url: item.file.webViewLink });
    }
  });

  const builder = SpreadsheetApp.newRichTextValue().setText(text);
  links.forEach(l => builder.setLinkUrl(l.start, l.end, l.url));
  range.setRichTextValue(builder.build());
}



