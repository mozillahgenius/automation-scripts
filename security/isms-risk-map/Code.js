/**********************************************************************
 *  Risk Tools all-in-one  ★v6.3 2025-05-03
 *  ------------------------------------------------------------------
 *  ・「カルテ_リスクマップ」から項目→大→中→小の従属プルダウン＆自動入力  
 *  ・リスクマップ_AUTO を自動生成（CIA別行表示＋中カテゴリ結合）  
 *  ・ヒートマップ_AUTO に Before/After 両方の  
 *    発生可能性＋影響度合計で 3 段階カラーを適用  
 *
 *  マスタ:  A=項目, B=大, C=中, D=小, E=リスク詳細, F=対応策例  
 *  入力:    A=項目, B=大, C=中, D=小, E=リスク概要,  
 *           F=発生可能性, G=インパクト, H=対応策の例,  
 *           I=管理後の発生可能性, J=管理後のインパクト  
 *********************************************************************/

/*=== 定数設定 ========================================================*/
const SETTINGS = {
  listSheet       : 'カルテ_リスクマップ',
  masterSheet     : 'リスクマップマスタ',
  masterRange     : 'A:F',
  mapSheet        : 'リスクマップ_AUTO',
  heatSheet       : 'ヒートマップ_AUTO',
  colRiskItem     : 1,   // A=リスク項目
  colBig          : 2,   // B=大カテゴリ
  colMid          : 3,   // C=中カテゴリ
  colSmall        : 4,   // D=小カテゴリ（機密性/完全性/可用性）
  colSummary      : 5,   // E=リスク概要
  colProbBefore   : 6,   // F=発生可能性
  colImpactBefore : 7,   // G=インパクト
  colAction       : 8,   // H=対応策の例
  colProbAfter    : 9,   // I=管理後の発生可能性
  colImpactAfter  : 10   // J=管理後のインパクト
};

const SCALE = {
  prob:   { '低':1, '中':3, '高':5 },
  impact: { '軽微':1, '中':3, '重大':5 }
};

/*=== onOpen：メニュー登録＋初期プルダウン生成 =======================*/
function onOpen(){
  SpreadsheetApp.getUi()
    .createMenu('🗺️ Risk Tools')
    .addItem('① プルダウン再生成','setupAllValidations')
    .addItem('② リスクマップ再生成','buildRiskMap')
    .addItem('③ ヒートマップ再生成','buildRiskHeatmap')
    .addSeparator()
    .addItem('すべて再生成','rebuildAll')
    .addToUi();
  setupAllValidations();
}

function rebuildAll(){
  setupAllValidations();
  buildRiskMap();
  buildRiskHeatmap();
}

/*=== setupAllValidations：項目プルダウン＋大カテゴリを項目依存で絞り込み =============*/
function setupAllValidations(){
  const ss   = SpreadsheetApp.getActive();
  const inSh = ss.getSheetByName(SETTINGS.listSheet);
  const maSh = ss.getSheetByName(SETTINGS.masterSheet);
  if(!inSh || !maSh) return;

  // マスタデータ取得
  const raw = maSh.getRange(SETTINGS.masterRange).getValues()
    .slice(1)
    .filter(r => r[0]);
  const uniq = arr => [...new Set(arr)].filter(String);

  // 項目リスト作成
  const items = uniq(raw.map(r => r[0]));
  const dvItem = SpreadsheetApp.newDataValidation()
    .requireValueInList(items, true).build();

  // データ検証: リスク項目列
  const lastRow = inSh.getMaxRows();
  const numRows = lastRow - 1;
  inSh.getRange(2, SETTINGS.colRiskItem, numRows)
    .clearDataValidations()
    .setDataValidation(dvItem);

  // 各行ごとに、選択されたリスク項目に紐づく大カテゴリを絞り込む
  for(let r = 2; r <= lastRow; r++){
    const itemVal = inSh.getRange(r, SETTINGS.colRiskItem).getValue();
    let majors = [];
    if(itemVal){
      majors = uniq(
        raw.filter(row => row[0] === itemVal)
           .map(row => row[1])
      );
    }
    const majorCell = inSh.getRange(r, SETTINGS.colBig);
    majorCell.clearDataValidations();
    if(majors.length){
      const dvMajor = SpreadsheetApp.newDataValidation()
        .requireValueInList(majors, true).build();
      majorCell.setDataValidation(dvMajor);
    }
  }
}

/*=== onEdit：従属プルダウン＋自動入力＋再生成 ========================*/
function onEdit(e){
  if(!e || !e.range) return;
  const ss = SpreadsheetApp.getActive();
  const sh = e.range.getSheet();
  if(sh.getName() === SETTINGS.masterSheet){
    setupAllValidations();
    return;
  }
  if(sh.getName() !== SETTINGS.listSheet) return;

  const row = e.range.getRow(), col = e.range.getColumn();
  // キャッシュ更新 (5分キャッシュ)
  if(!globalThis._cache || Date.now() - _cache.ts > 5*60*1000){
    const raw2 = ss.getSheetByName(SETTINGS.masterSheet)
      .getRange(SETTINGS.masterRange).getValues()
      .slice(1).filter(r => r[0]);
    globalThis._cache = {
      ts: Date.now(),
      data: raw2.map(r => ({
        item:    r[0], big: r[1], mid: r[2], small: r[3],
        summary: r[4], action: r[5]
      }))
    };
  }
  const M = _cache.data;
  const uniqArr = arr => [...new Set(arr)].filter(String);
  const setDV = (cell, list) =>
    cell.setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(list, true).build()
    );
  function clearSA(){
    sh.getRange(row, SETTINGS.colSummary)
      .clearContent().clearDataValidations();
    sh.getRange(row, SETTINGS.colAction)
      .clearContent().clearDataValidations();
  }
  function setMidDV(r){
    const mids = uniqArr(
      M.filter(o => o.big === sh.getRange(r, SETTINGS.colBig).getValue())
       .map(o => o.mid)
    );
    if(mids.length) setDV(sh.getRange(r, SETTINGS.colMid), mids);
    else            sh.getRange(r, SETTINGS.colMid).clearDataValidations();
  }
  function setSmallDV(r){
    const smls = uniqArr(
      M.filter(o =>
        o.big === sh.getRange(r, SETTINGS.colBig).getValue() &&
        o.mid === sh.getRange(r, SETTINGS.colMid).getValue()
      ).map(o => o.small)
    );
    if(smls.length) setDV(sh.getRange(r, SETTINGS.colSmall), smls);
    else            sh.getRange(r, SETTINGS.colSmall).clearDataValidations();
  }
  function applySA(){
    const big = sh.getRange(row, SETTINGS.colBig).getValue();
    const mid = sh.getRange(row, SETTINGS.colMid).getValue();
    const sml = sh.getRange(row, SETTINGS.colSmall).getValue();
    const recs = sml
      ? M.filter(o => o.big===big && o.mid===mid && o.small===sml)
      : M.filter(o => o.big===big && o.mid===mid);
    const sums = uniqArr(recs.map(o=>o.summary));
    const acts = uniqArr(recs.map(o=>o.action));
    const sc = sh.getRange(row, SETTINGS.colSummary);
    const ac = sh.getRange(row, SETTINGS.colAction);
    if(sums.length===1)    sc.setValue(sums[0]).clearDataValidations();
    else if(sums.length>1) { sc.clearContent(); setDV(sc,sums); }
    else                     sc.clearContent().clearDataValidations();
    if(acts.length===1)    ac.setValue(acts[0]).clearDataValidations();
    else if(acts.length>1) { ac.clearContent(); setDV(ac,acts); }
    else                     ac.clearContent().clearDataValidations();
  }

  if(col === SETTINGS.colRiskItem){
    const rec = M.find(o => o.item === e.range.getValue());
    if(rec){
      sh.getRange(row, SETTINGS.colBig)  .setValue(rec.big);
      sh.getRange(row, SETTINGS.colMid)  .setValue(rec.mid);
      sh.getRange(row, SETTINGS.colSmall).setValue(rec.small);
    }
    setMidDV(row); setSmallDV(row);
    clearSA(); applySA();
  }
  if(col === SETTINGS.colBig){
    sh.getRange(row, SETTINGS.colMid,1,2)
      .clearContent().clearDataValidations();
    clearSA(); setMidDV(row);
  }
  if(col === SETTINGS.colMid){
    sh.getRange(row, SETTINGS.colSmall)
      .clearContent().clearDataValidations();
    clearSA(); setSmallDV(row);
  }
  if(col === SETTINGS.colSmall){
    clearSA(); applySA();
  }

  buildRiskMap();
  buildRiskHeatmap();
}

/*=== リスクマップ_AUTO 生成（CIA３行＋中カテゴリ結合） ============*/
function buildRiskMap(){
  const ss  = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(SETTINGS.listSheet);
  if(!src) return;

  const rows = src.getDataRange().getValues().slice(1)
    .filter(r=>
      r[SETTINGS.colBig-1] &&
      r[SETTINGS.colMid-1] &&
      r[SETTINGS.colSmall-1] &&
      r[SETTINGS.colSummary-1]
    ).map(r=>({
      big:     r[SETTINGS.colBig-1],
      mid:     r[SETTINGS.colMid-1],
      small:   r[SETTINGS.colSmall-1],    // ex. "機密性"
      summary: r[SETTINGS.colSummary-1],
      action:  r[SETTINGS.colAction-1]||''
    }));

  const bigs   = [...new Set(rows.map(r=>r.big))];
  const header = ['中カテゴリ','CIA', ...bigs];
  const facets = ['機密性','完全性','可用性'];
  const mids   = [...new Set(rows.map(r=>r.mid))];
  const tbl    = [header];

  mids.forEach(midVal=>{
    facets.forEach(fac=>{
      const line = [midVal, fac];
      bigs.forEach(bigVal=>{
        const notes = rows.filter(o=>
            o.big===bigVal &&
            o.mid===midVal &&
            o.small===fac
        ).map(o=>{
          let txt = o.summary;
          if(o.action) txt += `（対応策: ${o.action}）`;
          return txt;
        });
        line.push(notes.join('\n'));
      });
      tbl.push(line);
    });
  });

  let sh = ss.getSheetByName(SETTINGS.mapSheet) || ss.insertSheet(SETTINGS.mapSheet);
  sh.clearContents();
  sh.getRange(1,1,tbl.length,tbl[0].length)
    .setValues(tbl)
    .setWrap(true)
    .setBorder(true,true,true,true,true,true);
  sh.getRange(1,1,1,tbl[0].length).setFontWeight('bold');
  sh.getRange(1,1,tbl.length,2).setFontWeight('bold');
  sh.autoResizeColumns(1,tbl[0].length);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);

  // 中カテゴリ列(第1列)を3行ずつマージ
  let r = 2;
  mids.forEach(_=>{
    sh.getRange(r,1,facets.length,1).mergeVertically();
    r += facets.length;
  });
}

/*=== ヒートマップ_AUTO 生成 =======================================*/
function buildRiskHeatmap(){
  const ss  = SpreadsheetApp.getActive();
  const src = ss.getSheetByName(SETTINGS.listSheet);
  if(!src) return;

  const toNum = v=>{
    if(v==null||v==='') return null;
    const n=Number(v);
    return isNaN(n)?(SCALE.prob[v]||SCALE.impact[v]):n;
  };
  const raw = src.getDataRange().getValues().slice(1).map(r=>({
    before:[ toNum(r[SETTINGS.colProbBefore-1]), toNum(r[SETTINGS.colImpactBefore-1]) ],
    after: [ toNum(r[SETTINGS.colProbAfter -1]), toNum(r[SETTINGS.colImpactAfter -1]) ],
    note:  `${r[SETTINGS.colSmall-1]}：${r[SETTINGS.colSummary-1]}（対応策: ${r[SETTINGS.colAction-1]||''}）`
  })).filter(o=>
    o.before[0]!=null && o.before[1]!=null &&
    o.after[0] !=null && o.after[1] !=null
  );

  const axis=[1,2,3,4,5];
  const countB={}, detailsB={}, countA={}, detailsA={};
  raw.forEach(o=>{
    const kB=`${o.before[0]}_${o.before[1]}`;
    const kA=`${o.after [0]}_${o.after [1]}`;
    countB[kB]=(countB[kB]||0)+1;
    countA[kA]=(countA[kA]||0)+1;
    (detailsB[kB]||(detailsB[kB]=[])).push(o.note);
    (detailsA[kA]||(detailsA[kA]=[])).push(o.note);
  });

  // Before table
  const tblB=[['発生\\影響',...axis]];
  axis.slice().reverse().forEach(p=>{
    const row=[p];
    axis.forEach(i=>row.push(countB[`${p}_${i}`]||''));
    tblB.push(row);
  });
  // After table
  const tblA=[['発生\\影響',...axis]];
  axis.slice().reverse().forEach(p=>{
    const row=[p];
    axis.forEach(i=>row.push(countA[`${p}_${i}`]||''));
    tblA.push(row);
  });

  let sh = ss.getSheetByName(SETTINGS.heatSheet) || ss.insertSheet(SETTINGS.heatSheet);
  sh.clearContents();

  // Before 出力
  sh.getRange(1,1).setValue('Before');
  const sb = 2;
  sh.getRange(sb,1,tblB.length,tblB[0].length)
    .setValues(tblB)
    .setBorder(true,true,true,true,true,true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sh.getRange(sb,1,1,axis.length+1).setFontWeight('bold');
  sh.getRange(sb,1,axis.length+1,1).setFontWeight('bold');

  // After 出力
  const sa = sb + tblB.length + 2;
  sh.getRange(sa-1,1).setValue('After');
  sh.getRange(sa,1,tblA.length,tblA[0].length)
    .setValues(tblA)
    .setBorder(true,true,true,true,true,true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sh.getRange(sa,1,1,axis.length+1).setFontWeight('bold');
  sh.getRange(sa,1,axis.length+1,1).setFontWeight('bold');

  // ノート設定（Before）
  const rev = axis.slice().reverse();
  tblB.slice(1).forEach((row, ri) => {
    row.slice(1).forEach((_, ci) => {
      const keyB = `${rev[ri]}_${axis[ci]}`;
      const cellB  = sh.getRange(sb + ri + 1, ci + 2);
      const notesB = detailsB[keyB];
      if (notesB && notesB.length) {
        cellB.setNote(notesB.join('\n'));
      } else {
        cellB.clearNote();
      }
    });
  });

  // ノート設定（After）
  tblA.slice(1).forEach((row, ri) => {
    row.slice(1).forEach((_, ci) => {
      const keyA = `${rev[ri]}_${axis[ci]}`;
      const cellA  = sh.getRange(sa + ri + 1, ci + 2);
      const notesA = detailsA[keyA];
      if (notesA && notesA.length) {
        cellA.setNote(notesA.join('\n'));
      } else {
        cellA.clearNote();
      }
    });
  });

  // 条件付き書式
  const bodyB = sh.getRange(sb+1,2,axis.length,axis.length);
  const bodyA = sh.getRange(sa+1,2,axis.length,axis.length);
  const rules = [
    // Before
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(B$2>=3,$A3+B$2<=5)')
      .setBackground('#FFF9EB').setRanges([bodyB]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(B$2>=3,$A3+B$2>=6,$A3+B$2<=7)')
      .setBackground('#FFE5CC').setRanges([bodyB]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND($A3+B$2>=8,$A3+B$2<=10)')
      .setBackground('#FFCCCC').setRanges([bodyB]).build(),
    // After
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(B$${sa}>=3,$A${sa+1}+B$${sa}<=5)`)
      .setBackground('#FFF9EB').setRanges([bodyA]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(B$${sa}>=3,$A${sa+1}+B$${sa}>=6,$A${sa+1}+B$${sa}<=7)`)
      .setBackground('#FFE5CC').setRanges([bodyA]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($A${sa+1}+B$${sa}>=8,$A${sa+1}+B$${sa}<=10)`)
      .setBackground('#FFCCCC').setRanges([bodyA]).build()
  ];
  sh.setConditionalFormatRules(rules);
}
