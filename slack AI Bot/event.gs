function onEditTrigger(e) {
    //SheetLog.log("onEditTrigger");
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    if(sheet.getName() === 'faq'){
      //SheetLog.log("onEditTrigger: row="+range.getRow()+" col="+range.getColumn());
      // チェックボックスC列
      if (range.getColumn() === 3){
        if(range.getValue() === true && sheet.getRange(range.getRow(),1)) {
          var keyword = sheet.getRange(range.getRow(), 1).getValue(); // キーワード取得
          var result = searchDriveLink(keyword,sheet.getRange(range.getRow(), 4));
          //SheetLog.log("onEditTrigger:"+JSON.stringify(result));
          range.setValue(false);
        }
      }
    }
  }
  