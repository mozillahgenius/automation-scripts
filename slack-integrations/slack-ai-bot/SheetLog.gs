var SheetLog = {
    log: function(message) {
     const as = SpreadsheetApp.getActiveSpreadsheet();
     try {
       const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('log');
       var now = new Date();    
       logSheet.appendRow([now, message]);
     }catch(e){
     }
   }
 }
 