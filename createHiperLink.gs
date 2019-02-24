function createHiperLink() {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets      = spreadSheet.getSheets();
  var newSheet    = sheetInsertOrFind("ハイパーリンク生成用");
  for (var num = 0; num < sheets.length; num++) {
    var text = createSheetLink(sheets[num], sheets[num].getName());
    newSheet.getRange(num + 1, 1).setValue(text);
  }
}

// ハイパーリンクの式を返却する関数
function createSheetLink(sheet, text) {
  return '=HYPERLINK("#gid=' + sheet.getSheetId() + '","' + text + '")'
}

// シートを見つける処理
function sheetInsertOrFind(name, index) {
  var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadSheet.getSheetByName(name);
  
  if (sheet == null) {
    sheet = (index != null) ? spreadSheet.insertSheet(name, index) : spreadSheet.insertSheet(name);
  }
  
  return sheet;
}