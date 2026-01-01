function doGet(e) {
  var page = e.parameter.page || 'index';
  return HtmlService.createTemplateFromFile(page)
      .evaluate()
      .setTitle('俄羅斯方塊專題')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 取得當前網頁部署的網址，用於頁面跳轉
function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

// 【CRUD】儲存分數
function saveScore(name, score) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Scores') || ss.insertSheet('Scores');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID', 'Name', 'Score', 'Date']);
  }
  var date = new Date().toLocaleString();
  sheet.appendRow([Utilities.getUuid(), name, score, date]);
  return "分數上傳成功！";
}
// 【CRUD】讀取排行榜
function getLeaderboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Scores');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  data.shift(); // 移除標題
  data.sort(function(a, b) { return b[2] - a[2]; }); // 分數由高到低
  return data.slice(0, 10);
}
// 【CRUD - Delete】 刪除特定 ID 的分數
function deleteScore(id) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Scores');
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) { // 比對 UUID
      sheet.deleteRow(i + 1);
      return "刪除成功";
    }
  }
  return "找不到該項目";
}

// 【CRUD - Update】 修改玩家名稱
function updatePlayerName(id, newName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Scores');
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i + 1, 2).setValue(newName); // 修改第二欄 (Name)
      return "修改成功";
    }
  }
  return "修改失敗";
}
