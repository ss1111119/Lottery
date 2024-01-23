function Lottery() {
  var token = "你的權杖";
  
  // 獲取活動的試算表
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!activeSpreadsheet) {
    // 如果試算表為 null，顯示錯誤訊息並退出函數
    console.error("無法取得試算表");
    return;
  }

  var sheet1 = activeSpreadsheet.getSheetByName("抽籤");
  var stock = [];
  
  // 從特定儲存格中獲取報酬率門檻值和中籤率門檻值
  var returnRateThreshold = sheet1.getRange("B1").getValue();
  var ballotRateThreshold = sheet1.getRange("D1").getValue();
  
  // 迴圈遍歷試算表的特定範圍
  for (i = 2; i < 10; i++) {
    // 從試算表中獲取股票相關資訊
    stock[0] = sheet1.getRange(i, 2).getValue();  // 股號名稱
    stock[1] = sheet1.getRange(i, 4).getValue();  // 申購日
    stock[2] = sheet1.getRange(i, 10).getValue(); // 報酬率
    stock[3] = sheet1.getRange(i, 13).getValue(); // 中籤率
    stock[4] = sheet1.getRange(i, 14).getValue(); // 備註
    
    // 檢查條件是否滿足
    if (stock[2] > returnRateThreshold && stock[3] > ballotRateThreshold && stock[4] == '申購中') { 
      // 滿足條件，組合通知訊息並發送 LINE 通知
      var message = "\n" + stock[0] + "\n申購日" + stock[1] + "\n" + "報酬率" + stock[2] + "% 中籤率" + stock[3] + "%";
      sendline(message, token);      
    }  
  }
}
