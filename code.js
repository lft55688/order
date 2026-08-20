// 1. 處理 GET 請求 (讀取資料)
function doGet(e) {
  var action = e.parameter.action;
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var response = { status: "success", data: null };

  if (action === 'getMenu') {
    var menuSheet = sheet.getSheetByName("Menu");
    if (menuSheet) {
      var data = menuSheet.getRange("A1").getValue();
      response.data = data ? JSON.parse(data) : [];
    } else {
      response.data = [];
    }
  } 
  else if (action === 'getOrders') {
    var orderSheet = sheet.getSheetByName("Orders");
    if (orderSheet) {
      var lastRow = orderSheet.getLastRow();
      if (lastRow > 1) {
        var rawData = orderSheet.getRange(2, 1, lastRow - 1, 6).getValues();
        var orders = rawData.map(function(row) {
          return {
            id: row[0],
            customerName: row[1],
            items: JSON.parse(row[2]),
            total: row[3],
            status: row[4],
            time: row[5]
          };
        });
        response.data = orders.reverse(); // 最新的在前面
      } else {
        response.data = [];
      }
    } else {
      response.data = [];
    }
  }

  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
}

// 2. 處理 POST 請求 (寫入資料)
function doPost(e) {
  var payload = JSON.parse(e.postData.contents);
  var action = payload.action;
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var response = { status: "success" };

  if (action === 'updateMenu') {
    var menuSheet = sheet.getSheetByName("Menu");
    if (!menuSheet) { menuSheet = sheet.insertSheet("Menu"); }
    menuSheet.getRange("A1").setValue(JSON.stringify(payload.data));
  } 
  else if (action === 'addOrder') {
    var orderSheet = sheet.getSheetByName("Orders");
    if (!orderSheet) {
      orderSheet = sheet.insertSheet("Orders");
      orderSheet.appendRow(["ID", "Customer Name", "Items JSON", "Total", "Status", "Time"]);
    }
    var order = payload.data;
    orderSheet.appendRow([order.id, order.customerName, JSON.stringify(order.items), order.total, order.status, order.time]);
  }

  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
}