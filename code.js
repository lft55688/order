// 1. 處理 GET 請求 (讀取資料)
function doGet(e) {
  var action = e.parameter.action;
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var response = { status: "success", data: null };

  try {
    if (action === 'getMenu') {
      var menuSheet = sheet.getSheetByName("Menu");
      if (menuSheet) {
        var lastRow = menuSheet.getLastRow();
        if (lastRow > 1) {
          var rawData = menuSheet.getRange(2, 1, lastRow - 1, 4).getValues();
          response.data = rawData.map(function(row) {
            return {
              id: row[0],
              name: row[1],
              isPublished: row[2] === true || row[2] === "true",
              data: row[3] ? JSON.parse(row[3]) : []
            };
          });
        } else {
          response.data = [];
        }
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
  } catch (error) {
    response = { status: "error", message: error.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
}

// 2. 處理 POST 請求 (寫入資料與 AI 辨識)
function doPost(e) {
  var response = { status: "success" };
  var lock = LockService.getScriptLock();
  
  if (lock.tryLock(5000)) {
    try {
      var payload = JSON.parse(e.postData.contents);
      var action = payload.action;
      var sheet = SpreadsheetApp.getActiveSpreadsheet();

      if (action === 'analyzeImage') {
        var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
        if (!apiKey) {
          throw new Error("尚未設定 GEMINI_API_KEY，請至專案設定 > 指令碼屬性中新增！");
        }

        var geminiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash:generateContent?key=" + apiKey;
        var geminiPayload = {
          contents: [{
            parts: [
              { text: "Extract menu items and prices. Return ONLY a valid JSON array. Structure: [{\"category\": \"Name\", \"items\": [{\"name\": \"Item\", \"price\": 150}]}]" },
              { inline_data: { mime_type: payload.mimeType, data: payload.base64Image } }
            ]
          }],
          generationConfig: { temperature: 0.1 }
        };

        var options = {
          method: "post",
          contentType: "application/json",
          payload: JSON.stringify(geminiPayload),
          muteHttpExceptions: true
        };

        var geminiRes = UrlFetchApp.fetch(geminiUrl, options);
        var jsonResponse = JSON.parse(geminiRes.getContentText());

        if (jsonResponse.error) {
          throw new Error("Gemini API 錯誤: " + jsonResponse.error.message);
        }

        var rawText = jsonResponse.candidates[0].content.parts[0].text;
        rawText = rawText.replace(/```json/g, '').replace(/```/g, '').trim();
        response.data = JSON.parse(rawText);
      }
      
      // 更新菜單
      else if (action === 'updateMenu') {
        var menuSheet = sheet.getSheetByName("Menu");
        if (!menuSheet) { 
          menuSheet = sheet.insertSheet("Menu"); 
          menuSheet.appendRow(["ID", "Name", "IsPublished", "Data JSON"]);
        }
        var lastRow = menuSheet.getLastRow();
        if (lastRow > 1) {
          menuSheet.getRange(2, 1, lastRow - 1, 4).clearContent();
        }
        var allMenus = payload.data;
        if (allMenus && allMenus.length > 0) {
          var rows = allMenus.map(function(m) {
            return [m.id, m.name, m.isPublished, JSON.stringify(m.data)];
          });
          menuSheet.getRange(2, 1, rows.length, 4).setValues(rows);
        }
      } 
      
      // 新增單筆訂單
      else if (action === 'addOrder') {
        var orderSheet = sheet.getSheetByName("Orders");
        if (!orderSheet) {
          orderSheet = sheet.insertSheet("Orders");
          orderSheet.appendRow(["ID", "Customer Name", "Items JSON", "Total", "Status", "Time"]);
        }
        var order = payload.data;
        orderSheet.appendRow([order.id, order.customerName, JSON.stringify(order.items), order.total, order.status, order.time]);
      }
      
      // 💡 覆蓋/修改所有訂單 (用於退換餐編輯)
      else if (action === 'updateOrders') {
        var orderSheet = sheet.getSheetByName("Orders");
        if (!orderSheet) {
          orderSheet = sheet.insertSheet("Orders");
          orderSheet.appendRow(["ID", "Customer Name", "Items JSON", "Total", "Status", "Time"]);
        }
        var lastRow = orderSheet.getLastRow();
        if (lastRow > 1) {
          orderSheet.deleteRows(2, lastRow - 1); // 刪除舊資料
        }
        
        var allOrders = payload.data;
        // 注意：前端傳來的是反轉過(最新的在前面)的，為了保持資料庫「最新的在最下面」，我們再次反轉寫入
        var reversedOrders = allOrders.slice().reverse(); 
        if (reversedOrders && reversedOrders.length > 0) {
          var rows = reversedOrders.map(function(o) {
            return [o.id, o.customerName, JSON.stringify(o.items), o.total, o.status, o.time];
          });
          orderSheet.getRange(2, 1, rows.length, 6).setValues(rows);
        }
      }

      // 清除所有訂單
      else if (action === 'clearOrders') {
        var orderSheet = sheet.getSheetByName("Orders");
        if (orderSheet) {
          var lastRow = orderSheet.getLastRow();
          if (lastRow > 1) {
            orderSheet.deleteRows(2, lastRow - 1);
          }
        }
      }
      else {
        throw new Error("接收到未知的 action 指令：'" + action + "'");
      }

    } catch (error) {
      response = { status: "error", message: error.toString() };
    } finally {
        lock.releaseLock();
    }
  } else {
    Logger.log("系統忙碌中，請稍後再試。");
  }
  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
}