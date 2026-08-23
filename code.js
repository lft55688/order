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
          // 讀取 A2:D (ID, Name, isPublished, Data JSON)
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
      
      // 更新菜單 (新邏輯：將每個菜單拆分存在獨立的 Row)
      else if (action === 'updateMenu') {
        var menuSheet = sheet.getSheetByName("Menu");
        if (!menuSheet) { 
          menuSheet = sheet.insertSheet("Menu"); 
          menuSheet.appendRow(["ID", "Name", "IsPublished", "Data JSON"]);
        }
        
        // 1. 清除舊資料 (保留第一行標題)
        var lastRow = menuSheet.getLastRow();
        if (lastRow > 1) {
          menuSheet.getRange(2, 1, lastRow - 1, 4).clearContent();
        }
        
        // 2. 寫入新資料 (一行一個菜單)
        var allMenus = payload.data;
        if (allMenus && allMenus.length > 0) {
          var rows = allMenus.map(function(m) {
            return [m.id, m.name, m.isPublished, JSON.stringify(m.data)];
          });
          menuSheet.getRange(2, 1, rows.length, 4).setValues(rows);
        }
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