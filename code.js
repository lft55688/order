// 1. 處理 GET 請求 (讀取資料)
function doGet(e) {
  var action = e.parameter.action;
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var response = { status: "success", data: null };

  try {
    if (action === 'getMenu') {
      var menuSheet = sheet.getSheetByName("Menu");
      if (menuSheet) {
        var data = menuSheet.getRange("A1").getValue();
        // 如果有資料就解析，沒資料就給空陣列
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
  } catch (error) {
    response = { status: "error", message: error.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
}

// 2. 處理 POST 請求 (寫入資料與 AI 辨識)
function doPost(e) {
  var response = { status: "success" };
  // 1. 取得指令碼鎖定物件
  var lock = LockService.getScriptLock();
  // 2. 嘗試取得鎖定，最多等待 5 秒 (5000 毫秒)
  if (lock.tryLock(5000)) {
    try {
      var payload = JSON.parse(e.postData.contents);
      var action = payload.action;
      var sheet = SpreadsheetApp.getActiveSpreadsheet();

      // 透過 GAS 呼叫 Gemini API 辨識圖片
      if (action === 'analyzeImage') {
        var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
        if (!apiKey) {
          throw new Error("尚未設定 GEMINI_API_KEY，請至專案設定 > 指令碼屬性中新增！");
        }

        // 💡 修正模型名稱：Google 目前最新的穩定版本為 gemini-1.5-flash
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

        // 清理 Markdown 格式
        var rawText = jsonResponse.candidates[0].content.parts[0].text;
        rawText = rawText.replace(/```json/g, '').replace(/```/g, '').trim();
        
        response.data = JSON.parse(rawText);
      }
      // 上傳/覆蓋菜單結構 (現在會存入包含多個菜單的陣列)
      else if (action === 'updateMenu') {
        var menuSheet = sheet.getSheetByName("Menu");
        if (!menuSheet) { menuSheet = sheet.insertSheet("Menu"); }
        menuSheet.getRange("A1").setValue(JSON.stringify(payload.data));
      } 
      // 新增訂單
      else if (action === 'addOrder') {
        var orderSheet = sheet.getSheetByName("Orders");
        if (!orderSheet) {
          orderSheet = sheet.insertSheet("Orders");
          orderSheet.appendRow(["ID", "Customer Name", "Items JSON", "Total", "Status", "Time"]);
        }
        var order = payload.data;
        orderSheet.appendRow([order.id, order.customerName, JSON.stringify(order.items), order.total, order.status, order.time]);
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
      // 3. 釋放鎖定
      lock.releaseLock();
    }
  } else {
    response = { status: "error", message: "系統忙碌中，請稍後再試。" };
  }
  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
}