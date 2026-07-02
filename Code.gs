/**
 * exYahoo Summer Party QR-Code 報名與簽到系統 - Apps Script 後端
 * 
 * 部署說明：
 * 1. 在 Google 試算表中點選「擴充功能」->「Apps Script」。
 * 2. 清空原本的程式碼，貼入此 Code.gs 檔案。
 * 3. 點選左側「檔案」旁的「+」新增 HTML 檔案，命名為「Index」。將 Index.html 的內容貼入。
 * 4. 點選專案設定（齒輪圖示），在「指令碼屬性」中新增以下屬性：
 *    - SECRET_TOKEN : 自訂一個複雜的字串（例如：exyahoo2026partysec）做為安全金鑰。
 * 5. 設定觸發器（左側時鐘圖示），新增觸發器：
 *    - 執行函式：handleEdit
 *    - 活動來源：試算表
 *    - 活動類型：編輯時
 * 6. 點選右上角「部署」->「新增部署」。
 *    - 類型選擇「網頁應用程式 (Web App)」。
 *    - 執行身分：您自己 (Me)。
 *    - 誰能存取：任何人 (Anyone)。
 *    - 複製部署完產生的 Web App URL。
 * 7. 在指令碼屬性中新增屬性：
 *    - WEB_APP_URL : 貼上剛剛複製的 Web App URL。
 */

// 取得指令碼屬性中的設定
function getSecretToken() {
  return PropertiesService.getScriptProperties().getProperty('SECRET_TOKEN') || 'default_secret_token';
}

function getWebAppUrl() {
  return PropertiesService.getScriptProperties().getProperty('WEB_APP_URL') || '';
}

/**
 * 當試算表被編輯時觸發 (需要設定為 Installable Trigger 才能有發信權限)
 */
function handleEdit(e) {
  var sheet = e.range.getSheet();
  
  // 只處理表單回應的工作表 (通常名稱為 "表單回應 1" 或類似)
  // 如果您的工作表名稱不同，可以修改此處，或是只要偵測欄位符合就處理
  var sheetName = sheet.getName();
  if (sheetName.indexOf("表單回應") === -1 && sheetName.indexOf("Responses") === -1 && sheetName !== "Sheet1") {
    // 若不是回應頁面就跳過，以策安全
    // 您也可以拔除此限制
  }
  
  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();
  
  // 避免首列標頭觸發
  if (row === 1) return;
  
  // 取得標頭欄位映射
  var headerMap = getHeaderMap(sheet);
  
  // 檢查是否包含必要的欄位，如果沒有，代表欄位名稱不符
  var statusCol = headerMap['對帳狀態'];
  var uuidCol = headerMap['Ticket UUID'];
  var emailCol = headerMap['電子郵件地址'] || headerMap['電子郵件'] || headerMap['Email'];
  var nameCol = headerMap['姓名'] || headerMap['聯絡人姓名'];
  
  if (!statusCol || !uuidCol || !emailCol || !nameCol) {
    Logger.log("找不到對應欄位，請檢查試算表首列的標頭名稱是否包含：'對帳狀態'、'Ticket UUID'、'電子郵件地址'、'姓名'");
    return;
  }
  
  // 判斷編輯的是否為「對帳狀態」這一欄
  if (col === statusCol) {
    var statusValue = range.getValue();
    
    // 當狀態變更為「匯款完成」
    if (statusValue === "匯款完成") {
      var currentUuid = sheet.getRange(row, uuidCol).getValue();
      
      // 如果已經有 UUID，代表之前已發過門票，避免重複發送
      if (currentUuid) {
        Logger.log("Row " + row + " 已經有 UUID，不重複發送門票。");
        return;
      }
      
      // 1. 產生 UUID
      var uuid = Utilities.getUuid();
      sheet.getRange(row, uuidCol).setValue(uuid);
      
      // 2. 取得聯絡人資訊
      var email = sheet.getRange(row, emailCol).getValue();
      var name = sheet.getRange(row, nameCol).getValue();
      
      if (!email) {
        Logger.log("Row " + row + " 沒有 Email 資料，無法寄送。");
        return;
      }
      
      // 3. 發送門票郵件
      try {
        sendTicketEmail(email, name, uuid);
        Logger.log("成功發送門票信件給：" + name + " (" + email + ")");
      } catch (err) {
        Logger.log("寄送信件失敗: " + err.toString());
      }
    }
  }
}

/**
 * 寄送門票電子郵件
 */
function sendTicketEmail(email, name, uuid) {
  var token = getSecretToken();
  var webAppUrl = getWebAppUrl();
  
  if (!webAppUrl) {
    throw new Error("尚未設定 WEB_APP_URL 指令碼屬性！請先部署 Web App 並填入屬性。");
  }
  
  // 組合簽到網址
  var checkInUrl = webAppUrl + "?uuid=" + uuid + "&token=" + token;
  
  // 組合 QR Code 圖片網址 (使用第三方 QR Server API，無流量限制)
  var qrCodeImageUrl = "https://api.qrserver.com/v1/create-qr-code/?size=250x250&data=" + encodeURIComponent(checkInUrl);
  
  var subject = "【exYahoo Summer Party】您的入場門票已核發！";
  
  // 漂亮的 HTML 信件範本
  var htmlBody = 
    "<div style='font-family: \"Helvetica Neue\", Helvetica, Arial, sans-serif; max-width: 500px; margin: 0 auto; padding: 25px; border: 1px solid #e5e7eb; border-radius: 16px; background-color: #ffffff; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);'>" +
      "<div style='text-align: center; margin-bottom: 20px;'>" +
        "<h2 style='color: #4f46e5; margin: 0; font-size: 24px; font-weight: 700; letter-spacing: -0.025em;'>exYahoo Summer Party</h2>" +
        "<p style='color: #6b7280; font-size: 14px; margin: 5px 0 0 0;'>活動入場憑證</p>" +
      "</div>" +
      "<div style='border-top: 1px dashed #e5e7eb; border-bottom: 1px dashed #e5e7eb; padding: 20px 0; text-align: center;'>" +
        "<p style='font-size: 16px; color: #1f2937; margin: 0 0 10px 0;'>親愛的 <strong>" + name + "</strong> 您好：</p>" +
        "<p style='font-size: 14px; color: #4b5563; margin: 0 0 20px 0; line-height: 1.5;'>我們已確認您的匯款，報名成功！<br>請妥善保存下方 QR Code，於活動當天進場時出示掃描。</p>" +
        "<div style='background-color: #f9fafb; padding: 15px; display: inline-block; border-radius: 12px; margin-bottom: 15px;'>" +
          "<img src='" + qrCodeImageUrl + "' alt='Entry QR Code' style='width: 200px; height: 200px; display: block;' />" +
        "</div>" +
        "<div style='font-family: monospace; font-size: 11px; color: #9ca3af; margin-top: 5px;'>Ticket ID: " + uuid + "</div>" +
      "</div>" +
      "<div style='margin-top: 20px; font-size: 13px; color: #6b7280; line-height: 1.6;'>" +
        "<p style='margin: 0 0 5px 0;'><strong>📅 活動時間：</strong> 2026年 Summer (詳細時間請參照活動公告)</p>" +
        "<p style='margin: 0 0 5px 0;'><strong>📍 活動地點：</strong> exYahoo Summer Party 會場</p>" +
        "<p style='margin: 0; color: #ef4444;'>⚠️ 注意事項：本票券僅限一人單次入場使用，請勿將 QR Code 轉傳給他人。</p>" +
      "</div>" +
    "</div>";

  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: htmlBody
  });
}

/**
 * 處理 Web App GET 請求 (手機掃描 QR Code 後的跳轉頁面)
 */
function doGet(e) {
  var uuid = e.parameter.uuid;
  var token = e.parameter.token;
  
  // 1. 安全驗證：檢查 Token 是否正確
  var expectedToken = getSecretToken();
  if (!token || token !== expectedToken) {
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('無權限 - exYahoo Check-in')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  
  // 建立回傳 Template
  var template = HtmlService.createTemplateFromFile('Index');
  template.uuid = uuid || "";
  template.token = token;
  
  // 若沒有帶入 UUID
  if (!uuid) {
    template.status = "error";
    template.message = "無效的門票連結 (缺少 UUID)";
    template.attendeeName = "";
    template.checkInTime = "";
    return template.evaluate()
      .setTitle('簽到失敗 - exYahoo')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // 2. 尋找 UUID 並進行簽到處理
  var result = processCheckIn(uuid);
  template.status = result.status;
  template.message = result.message;
  template.attendeeName = result.name;
  template.checkInTime = result.time;
  
  var title = "簽到結果 - exYahoo";
  if (result.status === "success") title = "簽到成功 - " + result.name;
  else if (result.status === "warning") title = "重複簽到 - " + result.name;
  else title = "簽到失敗";
  
  return template.evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 簽到核心 logic (搜尋 Sheet 並更新狀態)
 */
function processCheckIn(uuid) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; // 預設使用第一個工作表
  var headerMap = getHeaderMap(sheet);
  
  var uuidCol = headerMap['Ticket UUID'];
  var nameCol = headerMap['姓名'] || headerMap['聯絡人姓名'];
  var statusCol = headerMap['入場狀態'];
  var timeCol = headerMap['入場時間'];
  
  if (!uuidCol || !nameCol || !statusCol || !timeCol) {
    return {
      status: "error",
      name: "",
      time: "",
      message: "試算表欄位設定有誤，請確認是否包含 'Ticket UUID'、'姓名'、'入場狀態'、'入場時間' 等欄位。"
    };
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return { status: "error", name: "", time: "", message: "試算表中無任何報名資料。" };
  }
  
  // 取得整欄 UUID
  var uuidRange = sheet.getRange(2, uuidCol, lastRow - 1, 1);
  var uuidValues = uuidRange.getValues();
  
  var targetRow = -1;
  for (var i = 0; i < uuidValues.length; i++) {
    if (uuidValues[i][0].toString().trim() === uuid.trim()) {
      targetRow = i + 2; // 因為從第2列開始，且 index 從 0 開始
      break;
    }
  }
  
  if (targetRow === -1) {
    return { status: "error", name: "", time: "", message: "找不到此門票，可能是無效的 UUID。" };
  }
  
  // 取得來賓姓名與當前入場狀態
  var name = sheet.getRange(targetRow, nameCol).getValue();
  var checkInStatus = sheet.getRange(targetRow, statusCol).getValue();
  
  // 格式化當前時間
  var now = new Date();
  var formattedTime = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  
  // 判斷入場狀態
  if (!checkInStatus || checkInStatus.toString().trim() === "") {
    // 尚未入場 -> 更新為已入場
    sheet.getRange(targetRow, statusCol).setValue("已入場");
    sheet.getRange(targetRow, timeCol).setValue(formattedTime);
    
    return {
      status: "success",
      name: name,
      time: formattedTime,
      message: "簽到成功！歡迎光臨 exYahoo Summer Party 🎉"
    };
  } else {
    // 已經入場過 -> 顯示警告與首次入場時間
    var originalTime = sheet.getRange(targetRow, timeCol).getValue();
    var formattedOriginalTime = originalTime;
    if (originalTime instanceof Date) {
      formattedOriginalTime = Utilities.formatDate(originalTime, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    }
    
    return {
      status: "warning",
      name: name,
      time: formattedOriginalTime,
      message: "此門票先前已完成簽到！請勿重複使用。"
    };
  }
}

/**
 * 輔助函式：取得標頭與欄位編號的對照表
 */
function getHeaderMap(sheet) {
  var lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) return {};
  
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var headerName = headers[i].toString().trim();
    if (headerName) {
      map[headerName] = i + 1; // 1-based index
    }
  }
  return map;
}
