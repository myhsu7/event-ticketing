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

// ==========================================
// 系統全域設定
// ==========================================

// 測試模式開關：若設定為 true，則所有手動執行或觸發編輯時，都只會處理「備註」欄位含有 "test" (不分大小寫) 的行數。
var GLOBAL_TEST_MODE = false; 

// ==========================================
// 活動行事曆資訊設定
// ==========================================
var EVENT_TITLE = "2026退虎會 exYahoo Summer Party";
var EVENT_START = "2026-07-24T18:00:00+08:00";
var EVENT_END = "2026-07-24T21:00:00+08:00";
var EVENT_LOCATION = "台北松山意舍酒店 17 樓 Que 原木燒烤餐廳";
var EVENT_DESCRIPTION = "歡迎參加 exYahoo Summer Party！請憑門票信中的 QR Code 掃描入場。";
var CALENDAR_ID = "dd4d904608c5422751a1d50c122fabe01106eac06778b8679defd8757ca67043@group.calendar.google.com"; // 使用退虎會專用日曆 ID 

// 取得指令碼屬性中的設定
function getSecretToken() {
  return PropertiesService.getScriptProperties().getProperty('SECRET_TOKEN') || 'default_secret_token';
}

function getWebAppUrl() {
  return PropertiesService.getScriptProperties().getProperty('WEB_APP_URL') || '';
}

/**
 * 手動執行：發送未發送的門票 (生產環境)
 */
function runSendTicketsProduction() {
  processPendingTickets(false);
}

/**
 * 手動執行：測試發送門票 (僅處理「備註」含有 'test' 的資料)
 * 在測試模式下，即使該列已經有 Ticket UUID，也會重新產生並發送，以便重複測試。
 */
function runSendTicketsTest() {
  processPendingTickets(true);
}

/**
 * 當試算表被編輯時觸發 (需要設定為 Installable Trigger 才能有發信權限)
 */
function handleEdit(e) {
  var sheet = e.range.getSheet();
  
  // 只處理名稱為「匯款對帳」的工作表
  var sheetName = sheet.getName();
  if (sheetName !== "匯款對帳") {
    return;
  }
  
  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();
  
  // 避免首列標頭觸發
  if (row === 1) return;
  
  // 取得標頭欄位映射
  var headerMap = getHeaderMap(sheet);
  
  var statusCol = headerMap['對帳狀態'];
  var uuidCol = headerMap['Ticket UUID'];
  var emailCol = headerMap['電子郵件地址'] || headerMap['電子郵件'] || headerMap['Email'];
  var nameCol = headerMap['姓名'] || headerMap['聯絡人姓名'];
  var noteCol = headerMap['備註'] || headerMap['Remarks'] || headerMap['備註欄'];
  var ticketStatusCol = headerMap['門票發送狀態'] || headerMap['門票狀態'] || headerMap['入場券發送狀態'];
  
  if (!statusCol || !uuidCol || !emailCol || !nameCol) {
    Logger.log("找不到對應欄位，請檢查試算表首列的標頭名稱是否包含：'對帳狀態'、'Ticket UUID'、'電子郵件地址'、'姓名'");
    return;
  }
  
  // 判斷編輯的是否為「對帳狀態」這一欄
  if (col === statusCol) {
    var statusValue = range.getValue();
    
    // 當狀態變更為「匯款完成」
    if (statusValue === "匯款完成") {
      
      // 1. 測試模式過濾：若全域測試模式開啟，但該列備註不含 "test"，則跳過
      var noteValue = noteCol ? sheet.getRange(row, noteCol).getValue().toString().toLowerCase() : "";
      var isTestRow = noteValue.indexOf("test") !== -1;
      
      if (GLOBAL_TEST_MODE && !isTestRow) {
        Logger.log("[測試模式] 跳過未標記 'test' 的列: " + row);
        return;
      }
      
      var currentUuid = sheet.getRange(row, uuidCol).getValue();
      var currentTicketStatus = ticketStatusCol ? sheet.getRange(row, ticketStatusCol).getValue() : "";
      
      // 如果已經有 UUID 或已經標記為「已發送」，且「不是測試列」，代表之前已發過門票，避免重複發送
      if ((currentUuid || currentTicketStatus === "已發送") && !isTestRow) {
        Logger.log("Row " + row + " 已經發送過門票，跳過處理。");
        return;
      }
      
      // 2. 產生 UUID 并更新發送狀態
      var uuid = Utilities.getUuid();
      sheet.getRange(row, uuidCol).setValue(uuid);
      if (ticketStatusCol && !isTestRow) {
        sheet.getRange(row, ticketStatusCol).setValue("已發送");
      }
      
      // 3. 取得聯絡人資訊
      var email = sheet.getRange(row, emailCol).getValue();
      var name = sheet.getRange(row, nameCol).getValue();
      
      if (!email) {
        Logger.log("Row " + row + " 沒有 Email 資料，無法寄送。");
        return;
      }
      
      // 4. 發送門票郵件
      try {
        sendTicketEmail(email, name, uuid);
        Logger.log((isTestRow ? "[測試] " : "") + "成功發送門票信件給：" + name + " (" + email + ")");
      } catch (err) {
        Logger.log("寄送信件失敗: " + err.toString());
      }
      
      // 5. 邀請至 Google 日曆 (透過 Google 日曆發送邀請，帶入 uuid 以整合 QR Code)
      inviteAttendeeToCalendar(email, name, uuid);
    }
  }
}

/**
 * 批次處理未發送的門票 (支援測試模式與正式模式)
 */
function processPendingTickets(isTestMode) {
  var sheet = getActiveResponseSheet();
  var headerMap = getHeaderMap(sheet);
  
  // 偵錯日誌：印出目前抓取的工作表與所有讀取到的標頭
  Logger.log("【偵錯】目前執行的工作表分頁: '" + sheet.getName() + "'");
  Logger.log("【偵錯】該分頁讀取到的所有標頭: " + JSON.stringify(Object.keys(headerMap)));
  
  var statusCol = headerMap['對帳狀態'];
  var uuidCol = headerMap['Ticket UUID'];
  var emailCol = headerMap['電子郵件地址'] || headerMap['電子郵件'] || headerMap['Email'];
  var nameCol = headerMap['姓名'] || headerMap['聯絡人姓名'];
  var noteCol = headerMap['備註'] || headerMap['Remarks'] || headerMap['備註欄'];
  var ticketStatusCol = headerMap['門票發送狀態'] || headerMap['門票狀態'] || headerMap['入場券發送狀態'];
  
  if (!statusCol || !uuidCol || !emailCol || !nameCol) {
    Logger.log("找不到對應欄位，請確認標頭包含：'對帳狀態'、'Ticket UUID'、'電子郵件地址'、'姓名'");
    return;
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log("沒有資料可處理。");
    return;
  }
  
  var count = 0;
  for (var row = 2; row <= lastRow; row++) {
    var noteValue = noteCol ? sheet.getRange(row, noteCol).getValue().toString().toLowerCase() : "";
    var isTestRow = noteValue.indexOf("test") !== -1;
    
    // 如果是測試模式，只處理備註含 "test" 的列
    if (isTestMode && !isTestRow) {
      continue;
    }
    
    // 如果是正式模式，且開啟了全域測試，安全起見只處理測試列
    if (!isTestMode && GLOBAL_TEST_MODE && !isTestRow) {
      continue;
    }
    
    var statusValue = sheet.getRange(row, statusCol).getValue();
    if (statusValue === "匯款完成") {
      var currentUuid = sheet.getRange(row, uuidCol).getValue();
      var currentTicketStatus = ticketStatusCol ? sheet.getRange(row, ticketStatusCol).getValue() : "";
      
      // 非測試列若已有 UUID 或已標記「已發送」則跳過
      if ((currentUuid || currentTicketStatus === "已發送") && !isTestRow) {
        continue;
      }
      
      // 產生或沿用 UUID (測試列每次都產生新 UUID 方便重複測試)
      var uuid = (isTestRow) ? Utilities.getUuid() : (currentUuid || Utilities.getUuid());
      sheet.getRange(row, uuidCol).setValue(uuid);
      if (ticketStatusCol && !isTestRow) {
        sheet.getRange(row, ticketStatusCol).setValue("已發送");
      }
      
      var email = sheet.getRange(row, emailCol).getValue();
      var name = sheet.getRange(row, nameCol).getValue();
      
      if (email) {
        try {
          sendTicketEmail(email, name, uuid);
          Logger.log((isTestRow ? "[測試] " : "[正式] ") + "批次發送門票給：" + name + " (" + email + ")");
          count++;
        } catch (err) {
          Logger.log("批次寄送失敗 Row " + row + ": " + err.toString());
        }
        
        // 批次執行時也自動發送日曆邀請
        inviteAttendeeToCalendar(email, name, uuid);
      }
    }
  }
  Logger.log("執行完畢，共處理 " + count + " 筆資料。");
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
  
  // 產生個人行事曆快速加入連結
  var gCalLink = getGoogleCalendarLink(EVENT_TITLE, EVENT_START, EVENT_END, EVENT_DESCRIPTION, EVENT_LOCATION);
  var yCalLink = getYahooCalendarLink(EVENT_TITLE, EVENT_START, EVENT_END, EVENT_DESCRIPTION, EVENT_LOCATION);
  
  // 格式化活動時間顯示 (以台灣時區為準)
  var startDate = new Date(EVENT_START);
  var endDate = new Date(EVENT_END);
  var formattedEventTime = Utilities.formatDate(startDate, "GMT+8", "yyyy年 MM月 dd日 HH:mm") + 
                           " - " + 
                           Utilities.formatDate(endDate, "GMT+8", "HH:mm");
  
  var subject = "【exYahoo Summer Party】您的入場門票已核發！";
  
  // 漂亮的 HTML 信件範本 (包含 Google 與 Yahoo 日曆加入按鈕)
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
      
      // 行事曆快速加入按鈕區
      "<div style='margin: 20px 0; padding: 18px; background-color: #f9fafb; border-radius: 12px; text-align: center; border: 1px solid #f3f4f6;'>" +
        "<p style='font-size: 13px; color: #4b5563; font-weight: 600; margin: 0 0 12px 0;'>📅 快速加入我的個人行事曆：</p>" +
        "<div style='margin-bottom: 8px;'>" +
          "<a href='" + gCalLink + "' target='_blank' style='display: inline-block; background-color: #4285f4; color: #ffffff; padding: 8px 14px; border-radius: 8px; font-size: 12px; text-decoration: none; font-weight: 600; margin: 0 5px;'>+ Google 日曆</a>" +
          "<a href='" + yCalLink + "' target='_blank' style='display: inline-block; background-color: #6001d2; color: #ffffff; padding: 8px 14px; border-radius: 8px; font-size: 12px; text-decoration: none; font-weight: 600; margin: 0 5px;'>+ Yahoo 日曆</a>" +
        "</div>" +
        "<p style='font-size: 11px; color: #9ca3af; margin: 0;'>（Gmail 使用者亦將自動收到正式 Google 日曆受邀通知）</p>" +
      "</div>" +
      
      "<div style='margin-top: 20px; font-size: 13px; color: #6b7280; line-height: 1.6;'>" +
        "<p style='margin: 0 0 5px 0;'><strong>📅 活動時間：</strong> " + formattedEventTime + "</p>" +
        "<p style='margin: 0 0 5px 0;'><strong>📍 活動地點：</strong> " + EVENT_LOCATION + "</p>" +
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
  else if (result.status === "early") title = "尚未開放 - " + result.name;
  else title = "簽到失敗";
  
  return template.evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 簽到核心 logic (搜尋 Sheet 並更新狀態)
 */
function processCheckIn(uuid) {
  var sheet = getActiveResponseSheet();
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
  var noteCol = headerMap['備註'] || headerMap['Remarks'] || headerMap['備註欄'];
  var noteValue = noteCol ? sheet.getRange(targetRow, noteCol).getValue().toString().toLowerCase() : "";
  var isTestRow = noteValue.indexOf("test") !== -1;
  
  // 3. 檢查是否已開放簽到 (活動開始前 1 小時開放) - 測試列除外
  var now = new Date();
  var eventStartTime = new Date(EVENT_START);
  var checkInOpenTime = new Date(eventStartTime.getTime() - (60 * 60 * 1000));
  
  if (now < checkInOpenTime && !isTestRow) {
    var formattedOpenTime = Utilities.formatDate(checkInOpenTime, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
    return {
      status: "early",
      name: name,
      time: "",
      message: "本活動尚未開放簽到！\n開放簽到時間為：" + formattedOpenTime + " 起。\n期待您的光臨！"
    };
  }
  
  // 格式化當前時間
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
      map[headerName.toLowerCase()] = i + 1; // case-insensitive mapping
    }
  }
  return map;
}

/**
 * 自動搜尋所有分頁，尋找真正包含「對帳狀態」的報名工作表。
 * 避免因為試算表有空白的首個工作表 (Sheet1) 導致系統抓錯分頁。
 */
function getActiveResponseSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("匯款對帳");
  if (sheet) {
    return sheet;
  }
  // 備用：若找不到名為「匯款對帳」的分頁，則使用第一個分頁
  return ss.getSheets()[0];
}

/**
 * 取得或建立主活動 (Master Event) 並將使用者加為受邀者
 * 這會對 Google 帳號與非 Google 帳號 (如 Yahoo) 自動發出 iCal 日曆邀請信。
 */
function inviteAttendeeToCalendar(email, name, uuid) {
  try {
    var calendar = CalendarApp.getCalendarById(CALENDAR_ID) || CalendarApp.getDefaultCalendar();
    
    // 取得 Web App URL 與 Token 用以產生 QR Code 連結
    var token = getSecretToken();
    var webAppUrl = getWebAppUrl();
    var checkInUrl = webAppUrl ? (webAppUrl + "?uuid=" + uuid + "&token=" + token) : "";
    
    // 產生 QR Code 圖片連結
    var qrCodeImageUrl = "https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" + encodeURIComponent(checkInUrl);
    
    // 產生個人化的活動說明 (將專屬 QR Code 連結寫入日曆說明中)
    var personalDesc = "親愛的 " + name + " 您好：\n\n" +
                       "歡迎參加 2026退虎會 exYahoo Summer Party！\n" +
                       "本活動憑 QR Code 掃描入場，您可以在手機日曆中直接點擊下方連結出示您的入場券：\n\n" +
                       "🎫 您的專屬門票 QR Code 連結：\n" + qrCodeImageUrl + "\n\n" +
                       "--- \n" +
                       "門票編號 (UUID): " + uuid + "\n" +
                       "活動地點: " + EVENT_LOCATION + "\n" +
                       "活動時間: " + EVENT_START.substring(0,10) + " 18:00 - 21:00\n\n" +
                       "期待與您相見！";
    
    // 為每位來賓建立獨立的日曆活動，以夾帶個別的專屬 QR Code 連結
    var event = calendar.createEvent(EVENT_TITLE, new Date(EVENT_START), new Date(EVENT_END), {
      location: EVENT_LOCATION,
      description: personalDesc
    });
    
    // 確保來賓隱私：關閉「查看受邀者名單」、「修改活動」與「邀請他人」功能
    event.setGuestsCanSeeGuests(false);
    event.setGuestsCanModify(false);
    event.setGuestsCanInviteOthers(false);
    
    // 將該來賓加入此專屬活動
    event.addGuest(email);
    Logger.log("已成功為來賓建立專屬日曆活動並發送邀請：" + name + " (" + email + ")");
  } catch (err) {
    Logger.log("日曆邀請失敗: " + err.toString());
  }
}

/**
 * 產生 Yahoo 日曆的新增活動連結
 */
function getYahooCalendarLink(title, startIso, endIso, desc, loc) {
  var st = convertToUtcString(startIso);
  var et = convertToUtcString(endIso);
  return "https://calendar.yahoo.com/?v=60&view=d&type=20" +
         "&title=" + encodeURIComponent(title) +
         "&st=" + st +
         "&et=" + et +
         "&desc=" + encodeURIComponent(desc) +
         "&in_loc=" + encodeURIComponent(loc);
}

/**
 * 產生 Google 日曆的新增活動連結
 */
function getGoogleCalendarLink(title, startIso, endIso, desc, loc) {
  var st = convertToUtcString(startIso);
  var et = convertToUtcString(endIso);
  return "https://calendar.google.com/calendar/render?action=TEMPLATE" +
         "&text=" + encodeURIComponent(title) +
         "&dates=" + st + "/" + et +
         "&details=" + encodeURIComponent(desc) +
         "&location=" + encodeURIComponent(loc);
}

/**
 * 輔助函式：將 ISO 字串轉換成 UTC 格式 (YYYYMMDDTHHMMSSZ)
 */
function convertToUtcString(isoString) {
  var date = new Date(isoString);
  var pad = function(n) { return (n < 10 ? '0' : '') + n; };
  return date.getUTCFullYear() +
         pad(date.getUTCMonth() + 1) +
         pad(date.getUTCDate()) +
         'T' +
         pad(date.getUTCHours()) +
         pad(date.getUTCMinutes()) +
         pad(date.getUTCSeconds()) +
         'Z';
}


// ==========================================
// 兆豐銀行自動對帳系統設定
// ==========================================
var SPREADSHEET_ID = "1g-6SaVCkIZASiZO11Pb9mAsktOOB6rZkWi6bqa-IUQg";

// Telegram 異常通知設定 (選填，留空則不發送)
var TELEGRAM_BOT_TOKEN = "8551328354:AAF2jD02kYnILVK3jZiB7n55SSJ3_aWeLhw"; // 兆豐自動對帳機器人 Token
var TELEGRAM_CHAT_ID = "7594339427";   // 明彥個人 Chat ID

// 異常通知電子信箱 (選填，留空則預設寄給執行腳本的您本人)
var ADMIN_NOTIFY_EMAIL = "mingyen@gmail.com"; 

/**
 * 建立自動對帳的三個定時觸發器 (12:00, 18:00, 21:00)
 * 您可以在 Apps Script 編輯器中手動執行此函式一次即可完成設定。
 */
function setupAutoReconciliationTrigger() {
  // 清除舊的同名觸發器，防重複
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "runAutoReconciliation") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // 建立每日 12:00, 18:00, 21:00 觸發器 (Google 會在該整點的前後數分鐘內觸發)
  ScriptApp.newTrigger("runAutoReconciliation")
    .timeBased()
    .everyDays(1)
    .atHour(12)
    .create();
    
  ScriptApp.newTrigger("runAutoReconciliation")
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .create();
    
  ScriptApp.newTrigger("runAutoReconciliation")
    .timeBased()
    .everyDays(1)
    .atHour(21)
    .create();
    
  Logger.log("已成功建立 12:00, 18:00, 21:00 的自動對帳時間驅動觸發器！");
}

/**
 * 自動對帳系統主入口 (由定時觸發器呼叫)
 */
function runAutoReconciliation() {
  Logger.log("--- 開始執行自動對帳流程 ---");
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 1. 讀取 Gmail 信件
  var query = 'from:service@emailer.megabank.com.tw subject:"兆豐銀行數位存款" subject:"轉入交易" is:unread';
  var threads = GmailApp.search(query);
  Logger.log("搜尋到未讀信件 Thread 數: " + threads.length);
  
  var bankTransactions = [];
  var anomalies = [];
  
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      if (message.isUnread()) {
        var body = message.getPlainBody();
        // 如果純文字內容為空或是 HTML 預設的 fallback 訊息，則解析 HTML 內文並轉為純文字
        if (!body || body.indexOf("HTML") !== -1 || body.indexOf("格式信件") !== -1 || body.indexOf("讀信程式") !== -1) {
          body = convertHtmlToPlainText(message.getBody());
        }
        var msgId = message.getId();
        
        var tx = parseMegaBankEmail(body);
        if (tx.isValid) {
          tx.messageId = msgId;
          bankTransactions.push(tx);
          Logger.log("解析成功 -> 金額: " + tx.amount + ", 後五碼: " + tx.lastFive + " (Msg ID: " + msgId + ")");
        } else {
          var errorMsg = "信件解析異常 (Msg ID: " + msgId + ") - 原因: " + tx.errorReason + " (本文摘要: " + body.substring(0, 100).replace(/\n/g, " ") + ")";
          anomalies.push(errorMsg);
          Logger.log(errorMsg);
        }
        
        // 標示為已讀，防止重複處理
        message.markRead();
      }
    }
  }
  
  if (bankTransactions.length === 0 && anomalies.length === 0) {
    Logger.log("無任何新交易或異常，對帳結束。");
    return;
  }
  
  // 2. 依「後五碼」將交易分組
  var bankTxGrouped = {};
  for (var k = 0; k < bankTransactions.length; k++) {
    var tx = bankTransactions[k];
    if (!bankTxGrouped[tx.lastFive]) {
      bankTxGrouped[tx.lastFive] = [];
    }
    bankTxGrouped[tx.lastFive].push(tx);
  }
  
  var regSheet = ss.getSheetByName("報名名單");
  var formSheet = ss.getSheetByName("表單回覆 1");
  
  if (!regSheet || !formSheet) {
    var sheetError = "找不到必要的試算表分頁：'報名名單' 或 '表單回覆 1'";
    anomalies.push(sheetError);
    sendReconciliationAlert(anomalies);
    return;
  }
  
  var regHeaderMap = getHeaderMap(regSheet);
  var formHeaderMap = getHeaderMap(formSheet);
  
  var anyChanges = false;
  
  // 3. 遍歷每組後五碼，進行比對與寫入
  for (var lastFive in bankTxGrouped) {
    var txs = bankTxGrouped[lastFive];
    var B = txs.length; // 銀行通知筆數
    
    // A. 檢查「報名名單」是否存在此後五碼 (J 欄，index 10)
    var regMatches = findRowsInSheet(regSheet, 10, lastFive);
    var M = regMatches.length; // 報名名單筆數
    
    if (M > 0) {
      // 情況 A：已存在於報名名單
      if (B === M) {
        // 筆數相符，將這些列的 N 欄 (匯款完成，Col 14) 改為「匯款完成」
        for (var r = 0; r < regMatches.length; r++) {
          regSheet.getRange(regMatches[r], 14).setValue("匯款完成");
        }
        Logger.log("後五碼 [" + lastFive + "] 已存在報名名單且筆數相符 (" + B + " 筆)，已更新狀態為匯款完成。");
        anyChanges = true;
      } else {
        // 筆數不符，列為異常
        var countError = "後五碼 [" + lastFive + "] 筆數不符：銀行通知有 " + B + " 筆，但報名名單中有 " + M + " 筆。已跳過，需人工核對。";
        anomalies.push(countError);
        Logger.log(countError);
      }
    } else {
      // 情況 B：不在報名名單，去「表單回覆 1」比對 (I 欄，index 9)
      var formMatches = findRowsInSheet(formSheet, 9, lastFive);
      var F = formMatches.length; // 表單回覆筆數
      
      if (F > 0) {
        // 排序表單回覆 (依 L 欄「匯款時間」由早到晚)
        formMatches.sort(function(rowA, rowB) {
          var timeA = getValAsDate(formSheet.getRange(rowA, 12).getValue());
          var timeB = getValAsDate(formSheet.getRange(rowB, 12).getValue());
          return timeA.getTime() - timeB.getTime();
        });
        
        if (B === F) {
          // 筆數相符，全數寫入 (Append) 報名名單
          for (var f = 0; f < formMatches.length; f++) {
            var formRowValues = formSheet.getRange(formMatches[f], 1, 1, formSheet.getLastColumn()).getValues()[0];
            
            // 取得目前的最高序號，並 + 1
            var nextSerial = getNextSerialNumber(regSheet);
            var nextSerialStr = ("00" + nextSerial).slice(-3); // 補零至三碼
            
            // 解析並組裝寫入「報名名單」的欄位
            var parsed = getFormRowData(formRowValues, formHeaderMap);
            var newRowValues = [
              parsed.timestamp,                     // A: 時間戳記
              "'" + nextSerialStr,                  // B: 序號 (文字格式)
              parsed.name,                          // C: 姓名
              parsed.email,                         // D: Email
              parsed.phone,                         // E: 手機號碼
              parsed.dept,                          // F: 前 Yahoo 部門/團隊
              parsed.nickname,                      // G: 任職年份 / 暱稱
              parsed.food,                          // H: 餐飲備註
              "是",                                 // I: 是否已完成匯款
              parsed.lastFive,                      // J: 匯款帳號末五碼
              parsed.amount,                        // K: 匯款金額
              parsed.time,                          // L: 匯款時間
              parsed.screenshot,                    // M: 匯款截圖連結
              "匯款完成",                           // N: 匯款完成
              "",                                   // O: 空白
              ""                                    // P: 空白
            ];
            
            // 寫入報名名單
            var newRowIdx = regSheet.getLastRow() + 1;
            var range = regSheet.getRange(newRowIdx, 1, 1, newRowValues.length);
            range.getCell(1, 2).setNumberFormat("@"); // 確保序號欄為文字格式
            range.setValues([newRowValues]);
          }
          Logger.log("後五碼 [" + lastFive + "] 已成功從表單匯入報名名單且筆數相符 (" + B + " 筆)。");
          anyChanges = true;
        } else {
          // 筆數不符，列為異常
          var formCountError = "後五碼 [" + lastFive + "] 筆數不符：銀行通知有 " + B + " 筆，但表單回覆中有 " + F + " 筆。已跳過，需人工核對。";
          anomalies.push(formCountError);
          Logger.log(formCountError);
        }
      } else {
        // 找不到對應的報名或表單回覆
        var notFoundError = "後五碼 [" + lastFive + "] 無法在「報名名單」與「表單回覆 1」中找到任何報名資料。已跳過，需人工核對。";
        anomalies.push(notFoundError);
        Logger.log(notFoundError);
      }
    }
  }
  
  // 4. 同步「匯款對帳」分頁的 H:M 欄位
  if (anyChanges || bankTransactions.length > 0) {
    Logger.log("正在重新同步「匯款對帳」分頁...");
    syncPaymentReconciliationSheet(ss);
    Logger.log("「匯款對帳」分頁同步完成。");
    
    // 5. 自動觸發寄送門票與發送日曆邀請 (針對新標記為「匯款完成」且尚未發票的人)
    Logger.log("自動啟動批次發信與日曆邀請流程...");
    processPendingTickets(false);
  }
  
  // 6. 若有異常，發送通知
  if (anomalies.length > 0) {
    sendReconciliationAlert(anomalies);
  }
  
  Logger.log("--- 自動對帳流程執行完畢 ---");
}

/**
 * 解析兆豐銀行轉入通知信件
 */
function parseMegaBankEmail(body) {
  var tx = { isValid: false, amount: 0, lastFive: "", errorReason: "" };
  
  // 1. 擷取金額
  var amountRegex = /(?:轉入金額|金額)\s*[:：]\s*(?:NT\$)?\s*([0-9,.]+)/i;
  var amountMatch = body.match(amountRegex);
  if (amountMatch) {
    tx.amount = parseFloat(amountMatch[1].replace(/,/g, ''));
  } else {
    tx.errorReason = "無法解析交易金額";
    return tx;
  }
  
  // 2. 檢查金額是否為 1000 或 1000.00
  if (tx.amount !== 1000) {
    tx.errorReason = "交易金額不符 (金額為 " + tx.amount + " 元，應為 1000 元)";
    return tx;
  }
  
  // 3. 擷取交易備註後五碼
  var remarkRegex = /交易備註（部份內容）\s*[:：]\s*([^\n\r]+)/;
  var remarkMatch = body.match(remarkRegex);
  if (remarkMatch) {
    var remarkText = remarkMatch[1].toString().trim();
    // 剔除非數字字元
    var digits = remarkText.replace(/\D/g, '');
    if (digits.length >= 5) {
      tx.lastFive = digits.slice(-5);
      tx.isValid = true;
    } else if (digits.length > 0) {
      tx.lastFive = digits; // 不足 5 碼但有數字
      tx.errorReason = "交易備註數字不足五碼 (僅有 '" + digits + "')";
    } else {
      tx.errorReason = "交易備註無任何數字 (備註為 '" + remarkText + "')";
    }
  } else {
    tx.errorReason = "找不到 '交易備註（部份內容）' 關鍵字";
  }
  
  return tx;
}

/**
 * 取得「報名名單」目前最高序號，用以累加
 */
function getNextSerialNumber(regSheet) {
  var lastRow = regSheet.getLastRow();
  var maxSerial = 0;
  if (lastRow > 1) {
    var serialValues = regSheet.getRange(2, 2, lastRow - 1, 1).getValues();
    for (var r = 0; r < serialValues.length; r++) {
      var val = serialValues[r][0].toString().trim();
      var num = parseInt(val, 10);
      if (!isNaN(num) && num > maxSerial) {
        maxSerial = num;
      }
    }
  }
  return maxSerial;
}

/**
 * 輔助搜尋特定欄位值符合指定資料的列號陣列 (1-based row indices)
 */
function findRowsInSheet(sheet, colIndex, targetValue) {
  var lastRow = sheet.getLastRow();
  var matches = [];
  if (lastRow <= 1) return matches;
  
  var values = sheet.getRange(2, colIndex, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0].toString().trim() === targetValue.trim()) {
      matches.push(i + 2); // 調整回 1-based 列號
    }
  }
  return matches;
}

/**
 * 解析 Form Responses 的各欄位內容 (動態比對標頭名稱，相容度最高)
 */
function getFormRowData(rowValues, formHeaderMap) {
  var getVal = function(names, defaultCol) {
    for (var i = 0; i < names.length; i++) {
      var col = formHeaderMap[names[i].toLowerCase()];
      if (col) return rowValues[col - 1];
    }
    if (defaultCol && defaultCol <= rowValues.length) {
      return rowValues[defaultCol - 1];
    }
    return "";
  };
  
  return {
    timestamp: getVal(['時間戳記'], 1),
    name: getVal(['姓名', '您的姓名'], 2),
    email: getVal(['電子郵件地址', 'email', '電子郵件', '電子信箱'], 15), // 優先找 O 欄 (Col 15)
    phone: getVal(['手機號碼', '手機', '電話', '聯絡電話'], 4),
    dept: getVal(['前 yahoo 部門/團隊', '部門', '團隊'], 5),
    nickname: getVal(['任職年份 / 暱稱', '任職年份', '暱稱'], 6),
    food: getVal(['餐飲備註', '素食', '餐飲'], 7),
    lastFive: getVal(['匯款帳號末五碼', '匯款帳號後5碼', '帳號後5碼', '後五碼'], 9), // I 欄 (Col 9)
    amount: getVal(['匯款金額', '金額'], 11), // K 欄 (Col 11)
    time: getVal(['匯款時間', '時間'], 12), // L 欄 (Col 12)
    screenshot: getVal(['匯款截圖連結', '匯款截圖', '截圖'], 13)
  };
}

/**
 * 同步「匯款對帳」分頁的 H:M 欄位
 */
function syncPaymentReconciliationSheet(ss) {
  var reconSheet = ss.getSheetByName("匯款對帳");
  var regSheet = ss.getSheetByName("報名名單");
  var formSheet = ss.getSheetByName("表單回覆 1");
  
  if (!reconSheet || !regSheet || !formSheet) return;
  
  var reconLastRow = reconSheet.getLastRow();
  if (reconLastRow <= 1) return;
  
  var reconData = reconSheet.getRange(2, 1, reconLastRow - 1, 13).getValues(); // 讀取 A-M 欄
  
  // 1. 將「匯款對帳」依 G 欄 (存摺備註，即後五碼，Col 7) 分組
  var bankGroups = {};
  for (var i = 0; i < reconData.length; i++) {
    var rowNum = i + 2;
    var rowValues = reconData[i];
    var lastFive = rowValues[7 - 1].toString().trim();
    if (!lastFive) continue;
    
    if (!bankGroups[lastFive]) {
      bankGroups[lastFive] = [];
    }
    bankGroups[lastFive].push({ rowNum: rowNum, values: rowValues });
  }
  
  // 2. 將「報名名單」依 J 欄 (匯款帳號末五碼，Col 10) 分組
  var regLastRow = regSheet.getLastRow();
  var regData = regLastRow > 1 ? regSheet.getRange(2, 1, regLastRow - 1, 14).getValues() : [];
  
  var regGroups = {};
  for (var j = 0; j < regData.length; j++) {
    var rowNum = j + 2;
    var rowValues = regData[j];
    var lastFive = rowValues[10 - 1].toString().trim();
    if (!lastFive) continue;
    
    if (!regGroups[lastFive]) {
      regGroups[lastFive] = [];
    }
    regGroups[lastFive].push({ rowNum: rowNum, values: rowValues });
  }
  
  // 3. 將「表單回覆 1」依 I 欄 (匯款帳號末五碼，Col 9) 分組
  var formLastRow = formSheet.getLastRow();
  var formData = formLastRow > 1 ? formSheet.getRange(2, 1, formLastRow - 1, 15).getValues() : [];
  
  var formGroups = {};
  for (var f = 0; f < formData.length; f++) {
    var rowNum = f + 2;
    var rowValues = formData[f];
    var lastFive = rowValues[9 - 1].toString().trim();
    if (!lastFive) continue;
    
    if (!formGroups[lastFive]) {
      formGroups[lastFive] = [];
    }
    formGroups[lastFive].push({ rowNum: rowNum, values: rowValues });
  }
  
  // 4. 開始一對一比對與寫入
  for (var lastFive in bankGroups) {
    var reconGroup = bankGroups[lastFive];
    var regGroup = regGroups[lastFive] || [];
    var formGroup = formGroups[lastFive] || [];
    
    var X = reconGroup.length;
    var Y = regGroup.length;
    
    // 排序名單與表單（依匯款時間）
    regGroup.sort(function(a, b) {
      return getValAsDate(a.values[12 - 1]).getTime() - getValAsDate(b.values[12 - 1]).getTime();
    });
    formGroup.sort(function(a, b) {
      return getValAsDate(a.values[12 - 1]).getTime() - getValAsDate(b.values[12 - 1]).getTime();
    });
    
    if (Y === 0) {
      // 找不到對應
      for (var k = 0; k < X; k++) {
        var rowNum = reconGroup[k].rowNum;
        reconSheet.getRange(rowNum, 8, 1, 3).setValues([["", "", ""]]); // 清空 H, I, J
        reconSheet.getRange(rowNum, 11).setValue("找不到對應：" + lastFive); // K
      }
    } else if (X === Y) {
      // 筆數完全相符 -> 進行一對一寫入
      for (var k = 0; k < X; k++) {
        var rowNum = reconGroup[k].rowNum;
        var serial = regGroup[k].values[2 - 1]; // B: 序號
        var name = regGroup[k].values[3 - 1];   // C: 姓名
        
        // J 欄 Email 優先去「表單回覆 1」的 O 欄 (Col 15) 找，找不到才用「報名名單」D 欄 (Col 4)
        var email = "";
        if (k < formGroup.length && formGroup[k].values[15 - 1]) {
          email = formGroup[k].values[15 - 1];
        } else {
          email = regGroup[k].values[4 - 1];
        }
        
        // 寫入 H:K 欄，且強制保留 L 欄 (UUID) 與 M 欄 (門票發送狀態)
        var writeRange = reconSheet.getRange(rowNum, 8, 1, 4);
        writeRange.getCell(1, 1).setNumberFormat("@"); // 強制 B 欄（此處為 H 欄）為文字格式
        writeRange.setValues([["'" + serial, name, email, "匯款完成"]]);
      }
    } else {
      // 筆數不符
      for (var k = 0; k < X; k++) {
        var rowNum = reconGroup[k].rowNum;
        reconSheet.getRange(rowNum, 8, 1, 3).setValues([["", "", ""]]); // 清空 H, I, J
        reconSheet.getRange(rowNum, 11).setValue("筆數不符：" + lastFive); // K
      }
    }
  }
}

/**
 * 輔助函式：取得數值的 Date 物件格式
 */
function getValAsDate(val) {
  if (val instanceof Date) return val;
  if (!val) return new Date(0);
  var parsed = Date.parse(val.toString());
  if (!isNaN(parsed)) return new Date(parsed);
  return new Date(0);
}

/**
 * 發送異常警告通知 (Email & Telegram)
 */
function sendReconciliationAlert(anomalies) {
  if (anomalies.length === 0) return;
  
  var subject = "⚠️ 【2026退虎會】自動對帳系統異常通知！";
  var body = "自動對帳系統在最近一次執行中發現以下異常，請立即前往試算表人工確認：\n\n" +
             anomalies.map(function(err, idx) { return (idx + 1) + ". " + err; }).join("\n") +
             "\n\n對帳試算表連結：https://docs.google.com/spreadsheets/d/" + SPREADSHEET_ID + "/edit";
  
  // 1. 發送 Email 通知
  var emailTo = ADMIN_NOTIFY_EMAIL || Session.getActiveUser().getEmail();
  if (emailTo) {
    try {
      MailApp.sendEmail(emailTo, subject, body);
      Logger.log("已發送異常通知信件給：" + emailTo);
    } catch (e) {
      Logger.log("發送通知信件失敗: " + e.toString());
    }
  }
  
  // 2. 發送 Telegram 訊息 (若有設定 Token)
  if (TELEGRAM_BOT_TOKEN && TELEGRAM_CHAT_ID) {
    try {
      var url = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN + "/sendMessage";
      var payload = {
        chat_id: TELEGRAM_CHAT_ID,
        text: "⚠️ *【2026退虎會】對帳異常通知*\n\n" + anomalies.join("\n") + "\n\n[點我打開試算表](https://docs.google.com/spreadsheets/d/" + SPREADSHEET_ID + "/edit)",
        parse_mode: "Markdown"
      };
      
      UrlFetchApp.fetch(url, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      Logger.log("已發送 Telegram 異常通知。");
    } catch (e) {
      Logger.log("發送 Telegram 通知失敗: " + e.toString());
    }
  }
}


/**
 * 將 HTML 格式信件轉換為乾淨的純文字以供解析
 */
function convertHtmlToPlainText(html) {
  if (!html) return "";
  
  // 先濾除整個 <style> 與 <script> 區塊 (包含內部的 CSS 樣式與 JS 代碼)
  var text = html
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<script[\s\S]*?<\/script>/gi, "");
    
  // 將常見的換行與表格標籤換成新行或空白，保持格式結構
  text = text
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/tr>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<\/td>/gi, " ");
    
  // 濾除所有 HTML 標籤
  text = text.replace(/<\/?[^>]+(>|$)/g, "");
  
  // 解碼 HTML 特殊字元
  text = text
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"');
    
  return text;
}