const SHEET_NAME = "AI協會技術需求";

// 修正：HEADERS 將 "訊息" 替換成 "上傳檔案"
const HEADERS = ["編號", "送出時間", "名字", "單位", "電子郵件", "主要需求", "上傳檔案", "需求日期", "來源頁面"];

const OWNER_EMAIL = "cjhiwen2013@gmail.com";
const PARENT_FOLDER_ID = "1OIKzDCOdm_LYvfuEsG7bOZdONfGEixI7";

/**
 * 處理 GET 請求
 */
function doGet() {
  try {
    const sheet = getBackendSheet_();
    return json_({
      ok: true,
      message: "AI協會技術需求後端已啟用",
      spreadsheetUrl: sheet.getParent().getUrl(),
      sheetName: sheet.getName()
    });
  } catch (error) {
    return json_({ ok: false, message: error.message });
  }
}

/**
 * 處理 POST 請求
 */
function doPost(e) {
  try {
    var jsonString = e.postData.contents;
    var data = JSON.parse(jsonString);
    var sheet = getBackendSheet_();
    
    // 產生序列號 (格式如: T2026061501)
    var serialNumber = generateSerialNumber_(sheet);
    
    // 1. 在雲端硬碟建立專屬資料夾 (名稱格式為: "編號+填表人姓名")
    var parentFolder = DriveApp.getFolderById(PARENT_FOLDER_ID);
    var folderName = serialNumber + "_" + data.name;
    var subFolder = parentFolder.createFolder(folderName);
    var folderUrl = subFolder.getUrl();
    
    // 2. 如果有上傳檔案，將其存入專屬資料夾
    var fileUrl = "";
    if (data.fileData && data.fileName) {
      var decoded = Utilities.base64Decode(data.fileData);
      var blob = Utilities.newBlob(decoded, data.fileType, data.fileName);
      var file = subFolder.createFile(blob);
      // 設定共用權限為知道連結的人即可檢視
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = file.getUrl();
    }
    
    // 寫入資料（首欄為序列號，上傳檔案欄位寫入檔案連結，無檔案則寫入資料夾連結）
    sheet.appendRow([
      serialNumber,     // 1. 序列號
      data.submittedAt, // 2. 送出時間
      data.name,        // 3. 名字
      data.unit,        // 4. 單位
      data.email,       // 5. 電子郵件
      data.mainDemand,  // 6. 主要需求
      fileUrl || folderUrl, // 7. 上傳檔案
      data.demandDate,  // 8. 需求日期
      data.source       // 9. 來源頁面
    ]);
    
    // 執行寄信 (帶入序列號, 檔案連結, 資料夾連結)
    notifyOwner_(data, sheet, serialNumber, folderUrl, fileUrl);
    
    return json_({ "status": "success" });
  } catch (error) {
    return json_({ "status": "error", "message": error.toString() });
  }
}

/**
 * 產生每日流水號 (例如: T2026061501, 隔日重置為 01)
 */
function generateSerialNumber_(sheet) {
  const now = new Date();
  // 取得台灣時間的 YYYYMMDD
  const dateStr = Utilities.formatDate(now, "Asia/Taipei", "yyyyMMdd");
  const prefix = "T" + dateStr;
  
  let nextSeq = 1;
  const lastRow = sheet.getLastRow();
  
  if (lastRow > 1) {
    // 讀取 A 欄（不包含標題列）的所有值
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let maxSeq = 0;
    
    for (let i = 0; i < values.length; i++) {
      const val = String(values[i][0]);
      if (val.startsWith(prefix)) {
        // 取得後面 2 位流水號
        const seqStr = val.substring(prefix.length);
        const seqNum = parseInt(seqStr, 10);
        if (!isNaN(seqNum) && seqNum > maxSeq) {
          maxSeq = seqNum;
        }
      }
    }
    nextSeq = maxSeq + 1;
  }
  
  // 補足兩位數流水號 (例如 1 -> "01")
  const seqStr = String(nextSeq).padStart(2, '0');
  return prefix + seqStr;
}

/**
 * 分開寄信給管理員與填表人
 */
function notifyOwner_(payload, sheet, serialNumber, folderUrl, fileUrl) {
  const ownerEmail = OWNER_EMAIL || Session.getEffectiveUser().getEmail();
  const submitterEmail = payload.email ? payload.email.trim() : "";
  
  // 共用欄位內容 (最上方加入序列號)
  const commonFields = [
    `編號：${serialNumber}`,
    `送出時間：${payload.submittedAt || ""}`,
    `名字：${payload.name || ""}`,
    `單位：${payload.unit || ""}`,
    `電子郵件：${payload.email || ""}`,
    `主要需求：${payload.mainDemand || ""}`,
    `雲端專屬資料夾：${folderUrl}`,
    `上傳檔案連結：${fileUrl || "無上傳檔案"}`,
    `需求日期：${payload.demandDate || ""}`,
    `來源頁面：${payload.source || ""}`
  ].join("\n");

  // 1. 寄給管理員 (主旨與內文包含單號，且含試算表連結)
  if (ownerEmail) {
    try {
      const adminSubject = `【管理員通知】TAIAA需求單 ${serialNumber}：收到新填寫`;
      const adminBody = [
        "您好：",
        "",
        "表單收到一筆新的填寫資料，內容如下：",
        "----------------------------------------",
        commonFields,
        "----------------------------------------",
        "",
        `Google 試算表連結（僅限管理員）：${sheet.getParent().getUrl()}`
      ].join("\n");
      
      GmailApp.sendEmail(ownerEmail, adminSubject, adminBody);
    } catch (err) {
      Logger.log("管理員信件發送失敗: " + err.toString());
    }
  }

  // 2. 寄給填表人 (主旨與內文包含單號，不含試算表連結)
  if (submitterEmail) {
    try {
      const userSubject = `【TAIAA 技術需求表單】您的需求單號 ${serialNumber} 確認信`;
      const userBody = [
        "您好：",
        "",
        "感謝您填寫 TAIAA 技術需求表單。以下是您的填寫內容備份，我們將盡快與您聯繫：",
        "----------------------------------------",
        commonFields,
        "----------------------------------------",
        "",
        "此信件為系統自動發送，請勿直接回覆。"
      ].join("\n");
      
      GmailApp.sendEmail(submitterEmail, userSubject, userBody);
    } catch (err) {
      Logger.log("填表人信件發送失敗: " + err.toString());
    }
  }
}

/**
 * 初始化工作表
 */
function setupSheet() {
  const sheet = getBackendSheet_();
  try {
    SpreadsheetApp.getUi().alert(`後台試算表已設定完成：${sheet.getParent().getUrl()}`);
  } catch (e) {
    Logger.log(`後台試算表已設定完成：${sheet.getParent().getUrl()}`);
  }
}

/**
 * 取得或建立工作表
 */
function getBackendSheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setBackground("#176b87")
      .setFontColor("#ffffff")
      .setFontWeight("bold");
    sheet.autoResizeColumns(1, HEADERS.length);
  }

  return sheet;
}

/**
 * 將資料封裝為 JSON
 */
function json_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}