/**
 * GAS 後端 HTTP API 入口點 (請手動貼至您的 GAS 專案中)
 * 根據前後端分離架構設計，接聽前端發送過來的 JSON Payload
 */

function doPost(e) {
  // 設定 CORS Headers 允許前端存取 (如果是 Web App 執行身份為使用者，通常 GAS 會自帶，但建議回傳 JSON)
  const headers = { "Access-Control-Allow-Origin": "*" };
  
  try {
    if (!e.postData || !e.postData.contents) {
      return createJsonResponse({ success: false, message: "無效的請求：缺少 Payload" });
    }

    // 解析前端傳來的 JSON
    const request = JSON.parse(e.postData.contents);
    const action = request.action;
    const payload = request.payload || {};

    let result = {};

    // 依據 action 導向對應的邏輯處理 (原先的 google.script.run 函數)
    switch (action) {
      case 'init':
        // 需實作：驗證身分並回傳 initialData
        const roleData = checkUserPermissions_(payload.email); 
        result = getInitialDataForUser_(roleData); 
        break;

      case 'create_case':
        // 需實作：呼叫原有 registerInspection() 等邏輯
        result = registerInspection(payload);
        break;

      case 'update_case':
        result = updateCaseDetails(payload.caseId, payload.details, payload.modifier);
        break;

      case 'upload_file':
        result = uploadInspectionFile(
          { base64Data: payload.fileBase64, fileName: payload.fileName, mimeType: "application/pdf" /* 需動態 */ }, 
          payload.caseId, 
          payload.stage, 
          payload.modifier
        );
        break;

      case 'skip_stage3':
        result = skipStage3Upload(payload.caseId, payload.reason, payload.modifier);
        break;

      case 'get_history':
        result = getFileHistory(payload.caseId);
        break;

      case 'manual_remind':
        const summary = getOverdueAuditSummary().data;
        result = sendManualReminders(summary);
        break;

      default:
        throw new Error("未知的 API action: " + action);
    }

    // 回傳成功結果
    return createJsonResponse(result);

  } catch (error) {
    // 捕獲所有例外錯誤並確保以 JSON 格式回傳給前端
    return createJsonResponse({ success: false, message: error.message });
  }
}

/**
 * 處理 CORS 預檢請求 (OPTIONS)
 */
function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT); 
    // 注意：GAS 對 doOptions 支援有限，通常部署設定為"所有人"時，GAS 基礎設施會自動處理。
}

/**
 * 輔助方法：封裝 JSON 回應格式
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// -------------------------------------------------------------
// [實作] 權限驗證邏輯與工作表自動生成機制
// -------------------------------------------------------------

/**
 * 方法一：手動執行（推薦，可立刻看到結果）
 * 1. 將 GAS 程式碼貼到您的 Google Apps Script 編輯器中後。
 * 2. 在編輯器上方的工具列，有一個下拉式選單（通常預設顯示 doGet 或 doPost）。
 * 3. 點開那個下拉選單，找到並選擇 getOrCreatePermissionsSheet_。
 * 4. 點擊旁邊的 「執行」 按鈕。
 * 5. （如果是第一次執行，可能會跳出權限審查要求，請允許授權）。
 * 6. 回到您的 Google Sheets 試算表，您就會看到系統已經瞬間建立好「使用者權限」分頁，並且把 Admin、SafetyUploader、DepartmentUploader 這三組範例帳號都自動建好了！
 * 
 * 自動建立或取得「使用者權限」工作表，並寫入預設設定 
 */
function getOrCreatePermissionsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = '使用者權限';
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    // 若該工作表不存在，自動建立新的
    sheet = ss.insertSheet(sheetName);
    
    // 設定第一列標題列
    const headers = [['信箱 (Email)', '姓名 (Name)', '角色 (Role)', '所屬部門 (Department)', '啟用狀態 (Active)']];
    sheet.getRange(1, 1, 1, headers[0].length)
         .setValues(headers)
         .setBackground('#f3f4f6')
         .setFontWeight('bold');
    
    // 設定預設的三組不同層級帳號範例，您可以修改信箱以符合貴單位需要
    const defaultData = [
      ['d602tpc@gmail.com', '管理員(範例)', 'Admin', '工安組', true],
      ['safety@example.com', '工安人員(範例)', 'SafetyUploader', '職安衛中心', true],
      ['dept@example.com', '部門人員(範例)', 'DepartmentUploader', '承攬商甲', true]
    ];
    sheet.getRange(2, 1, defaultData.length, defaultData[0].length).setValues(defaultData);
    
    // 第三欄：建立下拉選單方便管理員選擇角色
    const roleRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Admin', 'SafetyUploader', 'DepartmentUploader'], true)
      .setAllowInvalid(false).build();
    sheet.getRange(2, 3, 500, 1).setDataValidation(roleRule);
    
    // 第五欄：建立核取方塊代表是否啟用
    const activeRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    sheet.getRange(2, 5, 500, 1).setDataValidation(activeRule);
    
    // 美化工作表：凍結頂部列並自動調整欄寬
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 5);
  }
  return sheet;
}

/**
 * 實際比對信箱是否註冊於系統，以及狀態是否啟用
 */
function checkUserPermissions_(email) {
  if (!email) throw new Error("未提供電子郵件或 Token 驗證失敗。");
  
  const sheet = getOrCreatePermissionsSheet_();
  // 取得除了標題以外的所有資料 (從第 2 列開始)
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1 || 1, 5).getValues();
  
  for (let i = 0; i < data.length; i++) {
    const rowEmail = data[i][0];
    const rowName = data[i][1];
    const rowRole = data[i][2];
    const rowDept = data[i][3];
    const isActive = data[i][4]; // 核取方塊為 boolean (true/false)
    
    // 若信箱相符 (忽略大小寫)
    if (rowEmail && String(rowEmail).trim().toLowerCase() === String(email).trim().toLowerCase()) {
      if (isActive !== true) {
        throw new Error("您的帳號已被停權，請聯絡系統管理員。");
      }
      return { 
        email: email, 
        name: rowName, 
        role: rowRole, 
        department: rowDept 
      };
    }
  }
  
  throw new Error(`系統中找不到信箱 ${email} 的權限記錄，請請管理員去「使用者權限」分頁新增您的信箱並打勾啟用。`);
}

function getInitialDataForUser_(roleData) {
  const baseData = getInitialData(); // 呼叫您的現有函數 (原先取得 cases/projects)
  if (!baseData.success) throw new Error(baseData.message);
  
  // TODO: 如果需要，您也可以在此加上角色資料過濾機制
  // 例如：若 roleData.role === 'DepartmentUploader'，則過濾 baseData.records 只保留該部門案件

  return { 
    success: true, 
    data: {
      email: roleData.email,
      name: roleData.name,
      role: roleData.role,
      department: roleData.department,
      cases: baseData.records || [],
      projects: baseData.projects
    }
  };
}
