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
// [示範] 權限驗證邏輯框架 (需對照使用者權限工作表)
// -------------------------------------------------------------
function checkUserPermissions_(email) {
  // TODO: 從 Spreadsheet 讀取名單，比對 email
  // return { email: email, role: 'Admin', department: '工安組', isActive: true };
  if (!email) throw new Error("未提供電子郵件。");
  return { email: email, role: 'SafetyUploader', department: '工安部門' }; 
}

function getInitialDataForUser_(roleData) {
  const baseData = getInitialData(); // 呼叫現有函數取得資料
  if (!baseData.success) throw new Error(baseData.message);
  
  // TODO: 依照 roleData.department 過濾 cases 清單
  // 若為 DepartmentUploader，僅回傳該部門的案件

  return { 
    success: true, 
    data: {
      role: roleData.role,
      department: roleData.department,
      cases: baseData.records || [], // 需將 records Map 成前端需要的屬性格式
      projects: baseData.projects
    }
  };
}
