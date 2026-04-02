/**
 * 工安查核管理系統 - GAS 全功能重構版 (前端分離架構 API 端)
 * @version 5.0
 * @description 整合 GitHub Pages 前端架構，包含完整的 doPost(e) JSON 介接、
 *              Google ID Token 驗證機制與自動化「使用者權限」管理功能。
 */

// ==================== 全域設定 ====================
// ==================== 請修改為您的實際 ID ====================
const SPREADSHEET_ID = '14ehLXBeIum-1QVsfp0vASWwt02iWN1HmFJZQ45XhtQU';
const MAIN_DRIVE_FOLDER_ID = '1ZbOl6wqEXbhDnSpnINyoolpetY8YDyh-';

const SHEET_AUDIT_LIST = '查核列表';
const SHEET_PROJECT_DB = '下拉選單';
const SHEET_MAIL_LIST = '信件寄送列表';
const SHEET_CHANGE_LOG = '變更記錄';
const SHEET_TEMPLATES = '範例檔案';
const SHEET_SYSTEM_SETTINGS = '系統設定';
const SHEET_ANNUAL_PLAN = '年度計畫';
const SHEET_DEFICIENCY_DB = '缺失清單';
const SHEET_USER_PERMISSIONS = '使用者權限';
const SHEET_FILE_HISTORY = '檔案歷程';
const SHEET_DEPT_ACCOUNTS = '帳號管理';
const SHEET_DEPT_LIST = '部門清單';
const SHEET_UPLOAD_TOKENS = '上傳Token';

// ── 測試模式設定（優先讀 Script Properties，讀不到則用下方預設值）──
const _props = PropertiesService.getScriptProperties();
const TEST_EMAIL = _props.getProperty('TEST_EMAIL') || 'clinlion418@gmail.com';
// 系統前台連結（信件中「前往系統」按鈕）
const SYSTEM_URL = _props.getProperty('SYSTEM_URL') || 'https://d602-tech.github.io/Safety/';

const STATUS = {
  STAGE1: '第1階段-已登錄',
  STAGE2: '第2階段-改善單已上傳',
  STAGE3: '第3階段-工作隊版已處理',
  STAGE4: '第4階段-已結案'
};

// ==================== 前端網頁介面 (防呆) ====================
function doGet() {
  return ContentService.createTextOutput("此系統已升級為前後端分離架構，請透過 GitHub 前端專案頁面登入操作。");
}

function doOptions(e) {
  return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT); 
}

// ==================== API 進入點 (接聽前端 JSON) ====================
function doPost(e) {
  const headers = { "Access-Control-Allow-Origin": "*" };
  
  try {
    if (!e.postData || !e.postData.contents) {
      return createJsonResponse({ success: false, message: "無效的請求：缺少 Payload" });
    }

    const request = JSON.parse(e.postData.contents);
    const action = request.action;
    const payload = request.payload || {};
    if (action === 'get_upload_link') {
        const audit = getAuditRecords_().find(function(a) { return a.id == payload.caseId; });
        if (!audit) throw new Error("找不到案件");
        const contractorEmail = audit['承辦人Email'] || audit['承辦人電子信箱'];
        const dueDate = audit['最晚應核章日期'];
        if (!contractorEmail || !dueDate) throw new Error("案件資料不全 (需有承辦人Email與截止日)");
        const token = generateUploadToken_(payload.caseId, contractorEmail, dueDate);
        return createJsonResponse({ success: true, url: getUploadPageUrl_(token) });
    }
    
    // 【公開 API】允許未登入存取特定功能
    if (action === 'get_public_cases') {
        return createJsonResponse(getPublicCases_());
    }
    if (action === 'verify_upload_token') {
        return createJsonResponse(verifyUploadToken_(payload.token));
    }
    if (action === 'token_upload_file') {
        return createJsonResponse(tokenUploadFile_(payload));
    }
    
    // 【防護網一：驗證前端傳來的 Token】
    // 取得 Google JWT ID Token，並透過 Google API 驗證解析出真實 Email
    const idToken = payload.token;
    if (!idToken) throw new Error("缺少登入驗證憑證 (Token)，請重新登入。");
    const verifiedEmail = verifyGoogleToken_(idToken);

    // 【防護網二：對比「使用者權限」表中的權限狀態】
    const roleData = checkUserPermissions_(verifiedEmail);

    let result = {};

    // 依據 action 導向對應的邏輯處理，並將身分資料帶入防護
    switch (action) {
      case 'init':
        result = getInitialDataForUser_(roleData); 
        break;

      case 'create_case':
        if(roleData.role !== 'Admin' && roleData.role !== 'SafetyUploader') {
            throw new Error("您的權限無法登錄案件。");
        }
        payload.inspector = roleData.email; // 強制覆寫為驗證過的信箱
        payload.modifier = roleData.email;
        result = registerInspection(payload);
        break;

      case 'update_case':
        result = updateCaseDetails(payload.caseId, payload.details, roleData.email);
        break;

      case 'upload_file':
        result = uploadInspectionFile(
          { base64Data: payload.fileBase64, fileName: payload.fileName, mimeType: "application/pdf" }, 
          payload.caseId, 
          payload.stage, 
          roleData // 傳遞完整的角色資料進行驗證
        );
        break;

      case 'skip_stage3':
        result = skipStage3Upload(payload.caseId, payload.reason, roleData.email);
        break;

      case 'get_history':
        result = getFileHistory(payload.caseId);
        break;

      case 'manual_remind':
        if(roleData.role !== 'Admin') {
            throw new Error("只有 Admin 有權限設定或手動觸發稽催信件。");
        }
        const summary = getOverdueAuditSummary().data;
        result = sendManualReminders(summary);
        break;

      case 'preview_case_reminder':
        if(roleData.role !== 'Admin' && roleData.role !== 'SafetyUploader') {
            throw new Error("僅限管理員或工安人員可預覽稽催信件。");
        }
        result = previewCaseReminder_(payload.caseId);
        break;

      case 'send_case_reminder':
        if(roleData.role !== 'Admin' && roleData.role !== 'SafetyUploader') {
            throw new Error("僅限管理員或工安人員可發送稽催信件。");
        }
        result = sendCaseReminder_(payload);
        break;

      case 'get_users':
        if(roleData.role !== 'Admin') {
            throw new Error("只有 Admin 有權限查看使用者列表。");
        }
        result = getAllUsers_();
        break;

      case 'add_project':
        if(roleData.role !== 'Admin') throw new Error("無權限執行此操作。");
        result = addProject_(payload);
        break;

      case 'update_project':
        if(roleData.role === 'DepartmentUploader') throw new Error("無權限執行此操作。");
        result = updateProject_(payload.serial, payload);
        break;

      case 'delete_project':
        if(roleData.role !== 'Admin') throw new Error("無權限執行此操作。");
        result = deleteProject_(payload.serial);
        break;

      case 'get_deficiencies':
        result = getDeficiencies_(roleData);
        break;

      case 'update_deficiency':
        result = updateDeficiency_(payload, roleData.email);
        break;

      case 'delete_case':
        if(roleData.role === 'DepartmentUploader') throw new Error("無權限執行此操作。");
        result = deleteCase_(payload.caseId, payload.reason, roleData.email);
        break;

      case 'setup_system':
        if(roleData.role !== 'Admin') throw new Error("唯有管理員可進行系統初始化。");
        result = setupSystem_(payload.mode, payload.year, roleData.email);
        break;

      case 'get_system_metadata':
        if(roleData.role !== 'Admin') throw new Error("唯有管理員可調閱系統統計資訊。");
        result = getSystemMetadata_(roleData);
        break;

      case 'test_send_email':
        if(roleData.role !== 'Admin') throw new Error("只有 Admin 可發送測試信件。");
        result = sendTestEmail_();
        break;

      case 'run_daily_reminder':
        if(roleData.role !== 'Admin') throw new Error("只有 Admin 可手動觸發稽催。");
        result = runDailyReminderJob_();
        break;

      case 'setup_trigger':
        if(roleData.role !== 'Admin') throw new Error("只有 Admin 可設定觸發器。");
        result = setupDailyTrigger_();
        break;

      case 'register_dept_account':
        if(roleData.role !== 'Admin') throw new Error("只有 Admin 可管理帳號。");
        result = registerDeptAccount_(payload);
        break;

      case 'get_dept_accounts':
        if(roleData.role !== 'Admin') throw new Error("只有 Admin 可讀取帳號清單。");
        result = getDeptAccounts_();
        break;

      case 'delete_deficiency':
        if(roleData.role !== 'Admin') {
            throw new Error("稽催功能僅限管理員。");
        }
        result = manualRemindNotifications();
        break;

      case 'batch_add_deficiencies':
        result = batchAddDeficiencies_(payload, roleData.email);
        break;

      case 'save_user':
        if(roleData.role !== 'Admin') {
            throw new Error("使用者管理僅限管理員。");
        }
        result = saveUser_(payload);
        break;

      case 'get_dept_members':
        result = getDeptMembers_();
        break;

      case 'save_dept_member':
        if(roleData.role !== 'Admin') throw new Error("無權限管理部門清單。");
        result = saveDeptMember_(payload);
        break;

      case 'delete_dept_member':
        if(roleData.role !== 'Admin') throw new Error("無權限管理部門清單。");
        result = deleteDeptMember_(payload.id);
        break;

      default:
        throw new Error("未定義的 Action: " + action);
    }

    return createJsonResponse(result);

  } catch (error) {
    return createJsonResponse({ success: false, message: error.message });
  }
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


// ==================== 資安與權限核心機制 ====================

/** 利用 Google 伺服器解析並驗證 Oauth JWT Token 的真偽 */
function verifyGoogleToken_(token) {
  // 注意：高頻使用下，可改用 jwt 解析套件或 Google Cloud 身分服務
  const url = "https://oauth2.googleapis.com/tokeninfo?id_token=" + token;
  const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    throw new Error("無效的登入憑證或逾時，請重新整理頁面。");
  }
  const data = JSON.parse(response.getContentText());
  return data.email; 
}

/**
 * 方法一：手動執行（推薦，可立刻看到結果）
 * 1. 將 GAS_Backend_Merged.js 的所有程式碼貼到您的 Google Apps Script 編輯器中後。
 * 2. 在編輯器上方的工具列，有一個下拉式選單（通常預設顯示 doGet 或 doPost）。
 * 3. 點開那個下拉選單，找到並選擇 getOrCreatePermissionsSheet。
 * 4. 點擊旁邊的 「執行」 按鈕。
 * 5. （如果是第一次執行，可能會跳出權限審查要求，請允許授權）。
 * 6. 回到您的 Google Sheets 試算表，您就會看到系統已經瞬間建立好「使用者權限」分頁，並且把 Admin、SafetyUploader、DepartmentUploader 這三組範例帳號都自動建好了！
 * 
 * 自動建立或取得「使用者權限」工作表，並寫入預設設定 
 */
function getOrCreatePermissionsSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetName = '使用者權限';
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    const headers = [['信箱 (Email)', '姓名 (Name)', '角色 (Role)', '所屬部門 (Department)', '啟用狀態 (Active)']];
    sheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setBackground('#f3f4f6').setFontWeight('bold');
    
    const defaultData = [
      ['d602tpc@gmail.com', '管理員', 'Admin', '工安組', true],
      ['safety@example.com', '工安人員', 'SafetyUploader', '主管機關', true],
      ['dept@example.com', '部門人員', 'DepartmentUploader', '工務部', true]
    ];
    sheet.getRange(2, 1, defaultData.length, defaultData[0].length).setValues(defaultData);
    
    // 第三欄：建立下拉選單角色
    const roleRule = SpreadsheetApp.newDataValidation().requireValueInList(['Admin', 'SafetyUploader', 'DepartmentUploader'], true).setAllowInvalid(false).build();
    sheet.getRange(2, 3, 500, 1).setDataValidation(roleRule);
    
    // 第五欄：建立核取方塊
    const activeRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    sheet.getRange(2, 5, 500, 1).setDataValidation(activeRule);
    
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 5);
  }
  return sheet;
}

/** 比對信箱是否註冊及是否啟用 */
function checkUserPermissions_(email) {
  const sheet = getOrCreatePermissionsSheet();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1 || 1, 5).getValues();
  
  for (let i = 0; i < data.length; i++) {
    const rowEmail = data[i][0];
    const rowName = data[i][1];
    const rowRole = data[i][2];
    const rowDept = data[i][3];
    const isActive = data[i][4];
    
    if (rowEmail && String(rowEmail).trim().toLowerCase() === String(email).trim().toLowerCase()) {
      if (isActive !== true) {
        throw new Error("您的帳號已被停權，請聯絡系統管理員。");
      }
      return { email: email, name: rowName, role: rowRole, department: rowDept };
    }
  }
  throw new Error('系統中找不到信箱 ' + email + '的權限記錄，請管理員於「使用者權限」工作表中新增您的資料並打勾啟用。');
}

/** 取得所有使用者清單 (僅限 Admin) */
function getAllUsers_() {
  try {
    const sheet = getOrCreatePermissionsSheet();
    if (sheet.getLastRow() < 2) return { success: true, data: [] };
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    const users = values.map(row => ({
      email: row[0],
      name: row[1],
      role: row[2],
      department: row[3],
      active: row[4]
    }));
    return { success: true, data: users };
  } catch (e) {
    throw new Error("無法讀取使用者清單: " + e.message);
  }
}

function getInitialDataForUser_(roleData) {
  const baseData = getInitialData(); 
  if (!baseData.success) throw new Error(baseData.message);
  
  // 若為 DepartmentUploader，則只能看到自己所屬部門的案件
  let userCases = baseData.records || [];
  if (roleData.role === 'DepartmentUploader') {
    userCases = userCases.filter(c => c['主辦部門'] === roleData.department);
  }

  return { 
    success: true, 
    data: {
      email: roleData.email,
      name: roleData.name,
      role: roleData.role,
      department: roleData.department,
      cases: userCases,
      projects: baseData.projects,
      deptMembers: getDeptMembers_().data
    }
  };
}


// ==================== 主要業務邏輯 (源自 V4.2 重構版) ====================

function getInitialData() {
  try {
    return {
      success: true,
      projects: getProjectOptions_(),
      records: getAuditRecords_(),
      templates: getTemplateLinks_()
    };
  } catch (e) {
    return { success: false, message: "初始化資料載入失敗: " + e.message };
  }
}

function registerInspection(data) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_AUDIT_LIST);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // 初始化確保欄位存在
    const requiredHeaders = ['查核日期', '案件ID', '工程名稱', '工程簡稱', '承攬商', '主辦部門', '最晚應核章日期', '辦理狀態', '修改人員', '填表人', '查核領隊', '查核成員', '承辦人員姓名', '承辦人Email', '承辦課長姓名', '課長Email', 'S2員工查核檔案位置', 'S2廠商查核檔案位置', 'S3廠商及員工改善後核章檔案位置', 'S4結案檔案位置'];
    requiredHeaders.forEach(function(req) {
        if (headers.indexOf(req) === -1) {
            headers.push(req);
            sheet.getRange(1, headers.length).setValue(req);
        }
    });

    const caseId = new Date().getTime().toString() + Math.random().toString(36).substring(2, 8);
    const auditDate = new Date(data.auditDate);
    const dueDate = new Date(auditDate);
    dueDate.setDate(auditDate.getDate() + 14);

    const newRow = {};
    headers.forEach(function(header) {
      switch (header) {
        case '案件ID': newRow[header] = caseId; break;
        case '查核日期': newRow[header] = auditDate; break;
        case '工程名稱': newRow[header] = data.name; break;
        case '工程簡稱': newRow[header] = data.abbr; break;
        case '承攬商': newRow[header] = data.contractor; break;
        case '主辦部門': newRow[header] = data.department; break;
        case '最晚應核章日期': newRow[header] = dueDate; break;
        case '辦理狀態': newRow[header] = STATUS.STAGE1; break;
        case '查核人員': newRow[header] = data.inspector; break;
        case '填表人': newRow[header] = data.inspector || ""; break;
        case '查核領隊': newRow[header] = data.auditLeader; break;
        case '查核成員': newRow[header] = data.auditMembers; break;
        case '承辦人員姓名': newRow[header] = data.contractorName; break;
        case '承辦人Email': newRow[header] = data.contractorEmail; break;
        case '承辦課長姓名': newRow[header] = data.contractorManagerTitle; break;
        case '課長Email': newRow[header] = data.contractorManagerEmail; break;
        default: newRow[header] = ""; break;
      }
    });

    const rowValues = headers.map(function(header) { return newRow[header]; });
    sheet.appendRow(rowValues);
    logChange_(caseId, data.abbr, data.modifier, '建立案件', '查核日期: ' + data.auditDate, '', '');
    return { success: true, message: "案件登錄成功！", records: getAuditRecords_() };
  } catch (e) {
    throw new Error("案件登錄失敗: " + e.message);
  }
}

/**
 * 更新案件資料 (結案日期與人員等)
 */
function updateCaseDetails(caseId, details, modifier) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_AUDIT_LIST);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 初始化確保欄位存在
    const requiredHeaders = ['案件ID', '查核日期', '工程名稱', '工程簡稱', '承攬商', '主辦部門', '最晚應核章日期', '辦理狀態', '查核人員', '修改人員', '結案日期', '查核領隊', '查核成員', '承辦人姓名', '承辦人電子信箱', '承辦課長職稱', '承辦課長電子信箱'];
    requiredHeaders.forEach(function(req) {
        if (headers.indexOf(req) === -1) {
            headers.push(req);
            sheet.getRange(1, headers.length).setValue(req);
        }
    });

    const caseIdCol = headers.indexOf('案件ID');
    if (caseIdCol === -1) throw new Error('找不到案件ID欄位。');

    const caseIds = sheet.getRange(2, caseIdCol + 1, sheet.getLastRow() - 1, 1).getValues().flat();
    const rowIdx = caseIds.findIndex(function(id) { return id == caseId; }) + 2;
    if (rowIdx < 2) throw new Error('找不到對應的案件ID: ' + caseId);

    const projectAbbr = sheet.getRange(rowIdx, headers.indexOf('工程簡稱') + 1).getValue();

    // 更新指定的欄位
    if (details.inspector !== undefined) {
        var fillerCol = headers.indexOf('填表人');
        if (fillerCol === -1) fillerCol = headers.indexOf('查核人員');
        sheet.getRange(rowIdx, fillerCol + 1).setValue(details.inspector);
    }
    if (details.auditLeader !== undefined) sheet.getRange(rowIdx, headers.indexOf('查核領隊') + 1).setValue(details.auditLeader);
    if (details.auditMembers !== undefined) sheet.getRange(rowIdx, headers.indexOf('查核成員') + 1).setValue(details.auditMembers);
    
    // 承辦人資訊 (相容新舊標語)
    if (details.contractorName !== undefined) {
        var col = headers.indexOf('承辦人員姓名');
        if (col === -1) col = headers.indexOf('承辦人姓名');
        sheet.getRange(rowIdx, col + 1).setValue(details.contractorName);
    }
    if (details.contractorEmail !== undefined) {
        var col = headers.indexOf('承辦人Email');
        if (col === -1) col = headers.indexOf('承辦人電子信箱');
        sheet.getRange(rowIdx, col + 1).setValue(details.contractorEmail);
    }
    if (details.contractorManagerTitle !== undefined) {
        var col = headers.indexOf('承辦課長姓名');
        if (col === -1) col = headers.indexOf('承辦課長職稱');
        sheet.getRange(rowIdx, col + 1).setValue(details.contractorManagerTitle);
    }
    if (details.contractorManagerEmail !== undefined) {
        var col = headers.indexOf('課長Email');
        if (col === -1) col = headers.indexOf('承辦課長電子信箱');
        sheet.getRange(rowIdx, col + 1).setValue(details.contractorManagerEmail);
    }
    
    if (details.closeDate !== undefined) sheet.getRange(rowIdx, headers.indexOf('結案日期') + 1).setValue(details.closeDate ? new Date(details.closeDate) : "");

    sheet.getRange(rowIdx, headers.indexOf('修改人員') + 1).setValue(modifier);
    
    logChange_(caseId, projectAbbr, modifier, '資料更新', '修改案件人員與結案資料', '', '');
    
    return { success: true, message: "案件資料已更新", records: getAuditRecords_() };
  } catch (e) {
    throw new Error("更新失敗: " + e.message);
  }
}

function uploadInspectionFile(fileInfo, caseId, stage, roleData) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_AUDIT_LIST);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const caseIdCol = headers.indexOf('案件ID');
    
    const caseIds = sheet.getRange(2, caseIdCol + 1, sheet.getLastRow() - 1, 1).getValues().flat();
    const rowIdx = caseIds.findIndex(function(id) { return id == caseId; }) + 2;
    if (rowIdx < 2) throw new Error('找不到對應的案件ID: ' + caseId);

    const rowData = sheet.getRange(rowIdx, 1, 1, headers.length).getValues()[0];
    const auditData = {};
    headers.forEach(function(h, i) { auditData[h] = rowData[i]; });

    // 【權限驗證核心】
    const userEmail = roleData.email;
    const userRole = roleData.role;
    const userDept = roleData.department;

    if (userRole === 'DepartmentUploader') {
        if (stage !== 'stage3') {
            throw new Error("部門人員僅能上傳「工作隊核章版」(第3階段)。");
        }
        if (auditData['主辦部門'] !== userDept) {
            throw new Error("您無權上傳非所屬部門的案件資料。");
        }
    } else if (userRole !== 'Admin' && userRole !== 'SafetyUploader') {
        throw new Error("您的權限無法執行上傳操作。");
    }

    const fileExtension = fileInfo.fileName.includes('.') ? '.' + fileInfo.fileName.split('.').pop() : '.pdf';
    const newFileName = generateFileName_(stage, auditData, fileExtension);

    const auditDate = new Date(auditData['查核日期']);
    const formattedDate = Utilities.formatDate(auditDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const folderName = formattedDate + '_' + auditData['工程簡稱'];
    const rootFolder = DriveApp.getFolderById(MAIN_DRIVE_FOLDER_ID);

    var targetFolderIterator = rootFolder.getFoldersByName(folderName);
    const targetFolder = targetFolderIterator.hasNext() ? targetFolderIterator.next() : rootFolder.createFolder(folderName);

    const blob = Utilities.newBlob(Utilities.base64Decode(fileInfo.base64Data), fileInfo.mimeType, newFileName);
    const uploadedFile = targetFolder.createFile(blob);
    const fileUrl = uploadedFile.getUrl();
    const fileId = uploadedFile.getId();

    // 【自動化抓取】上傳 S2廠商 時提取缺失
    if (stage === 'stage2c') {
      try {
        extractDeficienciesFromS2_(fileId, caseId, auditData['工程簡稱'], auditData['主辦部門'], userEmail);
      } catch (e) {
        console.error("S2 缺失自動抓取失敗:", e);
      }
    }

    const currentStatus = auditData['辦理狀態'] || '';

    // 【智慧狀態更新】僅在狀態「前進」時更新
    let newStatus = '';
    const statusOrder = [STATUS.STAGE1, STATUS.STAGE2, STATUS.STAGE3, STATUS.STAGE4];
    const currentIndex = statusOrder.indexOf(currentStatus);

    if ((stage === 'stage2e' || stage === 'stage2c') && currentIndex < 1) newStatus = STATUS.STAGE2;
    else if (stage === 'stage3' && currentIndex < 2) newStatus = STATUS.STAGE3;
    else if (stage === 'stage4' || stage === 'stage4e' || stage === 'stage4c') {
        newStatus = STATUS.STAGE4;
        const closeDateCol = headers.indexOf('結案日期');
        if (closeDateCol > -1) sheet.getRange(rowIdx, closeDateCol + 1).setValue(new Date());
    }

    if (newStatus !== '') {
        sheet.getRange(rowIdx, headers.indexOf('辦理狀態') + 1).setValue(newStatus);
    }
    
    // 將連結存回主清單對應欄位
    let urlColNum = -1;
    let targetHeader = '';
    
    if (stage === 'stage2e') {
        targetHeader = 'S2員工查核檔案位置';
    } else if (stage === 'stage2c') {
        targetHeader = 'S2廠商查核檔案位置';
    } else if (stage === 'stage3') {
        targetHeader = 'S3廠商及員工改善後核章檔案位置';
    } else if (stage === 'stage4' || stage === 'stage4e' || stage === 'stage4c') {
        targetHeader = 'S4結案檔案位置';
    }
    
    urlColNum = headers.indexOf(targetHeader) + 1;
    
    if (urlColNum > 0) {
        sheet.getRange(rowIdx, urlColNum).setValue(fileUrl);
    } else if (targetHeader !== '') {
        // 若欄位不存在，則動態新增
        const lastCol = sheet.getLastColumn();
        sheet.getRange(1, lastCol + 1).setValue(targetHeader);
        sheet.getRange(rowIdx, lastCol + 1).setValue(fileUrl);
        headers.push(targetHeader);
    }

    sheet.getRange(rowIdx, headers.indexOf('修改人員') + 1).setValue(userEmail);
    
    let stageDisplay = '';
    switch(stage) {
        case 'stage2e': stageDisplay = 'S2 員工改善單'; break;
        case 'stage2c': stageDisplay = 'S2 廠商改善單'; break;
        case 'stage3': stageDisplay = 'S3 廠商及員工'; break;
        case 'stage4e': 
        case 'stage4': stageDisplay = 'S4 員工結案版'; break;
        case 'stage4c': stageDisplay = 'S4 承攬商結案版'; break;
        default: stageDisplay = '檔案上傳';
    }
    logChange_(caseId, auditData['工程簡稱'], userEmail, '檔案上傳', stageDisplay, newFileName, fileUrl);
    
    // 【即時通知】S2 上傳完成後立即寄送通知信給承辦人（不等每日排程）
    if ((stage === 'stage2e' || stage === 'stage2c') && !hasNotificationSent_(caseId, 'NOTIFY_S2')) {
      try {
        const contractorEmail = auditData['承辦人Email'] || auditData['承辦人電子信箱'];
        if (contractorEmail) {
          // 重新組 audit 物件供 email template 使用
          const auditForEmail = {};
          headers.forEach(function(h, i) { auditForEmail[h] = auditData[h]; });
          auditForEmail.id = caseId;
          // 格式化日期
          ['查核日期', '最晚應核章日期'].forEach(function(key) {
            if (auditForEmail[key] instanceof Date) {
              auditForEmail[key] = Utilities.formatDate(auditForEmail[key], Session.getScriptTimeZone(), 'yyyy-MM-dd');
            }
          });
          // 產生安全上傳連結
          let s2UploadUrl = '';
          const dueDate = auditForEmail['最晚應核章日期'];
          if (dueDate) {
            try {
              const token = generateUploadToken_(caseId, contractorEmail, dueDate);
              s2UploadUrl = getUploadPageUrl_(token);
            } catch (te) { console.error('[Token] S2即時通知Token產生失敗:', te.message); }
          }

          GmailApp.sendEmail(
            contractorEmail,
            `第1階段，查核報告電子檔已上傳可先行轉知廠商修改，詳如說明，請查照`,
            '',
            { htmlBody: buildStage1EmailHtml_(auditForEmail, s2UploadUrl) }
          );
          markNotificationSent_(caseId, auditData['工程簡稱'], 'NOTIFY_S2');
          console.log(`📨 [S2上傳即時通知] ${auditData['工程簡稱']} → ${contractorEmail}`);
        } else {
          console.log(`⚠️ [S2上傳] ${auditData['工程簡稱']} - 承辦人Email為空，跳過即時通知`);
        }
      } catch (mailErr) {
        console.error(`❌ [S2即時通知失敗] ${auditData['工程簡稱']}: ${mailErr.message}`);
      }
    }

    return { success: true, message: '檔案 "' + newFileName + '" 上傳成功！', records: getAuditRecords_() };
  } catch (e) {
    throw new Error("檔案上傳失敗: " + e.message);
  }
}

function skipStage3Upload(caseId, reason, modifier) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_AUDIT_LIST);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const caseIdCol = headers.indexOf('案件ID');

    const caseIds = sheet.getRange(2, caseIdCol + 1, sheet.getLastRow() - 1, 1).getValues().flat();
    const rowIdx = caseIds.findIndex(function(id) { return id == caseId; }) + 2;

    const projectAbbr = sheet.getRange(rowIdx, headers.indexOf('工程簡稱') + 1).getValue();

    sheet.getRange(rowIdx, headers.indexOf('辦理狀態') + 1).setValue(STATUS.STAGE3);
    sheet.getRange(rowIdx, headers.indexOf('修改人員') + 1).setValue(modifier);

    logChange_(caseId, projectAbbr, modifier, '狀態變更', '跳過Stage3上傳', reason, '');
    return { success: true, message: "已記錄理由，案件進入下一階段。", records: getAuditRecords_() };
  } catch (e) {
    throw new Error("操作失敗: " + e.message);
  }
}

function getFileHistory(caseId) {
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_CHANGE_LOG);
        if (sheet.getLastRow() <= 1) return { success: true, data: [] };

        const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
        const headers = ["修改日期", "案件ID", "工程簡稱", "修改人員", "狀態", "說明", "檔案名稱", "檔案位置"];
        
        const history = data.map(function(row) {
            var entry = {};
            headers.forEach(function(header, i) { entry[header] = row[i]; });
            return entry;
        })
        .filter(function(entry) {
            return entry["案件ID"] == caseId && entry["狀態"] === "檔案上傳" && entry["檔案位置"];
        })
        .map(function(entry) {
            return {
                timestamp: Utilities.formatDate(new Date(entry["修改日期"]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"),
                modifier: entry["修改人員"],
                description: entry["說明"],
                fileName: entry["檔案名稱"],
                fileUrl: entry["檔案位置"]
            };
        })
        .sort(function(a, b) { return new Date(b.timestamp) - new Date(a.timestamp); });

        return { success: true, data: history };
    } catch (e) {
        throw new Error("獲取檔案歷史失敗: " + e.message);
    }
}

// ==================== 郵件系統 ====================

/** 刪除案件邏輯：刪除該列並紀錄 */
function deleteCase_(caseId, reason, modifierName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_AUDIT_LIST);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIndex = headers.indexOf('案件ID');
    const abbrIndex = headers.indexOf('工程簡稱');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] == caseId) {
        // Record deletion history
        const projectAbbr = abbrIndex !== -1 ? data[i][abbrIndex] : '未知';
        const reasonStr = reason ? `(理由: ${reason})` : '';
        
        // 使用統一紀錄函式
        logChange_(caseId, projectAbbr, modifierName, "案件刪除", "已徹底刪除此案件 " + reasonStr, "", "");
        
        // 額外記錄到檔案歷程 (相容舊設計需求)
        const historySheet = getOrCreateFileHistorySheet_();
        historySheet.appendRow([
          new Date(),
          modifierName,
          caseId,
          "案件刪除 " + reasonStr,
          "系統紀錄",
          "",
          "已刪除"
        ]);
        
        // Delete the row
        sheet.deleteRow(i + 1);
        
        return { success: true, message: "案件已刪除", records: getAuditRecords_() };
      }
    }
    return { success: false, message: "找不到指定的案件: " + caseId };
  } catch(e) {
    throw new Error("刪除案件失敗: " + e.message);
  }
}

/** 
 * 系統初始化：自動生成所有必要的分頁與標題列
 * 建議第一次使用或發現分頁遺失時執行一次
 */
/**
 * 取得系統統計與備份資訊 (供管理員初始化確認視窗使用)
 */
function getSystemMetadata_(roleData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const auditSheet = ss.getSheetByName(SHEET_AUDIT_LIST);
    const defSheet = ss.getSheetByName(SHEET_DEFICIENCY_DB);
    
    const caseCount = auditSheet ? auditSheet.getLastRow() - 1 : 0;
    const defCount = defSheet ? defSheet.getLastRow() - 1 : 0;
    
    // 取得最後備份時間 (搜尋 Drive 分頁中 Backup_ 開頭的資料夾內容)
    let lastBackup = "無紀錄";
    const rootFolder = DriveApp.getFolderById(MAIN_DRIVE_FOLDER_ID);
    const backups = rootFolder.getFoldersByName("System_Backups");
    if (backups.hasNext()) {
      const backupFolder = backups.next();
      const files = backupFolder.getFiles();
      let latestTime = 0;
      while (files.hasNext()) {
        const f = files.next();
        const t = f.getDateCreated().getTime();
        if (t > latestTime) {
          latestTime = t;
          lastBackup = Utilities.formatDate(f.getDateCreated(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        }
      }
    }

    return {
      success: true,
      data: {
        baselineYear: "115年度",
        operator: roleData.email,
        caseCount: Math.max(0, caseCount),
        deficiencyCount: Math.max(0, defCount),
        lastBackup: lastBackup
      }
    };
  } catch (e) {
    return { success: false, message: "取得系統資訊失敗: " + e.message };
  }
}

/**
 * 系統初始化 (支援同步欄位與全系統重置模式)
 */
function setupSystem_(mode, year, operator) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  if (mode === 'reset') {
    // 執行資料重置模式
    try {
      // 1. 先執行備份
      const backupResult = backupSystemData_(ss, operator);
      if (!backupResult.success) throw new Error("備份失敗，中止重置: " + backupResult.message);
      
      // 2. 清除資料 (保留標題)
      const sheetsToClear = [SHEET_AUDIT_LIST, SHEET_DEFICIENCY_DB, SHEET_CHANGE_LOG, SHEET_FILE_HISTORY];
      sheetsToClear.forEach(name => {
        const sheet = ss.getSheetByName(name);
        if (sheet && sheet.getLastRow() > 1) {
          sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
        }
      });
      
      return { success: true, message: "系統資料已成功備份並重置 (模式：資料清空)。" };
    } catch (e) {
      return { success: false, message: "資料重置失敗: " + e.message };
    }
  }

  // 模式 A: 模式：欄位檢查 / 同步 (原邏輯)
  const sheetsToCreate = [
    { name: SHEET_AUDIT_LIST, headers: ['查核日期', '結案日期', '工程名稱', '工程簡稱', '承攬商', '主辦部門', '最晚應核章日期', '案件ID', '辦理狀態', '工作天數計算', '分數', '填表人', '查核領隊', '查核成員', '承辦人員姓名', '承辦人Email', '承辦課長姓名', '課長Email', 'S2員工查核檔案位置', 'S2廠商查核檔案位置', 'S3廠商及員工改善後核章檔案位置', 'S4結案檔案位置', '修改人員', '系統初始化備註'], color: '#f3f4f6' },
    { name: SHEET_PROJECT_DB, headers: ['流水號', '工程簡稱', '工程名稱', '承攬商', '主辦部門', '承辦人姓名', '承辦人電子信箱', '承辦課長職稱', '承辦課長電子信箱'], color: '#f3f4f6' },
    { name: SHEET_DEFICIENCY_DB, headers: ['缺失ID', '案件ID', '工程簡稱', '缺失內容', '主辦部門', '改善期限', '狀態', '錄入者'], color: '#fef3c7' },
    { name: SHEET_CHANGE_LOG, headers: ['修改日期', '案件ID', '工程簡稱', '修改人員', '狀態', '說明', '檔案名稱', '檔案位置'], color: '#eff6ff' },
    { name: SHEET_FILE_HISTORY, headers: ['異動日期', '異動人員', '案件ID', '異動內容', '檔案類型', '檔案連結', '狀態'], color: '#d1fae5' },
    { name: SHEET_USER_PERMISSIONS, headers: ['信箱 (Email)', '姓名 (Name)', '角色 (Role)', '所屬部門 (Department)', '啟用狀態 (Active)'], color: '#f3f4f6' },
    { name: SHEET_DEPT_ACCOUNTS, headers: ['部門名稱', '承辦人姓名', '承辦人Email', '課長姓名', '課長Email', '備註', '建立日期'], color: '#e0f2fe' },
    { name: SHEET_DEPT_LIST, headers: ['ID', '主辦部門', '職稱', '姓名', '信箱'], color: '#fdf2f8' },
    { name: SHEET_UPLOAD_TOKENS, headers: ['Token', '案件ID', '授權Email', '建立時間', '到期時間', '狀態'], color: '#e0e7ff' }
  ];

  sheetsToCreate.forEach(cfg => {
    let sheet = ss.getSheetByName(cfg.name);
    if (!sheet) {
      sheet = ss.insertSheet(cfg.name);
    }
    // 檢查欄位是否完整
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, cfg.headers.length).setValues([cfg.headers]).setBackground(cfg.color).setFontWeight('bold');
      sheet.setFrozenRows(1);
    } else {
        // 更新標題列 (僅在欄位數不符時嘗試更新，或擴增)
        const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        cfg.headers.forEach((h, idx) => {
            if (currentHeaders.indexOf(h) === -1) {
                const nextCol = sheet.getLastColumn() + 1;
                sheet.getRange(1, nextCol).setValue(h).setBackground(cfg.color).setFontWeight('bold');
            }
        });
    }
  });

  return { success: true, message: "系統標題列與欄位同步檢視完成。" };
}

/**
 * 內部備份邏輯：建立當前試算表的複本存入 System_Backups 資料夾
 */
function backupSystemData_(ss, operator) {
  try {
    const rootFolder = DriveApp.getFolderById(MAIN_DRIVE_FOLDER_ID);
    let backupFolderIterator = rootFolder.getFoldersByName("System_Backups");
    let backupFolder = backupFolderIterator.hasNext() ? backupFolderIterator.next() : rootFolder.createFolder("System_Backups");
    
    const timeStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmm");
    const backupName = "AutoBackup_" + timeStamp + "_" + operator.split('@')[0];
    DriveApp.getFileById(SPREADSHEET_ID).makeCopy(backupName, backupFolder);
    
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getOrCreateFileHistorySheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_FILE_HISTORY);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_FILE_HISTORY);
    const headers = [['異動日期', '異動人員', '案件ID', '異動內容', '檔案類型', '檔案連結', '狀態']];
    sheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setBackground('#d1fae5').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOverdueAuditSummary() {
    try {
        const overdueData = getOverdueData_();
        if (Object.keys(overdueData).length === 0) return { success: true, data: [] };
        const summary = Object.keys(overdueData).map(function(dept) {
            return {
                department: dept,
                recipients: overdueData[dept].recipients,
                cases: overdueData[dept].cases
            };
        });
        return { success: true, data: summary };
    } catch (e) {
        throw new Error("整理稽催摘要時發生錯誤: " + e.message);
    }
}

function sendManualReminders(summaryData) {
    try {
        const mailTemplate = getTemplateLinks_().mailTemplate;
        if (!mailTemplate) throw new Error("找不到郵件範本。");
        if (!summaryData || summaryData.length === 0) {
            return { success: true, message: "無逾期案件資料，未發送任何郵件。" };
        }

        var sentCount = 0;
        summaryData.forEach(function(deptInfo) {
            if (deptInfo.recipients && deptInfo.recipients.length > 0) {
                const subject = '【工安查核逾期稽催】' + deptInfo.department + '尚有 ' + deptInfo.cases.length + ' 件逾期案件';
                var tableRows = '';
                deptInfo.cases.forEach(function(c) {
                    tableRows += '<tr><td>' + c['工程簡稱'] + '</td><td>' + c['承攬商'] + '</td><td>' + c['查核日期'] + '</td><td style="color:red;">' + c['最晚應核章日期'] + '</td><td>' + c['辦理狀態'] + '</td></tr>';
                });
                var body = mailTemplate.replace('{{部門名稱}}', deptInfo.department).replace('{{案件列表}}', tableRows);
                const attachments = getAttachmentsForCases_(deptInfo.cases);

                GmailApp.sendEmail(deptInfo.recipients.join(','), subject, '', {
                    htmlBody: body,
                    attachments: attachments
                });
                sentCount++;
            }
        });
        return { success: true, message: "已成功發送 " + sentCount + " 封稽催郵件！" };
    } catch (e) {
        throw new Error("發送稽催郵件時發生錯誤: " + e.message);
    }
}

function checkAndSendNotifications() {
  const mailTemplate = getTemplateLinks_().mailTemplate;
  if (!mailTemplate) return console.error("找不到郵件範本，中止通知。");

  const reminderCases = getReminderCases_(); 
  if (Object.keys(reminderCases).length === 0) return;

  for (const dept in reminderCases) {
    const deptInfo = reminderCases[dept];
    if (deptInfo.recipients.length > 0) {
      const subject = '【工安查核案件提醒】' + dept + '有 ' + deptInfo.cases.length + ' 件待辦案件';
      var tableRows = '';
      deptInfo.cases.forEach(function (a) {
        tableRows += '<tr><td>' + a['工程簡稱'] + '</td><td>' + a['承攬商'] + '</td><td>' + a['查核日期'] + '</td><td style="color:red;">' + a['最晚應核章日期'] + '</td><td>' + a['辦理狀態'] + '</td></tr>';
      });

      var body = mailTemplate.replace('{{部門名稱}}', dept).replace('{{案件列表}}', tableRows);
      const attachments = getAttachmentsForCases_(deptInfo.cases);

      GmailApp.sendEmail(deptInfo.recipients.join(','), subject, '', { htmlBody: body, attachments: attachments });
    }
  }
}

// ==================== 輔助函數 (私有) ====================

function getProjectOptions_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_PROJECT_DB);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const reqHeaders = ['流水號', '工程簡稱', '工程名稱', '承攬商', '主辦部門', '承辦人姓名', '承辦人電子信箱', '承辦課長職稱', '承辦課長電子信箱'];
  
  if (headers.length < 9) {
    reqHeaders.forEach(req => {
      if (headers.indexOf(req) === -1) {
        headers.push(req);
        sheet.getRange(1, headers.length).setValue(req);
      }
    });
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  return data.map(function (row) {
    return { 
      serial: row[0], abbr: row[1], name: row[2], contractor: row[3], department: row[4],
      contractorName: row[5] || '', contractorEmail: row[6] || '', contractorManagerTitle: row[7] || '', contractorManagerEmail: row[8] || ''
    };
  }).filter(function (p) { return p.abbr; });
}

function addProject_(p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_PROJECT_DB);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PROJECT_DB);
    sheet.appendRow(['流水號', '工程簡稱', '工程名稱', '承攬商', '主辦部門', '承辦人姓名', '承辦人電子信箱', '承辦課長職稱', '承辦課長電子信箱']);
  }
  const serial = 'P' + new Date().getTime();
  sheet.appendRow([
    serial, p.abbr, p.name, p.contractor, p.department, 
    p.contractorName || '', p.contractorEmail || '', p.contractorManagerTitle || '', p.contractorManagerEmail || ''
  ]);
  return { success: true, projects: getProjectOptions_() };
}

function updateProject_(serial, updatedData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_PROJECT_DB);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == serial) {
      sheet.getRange(i + 1, 2, 1, 8).setValues([[
        updatedData.abbr, 
        updatedData.name, 
        updatedData.contractor, 
        updatedData.department,
        updatedData.contractorName || '',
        updatedData.contractorEmail || '',
        updatedData.contractorManagerTitle || '',
        updatedData.contractorManagerEmail || ''
      ]]);
      return { success: true, projects: getProjectOptions_() };
    }
  }
  return { success: false, message: "找不到該工程項目" };
}

function deleteProject_(serial) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_PROJECT_DB);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == serial) {
      sheet.deleteRow(i + 1);
      return { success: true, projects: getProjectOptions_() };
    }
  }
  return { success: false, message: "找不到該工程項目" };
}

function getOrCreateDeficiencySheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_DEFICIENCY_DB);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DEFICIENCY_DB);
    const headers = [['缺失ID', '案件ID', '工程簡稱', '缺失內容', '主辦部門', '改善期限', '狀態', '錄入者']];
    sheet.getRange(1, 1, 1, headers[0].length).setValues(headers).setBackground('#fef3c7').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getDeficiencies_(roleData) {
  const sheet = getOrCreateDeficiencySheet_();
  if (sheet.getLastRow() < 2) return { success: true, data: [] };
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  let list = data.map(row => ({
    id: row[0], caseId: row[1], abbr: row[2], content: row[3], department: row[4],
    deadline: row[5] instanceof Date ? Utilities.formatDate(row[5], Session.getScriptTimeZone(), 'yyyy-MM-dd') : row[5],
    status: row[6], creator: row[7]
  }));
  if (roleData.role === 'DepartmentUploader') {
    list = list.filter(d => d.department === roleData.department);
  }
  return { success: true, data: list };
}

function updateDeficiency_(p, email) {
  const sheet = getOrCreateDeficiencySheet_();
  const data = sheet.getDataRange().getValues();
  if (p.id) {
    // Update existing
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == p.id) {
        sheet.getRange(i + 1, 4, 1, 4).setValues([[p.content, p.department, p.deadline, p.status]]);
        return { success: true };
      }
    }
  } else {
    // Create new
    const newId = 'DEF' + new Date().getTime();
    sheet.appendRow([newId, p.caseId, p.abbr, p.content, p.department, p.deadline, '待改善', email]);
  }
  return { success: true };
}

function batchAddDeficiencies_(payload, email) {
  try {
    const sheet = getOrCreateDeficiencySheet_();
    const { items } = payload;
    if (!items || !items.length) throw new Error("無缺失明細");

    const now = new Date().getTime();
    const rows = items.map((item, idx) => {
      const newId = 'DEF' + (now + idx);
      return [newId, item.caseId, item.abbr, item.content, item.department, item.deadline, '待改善', email];
    });

    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return { success: true, message: `已成功匯入 ${rows.length} 項缺失` };
  } catch (e) {
    throw new Error("批次新增缺失失敗: " + e.message);
  }
}

function deleteDeficiency_(id) {
  const sheet = getOrCreateDeficiencySheet_();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
  const rowIdx = data.indexOf(id);
  if (rowIdx > -1) {
    sheet.deleteRow(rowIdx + 2);
    return { success: true };
  }
  throw new Error("找不到該缺失紀錄: " + id);
}

function saveUser_(payload) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_USER_PERMISSIONS);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.findIndex(h => h.includes('Email'));
    const nameCol = headers.findIndex(h => h.includes('姓名'));
    const roleCol = headers.findIndex(h => h.includes('角色') || h.includes('Role'));
    const deptCol = headers.findIndex(h => h.includes('部門'));
    const activeCol = headers.findIndex(h => h.includes('狀態') || h.includes('Active'));

    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
        if (data[i][emailCol] === payload.email) {
            rowIdx = i + 1;
            break;
        }
    }

    if (rowIdx > -1) {
        // Update
        sheet.getRange(rowIdx, nameCol + 1).setValue(payload.name);
        sheet.getRange(rowIdx, roleCol + 1).setValue(payload.role);
        sheet.getRange(rowIdx, deptCol + 1).setValue(payload.department);
        sheet.getRange(rowIdx, activeCol + 1).setValue(payload.active !== undefined ? payload.active : true);
    } else {
        // Add new
        const newRow = [];
        headers.forEach((h, i) => {
            if (i === emailCol) newRow.push(payload.email);
            else if (i === nameCol) newRow.push(payload.name);
            else if (i === roleCol) newRow.push(payload.role);
            else if (i === deptCol) newRow.push(payload.department);
            else if (i === activeCol) newRow.push(true);
            else newRow.push('');
        });
        sheet.appendRow(newRow);
    }
    return { success: true, message: "使用者資料已儲存" };
  } catch (e) {
    throw new Error("儲存使用者失敗: " + e.message);
  }
}

/** 取得公開案件 (僅限未結案，且不含敏感資料) */
function getPublicCases_() {
  try {
    const allRecords = getAuditRecords_();
    const incompleteCases = allRecords.filter(function(r) { return r['辦理狀態'] !== STATUS.STAGE4; });
    return {
      success: true,
      data: {
        cases: incompleteCases,
        projects: getProjectOptions_() // 下拉選單供搜尋過濾使用
      }
    };
  } catch (e) {
    return { success: false, message: "公用資料讀取失敗" };
  }
}

function getAuditRecords_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_AUDIT_LIST);
  if (sheet.getLastRow() <= 1) return [];

  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // 自動補齊欄位，保障所有的存取都能有這些標題
  const requiredHeaders = ['查核日期', '結案日期', '工程名稱', '工程簡稱', '承攬商', '主辦部門', '最晚應核章日期', '案件ID', '辦理狀態', '工作天數計算', '分數', '填表人', '查核領隊', '查核成員', '承辦人員姓名', '承辦人Email', '承辦課長姓名', '課長Email', 'S2員工查核檔案位置', 'S2廠商查核檔案位置', 'S3廠商及員工改善後核章檔案位置', 'S4結案檔案位置', '修改人員', '系統初始化備註'];
  let headersAppended = false;
  requiredHeaders.forEach(function(req) {
      if (headers.indexOf(req) === -1) {
          headers.push(req);
          sheet.getRange(1, headers.length).setValue(req);
          headersAppended = true;
      }
  });

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  const caseIdCol = headers.indexOf('案件ID');

  return values.map(function(row, index) {
    const record = {};
    headers.forEach(function(header, i) {
      var cellValue = row[i];
      if (header.includes('日期') && cellValue instanceof Date) {
        record[header] = Utilities.formatDate(cellValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        record[header] = cellValue;
      }
    });
    record.id = row[caseIdCol] ? row[caseIdCol].toString() : 'row_' + (index + 2); 
    return record;
  }).filter(function (r) { return r['查核日期']; });
}

function getTemplateLinks_() {
    try {
        const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_TEMPLATES);
        if (sheet.getLastRow() < 2) return {};
        const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
        const templates = {};
        data.forEach(function(row) {
            const name = row[0] ? row[0].toString().trim() : '';
            const value = row[1];
            if (name === "信件範例") templates.mailTemplate = value;
            if (name === "工安查核表單") templates.auditForm = value;
        });
        return templates;
    } catch (e) { return {}; }
}

function getMailList_() {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_MAIL_LIST);
    if (sheet.getLastRow() <= 1) return {};
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    const mailList = {};
    data.forEach(function(row) {
        const dept = row[0], email = row[2];
        if (dept && email) {
            if (!mailList[dept]) mailList[dept] = [];
            mailList[dept].push(email);
        }
    });
    return mailList;
}

function logChange_(caseId, projectAbbr, modifier, status, description, fileName, fileUrl) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_CHANGE_LOG);
    if (sheet.getRange("A1").getValue() !== '修改日期') {
        sheet.getRange("A1:H1").setValues([['修改日期', '案件ID', '工程簡稱', '修改人員', '狀態', '說明', '檔案名稱', '檔案位置']]);
    }
    sheet.appendRow([new Date(), caseId, projectAbbr, modifier, status, description, fileName || '', fileUrl || '']);
  } catch (e) {}
}

function generateFileName_(stage, auditData, extension) {
  const auditDate = Utilities.formatDate(new Date(auditData['查核日期']), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const projectAbbr = auditData['工程簡稱'];
  const contractor = auditData['承攬商'];
  var prefix = '';
  switch (stage) {
    case 'stage2e': prefix = '【S2_員工改善單】'; break;
    case 'stage2c': prefix = '【S2_廠商改善單】'; break;
    case 'stage3': prefix = '【S3_廠商核章】'; break;
    case 'stage4': 
    case 'stage4e': prefix = '【S4_結案-員工】'; break;
    case 'stage4c': prefix = '【S4_結案-承攬商】'; break;
    case 'report': prefix = '【完成報告】'; break;
  }
  return prefix + auditDate + '_' + projectAbbr + '_工安查核改善單(' + contractor + ')' + extension;
}

function getOverdueData_() {
    const allAudits = getAuditRecords_();
    const mailList = getMailList_();
    const timeZone = Session.getScriptTimeZone();
    const todayStr = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");

    const overdueAudits = allAudits.filter(function(audit) {
        const status = audit['辦理狀態'];
        const dueDate = audit['最晚應核章日期'];
        return status !== STATUS.STAGE4 && dueDate && todayStr > dueDate;
    });

    if (overdueAudits.length === 0) return {};
    const auditsByDept = {};
    overdueAudits.forEach(function(audit) {
        const dept = audit['主辦部門'];
        if (!auditsByDept[dept]) auditsByDept[dept] = { cases: [], recipients: [] };
        auditsByDept[dept].cases.push(audit);
    });

    Object.keys(auditsByDept).forEach(function(dept) {
        let caseRecipients = [];
        auditsByDept[dept].cases.forEach(function(c) {
             if (c['承辦人Email']) caseRecipients.push(c['承辦人Email']);
             if (c['承辦人電子信箱']) caseRecipients.push(c['承辦人電子信箱']);
             if (c['課長Email']) caseRecipients.push(c['課長Email']);
             if (c['承辦課長電子信箱']) caseRecipients.push(c['承辦課長電子信箱']);
        });
        const deptRecipients = mailList[dept] || [];
        const safetyTeamRecipients = mailList['工安組'] || [];
        auditsByDept[dept].recipients = [...new Set([...deptRecipients, ...safetyTeamRecipients, ...caseRecipients])].filter(Boolean);
    });

    return auditsByDept;
}

function getReminderCases_() {
    const allAudits = getAuditRecords_();
    const mailList = getMailList_();
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const casesToRemind = allAudits.filter(function(audit) {
        const status = audit['辦理狀態'];
        if (status === STATUS.STAGE4) return false;
        const dueDate = new Date(audit['最晚應核章日期']);
        dueDate.setHours(0, 0, 0, 0);
        const timeDiff = dueDate.getTime() - today.getTime();
        const daysRemaining = Math.round(timeDiff / (1000 * 3600 * 24));
        const isOverdue = daysRemaining < 0; 
        const isInWarningWindow = daysRemaining >= 0 && daysRemaining <= 5; 
        const isReminderDay = (5 - daysRemaining) % 2 === 0;
        return isOverdue || (isInWarningWindow && isReminderDay);
    });

    if (casesToRemind.length === 0) return {};
    const auditsByDept = {};
    casesToRemind.forEach(function(audit) {
        const dept = audit['主辦部門'];
        if (!auditsByDept[dept]) auditsByDept[dept] = { cases: [], recipients: [] };
        auditsByDept[dept].cases.push(audit);
    });
    for (const dept in auditsByDept) {
        let caseRecipients = [];
        auditsByDept[dept].cases.forEach(function(c) {
             if (c['承辦人Email']) caseRecipients.push(c['承辦人Email']);
             if (c['承辦人電子信箱']) caseRecipients.push(c['承辦人電子信箱']);
             if (c['課長Email']) caseRecipients.push(c['課長Email']);
             if (c['承辦課長電子信箱']) caseRecipients.push(c['承辦課長電子信箱']);
        });
        const recipients = mailList[dept] || [];
        const safetyTeamRecipients = mailList['工安組'] || [];
        auditsByDept[dept].recipients = [...new Set([...recipients, ...safetyTeamRecipients, ...caseRecipients])].filter(Boolean);
    }
    return auditsByDept;
}

function getAttachmentsForCases_(cases) {
    const logSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_CHANGE_LOG);
    if (logSheet.getLastRow() < 2) return [];
    const logData = logSheet.getDataRange().getValues();
    const headers = logData[0];
    const caseIdCol = headers.indexOf('案件ID');
    const descCol = headers.indexOf('說明');
    const urlCol = headers.indexOf('檔案位置');
    
    const attachments = [];
    cases.forEach(function(caseInfo) {
        var latestUrl = '';
        for (var i = logData.length - 1; i > 0; i--) {
            const logRow = logData[i];
            if (logRow[caseIdCol] == caseInfo.id && logRow[descCol].includes('階段2檔案')) {
                latestUrl = logRow[urlCol];
                break;
            }
        }
        if (latestUrl) {
            try {
                const fileIdMatch = latestUrl.match(/\/d\/(.+?)\//);
                if (fileIdMatch && fileIdMatch[1]) {
                    const file = DriveApp.getFileById(fileIdMatch[1]);
                    attachments.push(file.getBlob());
                }
            } catch (e) {}
        }
    });
    return attachments;
}


// ==================== 測試模式 ====================

/**
 * 【手動執行】發送測試信至 TEST_EMAIL，忽略真實收件人
 * 可在 GAS 編輯器中直接選擇此函數並點「執行」測試郵件格式
 */
function sendTestEmail_() {
  try {
    const fakeCase = {
      id: 'TEST001',
      '工程簡稱': '測試工程',
      '工程名稱': '範例測試工程名稱',
      '承攬商': '測試承攬商',
      '主辦部門': '測試部門',
      '最晚應核章日期': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      '查核日期': Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      '辦理狀態': STATUS.STAGE2,
      '承辦人姓名': '張測試',
      '承辦人電子信箱': TEST_EMAIL,
      '承辦課長職稱': '課長',
      '承辦課長電子信箱': TEST_EMAIL
    };

    // 測試三種模板
    GmailApp.sendEmail(TEST_EMAIL,
      '【測試信】第1階段 - S2已上傳即時通知',
      '',
      { htmlBody: buildStage1EmailHtml_(fakeCase) }
    );

    GmailApp.sendEmail(TEST_EMAIL,
      '【測試信】第2階段 - 到期前3日提醒',
      '',
      { htmlBody: buildStage2EmailHtml_(fakeCase, 3) }
    );

    GmailApp.sendEmail(TEST_EMAIL,
      '【測試信】第3階段 - 最後1日催辦',
      '',
      { htmlBody: buildStage3EmailHtml_(fakeCase, 1) }
    );

    return { success: true, message: `已成功發送 3 封測試信至 ${TEST_EMAIL}` };
  } catch (e) {
    throw new Error('測試信發送失敗: ' + e.message);
  }
}

// 公開包裝，方便 GAS 選單直接執行
function sendTestEmail() { sendTestEmail_(); }


// ==================== 三階段自動稽催 ====================

/**
 * 主控函數：每日定時觸發器呼叫此函數
 * 通知邏輯:
 *   Stage1: S2 已上傳 且 尚未發送過 Stage1 通知 → 寄承辦人（只寄一次）
 *   Stage2: S3 尚未上傳 且 距到期 3~2 天 → 每日寄承辦人
 *   Stage3: S3 尚未上傳 且 距到期 1~0 天 → 每日寄承辦人 + 課長
 *   逾期:  S3 尚未上傳 且 已超過截止日 → 每日寄承辦人 + 課長
 */
function runDailyReminderJob_() {
  const allAudits = getAuditRecords_();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  console.log('══════════════════════════════════════════════');
  console.log(`🔔 [稽催排程啟動] 執行時間: ${todayStr}`);
  console.log(`📊 全部案件數: ${allAudits.length}`);
  console.log('══════════════════════════════════════════════');

  const results = { stage1: 0, stage2: 0, stage3: 0, overdue: 0, skipped: 0, errors: [] };
  const sentLog = []; // 記錄所有發送的信件摘要

  allAudits.forEach(function(audit) {
    try {
      const status = audit['辦理狀態'];
      const abbr = audit['工程簡稱'] || '未知工程';

      if (status === STATUS.STAGE4) {
        results.skipped++;
        return; // 已結案略過
      }

      const dueDateStr = audit['最晚應核章日期'];
      if (!dueDateStr) {
        console.log(`⏭️ [略過] ${abbr} (ID:${audit.id}) - 無設定最晚應核章日期`);
        results.skipped++;
        return;
      }
      const dueDate = new Date(dueDateStr);
      dueDate.setHours(0, 0, 0, 0);
      const daysLeft = Math.round((dueDate - today) / 86400000);

      const s2Uploaded = (status === STATUS.STAGE2 || status === STATUS.STAGE3);
      const s3NotUploaded = (status === STATUS.STAGE2 || status === STATUS.STAGE1);

      const contractorEmail = audit['承辦人Email'] || audit['承辦人電子信箱'];
      const managerEmail = audit['課長Email'] || audit['承辦課長電子信箱'];

      console.log(`── 檢查案件: ${abbr} | 狀態: ${status} | 到期日: ${dueDateStr} | 剩餘天數: ${daysLeft} | 承辦人: ${contractorEmail || '(空)'} | 課長: ${managerEmail || '(空)'}`);

      // 產生安全上傳連結（S3 未上傳且有承辦人 Email 時）
      let uploadUrl = '';
      if (s3NotUploaded && contractorEmail && dueDateStr) {
        try {
          const token = generateUploadToken_(audit.id, contractorEmail, dueDateStr);
          uploadUrl = getUploadPageUrl_(token);
        } catch (tokenErr) {
          console.error(`⚠️ [Token產生失敗] ${abbr}: ${tokenErr.message}`);
        }
      }

      // ─ Stage 1：S2 已上傳，立即通知（只送一次，用 Change Log 防重複）
      if (status === STATUS.STAGE2 && !hasNotificationSent_(audit.id, 'NOTIFY_S2')) {
        if (contractorEmail) {
          GmailApp.sendEmail(
            contractorEmail,
            `第1階段，查核報告電子檔已上傳可先行轉知廠商修改，詳如說明，請查照`,
            '',
            { htmlBody: buildStage1EmailHtml_(audit, uploadUrl) }
          );
          markNotificationSent_(audit.id, audit['工程簡稱'], 'NOTIFY_S2');
          results.stage1++;
          const logEntry = `📨 [Stage1-S2上傳通知] ${abbr} → ${contractorEmail}`;
          sentLog.push(logEntry);
          console.log(logEntry);
        } else {
          console.log(`⚠️ [Stage1] ${abbr} - S2已上傳但承辦人Email為空，無法寄送`);
        }
      }

      // ─ Stage 2：S3 尚未上傳，距到期 3~2 天 → 每日寄承辦人
      if (s3NotUploaded && (daysLeft === 3 || daysLeft === 2)) {
        if (contractorEmail) {
          GmailApp.sendEmail(
            contractorEmail,
            `【工安查核】${audit['工程簡稱']} 距核章截止剩 ${daysLeft} 天，請速辦！`,
            '',
            { htmlBody: buildStage2EmailHtml_(audit, daysLeft, uploadUrl) }
          );
          results.stage2++;
          const logEntry = `📨 [Stage2-倒數${daysLeft}天] ${abbr} → ${contractorEmail}`;
          sentLog.push(logEntry);
          console.log(logEntry);
        } else {
          console.log(`⚠️ [Stage2] ${abbr} - 距到期${daysLeft}天但承辦人Email為空，無法寄送`);
        }
      }

      // ─ Stage 3：S3 尚未上傳，距到期 1~0 天 → 每日寄承辦人 + 課長
      if (s3NotUploaded && (daysLeft === 1 || daysLeft === 0)) {
        const recipients = [contractorEmail, managerEmail].filter(Boolean);
        if (recipients.length > 0) {
          const subjectText = daysLeft === 0
            ? `【到期提醒】${audit['工程簡稱']} 今日為核章截止日，請立即處理！`
            : `【緊急】${audit['工程簡稱']} 明日即為核章截止日，請立即處理！`;
          GmailApp.sendEmail(
            recipients.join(','),
            subjectText,
            '',
            { htmlBody: buildStage3EmailHtml_(audit, daysLeft, uploadUrl) }
          );
          results.stage3++;
          const logEntry = `📨 [Stage3-倒數${daysLeft}天] ${abbr} → ${recipients.join(', ')}`;
          sentLog.push(logEntry);
          console.log(logEntry);
        } else {
          console.log(`⚠️ [Stage3] ${abbr} - 距到期${daysLeft}天但收件人皆為空，無法寄送`);
        }
      }

      // ─ 逾期：S3 尚未上傳，已超過截止日 → 每日通知承辦人 + 課長
      if (s3NotUploaded && daysLeft < 0) {
        const daysOverdue = Math.abs(daysLeft);
        const recipients = [contractorEmail, managerEmail].filter(Boolean);
        if (recipients.length > 0) {
          GmailApp.sendEmail(
            recipients.join(','),
            `【逾期警告】${audit['工程簡稱']} 已逾期 ${daysOverdue} 天，請立即處理！`,
            '',
            { htmlBody: buildOverdueEmailHtml_(audit, daysOverdue, uploadUrl) }
          );
          results.overdue++;
          const logEntry = `📨 [逾期第${daysOverdue}天] ${abbr} → ${recipients.join(', ')}`;
          sentLog.push(logEntry);
          console.log(logEntry);
        } else {
          console.log(`⚠️ [逾期] ${abbr} - 已逾期${daysOverdue}天但收件人皆為空，無法寄送`);
        }
      }
    } catch (err) {
      const errMsg = `${audit.id} (${audit['工程簡稱'] || '?'}): ${err.message}`;
      results.errors.push(errMsg);
      console.error(`❌ [發送失敗] ${errMsg}`);
    }
  });

  // 彙整報告
  console.log('══════════════════════════════════════════════');
  console.log(`✅ [稽催排程完成] ${todayStr}`);
  console.log(`   Stage1(S2上傳通知): ${results.stage1} 封`);
  console.log(`   Stage2(倒數3~2天):  ${results.stage2} 封`);
  console.log(`   Stage3(倒數1~0天):  ${results.stage3} 封`);
  console.log(`   逾期通知:           ${results.overdue} 封`);
  console.log(`   略過(已結案/無資料): ${results.skipped} 筆`);
  console.log(`   發送失敗:           ${results.errors.length} 筆`);
  if (sentLog.length > 0) {
    console.log('── 本次發送明細 ──');
    sentLog.forEach(function(l) { console.log('   ' + l); });
  }
  if (results.errors.length > 0) {
    console.log('── 錯誤清單 ──');
    results.errors.forEach(function(e) { console.log('   ❌ ' + e); });
  }
  console.log('══════════════════════════════════════════════');

  return {
    success: true,
    message: `稽催完成：Stage1×${results.stage1} Stage2×${results.stage2} Stage3×${results.stage3} 逾期×${results.overdue}`,
    data: { sentLog: sentLog },
    errors: results.errors
  };
}

// 讓 GAS 觸發器可直接呼叫（不需要 roleData）
function runDailyReminderJob() { runDailyReminderJob_(); }

/** 防重複：檢查 Change Log 是否已有對應通知類型 */
function hasNotificationSent_(caseId, notifyType) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_CHANGE_LOG);
  if (!sheet || sheet.getLastRow() < 2) return false;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  return data.some(function(row) {
    return row[1] == caseId && row[4] === notifyType;
  });
}

/** 防重複：寫入已通知標記 */
function markNotificationSent_(caseId, abbr, notifyType) {
  logChange_(caseId, abbr, 'SYSTEM', notifyType, '自動通知已發送', '', '');
}

/**
 * 建立每日觸發器 (每天上午 8:00)
 * 只需執行一次；若已存在同名觸發器則略過
 */
function setupDailyTrigger_() {
  const funcName = 'runDailyReminderJob';
  const existing = ScriptApp.getProjectTriggers();
  const alreadySet = existing.some(function(t) {
    return t.getHandlerFunction() === funcName;
  });
  if (alreadySet) {
    return { success: true, message: '每日觸發器已存在，無需重複建立。' };
  }
  ScriptApp.newTrigger(funcName)
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
  return { success: true, message: '每日上午 8:00 觸發器建立完成。' };
}


// ==================== 安全上傳連結 (Token) ====================

/**
 * 產生安全上傳 Token
 * @param {string} caseId - 案件ID
 * @param {string} email - 授權的承辦人 Email
 * @param {string|Date} dueDate - 最晚應核章日期（Token 到期 = 此日期 + 7 天）
 * @returns {string} token
 */
function generateUploadToken_(caseId, email, dueDate) {
  const raw = caseId + email + new Date().getTime() + Math.random().toString(36);
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
  const token = digest.map(function(b) {
    return ('0' + ((b + 256) % 256).toString(16)).slice(-2);
  }).join('');

  // 計算到期時間：核章截止日 + 7 天
  const expiry = new Date(dueDate);
  expiry.setDate(expiry.getDate() + 7);

  // 取得或建立 Token 工作表
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_UPLOAD_TOKENS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_UPLOAD_TOKENS);
    sheet.getRange(1, 1, 1, 6).setValues([['Token', '案件ID', '授權Email', '建立時間', '到期時間', '狀態']]);
    sheet.setFrozenRows(1);
  }

  // 檢查是否已有該案件的有效 Token（避免重複產生）
  if (sheet.getLastRow() > 1) {
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][1] == caseId && data[i][5] === '有效') {
        const existingExpiry = new Date(data[i][4]);
        if (existingExpiry > new Date()) {
          console.log(`[Token] 案件 ${caseId} 已有有效 Token，直接沿用`);
          return data[i][0]; // 沿用現有 Token
        } else {
          // 已過期，標記
          sheet.getRange(i + 2, 6).setValue('已過期');
        }
      }
    }
  }

  sheet.appendRow([token, caseId, email, new Date(), expiry, '有效']);
  console.log(`[Token] 已為案件 ${caseId} 產生新 Token（到期：${expiry}）`);
  return token;
}

/**
 * 驗證 Token 並回傳案件摘要（公開 API，不需 Google 登入）
 */
function verifyUploadToken_(token) {
  if (!token) throw new Error('缺少上傳驗證碼 (Token)');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_UPLOAD_TOKENS);
  if (!sheet || sheet.getLastRow() < 2) throw new Error('無效的上傳連結');

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  let tokenRow = null;
  let tokenRowIdx = -1;

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === token) {
      tokenRow = data[i];
      tokenRowIdx = i + 2;
      break;
    }
  }

  if (!tokenRow) throw new Error('無效的上傳連結，請向工安組索取新的連結。');

  // 檢查狀態
  const status = tokenRow[5];
  if (status === '已過期') throw new Error('此上傳連結已過期，請向工安組索取新的連結。');

  // 檢查到期時間
  const expiry = new Date(tokenRow[4]);
  if (new Date() > expiry) {
    sheet.getRange(tokenRowIdx, 6).setValue('已過期');
    throw new Error('此上傳連結已過期，請向工安組索取新的連結。');
  }

  // Token 有效，取得案件資訊
  const caseId = tokenRow[1];
  const allAudits = getAuditRecords_();
  const audit = allAudits.find(function(a) { return a.id == caseId; });
  if (!audit) throw new Error('找不到對應的案件資料，案件可能已被刪除。');

  // 檢查案件狀態，已結案不允許上傳
  if (audit['辦理狀態'] === STATUS.STAGE4) {
    throw new Error('此案件已結案，無需再上傳。');
  }

  // 組裝 S2 下載連結
  const s2Links = [];
  const s2eUrl = audit['S2員工查核檔案位置'];
  const s2cUrl = audit['S2廠商查核檔案位置'];
  if (s2eUrl) {
    const fid = s2eUrl.includes('/d/') ? s2eUrl.split('/d/')[1].split('/')[0] : '';
    s2Links.push({ label: 'S2 員工改善單', viewUrl: s2eUrl, downloadUrl: fid ? 'https://drive.google.com/uc?export=download&id=' + fid : s2eUrl });
  }
  if (s2cUrl) {
    const fid = s2cUrl.includes('/d/') ? s2cUrl.split('/d/')[1].split('/')[0] : '';
    s2Links.push({ label: 'S2 廠商改善單', viewUrl: s2cUrl, downloadUrl: fid ? 'https://drive.google.com/uc?export=download&id=' + fid : s2cUrl });
  }

  return {
    success: true,
    data: {
      caseId: caseId,
      abbr: audit['工程簡稱'],
      name: audit['工程名稱'],
      contractor: audit['承攬商'],
      department: audit['主辦部門'],
      auditDate: audit['查核日期'],
      dueDate: audit['最晚應核章日期'],
      status: audit['辦理狀態'],
      s2Links: s2Links,
      authorizedEmail: tokenRow[2],
      tokenExpiry: Utilities.formatDate(expiry, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    }
  };
}

/**
 * Token 授權上傳 S3 檔案（公開 API，不需 Google 登入）
 */
function tokenUploadFile_(payload) {
  const token = payload.token;
  const fileBase64 = payload.fileBase64;
  const fileName = payload.fileName;

  if (!token || !fileBase64 || !fileName) throw new Error('缺少必要參數');

  // 先驗證 Token
  const verify = verifyUploadToken_(token);
  if (!verify.success) throw new Error(verify.message);

  const caseId = verify.data.caseId;
  const authorizedEmail = verify.data.authorizedEmail;

  // 使用與主系統相同的上傳邏輯
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_AUDIT_LIST);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const caseIdCol = headers.indexOf('案件ID');

  const caseIds = sheet.getRange(2, caseIdCol + 1, sheet.getLastRow() - 1, 1).getValues().flat();
  const rowIdx = caseIds.findIndex(function(id) { return id == caseId; }) + 2;
  if (rowIdx < 2) throw new Error('找不到對應的案件');

  const rowData = sheet.getRange(rowIdx, 1, 1, headers.length).getValues()[0];
  const auditData = {};
  headers.forEach(function(h, i) { auditData[h] = rowData[i]; });

  // 產生檔名
  const fileExtension = fileName.includes('.') ? '.' + fileName.split('.').pop() : '.pdf';
  const newFileName = generateFileName_('stage3', auditData, fileExtension);

  // 建立資料夾與上傳
  const auditDate = new Date(auditData['查核日期']);
  const formattedDate = Utilities.formatDate(auditDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const folderName = formattedDate + '_' + auditData['工程簡稱'];
  const rootFolder = DriveApp.getFolderById(MAIN_DRIVE_FOLDER_ID);

  var targetFolderIterator = rootFolder.getFoldersByName(folderName);
  const targetFolder = targetFolderIterator.hasNext() ? targetFolderIterator.next() : rootFolder.createFolder(folderName);

  const blob = Utilities.newBlob(Utilities.base64Decode(fileBase64), 'application/pdf', newFileName);
  const uploadedFile = targetFolder.createFile(blob);
  const fileUrl = uploadedFile.getUrl();

  // 更新狀態
  const currentStatus = auditData['辦理狀態'] || '';
  const statusOrder = [STATUS.STAGE1, STATUS.STAGE2, STATUS.STAGE3, STATUS.STAGE4];
  const currentIndex = statusOrder.indexOf(currentStatus);
  if (currentIndex < 2) {
    sheet.getRange(rowIdx, headers.indexOf('辦理狀態') + 1).setValue(STATUS.STAGE3);
  }

  // 寫入 S3 檔案連結
  const s3Col = headers.indexOf('S3廠商及員工改善後核章檔案位置');
  if (s3Col > -1) {
    sheet.getRange(rowIdx, s3Col + 1).setValue(fileUrl);
  }

  sheet.getRange(rowIdx, headers.indexOf('修改人員') + 1).setValue(authorizedEmail + ' (Token上傳)');

  logChange_(caseId, auditData['工程簡稱'], authorizedEmail + '(Token)', '檔案上傳', 'S3 廠商及員工 (安全連結上傳)', newFileName, fileUrl);

  console.log(`📤 [Token上傳成功] ${auditData['工程簡稱']} | 上傳者: ${authorizedEmail} | 檔案: ${newFileName}`);

  return {
    success: true,
    message: '檔案 "' + newFileName + '" 上傳成功！案件狀態已更新為第3階段。'
  };
}

/**
 * 產生上傳頁面的完整 URL
 */
function getUploadPageUrl_(token) {
  return SYSTEM_URL + 'upload.html?token=' + token;
}


// ==================== HTML 郵件模板 ====================

/** 共用 Header/Footer HTML */
function _emailHeader_(title, subtitle, accentColor) {
  accentColor = accentColor || '#1e40af';
  // 使用 SVG 替代 Emoji 以防破圖
  const iconSvg = `<svg width="48" height="48" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg" style="opacity:0.6;"><path d="M19 21H5C4.44772 21 4 20.5523 4 20V4C4 3.44772 4.44772 3 5 3H14L20 9V20C20 20.5523 19.5523 21 19 21Z" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><path d="M14 3V9H20" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><path d="M8 13H16" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><path d="M8 17H16" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
  
  return `
<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#f1f5f9;font-family:'Segoe UI',Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f1f5f9;">
<tr><td align="center" style="padding:32px 16px;">
<table width="620" cellpadding="0" cellspacing="0" style="background:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,0.10);">

<!-- Header -->
<tr><td style="background:${accentColor};padding:28px 36px;">
  <table width="100%" cellpadding="0" cellspacing="0">
    <tr>
      <td>
        <div style="font-size:11px;color:rgba(255,255,255,0.7);letter-spacing:2px;text-transform:uppercase;margin-bottom:6px;">工安查核管理系統</div>
        <div style="font-size:24px;font-weight:700;color:#ffffff;line-height:1.3;">${title}</div>
        <div style="font-size:14px;color:rgba(255,255,255,0.85);margin-top:6px;">${subtitle}</div>
      </td>
      <td align="right" style="width:60px;">${iconSvg}</td>
    </tr>
  </table>
</td></tr>
`;
}

function _emailFooter_(systemUrl, uploadUrl) {
  systemUrl = systemUrl || SYSTEM_URL;
  const uploadBtn = uploadUrl ? `
        <a href="${uploadUrl}" style="display:inline-block;background:#059669;color:#ffffff;text-decoration:none;font-size:13px;font-weight:600;padding:10px 22px;border-radius:6px;margin-right:8px;">📤 上傳核章版</a>
  ` : '';
  return `
<!-- Footer -->
<tr><td style="background:#f8fafc;padding:20px 36px;border-top:1px solid #e2e8f0;">
  <table width="100%" cellpadding="0" cellspacing="0">
    <tr>
      <td style="font-size:12px;color:#94a3b8;line-height:1.6;">
        此信件由工安查核管理系統自動發送，請勿直接回覆。<br>
        如有疑問請聯絡工安組。
      </td>
      <td align="right">
        ${uploadBtn}
        <a href="${systemUrl}" style="display:inline-block;background:#1e40af;color:#ffffff;text-decoration:none;font-size:13px;font-weight:600;padding:10px 22px;border-radius:6px;">前往系統</a>
      </td>
    </tr>
  </table>
</td></tr>

</table>
</td></tr></table>
</body></html>
`;
}

/** 共用案件資訊 Table */
function _caseInfoTable_(audit, uploadUrl) {
  const s2Url = audit['S2員工查核檔案位置'] || audit['S2廠商查核檔案位置'] || audit['第2階段連結'];
  // 建立直接下載連結：將 /file/d/.../view 轉換為 /uc?export=download&id=...
  let s2DownloadLink = '';
  if (s2Url && s2Url.includes('/d/')) {
    const fileId = s2Url.split('/d/')[1].split('/')[0];
    const downloadUrl = `https://drive.google.com/uc?export=download&id=${fileId}`;
    s2DownloadLink = `
    <tr>
      <td style="padding:10px 14px;font-size:12px;color:#64748b;font-weight:600;white-space:nowrap;width:110px;border-bottom:1px solid #e2e8f0;background:#f8fafc;">S2 改善單</td>
      <td style="padding:10px 14px;font-size:14px;color:#0f172a;border-bottom:1px solid #e2e8f0;">
        <a href="${downloadUrl}" style="color:#1e40af;font-weight:700;text-decoration:none;">📥 點此直接下載 S2 檔案</a>
      </td>
    </tr>`;
  }

  const uploadLinkRow = uploadUrl ? `
    <tr style="background:#f0fdf4;">
      <td style="padding:10px 14px;font-size:12px;color:#166534;font-weight:600;white-space:nowrap;width:110px;border-bottom:1px solid #e2e8f0;">安全上傳連結</td>
      <td style="padding:10px 14px;font-size:14px;color:#059669;border-bottom:1px solid #e2e8f0;">
        <a href="${uploadUrl}" style="color:#059669;font-weight:700;text-decoration:none;">📤 點此直接進入上傳頁面 (免登入)</a>
      </td>
    </tr>` : '';

  return `
<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin-top:16px;border:1px solid #e2e8f0;">
  ${uploadLinkRow}
  <tr style="background:#f8fafc;">
    <td style="padding:10px 14px;font-size:12px;color:#64748b;font-weight:600;white-space:nowrap;width:110px;border-bottom:1px solid #e2e8f0;">工程簡稱</td>
    <td style="padding:10px 14px;font-size:14px;color:#1e293b;font-weight:600;border-bottom:1px solid #e2e8f0;">${audit['工程簡稱'] || '-'}</td>
  </tr>
  <tr>
    <td style="padding:10px 14px;font-size:12px;color:#64748b;font-weight:600;white-space:nowrap;border-bottom:1px solid #e2e8f0;">工程名稱</td>
    <td style="padding:10px 14px;font-size:14px;color:#1e293b;border-bottom:1px solid #e2e8f0;">${audit['工程名稱'] || '-'}</td>
  </tr>
  <tr style="background:#f8fafc;">
    <td style="padding:10px 14px;font-size:12px;color:#64748b;font-weight:600;white-space:nowrap;border-bottom:1px solid #e2e8f0;">承攬商</td>
    <td style="padding:10px 14px;font-size:14px;color:#1e293b;border-bottom:1px solid #e2e8f0;">${audit['承攬商'] || '-'}</td>
  </tr>
  <tr>
    <td style="padding:10px 14px;font-size:12px;color:#64748b;font-weight:600;white-space:nowrap;border-bottom:1px solid #e2e8f0;">主辦部門</td>
    <td style="padding:10px 14px;font-size:14px;color:#1e293b;border-bottom:1px solid #e2e8f0;">${audit['主辦部門'] || '-'}</td>
  </tr>
  <tr style="background:#f8fafc;">
    <td style="padding:10px 14px;font-size:12px;color:#64748b;font-weight:600;white-space:nowrap;border-bottom:1px solid #e2e8f0;">核章截止日</td>
    <td style="padding:10px 14px;font-size:14px;color:#dc2626;font-weight:700;border-bottom:1px solid #e2e8f0;">${audit['最晚應核章日期'] || '-'}</td>
  </tr>
  ${s2DownloadLink}
</table>
`;
}

/**
 * Stage 1 HTML 模板：S2 已上傳，通知承辦人可先行下載轉知承攬商改善
 */
function buildStage1EmailHtml_(audit, uploadUrl) {
  const uploadHint = uploadUrl ? `<br>📤 或直接<a href="${uploadUrl}" style="color:#059669;font-weight:700;">點此上傳 S3 核章版</a>（免登入）` : '';
  return _emailHeader_(
    '改善單已上傳，可先行下載轉知承攬商改善',
    `${audit['主辦部門']} ｜ ${audit['工程簡稱']}`,
    '#0f766e'
  ) + `
<!-- Body -->
<tr><td style="padding:32px 36px;">
  <p style="font-size:15px;color:#334155;line-height:1.7;margin:0 0 16px;">
    <strong>${audit['承辦人員姓名'] || audit['承辦人姓名'] || '承辦人'}</strong> 您好，
  </p>
  <p style="font-size:15px;color:#334155;line-height:1.7;margin:0 0 20px;">
    <strong>「原始改善單」已上傳</strong>，
    承辦人員可<strong style="color:#0f766e;">先行轉知廠商改善</strong>，
    書面文件待核定後傳遞。
  </p>

  ${_caseInfoTable_(audit, uploadUrl)}

  <!-- 綠色提示框 -->
  <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:20px;">
    <tr><td style="background:#ecfdf5;border-left:4px solid #10b981;border-radius:0 8px 8px 0;padding:14px 18px;">
      <p style="margin:0;font-size:14px;color:#065f46;line-height:1.7;">
        📋 <strong>後續步驟：</strong>請確認改善內容正確後，
        上傳「S3 工作隊核章版」完成程序。${uploadHint}
      </p>
    </td></tr>
  </table>
</td></tr>
` + _emailFooter_(null, uploadUrl);
}

/**
 * Stage 2 HTML 模板：距到期 3 天，提醒速辦 + 核章日期警示
 */
function buildStage2EmailHtml_(audit, daysLeft, uploadUrl) {
  const dueDate = audit['最晚應核章日期'];
  return _emailHeader_(
    `距核章截止剩 ${daysLeft} 天，請速辦！`,
    `${audit['主辦部門']} ｜ ${audit['工程簡稱']}`,
    '#b45309'
  ) + `
<!-- Body -->
<tr><td style="padding:32px 36px;">

  <!-- 倒數警示 -->
  <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px;">
    <tr><td align="center" style="background:#fef3c7;border:2px solid #f59e0b;border-radius:10px;padding:20px;">
      <div style="font-size:14px;color:#92400e;font-weight:600;letter-spacing:1px;">核章截止倒數</div>
      <div style="font-size:52px;font-weight:900;color:#b45309;line-height:1;margin:8px 0;">${daysLeft}</div>
      <div style="font-size:16px;color:#92400e;font-weight:700;">天</div>
    </td></tr>
  </table>

  <p style="font-size:15px;color:#334155;line-height:1.7;margin:0 0 20px;">
    以下案件的<strong>「S3 工作隊核章版」尚未上傳</strong>，距截止日期僅剩 <strong style="color:#b45309;">${daysLeft} 天</strong>，請儘速辦理。
  </p>

  ${_caseInfoTable_(audit, uploadUrl)}

  <!-- 紅色重點提醒 -->
  <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:20px;">
    <tr><td style="background:#fef2f2;border-left:4px solid #ef4444;border-radius:0 8px 8px 0;padding:14px 18px;">
      <p style="margin:0;font-size:14px;color:#7f1d1d;line-height:1.7;">
        ⚠️ <strong>重要提醒：</strong>請特別注意<strong>第一個章的核章日期</strong>，
        核章日期必須在截止日 <strong>${dueDate}</strong> 之前，否則將視為逾期。
      </p>
    </td></tr>
  </table>
</td></tr>
` + _emailFooter_(null, uploadUrl);
}

/**
 * Stage 3 HTML 模板：距到期 1 天，同時寄承辦人+課長，強調最後機會
 */
function buildStage3EmailHtml_(audit, daysLeft, uploadUrl) {
  const dueDate = audit['最晚應核章日期'];
  return _emailHeader_(
    `【緊急】明日即為核章截止日！`,
    `${audit['主辦部門']} ｜ ${audit['工程簡稱']}`,
    '#991b1b'
  ) + `
<!-- Body -->
<tr><td style="padding:32px 36px;">

  <!-- 緊急警示橫幅 -->
  <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px;">
    <tr><td align="center" style="background:#fee2e2;border:2px solid #ef4444;border-radius:10px;padding:20px;">
      <div style="font-size:28px;margin-bottom:8px;">🚨</div>
      <div style="font-size:18px;font-weight:800;color:#991b1b;">最後 ${daysLeft} 天！請立即處理</div>
      <div style="font-size:13px;color:#7f1d1d;margin-top:6px;">此訊息已同步通知承辦人及課長</div>
    </td></tr>
  </table>

  <p style="font-size:15px;color:#334155;line-height:1.7;margin:0 0 20px;">
    以下案件的<strong>「S3 工作隊核章版」仍未上傳</strong>，明日 <strong style="color:#991b1b;">${dueDate}</strong> 即為最後截止日。
    請承辦人員<strong>立即聯繫相關單位</strong>完成核章程序。
  </p>

  ${_caseInfoTable_(audit, uploadUrl)}

  <!-- 雙重重點提醒 -->
  <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:20px;">
    <tr><td style="background:#fef2f2;border-left:4px solid #ef4444;border-radius:0 8px 8px 0;padding:14px 18px;margin-bottom:12px;">
      <p style="margin:0;font-size:14px;color:#7f1d1d;line-height:1.8;">
        🔴 <strong>第一點：</strong>請<strong>立即</strong>上傳 S3 工作隊核章版至系統。<br>
        🔴 <strong>第二點：</strong>請特別確認<strong>第一個章的核章日期</strong>必須在
        <strong>${dueDate}</strong> 當日或之前，核章日期不符將視為無效。
      </p>
    </td></tr>
  </table>
</td></tr>
` + _emailFooter_(null, uploadUrl);
}

/**
 * 逾期通知 HTML 模板：S3 尚未上傳，已超過截止日，每日通知承辦人 + 課長
 */
function buildOverdueEmailHtml_(audit, daysOverdue, uploadUrl) {
  const dueDate = audit['最晚應核章日期'];
  return _emailHeader_(
    `【逾期警告】已逾期 ${daysOverdue} 天！`,
    `${audit['主辦部門']} ｜ ${audit['工程簡稱']}`,
    '#4c1d95'
  ) + `
<!-- Body -->
<tr><td style="padding:32px 36px;">

  <!-- 逾期天數大圖示 -->
  <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:24px;">
    <tr><td align="center" style="background:#f5f3ff;border:2px solid #7c3aed;border-radius:10px;padding:20px;">
      <div style="font-size:14px;color:#4c1d95;font-weight:600;letter-spacing:1px;">⛔ 已逾期</div>
      <div style="font-size:64px;font-weight:900;color:#6d28d9;line-height:1;margin:8px 0;">${daysOverdue}</div>
      <div style="font-size:16px;color:#4c1d95;font-weight:700;">天</div>
      <div style="font-size:12px;color:#7c3aed;margin-top:6px;">截止日：${dueDate}</div>
    </td></tr>
  </table>

  <p style="font-size:15px;color:#334155;line-height:1.7;margin:0 0 20px;">
    以下案件的<strong>「S3 工作隊核章版」仍未上傳</strong>，已超過截止日期 <strong style="color:#6d28d9;">${daysOverdue} 天</strong>，
    現通知<strong>承辦人員及課長</strong>，請立即確認並儘速補件。
  </p>

  ${_caseInfoTable_(audit, uploadUrl)}

  <!-- 紫色警示框 -->
  <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:20px;">
    <tr><td style="background:#f5f3ff;border-left:4px solid #7c3aed;border-radius:0 8px 8px 0;padding:14px 18px;">
      <p style="margin:0;font-size:14px;color:#4c1d95;line-height:1.8;">
        ⛔ <strong>第一點：</strong>請<strong>立即</strong>上傳 S3 工作隊核章版至系統。<br>
        ⛔ <strong>第二點：</strong>核章日期若早於截止日 <strong>${dueDate}</strong> 仍可受理，請確認後盡快上傳。
      </p>
    </td></tr>
  </table>
</td></tr>
` + _emailFooter_(null, uploadUrl);
}

/** 個別案件稽催預覽 */
function previewCaseReminder_(caseId) {
    const allAudits = getAuditRecords_();
    const audit = allAudits.find(a => a.id === caseId);
    if (!audit) throw new Error("找不到案件：" + caseId);
    if (audit['辦理狀態'] === STATUS.STAGE4) throw new Error("案件已結案，無需稽催。");

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    let daysLeft = null;
    if (audit['最晚應核章日期']) {
        const dueDate = new Date(audit['最晚應核章日期']);
        dueDate.setHours(0, 0, 0, 0);
        daysLeft = Math.round((dueDate - today) / 86400000);
    }

    const s2Uploaded = !!(audit['S2員工查核檔案位置'] || audit['S2廠商查核檔案位置']);
    const s3Uploaded = !!audit['S3廠商及員工改善後核章檔案位置'];
    if (s3Uploaded) throw new Error("S3 已上傳，無需稽催。");

    let subject = '';
    let htmlBody = '';

    if (!s2Uploaded) {
        subject = `【工安查核進度提醒】${audit['工程簡稱']} 尚未上傳 S2改善單`;
        htmlBody = _emailHeader_('案件進度提醒', `${audit['主辦部門']} ｜ ${audit['工程簡稱']}`, '#ea580c') + `
        <tr><td style="padding:32px 36px;">
          <p style="font-size:15px;color:#334155;line-height:1.7;margin:0 0 20px;"><strong>${audit['承辦人姓名'] || '承辦人'}</strong> 您好，</p>
          <p style="font-size:15px;color:#334155;line-height:1.7;margin:0 0 20px;">系統提醒您，此案件<strong>「S2 原始改善單」尚未上傳</strong>，請儘速補齊相關照片與文件。</p>
          ${_caseInfoTable_(audit, null)}
        </td></tr>` + _emailFooter_();
    } else {
        if (daysLeft !== null && daysLeft <= 1) {
            subject = `【緊急】${audit['工程簡稱']} 核章截止日將至，請立即處理！`;
            htmlBody = buildStage3EmailHtml_(audit, daysLeft, null);
        } else {
            const displayDays = daysLeft !== null ? daysLeft : '未設定';
            subject = `【工安查核】${audit['工程簡稱']} 距核章截止剩 ${displayDays} 天，請速辦！`;
            htmlBody = buildStage2EmailHtml_(audit, displayDays);
        }
    }

    const contractorEmail = audit['承辦人Email'] || audit['承辦人電子信箱'] || '';
    const managerEmail = audit['課長Email'] || audit['承辦課長電子信箱'] || '';
    const defRecipients = [contractorEmail, managerEmail].filter(Boolean);

    return { success: true, subject, htmlBody, recipients: defRecipients.join(', '), caseId };
}

/** 個別案件發送稽催 */
function sendCaseReminder_(payload) {
    try {
        if (!payload.recipients) throw new Error("請填寫收件人");
        const attachments = []; // 若需要也可抓取該案件的 Log 夾帶附件
        
        GmailApp.sendEmail(payload.recipients, payload.subject, '', {
            htmlBody: payload.htmlBody,
            attachments: attachments
        });
        
        return { success: true, message: "稽催信件已順利寄出！" };
    } catch (e) {
        throw new Error("寄出失敗：" + e.message);
    }
}



// ==================== 帳號管理 ====================

/**
 * 寫入「帳號管理」分頁
 * payload: { deptName, contractorName, contractorEmail, managerName, managerEmail, note }
 */
function registerDeptAccount_(payload) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_DEPT_ACCOUNTS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_DEPT_ACCOUNTS);
      sheet.getRange(1, 1, 1, 7).setValues([[
        '部門名稱', '承辦人姓名', '承辦人Email', '課長姓名', '課長Email', '備註', '建立日期'
      ]]).setBackground('#e0f2fe').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // 若已有同部門資料則更新，否則新增
    const data = sheet.getDataRange().getValues();
    let existingRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === payload.deptName) {
        existingRow = i + 1;
        break;
      }
    }

    const rowData = [
      payload.deptName || '',
      payload.contractorName || '',
      payload.contractorEmail || '',
      payload.managerName || '',
      payload.managerEmail || '',
      payload.note || '',
      new Date()
    ];

    if (existingRow > -1) {
      sheet.getRange(existingRow, 1, 1, 7).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }

    return { success: true, message: `部門「${payload.deptName}」帳號資料已儲存。` };
  } catch (e) {
    throw new Error('帳號管理儲存失敗: ' + e.message);
  }
}

/** 讀取「帳號管理」分頁所有資料 */
function getDeptAccounts_() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_DEPT_ACCOUNTS);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, data: [] };

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
    const accounts = data.map(function(row) {
      return {
        deptName: row[0],
        contractorName: row[1],
        contractorEmail: row[2],
        managerName: row[3],
        managerEmail: row[4],
        note: row[5],
        createdAt: row[6] instanceof Date
          ? Utilities.formatDate(row[6], Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : row[6]
      };
    }).filter(function(r) { return r.deptName; });

    return { success: true, data: accounts };
  } catch (e) {
    throw new Error('讀取帳號管理失敗: ' + e.message);
  }
}

/**
 * 【自動解析】從 S2 檔案中透過 OCR 提取缺失項目
 */
function extractDeficienciesFromS2_(fileId, caseId, projectAbbr, deptName, email) {
  try {
    const file = DriveApp.getFileById(fileId);
    
    // 1. 透過 Drive API 建立 OCR 文字副本 (需要開啟 Drive Advanced Service)
    // 若未開啟則會報錯於此，建議於 README 註記
    const resource = {
      title: 'TEMP_OCR_' + caseId,
      mimeType: file.getMimeType()
    };
    
    // 使用 Drive API v2
    const tempFile = Drive.Files.insert(resource, file.getBlob(), { ocr: true });
    const doc = DocumentApp.openById(tempFile.id);
    const text = doc.getBody().getText();
    
    // 2. 定位標籤：「第 4 項：建議及應行改善事項」
    const keyword = "第 4 項：建議及應行改善事項";
    const startIdx = text.indexOf(keyword);
    if (startIdx === -1) {
      console.warn("S2 檔案中找不到指定關鍵字標籤");
      Drive.Files.remove(tempFile.id); // 刪除暫存檔
      return;
    }
    
    // 3. 擷取文字內容 (到下一個「第 X 項」或結尾)
    const contentArea = text.substring(startIdx + keyword.length);
    const endMatch = contentArea.match(/第 [0-9一二三四五] 項/);
    let extractedText = endMatch ? contentArea.substring(0, endMatch.index) : contentArea;
    
    // 4. 清洗與過濾
    const lines = extractedText.split('\n')
      .map(function(line) { return line.trim(); })
      .filter(function(line) { 
        // 過濾掉空值、標題重複、或太短的雜訊
        return line.length > 2 && !line.includes('缺失內容') && !line.includes('改善建議');
      });
      
    if (lines.length > 0) {
      const sheet = getOrCreateDeficiencySheet_();
      const now = new Date().getTime();
      const rows = lines.map(function(content, idx) {
        const newId = 'DEF' + (now + idx);
        // 預設改善期限為查核日期 + 7 天 (或依需求調整)
        const deadline = ""; 
        return [newId, caseId, projectAbbr, content, deptName, deadline, '待改善', email];
      });
      
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
      console.log(`已從 S2 自動匯入 ${lines.length} 項缺失。`);
    }
    
    // 5. 清理暫存檔
    Drive.Files.remove(tempFile.id);
    
  } catch (e) {
    console.error("OCR 解析失敗:", e.message);
    // 不中斷主流程
  }
}

// ==================== 部門清單管理 (第 10 次優化新增) ====================

function getDeptMembers_() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_DEPT_LIST);
    if (!sheet) return { success: true, data: [] };
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: true, data: [] };
    
    const headers = data[0];
    const members = data.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    }).filter(m => m['主辦部門'] && m['姓名']);
    
    return { success: true, data: members };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function saveDeptMember_(payload) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_DEPT_LIST);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_DEPT_LIST);
      sheet.appendRow(['ID', '主辦部門', '職稱', '姓名', '信箱']);
    }
    
    const data = sheet.getDataRange().getValues();
    const id = payload.ID || ('M' + new Date().getTime());
    const rowData = [id, payload['主辦部門'], payload['職稱'], payload['姓名'], payload['信箱']];
    
    let existingRow = -1;
    if (payload.ID) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === payload.ID) {
          existingRow = i + 1;
          break;
        }
      }
    }
    
    if (existingRow > -1) {
      sheet.getRange(existingRow, 1, 1, 5).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    
    return { success: true, message: "資料已儲存" };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteDeptMember_(id) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_DEPT_LIST);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "資料已刪除" };
      }
    }
    return { success: false, message: "找不到該資料" };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
