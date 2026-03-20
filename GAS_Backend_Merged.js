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

// ── 測試模式設定（優先讀 Script Properties，讀不到則用下方預設值）──
const _props = PropertiesService.getScriptProperties();
const TEST_EMAIL = _props.getProperty('TEST_EMAIL') || 'clinlion418@gmail.com';
// 系統前台連結（信件中「前往系統」按鈕）
const SYSTEM_URL = _props.getProperty('SYSTEM_URL') || 'https://your-github-pages-url.github.io/';

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

    // 【公開 API】允許未登入存取特定功能
    if (action === 'get_public_cases') {
        return createJsonResponse(getPublicCases_());
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
        result = setupSystem_();
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
      projects: baseData.projects
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
    const requiredHeaders = ['案件ID', '查核日期', '工程名稱', '工程簡稱', '承攬商', '主辦部門', '最晚應核章日期', '辦理狀態', '查核人員', '修改人員', '結案日期', '查核領隊', '查核成員', '承辦人姓名', '承辦人電子信箱', '承辦課長職稱', '承辦課長電子信箱'];
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
        case '修改人員': newRow[header] = data.modifier; break;
        case '查核領隊': newRow[header] = data.auditLeader; break;
        case '查核成員': newRow[header] = data.auditMembers; break;
        case '承辦人姓名': newRow[header] = data.contractorName; break;
        case '承辦人電子信箱': newRow[header] = data.contractorEmail; break;
        case '承辦課長職稱': newRow[header] = data.contractorManagerTitle; break;
        case '承辦課長電子信箱': newRow[header] = data.contractorManagerEmail; break;
        default: newRow[header] = "";
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
    if (details.inspector !== undefined) sheet.getRange(rowIdx, headers.indexOf('查核人員') + 1).setValue(details.inspector);
    if (details.auditLeader !== undefined) sheet.getRange(rowIdx, headers.indexOf('查核領隊') + 1).setValue(details.auditLeader);
    if (details.auditMembers !== undefined) sheet.getRange(rowIdx, headers.indexOf('查核成員') + 1).setValue(details.auditMembers);
    if (details.contractorName !== undefined) sheet.getRange(rowIdx, headers.indexOf('承辦人姓名') + 1).setValue(details.contractorName);
    if (details.contractorEmail !== undefined) sheet.getRange(rowIdx, headers.indexOf('承辦人電子信箱') + 1).setValue(details.contractorEmail);
    if (details.contractorManagerTitle !== undefined) sheet.getRange(rowIdx, headers.indexOf('承辦課長職稱') + 1).setValue(details.contractorManagerTitle);
    if (details.contractorManagerEmail !== undefined) sheet.getRange(rowIdx, headers.indexOf('承辦課長電子信箱') + 1).setValue(details.contractorManagerEmail);
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
        targetHeader = '第2階段連結-員工';
    } else if (stage === 'stage2c') {
        targetHeader = '第2階段連結-廠商';
    } else if (stage === 'stage3') {
        targetHeader = '第3階段連結';
    } else if (stage === 'stage4' || stage === 'stage4e') {
        targetHeader = '第4階段連結-員工';
    } else if (stage === 'stage4c') {
        targetHeader = '第4階段連結-承攬商';
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
        case 'stage3': stageDisplay = 'S3 廠商核章'; break;
        case 'stage4e': 
        case 'stage4': stageDisplay = 'S4 員工結案版'; break;
        case 'stage4c': stageDisplay = 'S4 承攬商結案版'; break;
        default: stageDisplay = '檔案上傳';
    }
    logChange_(caseId, auditData['工程簡稱'], userEmail, '檔案上傳', stageDisplay, newFileName, fileUrl);
    
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
function setupSystem_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetsToCreate = [
    { name: SHEET_AUDIT_LIST, headers: ['案件ID', '查核日期', '工程名稱', '工程簡稱', '承攬商', '主辦部門', '最晚應核章日期', '辦理狀態', '查核人員', '修改人員', '結案日期', '查核領隊', '查核成員', '承辦人姓名', '承辦人電子信箱', '承辦課長職稱', '承辦課長電子信箱', '第2階段連結-員工', '第2階段連結-廠商', '第3階段連結', '第4階段連結-員工', '第4階段連結-承攬商'], color: '#f3f4f6' },
    { name: SHEET_PROJECT_DB, headers: ['流水號', '工程簡稱', '工程名稱', '承攬商', '主辦部門', '承辦人姓名', '承辦人電子信箱', '承辦課長職稱', '承辦課長電子信箱'], color: '#f3f4f6' },
    { name: SHEET_DEFICIENCY_DB, headers: ['缺失ID', '案件ID', '工程簡稱', '缺失內容', '主辦部門', '改善期限', '狀態', '錄入者'], color: '#fef3c7' },
    { name: SHEET_CHANGE_LOG, headers: ['修改日期', '案件ID', '工程簡稱', '修改人員', '狀態', '說明', '檔案名稱', '檔案位置'], color: '#eff6ff' },
    { name: SHEET_FILE_HISTORY, headers: ['異動日期', '異動人員', '案件ID', '異動內容', '檔案類型', '檔案連結', '狀態'], color: '#d1fae5' },
    { name: SHEET_USER_PERMISSIONS, headers: ['信箱 (Email)', '姓名 (Name)', '角色 (Role)', '所屬部門 (Department)', '啟用狀態 (Active)'], color: '#f3f4f6' },
    { name: SHEET_DEPT_ACCOUNTS, headers: ['部門名稱', '承辦人姓名', '承辦人Email', '課長姓名', '課長Email', '備註', '建立日期'], color: '#e0f2fe' }
  ];

  sheetsToCreate.forEach(cfg => {
    let sheet = ss.getSheetByName(cfg.name);
    if (!sheet) {
      sheet = ss.insertSheet(cfg.name);
    }
    // 檢查標題列，若空白則寫入
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, cfg.headers.length).setValues([cfg.headers]).setBackground(cfg.color).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  });

  return { success: true, message: "系統分頁與標題列初始化完成。" };
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
  const requiredHeaders = ['案件ID', '查核日期', '工程名稱', '工程簡稱', '承攬商', '主辦部門', '最晚應核章日期', '辦理狀態', '查核人員', '修改人員', '結案日期', '查核領隊', '查核成員', '承辦人姓名', '承辦人電子信箱', '承辦課長職稱', '承辦課長電子信箱'];
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
        const deptRecipients = mailList[dept] || [];
        const safetyTeamRecipients = mailList['工安組'] || [];
        auditsByDept[dept].recipients = [...new Set([...deptRecipients, ...safetyTeamRecipients])];
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
        const recipients = mailList[dept] || [];
        const safetyTeamRecipients = mailList['工安組'] || [];
        auditsByDept[dept].recipients = [...new Set([...recipients, ...safetyTeamRecipients])];
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
 * 三階段邏輯:
 *   Stage1: S2 已上傳 且 尚未發送過 Stage1 通知 → 寄承辦人
 *   Stage2: S3 尚未上傳 且 距到期恰好 3 天 → 寄承辦人
 *   Stage3: S3 尚未上傳 且 距到期恰好 1 天 → 寄承辦人 + 課長
 */
function runDailyReminderJob_() {
  const allAudits = getAuditRecords_();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const results = { stage1: 0, stage2: 0, stage3: 0, errors: [] };

  allAudits.forEach(function(audit) {
    try {
      const status = audit['辦理狀態'];
      if (status === STATUS.STAGE4) return; // 已結案略過

      const dueDateStr = audit['最晚應核章日期'];
      if (!dueDateStr) return;
      const dueDate = new Date(dueDateStr);
      dueDate.setHours(0, 0, 0, 0);
      const daysLeft = Math.round((dueDate - today) / 86400000);

      const s2Uploaded = (status === STATUS.STAGE2 || status === STATUS.STAGE3);
      const s3NotUploaded = (status === STATUS.STAGE2 || status === STATUS.STAGE1);

      const contractorEmail = audit['承辦人電子信箱'];
      const managerEmail = audit['承辦課長電子信箱'];

      // ─ Stage 1：S2 已上傳，立即通知（只送一次，用 Change Log 防重複）
      if (status === STATUS.STAGE2 && !hasNotificationSent_(audit.id, 'NOTIFY_S2')) {
        if (contractorEmail) {
          GmailApp.sendEmail(
            contractorEmail,
            `【工安查核】${audit['工程簡稱']} 改善單已上傳，可線上修改`,
            '',
            { htmlBody: buildStage1EmailHtml_(audit) }
          );
          markNotificationSent_(audit.id, audit['工程簡稱'], 'NOTIFY_S2');
          results.stage1++;
        }
      }

      // ─ Stage 2：S3 尚未上傳，距到期 3 天
      if (s3NotUploaded && daysLeft === 3) {
        if (contractorEmail) {
          GmailApp.sendEmail(
            contractorEmail,
            `【工安查核】${audit['工程簡稱']} 距核章截止剩 3 天，請速辦！`,
            '',
            { htmlBody: buildStage2EmailHtml_(audit, 3) }
          );
          results.stage2++;
        }
      }

      // ─ Stage 3：S3 尚未上傳，距到期 1 天 → 承辦人 + 課長
      if (s3NotUploaded && daysLeft === 1) {
        const recipients = [contractorEmail, managerEmail].filter(Boolean);
        if (recipients.length > 0) {
          GmailApp.sendEmail(
            recipients.join(','),
            `【緊急】${audit['工程簡稱']} 明日即為核章截止日，請立即處理！`,
            '',
            { htmlBody: buildStage3EmailHtml_(audit, 1) }
          );
          results.stage3++;
        }
      }
    } catch (err) {
      results.errors.push(`${audit.id}: ${err.message}`);
    }
  });

  return {
    success: true,
    message: `稽催完成：Stage1×${results.stage1} Stage2×${results.stage2} Stage3×${results.stage3}`,
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

function _emailFooter_(systemUrl) {
  systemUrl = systemUrl || SYSTEM_URL;
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
function _caseInfoTable_(audit) {
  const s2Url = audit['第2階段連結'];
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

  return `
<table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin-top:16px;border:1px solid #e2e8f0;">
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
 * Stage 1 HTML 模板：S2 已上傳，通知承辦人可先行線上修改
 */
function buildStage1EmailHtml_(audit) {
  return _emailHeader_(
    '改善單已上傳，可先行線上修改',
    `${audit['主辦部門']} ｜ ${audit['工程簡稱']}`,
    '#0f766e'
  ) + `
<!-- Body -->
<tr><td style="padding:32px 36px;">
  <p style="font-size:15px;color:#334155;line-height:1.7;margin:0 0 16px;">
    您好，<strong>${audit['承辦人姓名'] || '承辦人'}</strong> 您好，
  </p>
  <p style="font-size:15px;color:#334155;line-height:1.7;margin:0 0 20px;">
    系統已偵測到以下案件的<strong>「S2 原始改善單」已上傳</strong>，
    承辦人員可<strong style="color:#0f766e;">先行至系統線上修改</strong>相關內容，
    紙本文件可於後續補送。
  </p>

  ${_caseInfoTable_(audit)}

  <!-- 綠色提示框 -->
  <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:20px;">
    <tr><td style="background:#ecfdf5;border-left:4px solid #10b981;border-radius:0 8px 8px 0;padding:14px 18px;">
      <p style="margin:0;font-size:14px;color:#065f46;line-height:1.7;">
        📋 <strong>後續步驟：</strong>請至系統確認改善內容正確後，
        再上傳「S3 工作隊核章版」完成程序。
      </p>
    </td></tr>
  </table>
</td></tr>
` + _emailFooter_();
}

/**
 * Stage 2 HTML 模板：距到期 3 天，提醒速辦 + 核章日期警示
 */
function buildStage2EmailHtml_(audit, daysLeft) {
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

  ${_caseInfoTable_(audit)}

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
` + _emailFooter_();
}

/**
 * Stage 3 HTML 模板：距到期 1 天，同時寄承辦人+課長，強調最後機會
 */
function buildStage3EmailHtml_(audit, daysLeft) {
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

  ${_caseInfoTable_(audit)}

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
` + _emailFooter_();
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
