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
        if(roleData.role !== 'Admin') throw new Error("無權限執行此操作。");
        result = deleteCase_(payload.id);
        break;

      case 'delete_deficiency':
        if(roleData.role !== 'Admin') throw new Error("無權限執行此操作。");
        result = deleteDeficiency_(payload.id);
        break;

      default:
        throw new Error("未知的 API 操作: " + action);
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

function uploadInspectionFile(fileInfo, caseId, stage, modifier) {
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
    const fileUrl = targetFolder.createFile(blob).getUrl();

    const currentStatus = auditData['辦理狀態'] || '';

    // 【智慧狀態更新】僅在狀態「前進」時更新
    let newStatus = '';
    const statusOrder = [STATUS.STAGE1, STATUS.STAGE2, STATUS.STAGE3, STATUS.STAGE4];
    const currentIndex = statusOrder.indexOf(currentStatus);

    if (stage === 'stage2' && currentIndex < 1) newStatus = STATUS.STAGE2;
    else if (stage === 'stage3' && currentIndex < 2) newStatus = STATUS.STAGE3;
    else if (stage === 'stage4') {
        newStatus = STATUS.STAGE4;
        const closeDateCol = headers.indexOf('結案日期');
        if (closeDateCol > -1) sheet.getRange(rowIdx, closeDateCol + 1).setValue(new Date());
    }

    if (newStatus !== '') {
        sheet.getRange(rowIdx, headers.indexOf('辦理狀態') + 1).setValue(newStatus);
    }
    
    // 將連結存回主清單對應欄位
    let urlColNum = -1;
    if (stage === 'stage2') urlColNum = headers.indexOf('第2階段連結') + 1;
    else if (stage === 'stage3') urlColNum = headers.indexOf('第3階段連結') + 1;
    else if (stage === 'stage4') urlColNum = headers.indexOf('第4階段連結') + 1;
    
    if (urlColNum > 0) {
        sheet.getRange(rowIdx, urlColNum).setValue(fileUrl);
    } else {
        // 若欄位不存在，則動態新增（通常發生在剛升級系統時）
        const newHeader = stage === 'stage2' ? '第2階段連結' : (stage === 'stage3' ? '第3階段連結' : '第4階段連結');
        const lastCol = sheet.getLastColumn();
        sheet.getRange(1, lastCol + 1).setValue(newHeader);
        sheet.getRange(rowIdx, lastCol + 1).setValue(fileUrl);
        // 更新 headers 變數供後續 log 使用
        headers.push(newHeader);
    }

    sheet.getRange(rowIdx, headers.indexOf('修改人員') + 1).setValue(userEmail);
    
    let stageDisplay = stage === 'report' ? '完成報告檔案' : '階段' + stage.replace(/\D/g, '') + '檔案';
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
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  return data.map(function (row) {
    return { serial: row[0], abbr: row[1], name: row[2], contractor: row[3], department: row[4] };
  }).filter(function (p) { return p.abbr; });
}

function addProject_(p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_PROJECT_DB);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PROJECT_DB);
    sheet.appendRow(['流水號', '工程簡稱', '工程名稱', '承攬商', '主辦部門']);
  }
  const serial = 'P' + new Date().getTime();
  sheet.appendRow([serial, p.abbr, p.name, p.contractor, p.department]);
  return { success: true, projects: getProjectOptions_() };
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

function deleteCase_(id) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_AUDIT_LIST);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
  const rowIdx = data.indexOf(id);
  if (rowIdx > -1) {
    sheet.deleteRow(rowIdx + 2);
    return { success: true, records: getAuditRecords_() };
  }
  throw new Error("找不到該案件: " + id);
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

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
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
    case 'stage2': prefix = '【原始改善單】'; break;
    case 'stage3': prefix = '【工作隊核章】'; break;
    case 'stage4': prefix = '【結案】'; break;
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
