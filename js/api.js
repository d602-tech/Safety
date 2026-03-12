/**
 * 負責與 GAS 後端溝通的 API 模組
 */
const GAS_URL = 'https://script.google.com/macros/s/AKfycby_Your_Id_Here/exec'; // ⚠️ 請替換為您 GAS 部署的 Web App URL

let _authToken = null; // 存放 Google 登入後核發的 JWT ID Token

const api = {
    /** 設置驗證 Token */
    setToken: (token) => {
        _authToken = token;
    },

    /** 清除驗證 Token */
    clearToken: () => {
        _authToken = null;
    },

    /** 通用 Fetch 封裝，自動夾帶 Token */
    request: async (action, payload = {}) => {
        try {
            // 在每一包 payload 中強制塞入 token 交由 GAS 驗證
            if (_authToken) {
                payload.token = _authToken;
            }

            const response = await fetch(GAS_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' }, // 避免觸發 CORS 複雜預檢
                body: JSON.stringify({ action, payload })
            });
            
            const result = await response.json();
            
            // 處理 GAS 傳回的業務邏輯錯誤
            if (!result.success) throw new Error(result.message || '伺服器未知的錯誤');
            
            return result;
        } catch (error) {
            console.error(`API Error [${action}]:`, error);
            throw error;
        }
    },

    // 1. 初始化資料 (下拉選單, 權限比對, 案件清單)
    init: () => api.request('init'),

    // 2. 登錄新案件
    createCase: (data) => api.request('create_case', data),

    // 3. 上傳檔案 (支援 base64)
    uploadFile: (caseId, stage, fileBase64, fileName, modifierName) => 
        api.request('upload_file', { caseId, stage, fileBase64, fileName, modifier: modifierName }),

    // 4. 略過 Stage3
    skipStage3: (caseId, reason, modifierName) => 
        api.request('skip_stage3', { caseId, reason, modifier: modifierName }),

    // 5. 取得案件歷史記錄
    getHistory: (caseId) => api.request('get_history', { caseId }),

    // 6. 手動稽催觸發
    manualRemind: () => api.request('manual_remind')
};
