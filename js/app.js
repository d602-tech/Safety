/**
 * 前端核心邏輯與狀態管理
 */

// ⚠️ 從環境變數讀取 (GitHub Actions 注入)，或是 fallback 回本地開發用的設定
const GOOGLE_CLIENT_ID = window.ENV && window.ENV.GOOGLE_CLIENT_ID ? window.ENV.GOOGLE_CLIENT_ID : "791038911460-8tfq97vhrvr4iaq5r3s1ti1abfpuddd9.apps.googleusercontent.com";

const app = {
    state: {
        user: null, // { email, role, department }
        cases: [],
        projects: [],
        departments: [],
        quickFilter: 'all' // all, active, overdue, closed
    },

    /** ======================== Google 身分驗證 ======================== */
    initAuth: () => {
        if (!window.google) {
            console.error("Google Identity 載入失敗");
            return;
        }
        google.accounts.id.initialize({
            client_id: GOOGLE_CLIENT_ID,
            callback: app.handleCredentialResponse
        });
        google.accounts.id.renderButton(
            document.getElementById("googleLoginBtn"),
            { theme: "outline", size: "large" }
        );
        // 若有舊 Token (如 localStorage)，可以看需求實作自動登入
    },

    handleCredentialResponse: async (response) => {
        const idToken = response.credential;
        api.setToken(idToken); // 將 JWT 儲存到 API 模組中
        document.getElementById('userNameDisplay').innerText = "驗證中...";
        document.getElementById('googleLoginBtn').classList.add('hidden');
        app.showLoading(true);

        try {
            // 向 GAS 後端傳遞 Token，進行解析、驗證信箱與比對權限
            const res = await api.init();

            // 根據 GAS 的回傳結果設定本地狀態
            app.state.user = {
                email: res.data.email,
                role: res.data.role,
                department: res.data.department
            };
            app.state.cases = res.data.cases || [];
            app.state.projects = res.data.projects || [];

            document.getElementById('userNameDisplay').innerText = `${app.state.user.email} (${app.state.user.role})`;
            document.getElementById('logoutBtn').classList.remove('hidden');

            app.applyRoleRestrictions();
            app.extractDepartments();
            app.updateStats();
            app.renderTable();
        } catch (e) {
            alert("登入失敗或您的帳號未啟用權限：\n" + e.message);
            app.logout();
        } finally {
            app.showLoading(false);
        }
    },

    logout: () => {
        api.clearToken();
        app.state.user = null;
        app.state.cases = [];
        document.getElementById('userNameDisplay').innerText = `尚未登入`;
        document.getElementById('logoutBtn').classList.add('hidden');
        document.getElementById('googleLoginBtn').classList.remove('hidden');
        app.applyRoleRestrictions();
        app.renderTable();
    },

    applyRoleRestrictions: () => {
        const btnNew = document.getElementById('btnNewCase');
        const btnRemind = document.getElementById('btnRemind');
        btnNew.classList.add('hidden');
        btnRemind.classList.add('hidden');

        if (!app.state.user) return;

        if (app.state.user.role === 'Admin' || app.state.user.role === 'SafetyUploader') {
            btnNew.classList.remove('hidden');
        }
        if (app.state.user.role === 'Admin') {
            btnRemind.classList.remove('hidden');
        }
    },

    /** ======================== 畫面渲染 ======================== */
    extractDepartments: () => {
        const depts = new Set(app.state.cases.map(c => c['主辦部門']).filter(Boolean));
        const select = document.getElementById('filterDepartment');
        select.innerHTML = '<option value="">所有主辦部門</option>';
        depts.forEach(d => {
            select.innerHTML += `<option value="${d}">${d}</option>`;
        });
    },

    clearFilters: () => {
        document.getElementById('filterDepartment').value = '';
        document.getElementById('filterStatus').value = '';
        app.setQuickFilter('all');
    },

    setQuickFilter: (type) => {
        app.state.quickFilter = type;
        app.renderTable();
    },

    renderTable: () => {
        const tbody = document.getElementById('caseListBody');
        tbody.innerHTML = '';

        if (!app.state.user) {
            tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;">請先完成 Google 登入</td></tr>`;
            return;
        }

        const deptFilter = document.getElementById('filterDepartment').value;
        const statusFilter = document.getElementById('filterStatus').value;
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        // 篩選邏輯
        const filteredCases = app.state.cases.filter(c => {
            if (deptFilter && c['主辦部門'] !== deptFilter) return false;
            if (statusFilter && c['辦理狀態'] !== statusFilter) return false;

            const dueDate = new Date(c['最晚應核章日期']);
            const isClosed = c['辦理狀態'] === '第4階段-已結案';

            if (app.state.quickFilter === 'active' && isClosed) return false;
            if (app.state.quickFilter === 'closed' && !isClosed) return false;
            if (app.state.quickFilter === 'overdue' && (isClosed || dueDate >= today)) return false;

            return true;
        });

        if (filteredCases.length === 0) {
            tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;">無符合條件的案件</td></tr>`;
            return;
        }

        filteredCases.forEach(c => {
            const dueDate = new Date(c['最晚應核章日期']);
            const isOverdue = c['辦理狀態'] !== '第4階段-已結案' && dueDate < today;
            const tr = document.createElement('tr');

            tr.innerHTML = `
                <td>${c['工程簡稱']}</td>
                <td>${c['承攬商']}</td>
                <td>${c['主辦部門']}</td>
                <td>${c['查核日期']}</td>
                <td style="${isOverdue ? 'color: var(--warning); font-weight:bold;' : ''}">${c['最晚應核章日期']}</td>
                <td><span class="badge ${isOverdue ? 'warning' : ''}">${c['辦理狀態']}</span></td>
                <td><button class="btn btn-outline" onclick="app.openManage('${c.id}')">管理</button></td>
            `;
            tbody.appendChild(tr);
        });
    },

    updateStats: () => {
        if (!app.state.cases) return;
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        document.getElementById('stat-total').innerText = app.state.cases.length;
        document.getElementById('stat-active').innerText = app.state.cases.filter(c => c['辦理狀態'] !== '第4階段-已結案').length;
        document.getElementById('stat-closed').innerText = app.state.cases.filter(c => c['辦理狀態'] === '第4階段-已結案').length;
        document.getElementById('stat-overdue').innerText = app.state.cases.filter(c => c['辦理狀態'] !== '第4階段-已結案' && new Date(c['最晚應核章日期']) < today).length;
    },

    /** ======================== Modal 與業務操作 ======================== */
    showLoading: (show) => {
        const l = document.getElementById('loading');
        if (show) l.classList.remove('hidden'); else l.classList.add('hidden');
    },

    openModal: (title, htmlContent) => {
        document.getElementById('modalTitle').innerText = title;
        document.getElementById('modalBody').innerHTML = htmlContent;
        document.getElementById('modalLoading').classList.add('hidden');
        document.getElementById('modalOverlay').classList.remove('hidden');
    },
    closeModal: () => {
        document.getElementById('modalOverlay').classList.add('hidden');
    },

    // 建立案件
    openNewCaseModal: () => {
        let projOptions = '<option value="">請選擇工程專案...</option>';
        app.state.projects.forEach(p => {
            projOptions += `<option value="${p.abbr}">${p.abbr} - ${p.contractor} (${p.department})</option>`;
        });

        app.openModal('登錄新案件', `
            <div style="display:flex; flex-direction:column; gap:12px;">
                <label>選擇工程：</label>
                <select id="newProj">${projOptions}</select>
                
                <label>查核日期：</label>
                <input type="date" id="newDate" value="${new Date().toISOString().split('T')[0]}">
                
                <button class="btn btn-primary" onclick="app.submitNewCase()" style="margin-top:10px;">確認登錄</button>
            </div>
        `);
    },

    submitNewCase: async () => {
        const projAbbr = document.getElementById('newProj').value;
        const date = document.getElementById('newDate').value;
        if (!projAbbr || !date) return alert("請填寫完整資訊");

        const projInfo = app.state.projects.find(p => p.abbr === projAbbr);

        document.getElementById('modalLoading').classList.remove('hidden');
        try {
            const payload = {
                name: projInfo.name,
                abbr: projInfo.abbr,
                contractor: projInfo.contractor,
                department: projInfo.department,
                auditDate: date,
                inspector: app.state.user.email,
                modifier: app.state.user.email
            };
            const result = await api.createCase(payload);
            app.state.cases = result.records;
            app.updateStats();
            app.renderTable();
            app.closeModal();
            alert("登錄成功！");
        } catch (e) {
            alert("登錄失敗: " + e.message);
        } finally {
            document.getElementById('modalLoading').classList.add('hidden');
        }
    },

    // 案件管理與上傳
    openManage: (id) => {
        const c = app.state.cases.find(x => x.id === id);
        if (!c) return;

        let content = `<div style="margin-bottom:16px;"><strong>狀態：</strong>${c['辦理狀態']}</div>`;
        content += `<button class="btn btn-outline" style="width:100%; margin-bottom:16px;" onclick="app.viewHistory('${id}')">📄 查看檔案歷史紀錄</button>`;

        const isSafetyUploader = (app.state.user.role === 'Admin' || app.state.user.role === 'SafetyUploader');

        if (c['辦理狀態'] === '第1階段-已登錄' && isSafetyUploader) {
            content += app.getUploadSection(id, 'stage2', '🔴 上傳 Stage2: 原始改善單 (PDF/Image)', 'image/*,application/pdf');
        }
        else if (c['辦理狀態'] === '第2階段-改善單已上傳' && isSafetyUploader) {
            content += app.getUploadSection(id, 'stage3', '🟡 上傳 Stage3: 工作隊版改善單 (PDF)', 'application/pdf');
            content += `
                <hr style="margin: 16px 0; border-color: var(--border);">
                <p style="font-size:0.9rem;">或工作隊無缺失可直接略過：</p>
                <input type="text" id="skipReason" placeholder="請輸入略過理由..." style="width:100%; box-sizing:border-box; margin-bottom:8px;">
                <button class="btn btn-warning" style="width:100%;" onclick="app.submitSkip('${id}')">略過 Stage3 上傳</button>
            `;
        }
        else if (c['辦理狀態'] === '第3階段-工作隊版已處理' && isSafetyUploader) {
            content += app.getUploadSection(id, 'stage4', '🟢 上傳 Stage4: 結案核章版 (PDF) 並結案', 'application/pdf');
        }

        // 獨立的完成報告上傳 (Stage3 以後開放)
        const canUploadReport = (c['辦理狀態'] === '第3階段-工作隊版已處理' || c['辦理狀態'] === '第4階段-已結案');
        if (canUploadReport) {
            content += `<hr style="margin: 16px 0; border-color: var(--border);">`;
            content += app.getUploadSection(id, 'report', '📄 補充上傳：完成報告 (不改變當前階段)', 'application/pdf');
        }

        app.openModal(`案件管理: ${c['工程簡稱']}`, content);
    },

    getUploadSection: (id, stage, label, accept) => {
        return `
            <div style="background:rgba(0,0,0,0.2); padding:12px; border-radius:6px; margin-bottom:8px;">
                <p style="margin-top:0; margin-bottom:8px; font-weight:500;">${label}</p>
                <input type="file" id="file_${stage}" accept="${accept}" style="width:100%; margin-bottom:8px;" />
                <button class="btn btn-primary" style="width:100%;" onclick="app.submitFile('${id}', '${stage}')">確認上傳</button>
            </div>
        `;
    },

    /** 檔案轉 Base64 */
    fileToBase64: (file) => new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => resolve(reader.result.split(',')[1]); // 移除 "data:image/png;base64," 前綴
        reader.onerror = error => reject(error);
    }),

    submitFile: async (id, stage) => {
        const fileInput = document.getElementById(`file_${stage}`);
        if (!fileInput.files.length) return alert("請先選擇檔案！");
        const file = fileInput.files[0];

        // 限制單檔大小 (例如 5MB)，超過以 GAS 可能會超時拋錯
        if (file.size > 5 * 1024 * 1024) return alert("檔案過大！請限制在 5MB 以內。");

        document.getElementById('modalLoading').classList.remove('hidden');
        try {
            const base64Data = await app.fileToBase64(file);
            const result = await api.uploadFile(id, stage, base64Data, file.name, app.state.user.email);
            app.state.cases = result.records || app.state.cases; // 若為完成報告可能沒回傳新 records，需處理
            // 重新初始化刷新完整狀態較安全
            const initRes = await api.init();
            app.state.cases = initRes.data.cases;

            app.updateStats();
            app.renderTable();
            app.closeModal();
            alert("上傳成功！");
        } catch (e) {
            alert("上傳失敗: " + e.message);
        } finally {
            document.getElementById('modalLoading').classList.add('hidden');
        }
    },

    submitSkip: async (id) => {
        const reason = document.getElementById('skipReason').value;
        if (!reason) return alert("請填寫略過理由！");

        document.getElementById('modalLoading').classList.remove('hidden');
        try {
            await api.skipStage3(id, reason, app.state.user.email);
            const initRes = await api.init(); // 刷新狀態
            app.state.cases = initRes.data.cases;
            app.updateStats();
            app.renderTable();
            app.closeModal();
            alert("已略過 Stage3。");
        } catch (e) {
            alert("處理失敗: " + e.message);
        } finally {
            document.getElementById('modalLoading').classList.add('hidden');
        }
    },

    viewHistory: async (id) => {
        document.getElementById('modalLoading').classList.remove('hidden');
        try {
            const res = await api.getHistory(id);
            const records = res.data;
            let html = `<ul>`;
            if (records.length === 0) html += `<li>尚無檔案記錄。</li>`;

            records.forEach(r => {
                html += `
                    <li style="margin-bottom:8px; padding-bottom:8px; border-bottom:1px solid var(--border);">
                        <strong>[${r.timestamp}]</strong> ${r.description} <br>
                        上傳者: ${r.modifier} <br>
                        <a href="${r.fileUrl}" target="_blank" style="color:var(--primary); text-decoration:none;">📎 下載 ${r.fileName}</a>
                    </li>
                `;
            });
            html += `</ul><button class="btn btn-outline" style="width:100%" onclick="app.openManage('${id}')">返回管理</button>`;

            document.getElementById('modalTitle').innerText = `檔案歷史紀錄`;
            document.getElementById('modalBody').innerHTML = html;
        } catch (e) {
            alert("無法獲取歷史紀錄：" + e.message);
        } finally {
            document.getElementById('modalLoading').classList.add('hidden');
        }
    },

    // 稽催測試
    triggerManualRemind: async () => {
        if (confirm("確定要觸發手動稽催通知嗎？系統將由後端直接抓取名單夾帶原版改善單並發信。")) {
            app.showLoading(true);
            try {
                const res = await api.manualRemind();
                alert(res.message || "發信成功");
            } catch (e) {
                alert("發信失敗: " + e.message);
            } finally {
                app.showLoading(false);
            }
        }
    }
};

// ======================== 載入入口點 ========================
window.onload = () => {
    app.initAuth();
};
