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
        users: [], // 管理員查看的人員清單
        currentView: 'cases', // cases, users
        quickFilter: 'all', // all, active, overdue, closed
        theme: localStorage.getItem('theme') || 'dark'
    },

    /** ======================== 初始化與身分驗證 ======================== */
    initAuth: () => {
        app.setTheme(app.state.theme); // 套用存儲的主題
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
            { theme: app.state.theme === 'light' ? 'outline' : 'filled_blue', size: "large", shape: "pill" }
        );
    },

    handleCredentialResponse: async (response) => {
        const idToken = response.credential;
        api.setToken(idToken);
        document.getElementById('userNameDisplay').innerText = "驗證中...";
        document.getElementById('googleLoginBtn').classList.add('hidden');
        app.showLoading(true);

        try {
            const res = await api.init();
            app.state.user = {
                email: res.data.email,
                role: res.data.role,
                department: res.data.department
            };
            app.state.cases = res.data.cases || [];
            app.state.projects = res.data.projects || [];

            document.getElementById('userNameDisplay').innerText = `${app.state.user.email}`;
            document.getElementById('logoutBtn').classList.remove('hidden');

            app.applyRoleRestrictions();
            app.extractDepartments();
            app.updateStats();
            app.renderTable();
            
            // 如果是 Admin，預載入使用者清單
            if (app.state.user.role === 'Admin') {
                app.fetchUsers();
            }
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
        app.state.users = [];
        document.getElementById('userNameDisplay').innerText = `尚未登入`;
        document.getElementById('logoutBtn').classList.add('hidden');
        document.getElementById('googleLoginBtn').classList.remove('hidden');
        app.toggleView('cases');
        app.applyRoleRestrictions();
        app.renderTable();
    },

    applyRoleRestrictions: () => {
        const btnNew = document.getElementById('btnNewCase');
        const btnRemind = document.getElementById('btnRemind');
        const btnAdminUsers = document.getElementById('btnAdminUsers');
        
        if (btnNew) btnNew.classList.add('hidden');
        if (btnRemind) btnRemind.classList.add('hidden');
        if (btnAdminUsers) btnAdminUsers.classList.add('hidden');

        if (!app.state.user) return;

        if ((app.state.user.role === 'Admin' || app.state.user.role === 'SafetyUploader') && btnNew) {
            btnNew.classList.remove('hidden');
        }
        if (app.state.user.role === 'Admin') {
            if (btnRemind) btnRemind.classList.remove('hidden');
            if (btnAdminUsers) btnAdminUsers.classList.remove('hidden');
        }
    },

    /** ======================== UI 功能切換 ======================== */
    
    setTheme: (theme) => {
        app.state.theme = theme;
        document.documentElement.setAttribute('data-theme', theme);
        localStorage.setItem('theme', theme);
        const selector = document.getElementById('themeSelector');
        if (selector) selector.value = theme;
    },

    toggleInstructions: () => {
        const card = document.getElementById('instructionCard');
        if (card) card.classList.toggle('active');
    },

    toggleView: (view) => {
        app.state.currentView = view;
        const vCases = document.getElementById('viewCases');
        const vUsers = document.getElementById('viewUsers');
        const btnUsers = document.getElementById('btnAdminUsers');
        const btnBack = document.getElementById('btnBackToCases');

        if (view === 'users') {
            if (vCases) vCases.classList.add('hidden');
            if (vUsers) vUsers.classList.remove('hidden');
            if (btnUsers) btnUsers.classList.add('hidden');
            if (btnBack) btnBack.classList.remove('hidden');
            app.renderUsers();
        } else {
            if (vCases) vCases.classList.remove('hidden');
            if (vUsers) vUsers.classList.add('hidden');
            if (btnUsers) btnUsers.classList.remove('hidden');
            if (btnBack) btnBack.classList.add('hidden');
            app.renderTable();
        }
    },

    /** ======================== 畫面渲染 ======================== */
    extractDepartments: () => {
        const depts = new Set(app.state.cases.map(c => c['主辦部門']).filter(Boolean));
        const select = document.getElementById('filterDepartment');
        if (!select) return;
        select.innerHTML = '<option value="">所有主辦部門</option>';
        depts.forEach(d => {
            select.innerHTML += `<option value="${d}">${d}</option>`;
        });
    },

    clearFilters: () => {
        const dept = document.getElementById('filterDepartment');
        const status = document.getElementById('filterStatus');
        if (dept) dept.value = '';
        if (status) status.value = '';
        app.setQuickFilter('all');
    },

    setQuickFilter: (type) => {
        app.state.quickFilter = type;
        if (app.state.currentView !== 'cases') app.toggleView('cases');
        app.renderTable();
    },

    renderTable: () => {
        const tbody = document.getElementById('caseListBody');
        if (!tbody) return;
        tbody.innerHTML = '';

        if (!app.state.user) {
            tbody.innerHTML = `<tr><td colspan="7" style="text-align:center; padding:40px; color:var(--text-muted);"> 👋 歡迎使用！請先完成 Google 登入以查看案件。</td></tr>`;
            return;
        }

        const deptFilter = document.getElementById('filterDepartment') ? document.getElementById('filterDepartment').value : '';
        const statusFilter = document.getElementById('filterStatus') ? document.getElementById('filterStatus').value : '';
        const today = new Date();
        today.setHours(0, 0, 0, 0);

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
            tbody.innerHTML = `<tr><td colspan="7" style="text-align:center; padding:40px; color:var(--text-muted);"> 沒有找到符合條件的案件 ☕ </td></tr>`;
            return;
        }

        filteredCases.forEach(c => {
            const dueDate = new Date(c['最晚應核章日期']);
            const isOverdue = c['辦理狀態'] !== '第4階段-已結案' && dueDate < today;
            const tr = document.createElement('tr');

            tr.innerHTML = `
                <td><b style="color:var(--primary)">${c['工程簡稱']}</b></td>
                <td>${c['承攬商']}</td>
                <td>${c['主辦部門']}</td>
                <td>${c['查核日期']}</td>
                <td style="${isOverdue ? 'color: var(--warning); border-left: 3px solid var(--warning);' : ''}">${c['最晚應核章日期']}</td>
                <td><span class="badge ${isOverdue ? 'warning' : (c['辦理狀態'].includes('結案') ? 'success' : '')}">${c['辦理狀態']}</span></td>
                <td><button class="btn btn-outline" style="padding: 4px 12px; font-size:0.8rem;" onclick="app.openManage('${c.id}')">管理</button></td>
            `;
            tbody.appendChild(tr);
        });
    },

    updateStats: () => {
        if (!app.state.cases) return;
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        const total = document.getElementById('stat-total');
        const active = document.getElementById('stat-active');
        const closed = document.getElementById('stat-closed');
        const overdue = document.getElementById('stat-overdue');

        if (total) total.innerText = app.state.cases.length;
        if (active) active.innerText = app.state.cases.filter(c => c['辦理狀態'] !== '第4階段-已結案').length;
        if (closed) closed.innerText = app.state.cases.filter(c => c['辦理狀態'] === '第4階段-已結案').length;
        if (overdue) overdue.innerText = app.state.cases.filter(c => c['辦理狀態'] !== '第4階段-已結案' && new Date(c['最晚應核章日期']) < today).length;
    },

    /** ======================== 管理員功能 ======================== */
    fetchUsers: async () => {
        try {
            const res = await api.getUsers();
            app.state.users = res.data;
        } catch (e) {
            console.error("無法同步帳號清單", e);
        }
    },

    renderUsers: () => {
        const tbody = document.getElementById('userListBody');
        if (!tbody) return;
        tbody.innerHTML = '';
        if (app.state.users.length === 0) {
            tbody.innerHTML = `<tr><td colspan="5" style="text-align:center; padding:20px;">載入帳號中...</td></tr>`;
            return;
        }
        app.state.users.forEach(u => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${u.name}</td>
                <td>${u.email}</td>
                <td><span class="badge">${u.role}</span></td>
                <td>${u.department}</td>
                <td>${u.active ? '🟢 啟用中' : '🔴 停權'}</td>
            `;
            tbody.appendChild(tr);
        });
    },

    /** ======================== Modal 與業務操作 ======================== */
    showLoading: (show) => {
        const l = document.getElementById('loading');
        if (l) {
            if (show) l.classList.remove('hidden'); else l.classList.add('hidden');
        }
    },

    openModal: (title, htmlContent) => {
        const t = document.getElementById('modalTitle');
        const b = document.getElementById('modalBody');
        const l = document.getElementById('modalLoading');
        const o = document.getElementById('modalOverlay');

        if (t) t.innerText = title;
        if (b) b.innerHTML = htmlContent;
        if (l) l.classList.add('hidden');
        if (o) o.classList.remove('hidden');
    },
    closeModal: () => {
        const o = document.getElementById('modalOverlay');
        if (o) o.classList.add('hidden');
    },

    // 建立案件
    openNewCaseModal: () => {
        let projOptions = '<option value="">請選擇工程專案...</option>';
        app.state.projects.forEach(p => {
            projOptions += `<option value="${p.abbr}">${p.abbr} - ${p.contractor} (${p.department})</option>`;
        });

        app.openModal('登錄新案件', `
            <div style="display:flex; flex-direction:column; gap:16px;">
                <div class="form-item">
                    <label style="display:block; margin-bottom:8px; font-weight:600;">選擇工程：</label>
                    <select id="newProj" style="width:100%; border:1px solid var(--border); background:var(--bg-input); color:var(--text); padding:8px; border-radius:6px;">${projOptions}</select>
                </div>
                <div class="form-item">
                    <label style="display:block; margin-bottom:8px; font-weight:600;">查核日期：</label>
                    <input type="date" id="newDate" style="width:100%; border:1px solid var(--border); background:var(--bg-input); color:var(--text); padding:8px; border-radius:6px;" value="${new Date().toISOString().split('T')[0]}">
                </div>
                <button class="btn btn-primary" onclick="app.submitNewCase()" style="margin-top:10px; justify-content:center;">確認登錄</button>
            </div>
        `);
    },

    submitNewCase: async () => {
        const projAbbr = document.getElementById('newProj').value;
        const date = document.getElementById('newDate').value;
        if (!projAbbr || !date) return alert("請填寫完整資訊");

        const projInfo = app.state.projects.find(p => p.abbr === projAbbr);

        const l = document.getElementById('modalLoading');
        if (l) l.classList.remove('hidden');
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
            if (l) l.classList.add('hidden');
        }
    },

    // 案件管理與上傳
    openManage: (id) => {
        const c = app.state.cases.find(x => x.id === id);
        if (!c) return;

        let content = `<div style="margin-bottom:20px; padding:15px; background:rgba(0,0,100,0.05); border-radius:12px;"><strong>當前階段：</strong><span class="badge">${c['辦理狀態']}</span></div>`;
        content += `<button class="btn btn-outline" style="width:100%; margin-bottom:16px;" onclick="app.viewHistory('${id}')">📄 查看檔案歷史紀錄</button>`;

        const isSafetyUploader = (app.state.user.role === 'Admin' || app.state.user.role === 'SafetyUploader');

        if (c['辦理狀態'] === '第1階段-已登錄' && isSafetyUploader) {
            content += app.getUploadSection(id, 'stage2', '🔴 上傳 Stage2: 原始改善單 (PDF/Image)', 'image/*,application/pdf');
        }
        else if (c['辦理狀態'] === '第2階段-改善單已上傳' && isSafetyUploader) {
            content += app.getUploadSection(id, 'stage3', '🟡 上傳 Stage3: 工作隊版改善單 (PDF)', 'application/pdf');
            content += `
                <div style="margin-top: 24px; padding-top:16px; border-top: 1px dashed var(--border);">
                    <p style="font-size:0.9rem; margin-bottom:10px; color:var(--text-muted);">或工作隊無缺失可直接略過：</p>
                    <input type="text" id="skipReason" placeholder="請輸入略過理由..." style="width:100%; margin-bottom:8px; border:1px solid var(--border); background:var(--bg-card); color:var(--text); padding:8px; border-radius:6px;">
                    <button class="btn btn-warning" style="width:100%; background:var(--warning); color:white;" onclick="app.submitSkip('${id}')">略過 Stage3 上傳</button>
                </div>
            `;
        }
        else if (c['辦理狀態'] === '第3階段-工作隊版已處理' && isSafetyUploader) {
            content += app.getUploadSection(id, 'stage4', '🟢 上傳 Stage4: 結案核章版 (PDF) 並結案', 'application/pdf');
        }

        const canUploadReport = (c['辦理狀態'] === '第3階段-工作隊版已處理' || c['辦理狀態'] === '第4階段-已結案');
        if (canUploadReport) {
            content += `<hr style="margin: 20px 0; border:none; border-top:1px solid var(--border);">`;
            content += app.getUploadSection(id, 'report', '📄 補充上傳：完成報告 (不改變當前階段)', 'application/pdf');
        }

        app.openModal(`案件管理: ${c['工程簡稱']}`, content);
    },

    getUploadSection: (id, stage, label, accept) => {
        return `
            <div style="background:var(--bg-card); padding:16px; border-radius:12px; margin-bottom:12px; border:1px solid var(--border);">
                <p style="margin-top:0; margin-bottom:12px; font-weight:600; font-size:0.9rem;">${label}</p>
                <input type="file" id="file_${stage}" accept="${accept}" style="width:100%; margin-bottom:12px; font-size:0.8rem; color:var(--text);" />
                <button class="btn btn-primary" style="width:100%; justify-content:center;" onclick="app.submitFile('${id}', '${stage}')">確認上傳並存檔</button>
            </div>
        `;
    },

    fileToBase64: (file) => new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => resolve(reader.result.split(',')[1]);
        reader.onerror = error => reject(error);
    }),

    submitFile: async (id, stage) => {
        const fileInput = document.getElementById(`file_${stage}`);
        if (!fileInput || !fileInput.files.length) return alert("請先選擇檔案！");
        const file = fileInput.files[0];
        if (file.size > 5 * 1024 * 1024) return alert("檔案過大！請限制在 5MB 以內。");

        const l = document.getElementById('modalLoading');
        if (l) l.classList.remove('hidden');
        try {
            const base64Data = await app.fileToBase64(file);
            const result = await api.uploadFile(id, stage, base64Data, file.name, app.state.user.email);
            const initRes = await api.init();
            app.state.cases = initRes.data.cases;
            app.updateStats();
            app.renderTable();
            app.closeModal();
            alert("上傳成功！");
        } catch (e) {
            alert("上傳失敗: " + e.message);
        } finally {
            if (l) l.classList.add('hidden');
        }
    },

    submitSkip: async (id) => {
        const reasonInput = document.getElementById('skipReason');
        const reason = reasonInput ? reasonInput.value : '';
        if (!reason) return alert("請填寫略過理由！");

        const l = document.getElementById('modalLoading');
        if (l) l.classList.remove('hidden');
        try {
            await api.skipStage3(id, reason, app.state.user.email);
            const initRes = await api.init();
            app.state.cases = initRes.data.cases;
            app.updateStats();
            app.renderTable();
            app.closeModal();
            alert("已略過 Stage3。");
        } catch (e) {
            alert("處理失敗: " + e.message);
        } finally {
            if (l) l.classList.add('hidden');
        }
    },

    viewHistory: async (id) => {
        const l = document.getElementById('modalLoading');
        if (l) l.classList.remove('hidden');
        try {
            const res = await api.getHistory(id);
            const records = res.data;
            let html = `<div style="max-height:300px; overflow-y:auto; margin-bottom:20px;">`;
            if (records.length === 0) html += `<p style="text-align:center; color:var(--text-muted);">尚無檔案記錄。</p>`;

            records.forEach(r => {
                html += `
                    <div style="margin-bottom:12px; padding:12px; background:var(--bg-card); border-radius:10px; border:1px solid var(--border);">
                        <div style="font-size:0.75rem; color:var(--text-muted); margin-bottom:4px;">${r.timestamp} • ${r.modifier}</div>
                        <div style="font-weight:600; margin-bottom:8px;">${r.description}</div>
                        <a href="${r.fileUrl}" target="_blank" class="btn btn-outline" style="padding:4px 10px; font-size:0.8rem; text-decoration:none;">📎 下載 ${r.fileName}</a>
                    </div>
                `;
            });
            html += `</div><button class="btn btn-primary" style="width:100%; justify-content:center;" onclick="app.openManage('${id}')">返回管理</button>`;

            const t = document.getElementById('modalTitle');
            const b = document.getElementById('modalBody');
            if (t) t.innerText = `檔案歷史紀錄`;
            if (b) b.innerHTML = html;
        } catch (e) {
            alert("無法獲取歷史紀錄：" + e.message);
        } finally {
            if (l) l.classList.add('hidden');
        }
    },

    triggerManualRemind: async () => {
        if (confirm("確定要在後端觸發手動稽催通知嗎？")) {
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

window.onload = () => {
    app.initAuth();
};
