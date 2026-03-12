/**
 * 前端核心邏輯與狀態管理 v5.2
 */

const GOOGLE_CLIENT_ID = window.ENV && window.ENV.GOOGLE_CLIENT_ID ? window.ENV.GOOGLE_CLIENT_ID : "791038911460-8tfq97vhrvr4iaq5r3s1ti1abfpuddd9.apps.googleusercontent.com";

const app = {
    state: {
        user: null,
        cases: [],
        projects: [],
        users: [],
        currentView: 'cases', // cases, users
        viewMode: localStorage.getItem('viewMode') || 'grid', // grid, list
        quickFilter: 'all',
        theme: localStorage.getItem('theme') || 'light'
    },

    /** ======================== 初始化與身分驗證 ======================== */
    initAuth: () => {
        app.setTheme(app.state.theme);
        app.toggleViewMode(app.state.viewMode);
        
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
            app.state.user = { email: res.data.email, role: res.data.role, department: res.data.department };
            app.state.cases = res.data.cases || [];
            app.state.projects = res.data.projects || [];

            document.getElementById('userNameDisplay').innerText = `${app.state.user.email}`;
            document.getElementById('logoutBtn').classList.remove('hidden');

            app.applyRoleRestrictions();
            app.extractDepartments();
            app.updateStats();
            app.renderView();
            
            if (app.state.user.role === 'Admin') app.fetchUsers();
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
        app.renderView();
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

    toggleViewMode: (mode) => {
        app.state.viewMode = mode;
        localStorage.setItem('viewMode', mode);
        
        const gridBtn = document.getElementById('viewGridBtn');
        const listBtn = document.getElementById('viewListBtn');
        const gridView = document.getElementById('viewCasesGrid');
        const listView = document.getElementById('viewCasesList');

        if (mode === 'grid') {
            gridBtn?.classList.add('active');
            listBtn?.classList.remove('active');
            gridView?.classList.remove('hidden');
            listView?.classList.add('hidden');
        } else {
            gridBtn?.classList.remove('active');
            listBtn?.classList.add('active');
            gridView?.classList.add('hidden');
            listView?.classList.remove('hidden');
        }
        app.renderView();
    },

    toggleView: (view) => {
        app.state.currentView = view;
        const vCasesGrid = document.getElementById('viewCasesGrid');
        const vCasesList = document.getElementById('viewCasesList');
        const vUsers = document.getElementById('viewUsers');
        const btnUsers = document.getElementById('btnAdminUsers');
        const btnBack = document.getElementById('btnBackToCases');

        if (view === 'users') {
            vCasesGrid?.classList.add('hidden');
            vCasesList?.classList.add('hidden');
            vUsers?.classList.remove('hidden');
            btnUsers?.classList.add('hidden');
            btnBack?.classList.remove('hidden');
            app.renderUsers();
        } else {
            vUsers?.classList.add('hidden');
            btnUsers?.classList.remove('hidden');
            btnBack?.classList.add('hidden');
            app.toggleViewMode(app.state.viewMode);
        }
    },

    /** ======================== 畫面渲染 ======================== */
    extractDepartments: () => {
        const depts = new Set(app.state.cases.map(c => c['主辦部門']).filter(Boolean));
        const select = document.getElementById('filterDepartment');
        if (!select) return;
        select.innerHTML = '<option value="">全部部門</option>';
        depts.forEach(d => {
            select.innerHTML += `<option value="${d}">${d}</option>`;
        });
    },

    setQuickFilter: (type) => {
        app.state.quickFilter = type;
        if (app.state.currentView !== 'cases') app.toggleView('cases');
        app.renderView();
    },

    renderView: () => {
        if (!app.state.user) {
            const emptyMsg = `<tr><td colspan="7" style="text-align:center; padding:40px; color:var(--text-muted);">👋 請先登入 Google 帳號</td></tr>`;
            const gridBody = document.getElementById('viewCasesGrid');
            const listBody = document.getElementById('caseListBody');
            if (gridBody) gridBody.innerHTML = `<div style="grid-column: 1/-1; text-align:center; padding:100px; color:var(--text-muted);">👋 請先登入 Google 帳號</div>`;
            if (listBody) listBody.innerHTML = emptyMsg;
            return;
        }

        const deptFilter = document.getElementById('filterDepartment')?.value || '';
        const statusFilter = document.getElementById('filterStatus')?.value || '';
        const today = new Date();
        today.setHours(0,0,0,0);

        const filtered = app.state.cases.filter(c => {
            if (deptFilter && c['主辦部門'] !== deptFilter) return false;
            if (statusFilter && c['辦理狀態'] !== statusFilter) return false;
            const isClosed = c['辦理狀態'] === '第4階段-已結案';
            const isOverdue = !isClosed && (new Date(c['最晚應核章日期']) < today);

            if (app.state.quickFilter === 'active' && isClosed) return false;
            if (app.state.quickFilter === 'closed' && !isClosed) return false;
            if (app.state.quickFilter === 'overdue' && !isOverdue) return false;
            return true;
        });

        if (app.state.viewMode === 'grid') {
            app.renderGrid(filtered);
        } else {
            app.renderList(filtered);
        }
    },

    renderGrid: (cases) => {
        const container = document.getElementById('viewCasesGrid');
        if (!container) return;
        container.innerHTML = '';

        if (cases.length === 0) {
            container.innerHTML = `<div style="grid-column: 1/-1; text-align:center; padding:100px; color:var(--text-muted);">☕ 沒有找到符合條件的案件</div>`;
            return;
        }

        cases.forEach(c => {
            const today = new Date();
            today.setHours(0,0,0,0);
            const dueDate = new Date(c['最晚應核章日期']);
            const isClosed = c['辦理狀態'] === '第4階段-已結案';
            const isOverdue = !isClosed && (dueDate < today);

            const card = document.createElement('div');
            card.className = 'case-card';
            
            // 模擬附件狀態 (實際應由 api 提供，此處先以辦理狀態模擬)
            const hasS2 = c['辦理狀態'] !== '第1階段-已登錄';
            const hasS3 = c['辦理狀態'].includes('第3階段') || isClosed;
            const hasS4 = isClosed;

            card.innerHTML = `
                <div class="card-header">
                    <h4>${c['工程簡稱']}</h4>
                    <span class="badge ${isOverdue ? 'warning' : (isClosed ? 'success' : 'badge-status')}">${c['辦理狀態']}</span>
                </div>
                <div class="card-body">
                    <div class="info-row"><i class="fas fa-building"></i> ${c['主辦部門']}</div>
                    <div class="info-row"><i class="fas fa-hard-hat"></i> ${c['承攬商']}</div>
                    <div class="info-row"><i class="fas fa-calendar-alt"></i> 查核：${c['查核日期']}</div>
                    <div class="info-row" style="${isOverdue ? 'color:var(--warning); font-weight:700;' : ''}">
                        <i class="fas fa-clock"></i> 限辦：${c['最晚應核章日期']}
                    </div>
                    
                    <div class="attachment-list">
                        <div class="attach-item ${hasS2 ? 'done' : ''}">
                            <span>📄 原始改善單</span>
                            <span class="attach-status">${hasS2 ? '<i class="fas fa-check-circle success"></i>' : '<i class="far fa-circle empty"></i>'}</span>
                        </div>
                        <div class="attach-item ${hasS3 ? 'done' : ''}">
                            <span>👤 工作隊核章版</span>
                            <span class="attach-status">${hasS3 ? '<i class="fas fa-check-circle success"></i>' : '<i class="far fa-circle empty"></i>'}</span>
                        </div>
                        <div class="attach-item ${hasS4 ? 'done' : ''}">
                            <span>✅ 最終結案版</span>
                            <span class="attach-status">${hasS4 ? '<i class="fas fa-check-circle success"></i>' : '<i class="far fa-circle empty"></i>'}</span>
                        </div>
                    </div>
                </div>
                <div class="card-footer">
                    <button class="btn btn-primary" onclick="app.openManage('${c.id}')">管理案件</button>
                    <button class="btn btn-outline" onclick="app.viewHistory('${c.id}')"><i class="fas fa-history"></i></button>
                </div>
            `;
            container.appendChild(card);
        });
    },

    renderList: (cases) => {
        const tbody = document.getElementById('caseListBody');
        if (!tbody) return;
        tbody.innerHTML = '';

        if (cases.length === 0) {
            tbody.innerHTML = `<tr><td colspan="7" style="text-align:center; padding:40px; color:var(--text-muted);">☕ 沒有找到符合條件的案件</td></tr>`;
            return;
        }

        cases.forEach(c => {
            const today = new Date();
            today.setHours(0,0,0,0);
            const dueDate = new Date(c['最晚應核章日期']);
            const isClosed = c['辦理狀態'] === '第4階段-已結案';
            const isOverdue = !isClosed && (dueDate < today);

            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><b>${c['工程簡稱']}</b></td>
                <td>${c['承攬商']}</td>
                <td>${c['主辦部門']}</td>
                <td>${c['查核日期']}</td>
                <td style="${isOverdue ? 'color:var(--warning); font-weight:bold;' : ''}">${c['最晚應核章日期']}</td>
                <td><span class="badge ${isOverdue ? 'warning' : (isClosed ? 'success' : 'badge-status')}">${c['辦理狀態']}</span></td>
                <td><button class="btn btn-outline" style="padding:4px 12px;" onclick="app.openManage('${c.id}')">管理</button></td>
            `;
            tbody.appendChild(tr);
        });
    },

    updateStats: () => {
        if (!app.state.cases) return;
        const today = new Date(); today.setHours(0,0,0,0);
        const total = app.state.cases.length;
        const active = app.state.cases.filter(c => c['辦理狀態'] !== '第4階段-已結案').length;
        const closed = app.state.cases.filter(c => c['辦理狀態'] === '第4階段-已結案').length;
        const overdue = app.state.cases.filter(c => c['辦理狀態'] !== '第4階段-已結案' && new Date(c['最晚應核章日期']) < today).length;

        document.getElementById('stat-total').innerText = total;
        document.getElementById('stat-active').innerText = active;
        document.getElementById('stat-closed').innerText = closed;
        document.getElementById('stat-overdue').innerText = overdue;
    },

    renderUsers: () => {
        const tbody = document.getElementById('userListBody');
        if (!tbody) return;
        tbody.innerHTML = app.state.users.length ? '' : `<tr><td colspan="5" style="text-align:center; padding:20px;">載入中...</td></tr>`;
        app.state.users.forEach(u => {
            tbody.innerHTML += `
                <tr>
                    <td>${u.name}</td>
                    <td>${u.email}</td>
                    <td><span class="badge badge-status">${u.role}</span></td>
                    <td>${u.department}</td>
                    <td>${u.active ? '🟢 啟用中' : '🔴 停權'}</td>
                </tr>
            `;
        });
    },

    /** 基本功能與 Modal (保留原有邏輯但優化視覺) */
    showLoading: (show) => { const l = document.getElementById('loading'); if(l) show ? l.classList.remove('hidden') : l.classList.add('hidden'); },
    
    openModal: (title, html) => {
        document.getElementById('modalTitle').innerText = title;
        document.getElementById('modalBody').innerHTML = html;
        document.getElementById('modalOverlay').classList.remove('hidden');
    },
    
    closeModal: () => { document.getElementById('modalOverlay').classList.add('hidden'); },

    openNewCaseModal: () => {
        let options = app.state.projects.map(p => `<option value="${p.abbr}">${p.abbr} - ${p.contractor}</option>`).join('');
        app.openModal('登錄查核案件', `
            <div style="display:flex; flex-direction:column; gap:20px;">
                <div><label style="font-weight:700; display:block; margin-bottom:8px;">選擇工程項：</label><select id="newProj" style="width:100%; padding:10px; border-radius:8px; border:1px solid var(--border); background:var(--bg-input); color:var(--text-main);">${options}</select></div>
                <div><label style="font-weight:700; display:block; margin-bottom:8px;">查核日期：</label><input type="date" id="newDate" value="${new Date().toISOString().split('T')[0]}" style="width:100%; padding:10px; border-radius:8px; border:1px solid var(--border); background:var(--bg-input); color:var(--text-main);"></div>
                <button class="btn btn-primary" onclick="app.submitNewCase()" style="justify-content:center; padding:14px;">確認登錄</button>
            </div>
        `);
    },

    submitNewCase: async () => {
        const pAbbr = document.getElementById('newProj').value;
        const date = document.getElementById('newDate').value;
        const pInfo = app.state.projects.find(p => p.abbr === pAbbr);
        try {
            const res = await api.createCase({ ...pInfo, auditDate: date, inspector: app.state.user.email });
            app.state.cases = res.records;
            app.updateStats(); app.renderView(); app.closeModal();
            alert("✅ 登錄成功");
        } catch(e) { alert("❌ 失敗: " + e.message); }
    },

    fetchUsers: async () => { try { const res = await api.getUsers(); app.state.users = res.data; if(app.state.currentView === 'users') app.renderUsers(); } catch(e){} },

    // ... 其他業務函數 (openManage, submitFile 等) 直接延用先前實作，視覺已由全域 CSS 控制
    openManage: (id) => {
        const c = app.state.cases.find(x => x.id === id);
        if (!c) return;
        let content = `<div style="margin-bottom:20px; padding:15px; background:rgba(0,120,255,0.05); border-radius:12px; font-weight:700;">狀態：${c['辦理狀態']}</div>`;
        const isSafety = (app.state.user.role === 'Admin' || app.state.user.role === 'SafetyUploader');

        if (c['辦理狀態'] === '第1階段-已登錄' && isSafety) {
            content += app.getUploadSection(id, 'stage2', '🔴 上傳原始改善單 (PDF/Image)');
        } else if (c['辦理狀態'] === '第2階段-改善單已上傳' && isSafety) {
            content += app.getUploadSection(id, 'stage3', '🟡 上傳工作隊核章版 (PDF)');
        } else if (c['辦理狀態'] === '第3階段-工作隊版已處理' && isSafety) {
            content += app.getUploadSection(id, 'stage4', '🟢 上傳結案完成版 (PDF)');
        }
        app.openModal(`案件管理: ${c['工程簡稱']}`, content);
    },

    getUploadSection: (id, stage, label) => `
        <div style="background:var(--bg-card); padding:16px; border-radius:12px; border:1px solid var(--border);">
            <p style="margin-top:0; font-weight:800;">${label}</p>
            <input type="file" id="file_${stage}" style="margin-bottom:12px; color:var(--text-main);" />
            <button class="btn btn-primary" style="width:100%; justify-content:center;" onclick="app.submitFile('${id}', '${stage}')">確認上傳</button>
        </div>
    `,

    submitFile: async (id, stage) => {
        const input = document.getElementById(`file_${stage}`);
        if(!input.files.length) return alert("請先選擇檔案");
        const file = input.files[0];
        try {
            const base64 = await app.fileToBase64(file);
            await api.uploadFile(id, stage, base64, file.name, app.state.user.email);
            const res = await api.init();
            app.state.cases = res.data.cases;
            app.updateStats(); app.renderView(); app.closeModal();
            alert("✅ 上傳成功");
        } catch(e) { alert("❌ 失敗: " + e.message); }
    },

    fileToBase64: (file) => new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.readAsDataURL(file);
        reader.onload = () => resolve(reader.result.split(',')[1]);
        reader.onerror = error => reject(error);
    }),

    viewHistory: async (id) => {
        try {
            const res = await api.getHistory(id);
            let html = res.data.map(r => `
                <div style="margin-bottom:12px; padding:12px; background:var(--bg-card); border-radius:10px; border:1px solid var(--border); font-size:0.8rem;">
                    <div style="color:var(--text-muted); margin-bottom:4px;">${r.timestamp}</div>
                    <div style="font-weight:700;">${r.description}</div>
                    <a href="${r.fileUrl}" target="_blank" style="color:var(--primary); text-decoration:none;">📎 下載附件</a>
                </div>
            `).join('') || '<p>尚無歷史紀錄</p>';
            app.openModal('檔案歷史紀錄', html + `<button class="btn btn-primary" style="width:100%; margin-top:10px;" onclick="app.closeModal()">關閉</button>`);
        } catch(e) { alert("無法獲取歷史: " + e.message); }
    }
};

window.onload = () => app.initAuth();
