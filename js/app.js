/**
 * 前端核心邏輯與狀態管理 v5.3
 */

const GOOGLE_CLIENT_ID = window.ENV && window.ENV.GOOGLE_CLIENT_ID ? window.ENV.GOOGLE_CLIENT_ID : "791038911460-8tfq97vhrvr4iaq5r3s1ti1abfpuddd9.apps.googleusercontent.com";

const app = {
    state: {
        user: null,
        cases: [],
        projects: [],
        users: [],
        deficiencies: [],
        currentView: 'cases', // cases, users, deficiencies, projects
        viewMode: localStorage.getItem('viewMode') || 'grid', // grid, list, calendar
        quickFilter: 'all',
        theme: localStorage.getItem('theme') || 'light',
        calDate: new Date()
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
            app.fetchDeficiencies();
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
        
        const views = {
            grid: document.getElementById('viewCasesGrid'),
            list: document.getElementById('viewCasesList'),
            calendar: document.getElementById('viewCalendar')
        };
        const btns = {
            grid: document.getElementById('viewGridBtn'),
            list: document.getElementById('viewListBtn'),
            calendar: document.getElementById('viewCalBtn')
        };

        Object.keys(views).forEach(key => {
            if (key === mode) {
                views[key]?.classList.remove('hidden');
                btns[key]?.classList.add('active');
            } else {
                views[key]?.classList.add('hidden');
                btns[key]?.classList.remove('active');
            }
        });
        
        if (app.state.currentView === 'cases') app.renderView();
    },

    toggleView: (view) => {
        app.state.currentView = view;
        const mainViews = {
            cases: [document.getElementById('viewCasesGrid'), document.getElementById('viewCasesList'), document.getElementById('viewCalendar')],
            users: document.getElementById('viewUsers'),
            deficiencies: document.getElementById('viewDeficiencies'),
            projects: document.getElementById('viewProjects')
        };
        const btnUsers = document.getElementById('btnAdminUsers');
        const btnProj = document.getElementById('btnProjMgmt');
        const btnDef = document.getElementById('btnDefList');
        const btnBack = document.getElementById('btnBackToCases');

        // Hide all
        Object.values(mainViews).flat().forEach(v => v?.classList.add('hidden'));

        if (view === 'cases') {
            btnBack?.classList.add('hidden');
            btnUsers?.classList.remove('hidden');
            btnProj?.classList.remove('hidden');
            btnDef?.classList.remove('hidden');
            app.toggleViewMode(app.state.viewMode);
        } else {
            document.getElementById('view' + view.charAt(0).toUpperCase() + view.slice(1))?.classList.remove('hidden');
            btnBack?.classList.remove('hidden');
            btnUsers?.classList.add('hidden');
            btnProj?.classList.add('hidden');
            btnDef?.classList.add('hidden');
            if (view === 'users') app.renderUsers();
            if (view === 'deficiencies') app.renderDeficiencies();
            if (view === 'projects') app.renderProjects();
        }
    },

    /** ======================== 畫面渲染 ======================== */
    renderView: () => {
        if (!app.state.user) return;
        const filtered = app.getFilteredCases();
        if (app.state.viewMode === 'grid') app.renderGrid(filtered);
        else if (app.state.viewMode === 'list') app.renderList(filtered);
        else if (app.state.viewMode === 'calendar') app.renderCalendar(filtered);
    },

    getFilteredCases: () => {
        const deptFilter = document.getElementById('filterDepartment')?.value || '';
        const statusFilter = document.getElementById('filterStatus')?.value || '';
        const todayStr = new Date().toISOString().split('T')[0];

        return app.state.cases.filter(c => {
            if (deptFilter && c['主辦部門'] !== deptFilter) return false;
            if (statusFilter && c['辦理狀態'] !== statusFilter) return false;
            const isClosed = c['辦理狀態'] === '第4階段-已結案';
            const isOverdue = !isClosed && (c['最晚應核章日期'] < todayStr);

            if (app.state.quickFilter === 'active' && isClosed) return false;
            if (app.state.quickFilter === 'closed' && !isClosed) return false;
            if (app.state.quickFilter === 'overdue' && !isOverdue) return false;
            return true;
        });
    },

    exportToCSV: () => {
        const cases = app.getFilteredCases();
        if(!cases.length) return alert("沒有資料可匯出");

        const headers = ["查核日期", "查核地點(工程簡稱)", "完成日期(結案日期)", "查核人員"];
        const rows = cases.map(c => [
            c['查核日期'] || '',
            c['工程簡稱'] || '',
            c['結案日期'] || '進行中',
            c['查核人員'] || ''
        ]);

        const csvContent = [headers, ...rows].map(e => e.join(",")).join("\n");
        const blob = new Blob(["\ufeff" + csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.setAttribute("href", url);
        link.setAttribute("download", `工安查核清單匯出_${new Date().toISOString().split('T')[0]}.csv`);
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    },

    /** 1. 卡片視圖 */
    renderGrid: (cases) => {
        const container = document.getElementById('viewCasesGrid');
        if (!container) return;
        container.innerHTML = cases.length ? '' : `<div style="grid-column:1/-1; text-align:center; padding:100px;">☕ 尚無案件</div>`;
        cases.forEach(c => {
            const todayStr = new Date().toISOString().split('T')[0];
            const isClosed = c['辦理狀態'] === '第4階段-已結案';
            const isOverdue = !isClosed && (c['最晚應核章日期'] < todayStr);
            const card = document.createElement('div');
            card.className = 'case-card';
            card.innerHTML = `
                <div class="card-header">
                    <h4>${c['工程簡稱']}</h4>
                    <span class="badge ${isOverdue ? 'warning' : (isClosed ? 'success' : 'badge-status')}">${c['辦理狀態']}</span>
                </div>
                <div class="card-body">
                    <div class="info-row"><i class="fas fa-building"></i> ${c['主辦部門']}</div>
                    <div class="info-row"><i class="fas fa-hard-hat"></i> ${c['承攬商']}</div>
                    <div class="info-row"><i class="fas fa-calendar-alt"></i> 查核：${c['查核日期']}</div>
                    <div class="info-row" style="${isOverdue ? 'color:var(--warning);font-weight:700;' : ''}"><i class="fas fa-clock"></i> 限辦：${c['最晚應核章日期']}</div>
                </div>
                <div class="card-footer">
                    <button class="btn btn-primary" onclick="app.openManage('${c.id}')">管理</button>
                    <button class="btn btn-outline" onclick="app.openAddDefModal('${c.id}', '${c['工程簡稱']}', '${c['主辦部門']}')"><i class="fas fa-plus"></i> 缺失</button>
                    <button class="btn btn-outline" onclick="app.viewHistory('${c.id}')"><i class="fas fa-history"></i></button>
                </div>
            `;
            container.appendChild(card);
        });
    },

    /** 2. 列表視圖 */
    renderList: (cases) => {
        const tbody = document.getElementById('caseListBody');
        if (!tbody) return;
        tbody.innerHTML = cases.length ? '' : `<tr><td colspan="7" style="text-align:center; padding:40px;">☕ 尚無案件</td></tr>`;
        cases.forEach(c => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><b>${c['工程簡稱']}</b></td>
                <td>${c['承攬商']}</td>
                <td>${c['主辦部門']}</td>
                <td>${c['查核日期']}</td>
                <td>${c['最晚應核章日期']}</td>
                <td><span class="badge ${c['辦理狀態'] === '第4階段-已結案' ? 'success' : 'badge-status'}">${c['辦理狀態']}</span></td>
                <td><button class="btn btn-outline" onclick="app.openManage('${c.id}')">管理</button></td>
            `;
            tbody.appendChild(tr);
        });
    },

    /** 3. 日曆視圖 (Vanilla JS 實作) */
    renderCalendar: (cases) => {
        const container = document.getElementById('calendarContainer');
        if (!container) return;
        const cur = app.state.calDate;
        const year = cur.getFullYear(), month = cur.getMonth();
        const firstDay = new Date(year, month, 1).getDay();
        const daysInMonth = new Date(year, month + 1, 0).getDate();
        const monthName = new Intl.DateTimeFormat('zh-TW', {month:'long'}).format(cur);

        let html = `
            <div class="cal-nav">
                <button class="btn btn-outline" onclick="app.changeMonth(-1)"><i class="fas fa-chevron-left"></i></button>
                <h3 style="margin:0">${year}年 ${monthName}</h3>
                <button class="btn btn-outline" onclick="app.changeMonth(1)"><i class="fas fa-chevron-right"></i></button>
            </div>
            <div class="cal-grid">
                <div class="cal-header">週日</div><div class="cal-header">週一</div>
                <div class="cal-header">週二</div><div class="cal-header">週三</div>
                <div class="cal-header">週四</div><div class="cal-header">週五</div>
                <div class="cal-header">週六</div>
        `;

        for(let i=0; i<firstDay; i++) html += `<div class="cal-day other-month"></div>`;
        for(let d=1; d<=daysInMonth; d++) {
            const dStr = `${year}-${String(month+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
            const dayCases = cases.filter(c => c['查核日期'] === dStr);
            const isToday = dStr === new Date().toISOString().split('T')[0];
            html += `
                <div class="cal-day ${isToday?'today':''}">
                    <div class="cal-date">${d}</div>
                    ${dayCases.map(c => `<div class="cal-event ${c['辦理狀態']==='第4階段-已結案'?'success':''}" onclick="app.openManage('${c.id}')">${c['工程簡稱']}</div>`).join('')}
                </div>
            `;
        }
        html += `</div>`;
        container.innerHTML = html;
    },
    changeMonth: (offset) => { app.state.calDate.setMonth(app.state.calDate.getMonth() + offset); app.renderView(); },

    /** 4. 缺失清單 */
    fetchDeficiencies: async () => { try { const res = await api.getDeficiencies(); app.state.deficiencies = res.data; if(app.state.currentView === 'deficiencies') app.renderDeficiencies(); } catch(e){} },
    renderDeficiencies: () => {
        const tbody = document.getElementById('defListBody');
        if (!tbody) return;
        tbody.innerHTML = app.state.deficiencies.length ? '' : `<tr><td colspan="6" style="text-align:center;">尚無缺失紀錄</td></tr>`;
        app.state.deficiencies.forEach(d => {
            tbody.innerHTML += `
                <tr>
                    <td>${d.abbr}</td>
                    <td>${d.content}</td>
                    <td>${d.department}</td>
                    <td>${d.deadline}</td>
                    <td><span class="badge ${d.status==='已改善'?'success':'warning'}">${d.status}</span></td>
                    <td><button class="btn btn-outline" onclick="app.openEditDef('${d.id}')">編輯</button></td>
                </tr>
            `;
        });
    },
    openAddDefModal: (caseId, abbr, dept) => {
        app.openModal('新增缺失紀錄', `
            <div style="display:flex;flex-direction:column;gap:15px;">
                <div><label>工程：</label> <b>${abbr}</b></div>
                <div><label>缺失描述：</label><textarea id="defContent" style="width:100%;height:80px;"></textarea></div>
                <div><label>改善期限：</label><input type="date" id="defDeadline" value="${new Date().toISOString().split('T')[0]}"></div>
                <div><label>負責部門：</label><input type="text" id="defDept" value="${dept}"></div>
                <button class="btn btn-primary" onclick="app.submitDeficiency('${caseId}', '${abbr}')">確認存檔</button>
            </div>
        `);
    },
    submitDeficiency: async (caseId, abbr) => {
        const p = { caseId, abbr, content: document.getElementById('defContent').value, deadline: document.getElementById('defDeadline').value, department: document.getElementById('defDept').value, status: '待改善' };
        try { await api.updateDeficiency(p); app.fetchDeficiencies(); app.closeModal(); alert("✅ 已新增缺失"); } catch(e) { alert(e.message); }
    },

    /** 5. 工程管理 */
    renderProjects: () => {
        const tbody = document.getElementById('projListBody');
        if (!tbody) return;
        tbody.innerHTML = app.state.projects.length ? '' : `<tr><td colspan="6" style="text-align:center;">尚無工程項目</td></tr>`;
        app.state.projects.forEach(p => {
            tbody.innerHTML += `
                <tr>
                    <td>${p.serial}</td>
                    <td><b>${p.abbr}</b></td>
                    <td>${p.name}</td>
                    <td>${p.contractor}</td>
                    <td>${p.department}</td>
                    <td><button class="btn btn-outline" onclick="app.deleteProject('${p.serial}')"><i class="fas fa-trash"></i></button></td>
                </tr>
            `;
        });
    },
    openNewProjectModal: () => {
        app.openModal('新增工程項目', `
            <div style="display:flex;flex-direction:column;gap:15px;">
                <input type="text" id="pAbbr" placeholder="工程簡稱 (例如: 台中電纜)">
                <input type="text" id="pName" placeholder="工程全名">
                <input type="text" id="pContractor" placeholder="承攬商名稱">
                <input type="text" id="pDept" placeholder="主辦部門">
                <button class="btn btn-primary" onclick="app.submitNewProject()">確認新增</button>
            </div>
        `);
    },
    submitNewProject: async () => {
        const p = { abbr: document.getElementById('pAbbr').value, name: document.getElementById('pName').value, contractor: document.getElementById('pContractor').value, department: document.getElementById('pDept').value };
        try { const res = await api.addProject(p); app.state.projects = res.projects; app.renderProjects(); app.closeModal(); } catch(e){ alert(e.message); }
    },
    deleteProject: async (serial) => {
       if(!confirm("確定要刪除此工程項目？")) return;
       try { const res = await api.deleteProject(serial); app.state.projects = res.projects; app.renderProjects(); } catch(e){ alert(e.message); }
    },

    /** 獲取使用者權限清單 */
    fetchUsers: async () => { try { const res = await api.getUsers(); app.state.users = res.data; if(app.state.currentView === 'users') app.renderUsers(); } catch(e){} },
    renderUsers: () => {
        const tbody = document.getElementById('userListBody');
        if (!tbody) return;
        tbody.innerHTML = app.state.users.length ? '' : `<tr><td colspan="5" style="text-align:center; padding:20px;">載入中...</td></tr>`;
        app.state.users.forEach(u => {
            tbody.innerHTML += `<tr><td>${u.name}</td><td>${u.email}</td><td><span class="badge badge-status">${u.role}</span></td><td>${u.department}</td><td>${u.active ? '🟢' : '🔴'}</td></tr>`;
        });
    },

    /** 統計與基礎共用函數 */
    updateStats: () => {
        const todayStr = new Date().toISOString().split('T')[0];
        document.getElementById('stat-total').innerText = app.state.cases.length;
        document.getElementById('stat-active').innerText = app.state.cases.filter(c => c['辦理狀態'] !== '第4階段-已結案').length;
        document.getElementById('stat-closed').innerText = app.state.cases.filter(c => c['辦理狀態'] === '第4階段-已結案').length;
        document.getElementById('stat-overdue').innerText = app.state.cases.filter(c => c['辦理狀態'] !== '第4階段-已結案' && c['最晚應核章日期'] < todayStr).length;
    },
    extractDepartments: () => {
        const depts = new Set(app.state.cases.map(c => c['主辦部門']).filter(Boolean));
        const select = document.getElementById('filterDepartment');
        if (select) { select.innerHTML = '<option value="">全部部門</option>'; depts.forEach(d => select.innerHTML += `<option value="${d}">${d}</option>`); }
    },
    setQuickFilter: (type) => { app.state.quickFilter = type; if (app.state.currentView !== 'cases') app.toggleView('cases'); app.renderView(); },
    showLoading: (show) => { const l = document.getElementById('loading'); if(l) show ? l.classList.remove('hidden') : l.classList.add('hidden'); },
    openModal: (title, html) => { document.getElementById('modalTitle').innerText = title; document.getElementById('modalBody').innerHTML = html; document.getElementById('modalOverlay').classList.remove('hidden'); },
    closeModal: () => { document.getElementById('modalOverlay').classList.add('hidden'); },

    // 案件管理與上傳邏輯 (延用上一版本並微調)
    openManage: (id) => {
        const c = app.state.cases.find(x => x.id === id);
        if (!c) return;
        let content = `<div style="margin-bottom:20px; padding:15px; background:rgba(0,120,255,0.05); border-radius:12px; font-weight:700;">狀態：${c['辦理狀態']}</div>`;
        const isSafety = (app.state.user.role === 'Admin' || app.state.user.role === 'SafetyUploader');
        if (c['辦理狀態'] === '第1階段-已登錄' && isSafety) content += app.getUploadSection(id, 'stage2', '🔴 上傳原始改善單');
        else if (c['辦理狀態'] === '第2階段-改善單已上傳' && isSafety) content += app.getUploadSection(id, 'stage3', '🟡 上傳工作隊核章版');
        else if (c['辦理狀態'] === '第3階段-工作隊版已處理' && isSafety) content += app.getUploadSection(id, 'stage4', '🟢 上傳結案完成版');
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
        } catch(e) { alert(e.message); }
    },
    fileToBase64: (file) => new Promise((resolve) => {
        const reader = new FileReader(); reader.readAsDataURL(file);
        reader.onload = () => resolve(reader.result.split(',')[1]);
    }),
    viewHistory: async (id) => {
        try {
            const res = await api.getHistory(id);
            let html = res.data.map(r => `<div style="margin-bottom:12px; padding:12px; background:var(--bg-card); border-radius:10px; border:1px solid var(--border); font-size:0.8rem;"><div style="color:var(--text-muted);">${r.timestamp}</div><div style="font-weight:700;">${r.description}</div><a href="${r.fileUrl}" target="_blank" style="color:var(--primary);">📎 下載</a></div>`).join('') || '<p>尚無紀錄</p>';
            app.openModal('歷史紀錄', html);
        } catch(e) { alert(e.message); }
    },
    openNewCaseModal: () => {
        let options = app.state.projects.map(p => `<option value="${p.abbr}">${p.abbr} - ${p.contractor}</option>`).join('');
        app.openModal('登錄查核案件', `<div style="display:flex;flex-direction:column;gap:15px;"><div>工程：<select id="newProj" style="width:100%">${options}</select></div><div>日期：<input type="date" id="newDate" value="${new Date().toISOString().split('T')[0]}"></div><button class="btn btn-primary" onclick="app.submitNewCase()">確認登錄</button></div>`);
    },
    submitNewCase: async () => {
        const pAbbr = document.getElementById('newProj').value; const date = document.getElementById('newDate').value;
        const pInfo = app.state.projects.find(p => p.abbr === pAbbr);
        try { const res = await api.createCase({ ...pInfo, auditDate: date }); app.state.cases = res.records; app.updateStats(); app.renderView(); app.closeModal(); alert("✅ 成功"); } catch(e) { alert(e.message); }
    }
};

window.onload = () => app.initAuth();
