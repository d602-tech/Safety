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
        calDate: new Date(),
        searchKeyword: ''
    },

    /** ======================== 全域通知與顯示 ======================== */
    showToast: (msg, type = 'success') => {
        const container = document.getElementById('toastContainer');
        if (!container) return;
        const toast = document.createElement('div');
        toast.className = `toast toast-${type}`;
        toast.innerHTML = `<i class="fas fa-${type === 'success' ? 'check-circle' : 'exclamation-circle'}"></i> <span>${msg}</span>`;
        container.appendChild(toast);
        setTimeout(() => {
            toast.classList.add('toast-out');
            setTimeout(() => toast.remove(), 300);
        }, 3000);
    },

    previewImage: (url) => {
        const lb = document.getElementById('lightbox');
        const img = document.getElementById('lightboxImg');
        if (lb && img) {
            img.src = url;
            lb.classList.remove('hidden');
        }
    },

    setModalLoading: (show) => {
        const modalLoading = document.getElementById('modalLoading');
        const submitBtns = document.querySelectorAll('.modal .btn-primary');
        if (modalLoading) show ? modalLoading.classList.remove('hidden') : modalLoading.classList.add('hidden');
        submitBtns.forEach(btn => btn.disabled = show);
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

        // 未登入狀態下先抓取預設公開資料
        if (!app.state.user) {
            app.fetchPublicData();
        }
    },

    fetchPublicData: async () => {
        app.showLoading(true);
        try {
            const res = await api.getPublicCases();
            app.state.cases = res.data.cases || [];
            app.state.projects = res.data.projects || [];
            app.extractDepartments();
            app.updateStats();
            app.renderView();
        } catch (e) {
            console.error("無法載入預設資料", e);
        } finally {
            app.showLoading(false);
        }
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
            app.showToast("登入成功");
        } catch (e) {
            app.showToast("登入失敗或帳號未啟用", "error");
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
        app.fetchPublicData(); // 登出後切換回訪客資料
    },

    applyRoleRestrictions: () => {
        const btnNew = document.getElementById('btnNewCase');
        const btnRemind = document.getElementById('btnRemind');
        const btnProjMgmt = document.getElementById('btnProjMgmt');
        const btnAdminUsers = document.getElementById('btnAdminUsers');
        
        if (btnNew) btnNew.classList.add('hidden');
        if (btnRemind) btnRemind.classList.add('hidden');
        if (btnProjMgmt) btnProjMgmt.classList.add('hidden');
        if (btnAdminUsers) btnAdminUsers.classList.add('hidden');

        if (!app.state.user) return;

        if ((app.state.user.role === 'Admin' || app.state.user.role === 'SafetyUploader') && btnNew) {
            btnNew.classList.remove('hidden');
        }
        if (app.state.user.role === 'Admin') {
            if (btnRemind) btnRemind.classList.remove('hidden');
            if (btnProjMgmt) btnProjMgmt.classList.remove('hidden');
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

    /** 取得案件的檔案狀態 HTML (包含圖示與下載連結) */
    getFileStatusHtml: (c) => {
        const canAccess = app.state.user && (
            app.state.user.role === 'Admin' || 
            app.state.user.role === 'SafetyUploader' || 
            (app.state.user.role === 'DepartmentUploader' && c['主辦部門'] === app.state.user.department)
        );

        const stages = [
            { key: '第2階段連結', label: 'S2', class: 's2', icon: 'fa-file-pdf' },
            { key: '第3階段連結', label: 'S3', class: 's3', icon: 'fa-file-pdf' },
            { key: '第4階段連結', label: 'S4', class: 's4', icon: 'fa-file-check' }
        ];

        return `
            <div class="file-status-icons">
                ${stages.map(s => {
                    const url = c[s.key];
                    if (url && canAccess) {
                        return `<a href="${url}" target="_blank" class="file-icon uploaded ${s.class}" title="已上傳 ${s.label}"><i class="fas ${s.icon}"></i> ${s.label}</a>`;
                    } else if (url && !canAccess) {
                        // 有檔案但沒權限下載，顯示鎖定圖示
                        return `<div class="file-icon uploaded" title="已上傳，但您無權下載" style="color:var(--text-muted);"><i class="fas fa-lock"></i> ${s.label}</div>`;
                    } else {
                        // 尚未上傳
                        return `<div class="file-icon missing" title="${s.label} 尚未上傳"><i class="fas ${s.icon}"></i> ${s.label}</div>`;
                    }
                }).join('')}
            </div>
        `;
    },

    getFilteredCases: () => {
        const deptFilter = document.getElementById('filterDepartment')?.value || '';
        const statusFilter = document.getElementById('filterStatus')?.value || '';
        const keyword = document.getElementById('keywordSearch')?.value.toLowerCase().trim() || '';
        const todayStr = new Date().toISOString().split('T')[0];

        return app.state.cases.filter(c => {
            if (deptFilter && c['主辦部門'] !== deptFilter) return false;
            if (statusFilter && c['辦理狀態'] !== statusFilter) return false;
            
            if (keyword) {
                const searchStr = `${c['工程簡稱']} ${c['承攬商']} ${c['查核人員']} ${c['案件ID']}`.toLowerCase();
                if (!searchStr.includes(keyword)) return false;
            }

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
        if(!cases.length) return app.showToast("沒有資料可匯出", "error");

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

    downloadExample: () => {
        const headers = ["工程簡稱", "查核日期", "承攬商", "主辦部門", "查核地點", "查核人員", "缺失描述", "改善期限"];
        const exampleRow = ["範例工程", new Date().toISOString().split('T')[0], "範例承攬商", "工安課", "工地現場", "管理員", "範例缺失內容", new Date().toISOString().split('T')[0]];
        const csvContent = [headers, exampleRow].map(e => e.join(",")).join("\n");
        const blob = new Blob(["\ufeff" + csvContent], { type: 'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.setAttribute("href", url);
        link.setAttribute("download", "工安查核範例.csv");
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
            const isAdmin = app.state.user && app.state.user.role === 'Admin';
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
                    <button class="btn btn-outline" onclick="app.viewHistory('${c.id}')"><i class="fas fa-history"></i></button>
                    ${isAdmin ? `<button class="btn btn-outline" style="color:var(--warning); border-color:var(--warning);" onclick="app.deleteCase('${c.id}')"><i class="fas fa-trash"></i></button>` : ''}
                </div>
                <!-- 檔案狀態指示器 (原本為 Admin 專用改為公開狀態顯示) -->
                <div style="padding: 12px 24px; border-top: 1px dashed var(--border);">
                    ${app.getFileStatusHtml(c)}
                </div>
            `;
            container.appendChild(card);
        });
    },

    /** 2. 列表視圖 */
    renderList: (cases) => {
        const tbody = document.getElementById('caseListBody');
        if (!tbody) return;
        const isAdmin = app.state.user && app.state.user.role === 'Admin';
        tbody.innerHTML = cases.length ? '' : `<tr><td colspan="7" style="text-align:center; padding:40px;">☕ 尚無案件</td></tr>`;
        cases.forEach(c => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><b>${c['工程簡稱']}</b></td>
                <td>${c['承攬商']}</td>
                <td>${c['主辦部門']}</td>
                <td>${c['查核日期']}</td>
                <td>${c['最晚應核章日期']}</td>
                <td>${app.getFileStatusHtml(c)}</td>
                <td><span class="badge ${c['辦理狀態'] === '第4階段-已結案' ? 'success' : 'badge-status'}">${c['辦理狀態']}</span></td>
                <td>
                    <div style="display:flex; gap:5px;">
                        <button class="btn btn-outline" onclick="app.openManage('${c.id}')">管理</button>
                        ${isAdmin ? `<button class="btn btn-outline" style="color:var(--warning); border-color:var(--warning); padding:8px 12px;" onclick="app.deleteCase('${c.id}')"><i class="fas fa-trash"></i></button>` : ''}
                    </div>
                </td>
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
        const isAdmin = app.state.user && app.state.user.role === 'Admin';
        tbody.innerHTML = app.state.deficiencies.length ? '' : `<tr><td colspan="6" style="text-align:center;">尚無缺失紀錄</td></tr>`;
        app.state.deficiencies.forEach(d => {
            tbody.innerHTML += `
                <tr>
                    <td>${d.abbr}</td>
                    <td>${d.content}</td>
                    <td>${d.department}</td>
                    <td>${d.deadline}</td>
                    <td><span class="badge ${d.status==='已改善'?'success':'warning'}">${d.status}</span></td>
                    <td>
                        <div style="display:flex; gap:5px;">
                            <button class="btn btn-outline" onclick="app.openEditDef('${d.id}')">編輯</button>
                            ${isAdmin ? `<button class="btn btn-outline" style="color:var(--warning); border-color:var(--warning); padding:8px 12px;" onclick="app.deleteDeficiency('${d.id}')"><i class="fas fa-trash"></i></button>` : ''}
                        </div>
                    </td>
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
        const content = document.getElementById('defContent').value.trim();
        const deadline = document.getElementById('defDeadline').value;
        const department = document.getElementById('defDept').value.trim();

        if (!content || !deadline || !department) return app.showToast("請填寫所有欄位", "error");

        app.setModalLoading(true);
        const p = { caseId, abbr, content, deadline, department, status: '待改善' };
        try { 
            await api.updateDeficiency(p); 
            app.fetchDeficiencies(); 
            app.closeModal(); 
            app.showToast("✅ 已新增缺失"); 
        } catch(e) { 
            app.showToast(e.message, "error"); 
        } finally {
            app.setModalLoading(false);
        }
    },

    /** 5. 工程管理 */
    renderProjects: () => {
        const tbody = document.getElementById('projListBody');
        if (!tbody) return;
        const isAdmin = app.state.user && app.state.user.role === 'Admin';
        tbody.innerHTML = app.state.projects.length ? '' : `<tr><td colspan="6" style="text-align:center;">尚無工程項目</td></tr>`;
        app.state.projects.forEach(p => {
            tbody.innerHTML += `
                <tr>
                    <td>${p.serial}</td>
                    <td><b>${p.abbr}</b></td>
                    <td>${p.name}</td>
                    <td>${p.contractor}</td>
                    <td>${p.department}</td>
                    <td>${isAdmin ? `<button class="btn btn-outline" style="color:var(--warning); border-color:var(--warning);" onclick="app.deleteProject('${p.serial}')"><i class="fas fa-trash"></i></button>` : ''}</td>
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
        const abbr = document.getElementById('pAbbr').value.trim();
        const name = document.getElementById('pName').value.trim();
        const contractor = document.getElementById('pContractor').value.trim();
        const department = document.getElementById('pDept').value.trim();

        if (!abbr || !name || !contractor || !department) return app.showToast("請填寫完整資訊", "error");

        app.setModalLoading(true);
        const p = { abbr, name, contractor, department };
        try { 
            const res = await api.addProject(p); 
            app.state.projects = res.projects; 
            app.renderProjects(); 
            app.closeModal(); 
            app.showToast("工程新增成功"); 
        } catch(e){ 
            app.showToast(e.message, "error"); 
        } finally {
            app.setModalLoading(false);
        }
    },
    deleteProject: async (serial) => {
       if(!confirm("確定要刪除此工程項目？")) return;
       const btn = event.currentTarget;
       if (btn) btn.disabled = true;
       try { 
           const res = await api.deleteProject(serial); 
           app.state.projects = res.projects; 
           app.renderProjects(); 
           app.showToast("已刪除"); 
       } catch(e){ 
           app.showToast(e.message, "error"); 
       } finally {
           if (btn) btn.disabled = false;
       }
    },

    /** 獲取使用者權限清單 */
    fetchUsers: async () => { try { const res = await api.getUsers(); app.state.users = res.data; if(app.state.currentView === 'users') app.renderUsers(); } catch(e){} },
    renderUsers: () => {
        const tbody = document.getElementById('userListBody');
        const container = document.getElementById('viewUsers');
        if (!tbody || !container) return;
        
        // Ensure Add User button exists
        if (!document.getElementById('btnAddUser')) {
            const btn = document.createElement('button');
            btn.id = 'btnAddUser';
            btn.className = 'btn btn-primary';
            btn.style.margin = '0 15px 15px';
            btn.innerHTML = '<i class="fas fa-plus"></i> 新增帳號';
            btn.onclick = () => app.openUserModal();
            container.insertBefore(btn, container.querySelector('.case-table'));
        }

        const validUsers = app.state.users.filter(u => u.email && u.email.trim() !== '');
        tbody.innerHTML = validUsers.length ? '' : `<tr><td colspan="6" style="text-align:center; padding:20px;">尚無權限資料</td></tr>`;
        validUsers.forEach(u => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${u.name}</td>
                <td>${u.email}</td>
                <td><span class="badge badge-status">${u.role}</span></td>
                <td>${u.department}</td>
                <td>${u.active ? '🟢 啟用' : '🔴 停用'}</td>
                <td><button class="btn btn-outline" style="padding:4px 10px; font-size:0.8rem;" onclick='app.openUserModal(${JSON.stringify(u)})'>編輯</button></td>
            `;
            tbody.appendChild(tr);
        });
    },

    openUserModal: (user = null) => {
        const title = user ? '修改權限' : '新增使用者帳號';
        const html = `
            <div style="display:flex; flex-direction:column; gap:15px;">
                <div><label>電子信箱：</label><input type="email" id="uEmail" value="${user ? user.email : ''}" ${user ? 'readonly style="background:var(--bg-input);"' : ''}></div>
                <div><label>姓名：</label><input type="text" id="uName" value="${user ? user.name : ''}"></div>
                <div><label>授權角色：</label>
                    <select id="uRole">
                        <option value="Admin" ${user && user.role === 'Admin' ? 'selected' : ''}>Admin (管理者)</option>
                        <option value="SafetyUploader" ${user && user.role === 'SafetyUploader' ? 'selected' : ''}>SafetyUploader (工安稽核)</option>
                        <option value="DepartmentUploader" ${user && user.role === 'DepartmentUploader' ? 'selected' : ''}>DepartmentUploader (部門填報)</option>
                    </select>
                </div>
                <div><label>所屬部門：</label><input type="text" id="uDept" value="${user ? user.department : ''}" placeholder="例如：工安組、工務段"></div>
                ${user ? `<div><label>帳號狀態：</label>
                    <select id="uActive">
                        <option value="true" ${user.active ? 'selected' : ''}>啟用</option>
                        <option value="false" ${!user.active ? 'selected' : ''}>停用</option>
                    </select>
                </div>` : ''}
                <button class="btn btn-primary" style="margin-top:10px; justify-content:center;" onclick="app.submitUser()">確認儲存</button>
            </div>
        `;
        app.openModal(title, html);
    },

    submitUser: async () => {
        const email = document.getElementById('uEmail').value.trim();
        const name = document.getElementById('uName').value.trim();
        const role = document.getElementById('uRole').value;
        const department = document.getElementById('uDept').value.trim();
        const activeElem = document.getElementById('uActive');
        const active = activeElem ? activeElem.value === 'true' : true;

        if (!email || !name || !department) return app.showToast("請填寫所有欄位", "error");
        
        app.setModalLoading(true);
        try {
            await api.saveUser({ email, name, role, department, active });
            app.showToast("使用者設定已更新");
            app.fetchUsers();
            app.closeModal();
        } catch (e) {
            app.showToast(e.message, "error");
        } finally {
            app.setModalLoading(false);
        }
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
    openManage: async (id) => {
        const c = app.state.cases.find(item => item.id == id);
        if (!c) return;
        
        // 流程說明區塊
        let content = `
            <div style="margin-bottom:20px; padding:15px; background:rgba(99,102,241,0.05); border-radius:14px; font-size:0.85rem; line-height:1.6; border:1px solid rgba(99,102,241,0.1);">
                <div style="font-weight:800; color:var(--primary); margin-bottom:8px;"><i class="fas fa-info-circle"></i> 案件辦理流程說明</div>
                <div style="display:grid; grid-template-columns:auto 1fr; gap:5px 12px;">
                    <b style="color:var(--warning);">S2:</b> <span>原始工安紀錄及改善清單 (查核人員上傳)</span>
                    <b style="color:var(--info);">S3:</b> <span>部門改善核章版 (受查部門上傳)</span>
                    <b style="color:var(--success);">S4:</b> <span>結案完成版 (工安組上傳)</span>
                </div>
            </div>
            <div style="margin-bottom:20px; padding:15px; background:rgba(0,120,255,0.05); border-radius:12px; font-weight:700;">
                當前狀態：<span style="color:var(--primary);">${c['辦理狀態']}</span>
            </div>
        `;

        const isSafety = (app.state.user.role === 'Admin' || app.state.user.role === 'SafetyUploader');
        const isDeptOwner = (app.state.user.role === 'DepartmentUploader' && c['主辦部門'] === app.state.user.department);
        const isAdmin = app.state.user.role === 'Admin';

        if (isAdmin) {
            // 管理者顯示所有階段上傳 (顏色管理)
            content += `<div style="display:flex; flex-direction:column; gap:16px;">
                ${app.getUploadSection(id, 'stage2', '🔴 補傳/更換：原始改善單 (S2)', 'var(--warning)', '請上傳 Word 檔，方便改善部門修改')}
                ${app.getUploadSection(id, 'stage3', '🟡 補傳/更換：工作隊核章版 (S3)', '#fbbf24')}
                ${app.getUploadSection(id, 'stage4', '🟢 補傳/更換：結案完成版 (S4)', 'var(--success)')}
            </div>`;
        } else {
            // 一般使用者依狀態顯示
            if (c['辦理狀態'] === '第1階段-已登錄' && isSafety) {
                content += app.getUploadSection(id, 'stage2', '🔴 上傳原始改善單', 'var(--warning)');
            } else if (c['辦理狀態'] === '第2階段-改善單已上傳' && (isSafety || isDeptOwner)) {
                if (isDeptOwner) {
                    content += `<div style="margin-bottom:15px; padding:12px; background:rgba(16,185,129,0.1); border-radius:10px; font-size:0.85rem; color:var(--success);">
                        <i class="fas fa-info-circle"></i> 您可以先點擊下方「歷史紀錄」下載原始改善單，核章後再於此處上傳。
                    </div>`;
                }
                content += app.getUploadSection(id, 'stage3', '🟡 上傳工作隊核章版', '#fbbf24');
            } else if (c['辦理狀態'] === '第3階段-工作隊版已處理' && isSafety) {
                content += app.getUploadSection(id, 'stage4', '🟢 上傳結案完成版', 'var(--success)');
            }
        }

        // 管理者快速下載區 (優化顯示)
        if (isAdmin && (c['第2階段連結'] || c['第3階段連結'] || c['第4階段連結'])) {
            content += `<div style="margin-top:20px; padding:15px; background:rgba(0,0,0,0.03); border-radius:12px; border:1px solid var(--border);">
                <div style="font-weight:800; margin-bottom:10px;"><i class="fas fa-folder-open"></i> 已上傳檔案存檔</div>
                <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
                    ${c['第2階段連結'] ? `<a href="${c['第2階段連結']}" target="_blank" class="btn btn-outline" style="font-size:0.75rem;"><i class="fas fa-file-pdf"></i> S2 原始單</a>` : ''}
                    ${c['第3階段連結'] ? `<a href="${c['第3階段連結']}" target="_blank" class="btn btn-outline" style="font-size:0.75rem;"><i class="fas fa-file-pdf"></i> S3 核章版</a>` : ''}
                    ${c['第4階段連結'] ? `<a href="${c['第4階段連結']}" target="_blank" class="btn btn-outline" style="font-size:0.75rem;"><i class="fas fa-file-check"></i> S4 結案版</a>` : ''}
                </div>
            </div>`;
        }
        
        content += `<button class="btn btn-outline" style="width:100%; margin-top:15px;" onclick="app.viewHistory('${id}')"><i class="fas fa-history"></i> 查看完整歷史紀錄</button>`;
        
        app.openModal(`案件管理: ${c['工程簡稱']}`, content);
    },
    getUploadSection: (id, stage, label, color, note = '') => `
        <div style="background:var(--bg-card); padding:16px; border-radius:12px; border:1px solid ${color || 'var(--border)'}; border-left-width:5px;">
            <p style="margin:0 0 8px; font-weight:800; color:${color || 'inherit'};">${label}</p>
            ${note ? `<p style="margin:0 0 12px; font-size:0.75rem; color:var(--text-muted);"><i class="fas fa-exclamation-circle"></i> ${note}</p>` : ''}
            <input type="file" id="file_${stage}" style="margin-bottom:12px; color:var(--text-main); font-size:0.85rem;" />
            <button class="btn" style="width:100%; justify-content:center; background:${color || 'var(--primary)'}; color:white;" onclick="app.submitFile('${id}', '${stage}')">確認上傳</button>
        </div>
    `,
    submitFile: async (id, stage) => {
        const input = document.getElementById(`file_${stage}`);
        if(!input.files.length) return app.showToast("請先選擇檔案", "error");
        const file = input.files[0];
        app.setModalLoading(true);
        try {
            const base64 = await app.fileToBase64(file);
            await api.uploadFile(id, stage, base64, file.name, app.state.user.email);
            const res = await api.init();
            app.state.cases = res.data.cases;
            app.updateStats(); app.renderView(); app.closeModal();
            app.showToast("✅ 上傳成功");
        } catch(e) { 
            app.showToast(e.message, "error"); 
        } finally {
            app.setModalLoading(false);
        }
    },
    fileToBase64: (file) => new Promise((resolve) => {
        const reader = new FileReader(); reader.readAsDataURL(file);
        reader.onload = () => resolve(reader.result.split(',')[1]);
    }),
    viewHistory: async (id) => {
        try {
            const res = await api.getHistory(id);
            let html = res.data.map(r => {
                const isImg = /\.(jpg|jpeg|png|gif|webp)$/i.test(r.fileUrl) || /\.(jpg|jpeg|png|gif|webp)$/i.test(r.fileName);
                return `
                <div style="margin-bottom:12px; padding:12px; background:var(--bg-card); border-radius:10px; border:1px solid var(--border); font-size:0.8rem;">
                    <div style="color:var(--text-muted);">${r.timestamp}</div>
                    <div style="font-weight:700;">${r.description}</div>
                    <div style="margin-top:8px; display:flex; gap:10px;">
                        <a href="${r.fileUrl}" target="_blank" style="color:var(--primary); text-decoration:none;"><i class="fas fa-download"></i> 下載</a>
                        ${isImg ? `<a href="javascript:void(0)" onclick="app.previewImage('${r.fileUrl}')" style="color:var(--success); text-decoration:none;"><i class="fas fa-eye"></i> 預覽</a>` : ''}
                    </div>
                </div>`;
            }).join('') || '<p>尚無紀錄</p>';
            app.openModal('歷史紀錄', html);
        } catch(e) { app.showToast(e.message, "error"); }
    },
    openNewCaseModal: () => {
        let options = app.state.projects.map(p => `<option value="${p.abbr}">${p.serial} - ${p.abbr} - ${p.name}</option>`).join('');
        app.openModal('登錄查核案件', `<div style="display:flex;flex-direction:column;gap:15px;"><div>工程：<select id="newProj" style="width:100%">${options}</select></div><div>日期：<input type="date" id="newDate" value="${new Date().toISOString().split('T')[0]}"></div><button class="btn btn-primary" onclick="app.submitNewCase()">確認登錄</button></div>`);
    },
    submitNewCase: async () => {
        const pAbbr = document.getElementById('newProj').value; 
        const date = document.getElementById('newDate').value;
        if (!pAbbr || !date) return app.showToast("工程與日期為必填", "error");
        
        app.setModalLoading(true);
        const pInfo = app.state.projects.find(p => p.abbr === pAbbr);
        try { 
            const res = await api.createCase({ ...pInfo, auditDate: date }); 
            app.state.cases = res.records; 
            app.updateStats(); app.renderView(); app.closeModal(); 
            app.showToast("✅ 案件登錄成功"); 
        } catch(e) { 
            app.showToast(e.message, "error"); 
        } finally {
            app.setModalLoading(false);
        }
    },

    deleteCase: async (id) => {
        if (!confirm("確定要刪除此案件？此操作不可恢復！")) return;
        try {
            const res = await api.deleteCase(id);
            app.state.cases = res.records;
            app.updateStats();
            app.renderView();
            app.showToast("案件已刪除");
        } catch (e) {
            app.showToast(e.message, "error");
        }
    },

    deleteDeficiency: async (id) => {
        if (!confirm("確定要刪除此缺失紀錄？")) return;
        try {
            await api.deleteDeficiency(id);
            app.fetchDeficiencies();
            app.showToast("缺失紀錄已刪除");
        } catch (e) {
            app.showToast(e.message, "error");
        }
    },

    triggerManualRemind: async () => {
        const btn = document.getElementById('btnRemind');
        if (btn) btn.disabled = true;
        app.showToast("正在處理稽催信件...", "success");
        try {
            await api.manualRemind();
            app.showToast("✅ 稽催信件已發送");
        } catch(e) {
            app.showToast(e.message, "error");
        } finally {
            if (btn) btn.disabled = false;
        }
    }
};

window.onload = () => app.initAuth();
