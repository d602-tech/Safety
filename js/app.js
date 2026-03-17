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
        searchKeyword: '',
        systemMode: localStorage.getItem('systemMode') || 'progress' // progress, tracking
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

    showGlobalProgress: (show, msg = "處理中...", percent = 0) => {
        const overlay = document.getElementById('globalProgress');
        const label = document.getElementById('gpLabel');
        const bar = document.getElementById('gpBar');
        const text = document.getElementById('gpText');
        if (!overlay) return;
        
        if (show) {
            label.innerText = msg;
            bar.style.width = `${percent}%`;
            text.innerText = `${Math.round(percent)}%`;
            overlay.classList.remove('hidden');
        } else {
            overlay.classList.add('hidden');
            // 重置
            bar.style.width = '0%';
            text.innerText = '0%';
        }
    },

    copyToClipboard: (text) => {
        navigator.clipboard.writeText(text).then(() => {
            app.showToast("📋 已複製到剪貼簿");
        }).catch(err => {
            console.error('無法複製文字: ', err);
            // Fallback
            const textArea = document.createElement("textarea");
            textArea.value = text;
            document.body.appendChild(textArea);
            textArea.select();
            try {
                document.execCommand('copy');
                app.showToast("📋 已複製到剪貼簿(Fallback)");
            } catch (err) {
                app.showToast("複製失敗", "error");
            }
            document.body.removeChild(textArea);
        });
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
        
        // 初始化模式按鈕狀態
        app.setSystemMode(app.state.systemMode);

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
            app.extractYears();
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
            app.extractYears();
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

    setSystemMode: (mode) => {
        app.state.systemMode = mode;
        localStorage.setItem('systemMode', mode);
        
        const btnReport = document.getElementById('modeReportBtn');
        const btnDef = document.getElementById('modeDefBtn');
        
        if (mode === 'progress') {
            btnReport?.classList.add('active');
            btnDef?.classList.remove('active');
        } else {
            btnReport?.classList.remove('active');
            btnDef?.classList.add('active');
        }
        
        app.renderView();
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

    /** 取得案件 S1-S4 進度條 HTML */
    getProgressHtml: (c) => {
        const s2 = !!c['第2階段連結'];
        const s3 = !!c['第3階段連結'];
        const s4e = !!c['第4階段連結-員工'];
        const s4c = !!c['第4階段連結-承攬商'];
        const isClosed = c['辦理狀態'] === '第4階段-已結案';

        return `
            <div class="progress-container">
                <div class="progress-label">
                    <span>階段進度</span>
                    <span>${isClosed ? '100%' : ((s4e && s4c) ? '80%' : (s4e || s4c ? '70%' : (s3 ? '50%' : (s2 ? '25%' : '0%'))))}</span>
                </div>
                <div class="progress-track">
                    <div class="progress-step step-s1 active" title="S1: 已登錄"></div>
                    <div class="progress-step step-s2 ${s2 || s3 || s4e || s4c ? 'active' : ''}" title="S2: 改善單"></div>
                    <div class="progress-step step-s3 ${s3 || s4e || s4c ? 'active' : ''}" title="S3: 核章版"></div>
                    <div class="progress-step step-s4 ${s4e || s4c || isClosed ? 'active' : ''}" title="S4: 結案 (員/承)"></div>
                </div>
            </div>
        `;
    },

    /** 取得案件的檔案狀態 HTML (包含圖示與下載連結) */
    getFileStatusHtml: (c) => {
        const canAccess = app.state.user && (
            app.state.user.role === 'Admin' || 
            app.state.user.role === 'SafetyUploader' || 
            (app.state.user.role === 'DepartmentUploader' && c['主辦部門'] === app.state.user.department)
        );

        const stages = [
            { key: '第2階段連結', label: 'S2', class: 's2', icon: 'fa-file-signature' },
            { key: '第3階段連結', label: 'S3', class: 's3', icon: 'fa-stamp' },
            { key: '第4階段連結-員工', label: 'S4員工', class: 's4', icon: 'fa-user-check' },
            { key: '第4階段連結-承攬商', label: 'S4承攬', class: 's4', icon: 'fa-building-circle-check' }
        ];

        return `
            <div class="file-status-icons">
                ${stages.map(s => {
                    const url = c[s.key];
                    if (url && canAccess) {
                        return `<a href="${url}" target="_blank" class="file-icon uploaded ${s.class}" title="已上傳 ${s.label}"><i class="fas ${s.icon}"></i> ${s.label}</a>`;
                    } else if (url && !canAccess) {
                        return `<div class="file-icon uploaded" title="已上傳，但您無權下載" style="color:var(--text-muted);"><i class="fas fa-lock"></i> ${s.label}</div>`;
                    } else {
                        return `<div class="file-icon missing" title="${s.label} 尚未上傳"><i class="fas ${s.icon}"></i> ${s.label}</div>`;
                    }
                }).join('')}
            </div>
        `;
    },

    getFilteredCases: () => {
        const yearFilter = document.getElementById('filterYear')?.value || '';
        const deptFilter = document.getElementById('filterDepartment')?.value || '';
        const statusFilter = document.getElementById('filterStatus')?.value || '';
        const keyword = document.getElementById('keywordSearch')?.value.toLowerCase().trim() || '';
        const todayStr = new Date().toISOString().split('T')[0];

        return app.state.cases.filter(c => {
            if (yearFilter && (!c['查核日期'] || !c['查核日期'].startsWith(yearFilter))) return false;
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
            const projInfo = app.state.projects.find(p => p.abbr === c['工程簡稱']);
            const snLabel = projInfo ? `${projInfo.serial} - ` : '';
            const hasFiles = !!(c['第2階段連結'] || c['第3階段連結'] || c['第4階段連結-員工'] || c['第4階段連結-承攬商']);
            const card = document.createElement('div');
            card.className = `case-card ${hasFiles ? 'has-files' : ''}`;
            card.setAttribute('data-dept', c['主辦部門'] || '');
            card.innerHTML = `
                <div class="card-header">
                    <h4 class="report-clickable" onclick="app.openManage('${c.id}')">${snLabel}${c['工程簡稱']}</h4>
                    <span class="badge ${isOverdue ? 'warning' : (isClosed ? 'success' : 'badge-status')}">${c['辦理狀態']}</span>
                </div>
                <div class="card-body">
                    <div class="info-row"><i class="fas fa-building"></i> ${c['主辦部門']}</div>
                    <div class="info-row"><i class="fas fa-hard-hat"></i> ${c['承攬商']}</div>
                    <div class="info-row"><i class="fas fa-calendar-alt"></i> 查核：${c['查核日期']}</div>
                    <div class="info-row" style="${isOverdue ? 'color:var(--warning);font-weight:700;' : ''}"><i class="fas fa-clock"></i> 限辦：${c['最晚應核章日期']}</div>
                    
                    ${app.state.systemMode === 'progress' ? app.getProgressHtml(c) : `
                        <div style="margin-top:10px; padding:10px; background:rgba(0,0,0,0.03); border-radius:10px; font-size:0.8rem;">
                            <i class="fas fa-exclamation-circle" style="color:var(--primary);"></i> 缺失數：${app.state.deficiencies.filter(d => d.caseId === c.id).length}
                        </div>
                    `}
                </div>
                <div class="card-footer">
                    <button class="btn btn-primary" onclick="app.openManage('${c.id}')">管理</button>
                    <button class="btn btn-outline" onclick="app.viewHistory('${c.id}')"><i class="fas fa-history"></i></button>
                    ${isAdmin ? `<button class="btn btn-outline" style="color:var(--warning); border-color:var(--warning);" onclick="app.deleteCase('${c.id}')"><i class="fas fa-trash"></i></button>` : ''}
                </div>
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
            const projInfo = app.state.projects.find(p => p.abbr === c['工程簡稱']);
            const snLabel = projInfo ? `${projInfo.serial} - ` : '';
            const hasFiles = !!(c['第2階段連結'] || c['第3階段連結'] || c['第4階段連結-員工'] || c['第4階段連結-承攬商']);
            const tr = document.createElement('tr');
            if (hasFiles) tr.classList.add('has-files');
            tr.innerHTML = `
                <td><b>${snLabel}${c['工程簡稱']}</b></td>
                <td>${c['承攬商']}</td>
                <td>${c['主辦部門']}</td>
                <td>${c['查核日期']}</td>
                <td>${c['最晚應核章日期']}</td>
                <td>
                    ${app.getFileStatusHtml(c)}
                    ${app.getProgressHtml(c)}
                </td>
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
                    ${dayCases.map(c => {
                        const projInfo = app.state.projects.find(p => p.abbr === c['工程簡稱']);
                        const snLabel = projInfo ? `${projInfo.serial} - ` : '';
                        return `<div class="cal-event ${c['辦理狀態']==='第4階段-已結案'?'success':''}" onclick="app.openManage('${c.id}')">${snLabel}${c['工程簡稱']}</div>`;
                    }).join('')}
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
                <div>
                    <label>缺失描述：</label>
                    <textarea id="defContent" style="width:100%;height:120px;" placeholder="輸入一項缺失。如有多項，請換行輸入（每一行將自動轉為一筆獨立缺失項目）。"></textarea>
                    <small style="color:var(--text-muted); display:block; margin-top:4px;"><i class="fas fa-info-circle"></i> 換行即為一項，支援批次輸入。</small>
                </div>
                <div><label>改善期限：</label><input type="date" id="defDeadline" value="${new Date().toISOString().split('T')[0]}"></div>
                <div><label>負責部門：</label><input type="text" id="defDept" value="${dept}"></div>
                <button class="btn btn-primary" onclick="app.submitDeficiency('${caseId}', '${abbr}')">確認存檔</button>
            </div>
        `);
    },
    submitDeficiency: async (caseId, abbr) => {
        const contentRaw = document.getElementById('defContent').value.trim();
        const deadline = document.getElementById('defDeadline').value;
        const department = document.getElementById('defDept').value.trim();

        if (!contentRaw || !deadline || !department) return app.showToast("請填寫所有欄位", "error");

        const lines = contentRaw.split('\n').map(l => l.trim()).filter(l => l !== '');
        
        app.setModalLoading(true);
        try { 
            if (lines.length > 1) {
                // Batch add
                const items = lines.map(line => ({
                    caseId, abbr, content: line, deadline, department, status: '待改善'
                }));
                await api.batchAddDeficiencies(items);
                app.showToast(`✅ 已批次新增 ${lines.length} 項缺失`);
            } else {
                // Single add
                const p = { caseId, abbr, content: lines[0], deadline, department, status: '待改善' };
                await api.updateDeficiency(p); 
                app.showToast("✅ 已新增缺失"); 
            }
            app.fetchDeficiencies(); 
            app.closeModal(); 
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
        if (select) { 
            const currentVal = select.value;
            select.innerHTML = '<option value="">全部工作隊</option>'; 
            depts.forEach(d => select.innerHTML += `<option value="${d}">${d}</option>`); 
            if (currentVal) select.value = currentVal;
        }
    },
    extractYears: () => {
        const years = new Set(app.state.cases.map(c => c['查核日期'] ? c['查核日期'].substring(0, 4) : null).filter(Boolean));
        const select = document.getElementById('filterYear');
        if (select) {
            const currentYear = new Date().getFullYear().toString();
            select.innerHTML = '<option value="">全部年度</option>';
            Array.from(years).sort((a,b) => b-a).forEach(y => {
                select.innerHTML += `<option value="${y}">${y} 年度</option>`;
            });
            // 預設為今年
            if (years.has(currentYear)) {
                select.value = currentYear;
            }
        }
    },
    setQuickFilter: (type) => { app.state.quickFilter = type; if (app.state.currentView !== 'cases') app.toggleView('cases'); app.renderView(); },
    showLoading: (show) => { const l = document.getElementById('loading'); if(l) show ? l.classList.remove('hidden') : l.classList.add('hidden'); },
    openModal: (title, html) => { document.getElementById('modalTitle').innerText = title; document.getElementById('modalBody').innerHTML = html; document.getElementById('modalOverlay').classList.remove('hidden'); },
    closeModal: () => { document.getElementById('modalOverlay').classList.add('hidden'); },

    // 案件管理與上傳邏輯 (延用上一版本並微調)
    openManage: async (id) => {
        const c = app.state.cases.find(item => item.id == id);
        if (!c) return;
        
        const isSafety = (app.state.user.role === 'Admin' || app.state.user.role === 'SafetyUploader');
        const isDeptOwner = (app.state.user.role === 'DepartmentUploader' && c['主辦部門'] === app.state.user.department);
        const isAdmin = app.state.user.role === 'Admin';
        const isLite = (app.state.user.role === 'DepartmentUploader');

        // 歸檔路徑說明 HTML
        const archivingHtml = `
            <div style="margin-top:20px; border-top:1px solid var(--border); padding-top:15px;">
                <div style="font-weight:700; font-size:0.85rem; color:var(--text-muted); margin-bottom:10px;"><i class="fas fa-archive"></i> 歸檔位置說明 (點擊圖示複製)</div>
                <div class="archiving-path-card">
                    <span><b>S2 原始單：</b>\\\\10.64.200.21\\d602\\部門資料夾\\d60269 工業安全衛生組\\開放區\\8 工安查核(表)</span>
                    <i class="fas fa-copy copy-btn" onclick="app.copyToClipboard('\\\\\\\\10.64.200.21\\\\d602\\\\部門資料夾\\\\d60269 工業安全衛生組\\\\開放區\\\\8 工安查核(表)')"></i>
                </div>
                <div class="archiving-path-card">
                    <span><b>S3, S4 核章/結案：</b>\\\\10.64.200.21\\d602\\部門資料夾\\d60269 工業安全衛生組\\部門共用區\\0.工安組線上公文查詢系統\\115年\\115年度--工安查核紀錄</span>
                    <i class="fas fa-copy copy-btn" onclick="app.copyToClipboard('\\\\\\\\10.64.200.21\\\\d602\\\\部門資料夾\\\\d60269 工業安全衛生組\\\\部門共用區\\\\0.工安組線上公文查詢系統\\\\115年\\\\115年度--工安查核紀錄')"></i>
                </div>
            </div>
        `;

        if (isLite) {
            // DepartmentUploader 的高階大畫面任務導向介面
            let liteHtml = `
                <div class="lite-flow-container">
                    <div class="lite-date-hero">
                        <div class="label"><i class="fas fa-calendar-check"></i> 查核日期</div>
                        <div class="value">${c['查核日期']}</div>
                        <div style="font-size:0.9rem; opacity:0.8; margin-top:8px;">工程：<b>${c['工程簡稱']}</b></div>
                    </div>
                    
                    <div class="lite-step-card ${c['第2階段連結'] ? 'active' : ''}">
                        <span class="step-badge">步驟 1</span>
                        <h4><i class="fas fa-file-word"></i> 下載第 2 階段原始單</h4>
                        <p>工安組已完成初勘登錄。請點擊上方按鈕下載 Word 格式原始單，進行後續改善複核作業。</p>
                        ${c['第2階段連結'] ? 
                            `<a href="${c['第2階段連結']}" target="_blank" class="btn btn-primary" style="width:100%; height:60px; font-size:1.1rem; justify-content:center; border-radius:15px;"><i class="fas fa-download"></i> 立即下載原始單 Word</a>` : 
                            `<div style="color:var(--warning); background:rgba(244,63,94,0.05); padding:15px; border-radius:12px; text-align:center; border:1px dashed var(--warning);"><i class="fas fa-clock"></i> 工安組作業中，尚未上傳原始單</div>`
                        }
                    </div>

                    <div class="lite-step-card ${(c['辦理狀態'] === '第2階段-改善單已上傳' || c['辦理狀態'] === '第3階段-工作隊版已處理') ? 'active' : ''}">
                        <span class="step-badge">步驟 2</span>
                        <h4><i class="fas fa-file-pdf"></i> 上傳第 3 階段核章版</h4>
                        <p>完成現場改善並核章後，請將掃描後的 PDF 檔案在此回傳。系統將自動通知工安組結案。</p>
                        ${(c['辦理狀態'] === '第2階段-改善單已上傳' || (c['辦理狀態'] === '第3階段-工作隊版已處理' && isDeptOwner)) ? 
                            app.getUploadSection(id, 'stage3', '立即上傳核章版 PDF', '#fbbf24', '', !!c['第3階段連結']) : 
                            `<div style="color:var(--text-muted); background:rgba(0,0,0,0.03); padding:15px; border-radius:12px; text-align:center; border:1px dashed var(--border);"><i class="fas fa-lock"></i> 目前尚未開放上傳 (需先完成步驟 1)</div>`
                        }
                    </div>

                    ${archivingHtml}
                    <button class="btn btn-outline" style="width:100%; margin-top:10px; height:50px; justify-content:center;" onclick="app.viewHistory('${id}')"><i class="fas fa-history"></i> 查看此案件歷史紀錄</button>
                </div>
            `;
            app.openModal('傳辦人員作業中心', liteHtml);
        } else {
            // Admin / SafetyUploader 的完整管理介面
            let html = `
                <div class="tabs-container">
                    <div class="tabs-header">
                        <button class="tab-btn active" onclick="app.switchTab(event, 'tabFiles')"><i class="fas fa-folder-open"></i> 檔案管理</button>
                        <button class="tab-btn" onclick="app.switchTab(event, 'tabDefs')"><i class="fas fa-list-ul"></i> 缺失項目</button>
                    </div>

                    <!-- 分頁一：檔案管理 -->
                    <div id="tabFiles" class="tab-content active">
                        <div style="margin-bottom:15px; padding:12px; background:rgba(99,102,241,0.05); border-radius:12px; font-size:0.8rem; border:1px solid rgba(99,102,241,0.1);">
                            <i class="fas fa-info-circle"></i> 當前狀態：<b style="color:var(--primary);">${c['辦理狀態']}</b>
                        </div>
                        <div class="manage-grid">
                            ${app.getUploadSection(id, 'stage2', 'S2 原始單', 'var(--warning)', '更換 Word 檔', !!c['第2階段連結'])}
                            ${app.getUploadSection(id, 'stage3', 'S3 核章版', '#fbbf24', '更換核章版', !!c['第3階段連結'])}
                            ${app.getUploadSection(id, 'stage4e', 'S4 結案(員工)', 'var(--success)', '更換員工版', !!c['第4階段連結-員工'])}
                            ${app.getUploadSection(id, 'stage4c', 'S4 結案(承)', 'var(--success)', '更換承攬商版', !!c['第4階段連結-承攬商'])}
                        </div>
                        
                        ${(c['第2階段連結'] || c['第3階段連結'] || c['第4階段連結-員工'] || c['第4階段連結-承攬商']) ? `
                        <div style="margin-top:15px; padding:12px; background:rgba(0,0,0,0.03); border-radius:12px; border:1px solid var(--border);">
                            <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(80px, 1fr)); gap:8px;">
                                ${c['第2階段連結'] ? `<a href="${c['第2階段連結']}" target="_blank" class="btn btn-outline" style="font-size:0.7rem; justify-content:center;"><i class="fas fa-file-signature"></i> S2</a>` : ''}
                                ${c['第3階段連結'] ? `<a href="${c['第3階段連結']}" target="_blank" class="btn btn-outline" style="font-size:0.7rem; justify-content:center;"><i class="fas fa-stamp"></i> S3</a>` : ''}
                                ${c['第4階段連結-員工'] ? `<a href="${c['第4階段連結-員工']}" target="_blank" class="btn btn-outline" style="font-size:0.7rem; justify-content:center;"><i class="fas fa-user-check"></i> S4員</a>` : ''}
                                ${c['第4階段連結-承攬商'] ? `<a href="${c['第4階段連結-承攬商']}" target="_blank" class="btn btn-outline" style="font-size:0.7rem; justify-content:center;"><i class="fas fa-building-circle-check"></i> S4承</a>` : ''}
                            </div>
                        </div>` : ''}
                        
                        ${archivingHtml}
                        <button class="btn btn-outline" style="width:100%; margin-top:20px;" onclick="app.viewHistory('${id}')"><i class="fas fa-history"></i> 查看歷史紀錄</button>
                    </div>

                    <!-- 分頁二：缺失項目 -->
                    <div id="tabDefs" class="tab-content">
                        <div class="modal-instruction" style="margin-bottom:10px;">
                            <i class="fas fa-keyboard"></i> <b>填寫說明：</b> 每一行代表一項缺失。如有多項可直接換行輸入，系統將自動拆分為獨立項目。
                        </div>
                        <div id="caseDefsList" style="margin-bottom:15px; max-height:220px; overflow-y:auto; border:1px solid var(--border); border-radius:12px; padding:5px; background:rgba(0,0,0,0.01);">
                            <!-- 加載該案件的缺失 -->
                            <div style="text-align:center; padding:10px; color:var(--text-muted);">載入中...</div>
                        </div>
                        <div style="background:var(--bg-input); padding:15px; border-radius:14px; border:1px solid var(--border);">
                            <div style="font-weight:700; margin-bottom:8px; font-size:0.85rem; color:var(--primary);">
                                <i class="fas fa-plus-circle"></i> 快速新增缺失
                            </div>
                            <textarea id="caseDefContent" style="width:100%; height:80px; margin-bottom:10px; border-radius:8px;" placeholder="範例：&#10;1. 現場人員未戴安全帽&#10;2. 施工架下方未設置防護網..."></textarea>
                            <div style="display:flex; gap:10px; align-items:center;">
                                <div style="flex:1;">
                                    <label style="font-size:0.75rem; color:var(--text-muted); display:block; margin-bottom:4px;">改善期限</label>
                                    <input type="date" id="caseDefDeadline" value="${new Date().toISOString().split('T')[0]}" style="width:100%; font-size:0.8rem;">
                                </div>
                                <button class="btn btn-primary" onclick="app.submitDeficiencyFromCase('${id}', '${c['工程簡稱']}', '${c['主辦部門']}')" style="align-self: flex-end; padding:10px 20px;">
                                    <i class="fas fa-paper-plane"></i> 確認新增
                                </button>
                            </div>
                        </div>
                    </div>
                </div>
            `;
            const projInfo = app.state.projects.find(p => p.abbr === c['工程簡稱']);
            const snLabel = projInfo ? `${projInfo.serial} - ` : '';
            app.openModal(`案件管理: ${snLabel}${c['工程簡稱']}`, html);
        }
        
        // 延時載入缺失清單（確保 DOM 已渲染）
        if (!isLite) setTimeout(() => app.renderCaseDeficiencies(id), 100);
    },

    switchTab: (e, tabId) => {
        const container = e.target.closest('.tabs-container');
        container.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        container.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
        e.target.classList.add('active');
        document.getElementById(tabId).classList.add('active');
    },
    renderCaseDeficiencies: (caseId) => {
        const listDiv = document.getElementById('caseDefsList');
        if (!listDiv) return;
        const defs = app.state.deficiencies.filter(d => d.caseId === caseId);
        if (defs.length === 0) {
            listDiv.innerHTML = `<div style="text-align:center; padding:20px; color:var(--text-muted); font-size:0.8rem;">☕ 此案件尚無缺失紀錄</div>`;
            return;
        }
        listDiv.innerHTML = defs.map(d => `
            <div style="padding:10px; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center;">
                <div style="font-size:0.8rem; flex:1;">
                    <div style="display:flex; align-items:center; gap:8px;">
                        <span style="font-weight:700;">${d.content}</span>
                        <span class="def-dept-tag">${d.department}</span>
                    </div>
                    <div style="color:var(--text-muted); font-size:0.75rem; margin-top:4px;">
                        期限：${d.deadline} | 狀態：<span style="color:${d.status === '已改善' ? 'var(--success)' : 'var(--warning)'}">${d.status}</span>
                    </div>
                </div>
                <div style="display:flex; gap:4px;">
                    <button class="btn" onclick="app.openEditDef('${d.id}')" style="padding:4px; color:var(--primary); background:none;"><i class="fas fa-edit"></i></button>
                    ${app.state.user.role === 'Admin' ? `<button class="btn" onclick="app.deleteDeficiency('${d.id}'); setTimeout(()=>app.renderCaseDeficiencies('${caseId}'), 1000);" style="padding:4px; color:var(--warning); background:none;"><i class="fas fa-trash"></i></button>` : ''}
                </div>
            </div>
        `).join('');
    },
    openEditDef: (id) => {
        const d = app.state.deficiencies.find(item => item.id == id);
        if (!d) return;
        
        app.openModal('編輯缺失紀錄', `
            <div style="display:flex; flex-direction:column; gap:15px;">
                <div><label>工程：</label> <b>${d.abbr}</b></div>
                <div><label>缺失內容：</label><textarea id="editDefContent" style="width:100%; height:100px;">${d.content}</textarea></div>
                <div><label>改善期限：</label><input type="date" id="editDefDeadline" value="${d.deadline}"></div>
                <div><label>主辦部門：</label><input type="text" id="editDefDept" value="${d.department}"></div>
                <div><label>狀態：</label>
                    <select id="editDefStatus">
                        <option value="待改善" ${d.status === '待改善' ? 'selected' : ''}>待改善</option>
                        <option value="已改善" ${d.status === '已改善' ? 'selected' : ''}>已改善</option>
                    </select>
                </div>
                <button class="btn btn-primary" onclick="app.submitEditDef('${id}', '${d.caseId}')">儲存修改</button>
            </div>
        `);
    },
    submitEditDef: async (id, caseId) => {
        const content = document.getElementById('editDefContent').value.trim();
        const deadline = document.getElementById('editDefDeadline').value;
        const department = document.getElementById('editDefDept').value.trim();
        const status = document.getElementById('editDefStatus').value;

        if (!content || !deadline || !department) return app.showToast("請填寫所有欄位", "error");

        app.setModalLoading(true);
        try {
            await api.updateDeficiency({ id, content, deadline, department, status });
            app.showToast("✅ 修改成功");
            await app.fetchDeficiencies();
            app.closeModal();
            // 如果是在案件管理開啟的，重新渲染該案件的缺失
            app.openManage(caseId); 
        } catch(e) { app.showToast(e.message, "error"); } finally { app.setModalLoading(false); }
    },
    openGlobalBatchAddModal: () => {
        // 全域批次新增：需選擇工程
        let options = app.state.cases.filter(c => c['辦理狀態'] !== '第4階段-已結案')
            .map(c => `<option value="${c.id}">${c['工程簡稱']} (${c['查核日期']})</option>`).join('');
        
        if (!options) return app.showToast("目前沒有進行中的案件可供新增缺失", "warning");

        app.openModal('全域批次新增缺失', `
            <div style="display:flex; flex-direction:column; gap:15px;">
                <div class="modal-instruction">
                    <i class="fas fa-lightbulb"></i> <b>使用小技巧：</b><br>
                    您可以一次貼上多行缺失內容，系統會自動為該案件建立多筆紀錄。主辦部門與承攬商將自動代入案件資訊。
                </div>
                <div>
                    <label>選擇案件：</label>
                    <select id="globalBatchCase" style="width:100%; height:45px;">${options}</select>
                </div>
                <div>
                    <label>缺失內容 (每一行一筆)：</label>
                    <textarea id="globalBatchContent" style="width:100%; height:150px;" placeholder="例如：&#10;滅火器過期&#10;走道堆置雜物..."></textarea>
                </div>
                <div>
                    <label>改善期限：</label>
                    <input type="date" id="globalBatchDeadline" value="${new Date().toISOString().split('T')[0]}">
                </div>
                <button class="btn btn-primary" onclick="app.submitGlobalBatch()" style="justify-content:center; padding:15px;">
                    <i class="fas fa-plus"></i> 確認批次新增
                </button>
            </div>
        `);
    },
    submitGlobalBatch: async () => {
        const caseId = document.getElementById('globalBatchCase').value;
        const contentRaw = document.getElementById('globalBatchContent').value.trim();
        const deadline = document.getElementById('globalBatchDeadline').value;
        
        if (!contentRaw || !deadline) return app.showToast("請填寫內容與期限", "error");
        
        const c = app.state.cases.find(item => item.id == caseId);
        const lines = contentRaw.split('\n').map(l => l.trim()).filter(l => l !== '');
        
        app.setModalLoading(true);
        try {
            const items = lines.map(line => ({ 
                caseId, abbr: c['工程簡稱'], content: line, deadline, department: c['主辦部門'], status: '待改善' 
            }));
            await api.batchAddDeficiencies(items);
            app.showToast(`✅ 已新增 ${lines.length} 筆缺失`);
            await app.fetchDeficiencies();
            app.closeModal();
            app.renderView();
        } catch(e) { app.showToast(e.message, "error"); } finally { app.setModalLoading(false); }
    },
    submitDeficiencyFromCase: async (caseId, abbr, dept) => {
        const contentRaw = document.getElementById('caseDefContent').value.trim();
        const deadline = document.getElementById('caseDefDeadline').value;
        if (!contentRaw || !deadline) return app.showToast("請輸入內容與期限", "error");

        const lines = contentRaw.split('\n').map(l => l.trim()).filter(l => l !== '');
        app.setModalLoading(true);
        try {
            if (lines.length > 1) {
                const items = lines.map(line => ({ caseId, abbr, content: line, deadline, department: dept, status: '待改善' }));
                await api.batchAddDeficiencies(items);
            } else {
                await api.updateDeficiency({ caseId, abbr, content: lines[0], deadline, department: dept, status: '待改善' });
            }
            app.showToast("✅ 新增成功");
            document.getElementById('caseDefContent').value = '';
            await app.fetchDeficiencies();
            app.renderCaseDeficiencies(caseId);
        } catch(e) { app.showToast(e.message, "error"); } finally { app.setModalLoading(false); }
    },
    getUploadSection: (id, stage, label, color, note = '', exists = false) => `
        <div class="upload-section ${exists ? 'has-file' : ''}" style="border-left: 5px solid ${color || 'var(--border)'}; position:relative;">
            ${exists ? `<span style="position:absolute; top:8px; right:8px; font-size:0.6rem; color:var(--success); background:rgba(16,185,129,0.1); padding:2px 6px; border-radius:4px;"><i class="fas fa-check"></i> 已存在</span>` : ''}
            <div class="upload-header" style="color:${color || 'inherit'}">
                <i class="fas fa-cloud-upload-alt"></i> ${label}
            </div>
            ${note ? `<p class="upload-note">${note}</p>` : ''}
            <div class="upload-actions">
                <input type="file" id="file_${stage}" style="width:100%; margin-bottom:12px; font-size:0.8rem;" />
                <button class="btn" style="width:100%; justify-content:center; background:${color || 'var(--primary)'}; color:white;" onclick="app.submitFile('${id}', '${stage}', ${exists})">
                    ${exists ? '替換現有檔案' : '確認上傳存檔'}
                </button>
            </div>
        </div>
    `,
    submitFile: async (id, stage, isReplace = false) => {
        const input = document.getElementById(`file_${stage}`);
        if(!input.files.length) return app.showToast("請先選擇檔案", "error");
        
        let reason = "";
        if (isReplace) {
            reason = prompt("⚠️ 您正在替換現有檔案，請輸入更換原因：");
            if (reason === null) return; // 使用者取消
            reason = reason.trim();
            if (!reason) return app.showToast("更換檔案必須輸入原因", "error");
        }

        const file = input.files[0];
        app.showGlobalProgress(true, "正在處理檔案...", 20);
        
        try {
            const base64 = await app.fileToBase64(file);
            app.showGlobalProgress(true, "檔案傳送中...", 60);
            
            await api.uploadFile(id, stage, base64, file.name, app.state.user.email, reason);
            
            app.showGlobalProgress(true, "同步資料中...", 90);
            const res = await api.init();
            app.state.cases = res.data.cases;
            
            app.showGlobalProgress(false);
            app.updateStats(); app.renderView(); app.closeModal();
            app.showToast("✅ 上傳成功");
            alert("✅ 檔案已完成上傳並同步成功！");
        } catch(e) { 
            app.showGlobalProgress(false);
            app.showToast(e.message, "error"); 
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
    },
    openBlankFormsModal: () => {
        const categories = [
            {
                title: "🏢 承攬商相關表單",
                forms: [
                    { name: "1. 承攬商工安查核紀錄表", id: "1Ld3Kpbdb7iKEb7jaRa6wTE-k7UX-ACw_" },
                    { name: "2. 附件3：承攬商外籍移工工作管理抽查表", id: "1G9dgYKIhSIZtgC9NK-H1lK6irvCQveJg" },
                    { name: "3. 附件4：防颱整備聯繫支援督導表 【每年5~11月必填】", id: "1-GN-OqB59cEKEd0SJ1ghcwkcFtwuXNY0" }
                ]
            },
            {
                title: "👷 自辦監造相關表單",
                forms: [
                    { name: "4. 自辦監造員工工安查核紀錄表", id: "1xFH7fxaR69GYnyEwK73KLH9TcjucdQ9O" }
                ]
            },
            {
                title: "🤝 委外監造相關表單",
                forms: [
                    { name: "5. 委外監造員工工安查核紀錄表", id: "12kXyBDtAddQYPzaGnKtZH2ZfOjWGxiCS" }
                ]
            }
        ];

        let html = `
            <div class="modal-instruction">
                <i class="fas fa-info-circle"></i> <b>空白表單下載說明：</b><br>
                點擊下方連結將直接下載為 <b>Word (.docx)</b> 格式。表單僅限下載使用，不支援線上編輯，以確保版面格式正確。
            </div>
        `;

        categories.forEach(cat => {
            html += `
                <div style="margin-bottom:20px;">
                    <div style="font-weight:700; color:var(--primary); margin-bottom:10px; padding-left:5px; border-left:4px solid var(--primary);">${cat.title}</div>
                    <div class="forms-grid">
            `;
            cat.forms.forEach(f => {
                html += `
                    <a href="https://docs.google.com/document/d/${f.id}/export?format=docx" 
                       class="form-card" 
                       target="_blank" 
                       title="點擊下載 Word 檔">
                        <span>${f.name}</span>
                        <i class="fas fa-download"></i>
                    </a>
                `;
            });
            html += `</div></div>`;
        });

        app.openModal('空白表單下載區', html);
    },

    openReportModal: () => {
        const projects = app.state.projects;
        const projectOptions = projects.map(p => `<option value="${p.abbr}">${p.abbr}</option>`).join('');
        const now = new Date();
        const firstDay = new Date(now.getFullYear(), now.getMonth(), 1).toISOString().split('T')[0];
        const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0).toISOString().split('T')[0];

        app.openModal('生成統計與缺失查詢報告', `
            <div style="display:flex; flex-direction:column; gap:15px;">
                <div class="report-field-group">
                    <div><label>起始日期：</label><input type="date" id="reportStart" value="${firstDay}" style="width:100%;"></div>
                    <div><label>結束日期：</label><input type="date" id="reportEnd" value="${lastDay}" style="width:100%;"></div>
                </div>
                <div class="report-field-group">
                    <div>
                        <label>篩選工程：</label>
                        <select id="reportProj" style="width:100%;">
                            <option value="">-- 全部工程 --</option>
                            ${projectOptions}
                        </select>
                    </div>
                    <div>
                        <label>結案狀態：</label>
                        <select id="reportStatus" style="width:100%;">
                            <option value="">-- 全部狀態 --</option>
                            <option value="已結案">僅列已結案</option>
                            <option value="進行中">僅列進行中</option>
                        </select>
                    </div>
                </div>
                <div>
                    <label>缺失內容關鍵字查詢：</label>
                    <input type="text" id="reportKeyword" placeholder="關鍵字 (如: 漏電, 安全帽...)" style="width:100%;">
                </div>
                <button class="btn btn-primary" onclick="app.runReport()" style="justify-content:center;">查詢並生成報告</button>
                <div id="reportResult" class="hidden" style="margin-top:15px; border-top:1px solid var(--border); padding-top:15px;">
                    <div id="reportSummary" style="display:grid; grid-template-columns:1fr 1fr; gap:10px; margin-bottom:15px;"></div>
                    <div id="reportDetail" style="max-height:300px; overflow-y:auto; font-size:0.85rem;"></div>
                    <button class="btn btn-outline" style="width:100%; margin-top:10px; justify-content:center;" onclick="app.printReport()">列印報告 / 導出 PDF</button>
                </div>
            </div>
        `);
    },
    runReport: () => {
        const start = document.getElementById('reportStart').value;
        const end = document.getElementById('reportEnd').value;
        const projAbbr = document.getElementById('reportProj').value;
        const statusFilter = document.getElementById('reportStatus').value;
        const keyword = document.getElementById('reportKeyword').value.trim().toLowerCase();

        // 1. 先篩選案件
        let filteredCases = app.state.cases.filter(c => {
            const date = c['查核日期'];
            if (!date) return false;
            const isDateMatch = date >= start && date <= end;
            const isProjMatch = projAbbr === '' || c['工程簡稱'] === projAbbr;
            
            let isStatusMatch = true;
            if (statusFilter === '已結案') isStatusMatch = c['辦理狀態'] === '第4階段-已結案';
            if (statusFilter === '進行中') isStatusMatch = c['辦理狀態'] !== '第4階段-已結案';

            return isDateMatch && isProjMatch && isStatusMatch;
        });

        // 2. 獲取篩選案件的缺失，並應用關鍵字過濾
        let filteredDefs = app.state.deficiencies.filter(d => filteredCases.find(c => c.id === d.caseId));
        if (keyword) {
            filteredDefs = filteredDefs.filter(d => d.content.toLowerCase().includes(keyword));
            // 反向過濾案件：僅保留包含關鍵字缺失的案件
            const matchingCaseIds = new Set(filteredDefs.map(d => d.caseId));
            filteredCases = filteredCases.filter(c => matchingCaseIds.has(c.id));
        }

        const summaryBox = document.getElementById('reportSummary');
        summaryBox.innerHTML = `
            <div style="padding:10px; background:rgba(99,102,241,0.1); border-radius:8px; text-align:center;">
                <div style="font-size:0.75rem; color:var(--text-muted);">查核次數</div>
                <div style="font-size:1.5rem; font-weight:800; color:var(--primary);">${filteredCases.length}</div>
            </div>
            <div style="padding:10px; background:rgba(245,158,11,0.1); border-radius:8px; text-align:center;">
                <div style="font-size:0.75rem; color:var(--text-muted);">符合缺失數</div>
                <div style="font-size:1.5rem; font-weight:800; color:var(--warning);">${filteredDefs.length}</div>
            </div>
        `;

        const detailBox = document.getElementById('reportDetail');
        if (filteredCases.length === 0) {
            detailBox.innerHTML = '<p style="text-align:center; padding:20px;">此條件下無相關資料</p>';
        } else {
            let html = '<table style="width:100%; border-collapse:collapse; font-size:0.8rem;">';
            html += '<thead style="background:var(--bg-input);"><tr><th style="text-align:left; padding:8px;">日期/工程</th><th style="text-align:left; padding:8px;">缺失內容明細</th></tr></thead><tbody>';
            filteredCases.forEach(c => {
                const caseDefs = filteredDefs.filter(d => d.caseId === c.id);
                const defsHtml = caseDefs.map(d => `• ${d.content}`).join('<br>');
                html += `<tr style="border-bottom:1px solid var(--border); vertical-align:top;">
                    <td style="padding:8px; white-space:nowrap;" class="report-clickable" onclick="app.openManage('${c.id}')">
                        <b>${c['查核日期']}</b><br>${c['工程簡稱']}
                    </td>
                    <td style="padding:8px;">${defsHtml || '<span style="color:var(--text-muted);">無符合關鍵字缺失</span>'}</td>
                </tr>`;
            });
            html += '</tbody></table>';
            detailBox.innerHTML = html;
        }

        document.getElementById('reportResult').classList.remove('hidden');
    },
    printReport: () => {
        const start = document.getElementById('reportStart').value;
        const end = document.getElementById('reportEnd').value;
        const summary = document.getElementById('reportSummary').innerHTML;
        const detail = document.getElementById('reportDetail').innerHTML;
        const win = window.open('', '_blank');
        win.document.write(`
            <html><head><title>工安查核統計報告</title>
            <style>
                body { font-family: sans-serif; padding: 40px; }
                table { width: 100%; border-collapse: collapse; margin-top: 20px; }
                th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
                th { background: #f4f4f4; }
                .header { text-align: center; margin-bottom: 30px; }
            </style>
            </head><body>
            <div class="header">
                <h2>工安查核統計報告</h2>
                <p>查詢區間：${start} ~ ${end}</p>
            </div>
            <div style="display:flex; gap:20px; justify-content:center; margin-bottom:20px;">${summary}</div>
            ${detail}
            <script>window.print();<\/script>
            </body></html>
        `);
        win.document.close();
    }
};

window.onload = () => app.initAuth();
