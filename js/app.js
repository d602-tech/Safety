/**
 * 前端核心邏輯與狀態管理 v5.4
 */

const GOOGLE_CLIENT_ID = window.ENV && window.ENV.GOOGLE_CLIENT_ID ? window.ENV.GOOGLE_CLIENT_ID : "791038911460-8tfq97vhrvr4iaq5r3s1ti1abfpuddd9.apps.googleusercontent.com";

const app = {
    showLoading: (show) => {
        const loader = document.getElementById('loading');
        if (loader) show ? loader.classList.remove('hidden') : loader.classList.add('hidden');
    },

    openModal: (title, html) => {
        const overlay = document.getElementById('modalOverlay');
        const modalTitle = document.getElementById('modalTitle');
        const modalBody = document.getElementById('modalBody');
        const modal = document.querySelector('.modal');
        
        if (overlay && modalTitle && modalBody) {
            modalTitle.innerText = title;
            modalBody.innerHTML = html;
            overlay.classList.remove('hidden');
            if (modal) modal.classList.remove('modal-lg'); 
        }
    },

    closeModal: () => {
        const overlay = document.getElementById('modalOverlay');
        if (overlay) overlay.classList.add('hidden');
    },

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
            app.state.quickFilter = 'active'; // 預設顯示未結案
            app.applyRoleRestrictions(); // 關鍵：確保初始化時就套用角色限制（包含訪客模式）
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
            app.applyRoleRestrictions(); // 關鍵：拉完資料後再次確保畫面依角色/訪客模式切換
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

            app.state.quickFilter = 'active'; // 登入後預設顯示未結案
            app.applyRoleRestrictions();
            app.extractYears();
            app.extractDepartments();
            
            // 先拉缺失清單，再更新統計（否則缺失次數永遠為 0）
            if (app.state.user.role === 'Admin') app.fetchUsers();
            await app.fetchDeficiencies();
            app.updateStats();
            app.renderView();
            
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
        
        // 先全部隱藏
        if (btnNew) btnNew.classList.add('hidden');
        if (btnRemind) btnRemind.classList.add('hidden');
        if (btnProjMgmt) btnProjMgmt.classList.add('hidden');
        if (btnAdminUsers) btnAdminUsers.classList.add('hidden');

        const workflow = document.querySelector('.workflow-section');
        const modeSwitch = document.querySelector('.mode-switch-wrapper');
        const deptStats = document.getElementById('deptStatsSection');
        const btnExport = document.querySelector('button[onclick="app.exportToCSV()"]');
        const btnExample = document.querySelector('button[onclick="app.downloadExample()"]');
        const filterDept = document.getElementById('filterDepartment');
        const filterYear = document.getElementById('filterYear');
        const filterStatus = document.getElementById('filterStatus');
        const toolbarSection = document.querySelector('.toolbar');
        const dashboard = document.querySelector('.dashboard');

        // 重置
        if (workflow) workflow.classList.remove('hidden');
        if (modeSwitch) modeSwitch.classList.remove('hidden');
        if (deptStats) deptStats.classList.add('hidden');
        if (btnExport) btnExport.classList.remove('hidden');
        if (btnExample) btnExample.classList.remove('hidden');
        if (filterDept) filterDept.classList.remove('hidden');
        if (filterYear) filterYear.classList.remove('hidden');
        if (filterStatus) filterStatus.classList.remove('hidden');
        if (toolbarSection) toolbarSection.classList.remove('hidden');
        if (dashboard) dashboard.classList.remove('hidden');

        const user = app.state.user;
        const pathRef = document.getElementById('pathReferenceSection');
        if (pathRef) pathRef.style.display = 'none';

        // 未登入訪客模式
        if (!user) {
            if (workflow) workflow.classList.add('hidden');
            if (modeSwitch) modeSwitch.classList.add('hidden');
            if (btnExport) btnExport.classList.add('hidden');
            if (filterDept) filterDept.classList.add('hidden');
            if (filterStatus) filterStatus.classList.add('hidden');
            if (filterYear) filterYear.classList.add('hidden');
            if (dashboard) dashboard.classList.add('hidden');
            if (toolbarSection) toolbarSection.classList.add('hidden');
            
            // 訪客時隱藏全域導覽列部分按鈕
            const navActions = document.querySelector('.nav-actions');
            if (navActions) navActions.classList.add('hidden');

            // 切換到訪客專屬畫面
            const guestView = document.getElementById('guestView');
            const mainView = document.getElementById('mainView');
            if (guestView) guestView.classList.remove('hidden');
            if (mainView) mainView.classList.add('hidden');
            app.renderGuestView();
            return;
        }
        // 登入後恢復全域導覽列與畫面切換
        const navActions = document.querySelector('.nav-actions');
        if (navActions) navActions.classList.remove('hidden');
        const guestView = document.getElementById('guestView');
        const mainView = document.getElementById('mainView');
        if (guestView) guestView.classList.add('hidden');
        if (mainView) mainView.classList.remove('hidden');

        if (user.role === 'DepartmentUploader') {
            if (workflow) workflow.classList.add('hidden');
            if (modeSwitch) modeSwitch.classList.add('hidden');
            if (deptStats) deptStats.classList.remove('hidden');
            if (btnExport) btnExport.classList.add('hidden');
            if (btnExample) btnExample.classList.add('hidden');
            if (filterDept) filterDept.classList.add('hidden');
        }

        if (user.role === 'Admin' || user.role === 'SafetyUploader') {
            if (btnNew) btnNew.classList.remove('hidden');
            if (pathRef) pathRef.style.display = 'block';
        }
        if (user.role === 'Admin') {
            if (btnRemind) btnRemind.classList.remove('hidden');
            if (btnProjMgmt) btnProjMgmt.classList.remove('hidden');
            if (btnAdminUsers) btnAdminUsers.classList.remove('hidden');
            const btnInitSystem = document.getElementById('btnInitSystem');
            const btnDeptAccounts = document.getElementById('btnDeptAccounts');
            const btnTestEmail = document.getElementById('btnTestEmail');
            const btnRunReminder = document.getElementById('btnRunReminder');
            const btnSetupTrigger = document.getElementById('btnSetupTrigger');
            if (btnInitSystem) btnInitSystem.classList.remove('hidden');
            if (btnDeptAccounts) btnDeptAccounts.classList.remove('hidden');
            if (btnTestEmail) btnTestEmail.classList.remove('hidden');
            if (btnRunReminder) btnRunReminder.classList.remove('hidden');
            if (btnSetupTrigger) btnSetupTrigger.classList.remove('hidden');
        }
    },

    renderGuestView: () => {
        const container = document.getElementById('guestView');
        if (!container) return;
        const todayStr = new Date().toISOString().split('T')[0];
        const activeCases = app.state.cases.filter(c => c['辦理狀態'] !== '第4階段-已結案');

        const caseCards = activeCases.map(c => {
            const diffDays = Math.ceil((new Date(c['最晚應核章日期']) - new Date(todayStr)) / (1000 * 60 * 60 * 24));
            const isOverdue = diffDays <= 0;
            const isCritical = diffDays <= 3;
            const isUrgent = diffDays <= 7;
            const urgencyClass = isCritical ? 'countdown-critical' : isUrgent ? 'countdown-urgent' : '';
            const urgencyColor = isCritical ? 'var(--danger)' : isUrgent ? '#f97316' : 'var(--warning)';
            return `
                <div class="guest-case-card ${urgencyClass}" style="border-top: 4px solid ${urgencyColor};">
                    <div class="guest-case-name">${c['工程簡稱']}</div>
                    <div class="guest-case-meta">
                        <span><i class="fas fa-building"></i> ${c['主辦部門']}</span>
                        <span><i class="fas fa-hard-hat"></i> ${c['承攬商']}</span>
                    </div>
                    <div class="countdown-hero ${urgencyClass}" style="margin-top:12px;">
                        <div class="label">⏳ 離最晚核章期限</div>
                        <div class="countdown-days">${isOverdue ? '已逾期' : diffDays + ' 天'}</div>
                        <div style="font-size:0.75rem; color:var(--text-muted); margin-top:2px;">${c['最晚應核章日期']}</div>
                    </div>
                </div>
            `;
        }).join('');

        container.innerHTML = `
            <div class="guest-page">
                <!-- Hero 登入區塊 -->
                <div class="guest-hero">
                    <div style="font-size:3rem;">🛡️</div>
                    <h1 class="guest-title">工安查核管理系統</h1>
                    <p class="guest-subtitle">查看即時查核進度與結案倒數。登入後可管理案件與上傳檔案。</p>
                    <div class="guest-actions">
                        <div id="guestLoginBtn" style="display:inline-block;"></div>
                        <button class="btn btn-outline" style="border-radius:50px; padding:10px 28px; font-size:1rem;" onclick="app.openBlankFormsModal()">
                            <i class="fas fa-file-download"></i> 空白表單下載
                        </button>
                    </div>
                </div>

                <!-- 未結案案件倶數卡片 -->
                <div class="guest-section-title">
                    <i class="fas fa-clock"></i> 目前進行中案件 (${activeCases.length} 項)
                </div>
                ${activeCases.length === 0 
                    ? '<div class="guest-empty">✅ 所有案件均已結案</div>'
                    : `<div class="guest-cases-grid">${caseCards}</div>`
                }
            </div>
        `;

        // 在 guestView 裡回度渲染 Google Login 按鈕
        const btnEl = document.getElementById('guestLoginBtn');
        if (btnEl && window.google) {
            google.accounts.id.renderButton(btnEl, {
                theme: app.state.theme === 'light' ? 'outline' : 'filled_blue',
                size: 'large', shape: 'pill', text: 'signin_with'
            });
        }
    },
    
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
            projects: document.getElementById('viewProjects'),
            deptAccounts: document.getElementById('viewDeptAccounts')
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
            // deptAccounts 特殊處理： view 名稱轉為 id
            const viewId = view === 'deptAccounts' ? 'viewDeptAccounts'
                         : 'view' + view.charAt(0).toUpperCase() + view.slice(1);
            document.getElementById(viewId)?.classList.remove('hidden');
            btnBack?.classList.remove('hidden');
            btnUsers?.classList.add('hidden');
            btnProj?.classList.add('hidden');
            btnDef?.classList.add('hidden');
            if (view === 'users') app.renderUsers();
            if (view === 'deficiencies') app.renderDeficiencies();
            if (view === 'projects') app.renderProjects();
            if (view === 'deptAccounts') app.renderDeptAccounts();
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
    getBadgeClass: (status) => {
        if (!status) return 'badge-status';
        if (status.includes('1')) return 'badge-status-s1';
        if (status.includes('2')) return 'warning';
        if (status.includes('3')) return 'badge-status-s3';
        if (status.includes('4')) return 'success';
        return 'badge-status';
    },
    
                        getProgressHtml: (c) => {
        const s2e = !!c['S2員工查核檔案位置'];
        const s2c = !!c['S2廠商查核檔案位置'];
        const s3 = !!c['S3廠商及員工改善後核章檔案位置'];
        const s4 = !!c['S4結案檔案位置'];
        const isClosed = c['辦理狀態'] === '第4階段-已結案';

        return `
            <div class="progress-container">
                <div class="progress-label">
                    <span>階段進度</span>
                    <span>${isClosed ? '100%' : (s4 ? '80%' : (s3 ? '50%' : (s2e || s2c ? '25%' : '0%')))}</span>
                </div>
                <div class="progress-track">
                    <div class="progress-step step-s1 active" title="S1: 已登錄"></div>
                    <div class="progress-step step-s2 ${s2e || s2c || s3 || s4 ? 'active' : ''}" title="S2: 員/廠改善單"></div>
                    <div class="progress-step step-s3 ${s3 || s4 ? 'active' : ''}" title="S3: 廠商及員工"></div>
                    <div class="progress-step step-s4 ${s4 || isClosed ? 'active' : ''}" title="S4: 結案 (廠商)"></div>
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
            { key: 'S2員工查核檔案位置', label: 'S2 員工', class: 's2', icon: 'fa-user-shield' },
            { key: 'S2廠商查核檔案位置', label: 'S2 廠商', class: 's2', icon: 'fa-business-time' },
            { key: 'S3廠商及員工改善後核章檔案位置', label: 'S3', class: 's3', icon: 'fa-file-signature' },
            { key: 'S4結案檔案位置', label: 'S4 結(廠商)', class: 's4', icon: 'fa-check-double' }
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
        }).sort((a, b) => {
            const aClosed = a['辦理狀態'] === '第4階段-已結案';
            const bClosed = b['辦理狀態'] === '第4階段-已結案';
            if (aClosed && !bClosed) return 1;
            if (!aClosed && bClosed) return -1;
            
            const aDate = new Date(a['最晚應核章日期']).getTime();
            const bDate = new Date(b['最晚應核章日期']).getTime();
            return aDate - bDate;
        });
    },

    copyToClipboard: (elementId) => {
        const textToCopy = document.getElementById(elementId)?.innerText;
        if (!textToCopy) return;
        navigator.clipboard.writeText(textToCopy).then(() => {
            app.showToast("路徑已複製！", "success");
        }).catch(err => {
            console.error('Copy failed', err);
            app.showToast("複製失敗，請手動選取複製", "error");
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

    confirmS2Download: (url) => {
        if (!url || url === 'undefined') return app.showToast('無法取得 S2 檔案連結', 'error');
        if (confirm('確認要下載此案件的 S2 檔案嗎？')) {
            window.open(url, '_blank');
        }
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
            const hasFiles = !!(c['\u7b2c2\u968e\u6bb5\u9023\u7d50-\u54e1\u5de5'] || c['\u7b2c2\u968e\u6bb5\u9023\u7d50-\u5ee0\u5546'] || c['\u7b2c3\u968e\u6bb5\u9023\u7d50'] || c['\u7b2c4\u968e\u6bb5\u9023\u7d50-\u54e1\u5de5'] || c['\u7b2c4\u968e\u6bb5\u9023\u7d50-\u627f\u652c\u5546']);
            
            // 倒數計時區塊 (針對 DeptUploader 或 未登入訪客 且未結案)
            let countdownHtml = '';
            let isLargeDeptCard = false;
            let s2Url = c['\u7b2c2\u968e\u6bb5\u9023\u7d50-\u54e1\u5de5'] || c['\u7b2c2\u968e\u6bb5\u9023\u7d50-\u5ee0\u5546'] || c['\u7b2c2\u968e\u6bb5\u9023\u7d50']; // 相容舊版或優先顯示員工版
            const isGuest = !app.state.user;
            const isDept = app.state.user && app.state.user.role === 'DepartmentUploader';

            if ((isDept || isGuest) && !isClosed) {
                isLargeDeptCard = true;
                const diffDays = Math.ceil((new Date(c['最晚應核章日期']) - new Date(todayStr)) / (1000 * 60 * 60 * 24));
                const isUrgent = diffDays <= 7;
                const isCritical = diffDays <= 3;
                const urgencyClass = isCritical ? 'countdown-critical' : isUrgent ? 'countdown-urgent' : '';
                
                let downloadBtnHtml = '';
                if (isDept) {
                    downloadBtnHtml = `
                        <div style="font-size:0.8rem; font-weight:700; background:var(--primary); color:white; padding:5px 12px; border-radius:20px; margin-top:5px; cursor:pointer;">
                            <i class="fas fa-download"></i> 快速下載 S2 原始單
                        </div>`;
                }

                countdownHtml = `
                    <div class="countdown-hero ${urgencyClass}" style="margin-bottom:15px;" ${isDept ? `onclick="app.confirmS2Download('${s2Url}')"` : ''}>
                        <div class="label">⏳ 離結案期限還剩</div>
                        <div class="value countdown-days">${diffDays > 0 ? diffDays : '0'} 天</div>
                        ${diffDays <= 0 ? '<div style="color:var(--danger); font-size:0.8rem; font-weight:700;">⚠️ 已逾期！</div>' : ''}
                        ${downloadBtnHtml}
                    </div>
                `;
            }

            let stageClass = 'stage-s1';
            let statusStr = c['辦理狀態'] || '';
            if (statusStr.includes('4')) stageClass = 'stage-s4';
            else if (statusStr.includes('3')) stageClass = 'stage-s3';
            else if (statusStr.includes('2')) stageClass = 'stage-s2';
            
            const card = document.createElement('div');
            card.className = `case-card ${hasFiles ? 'has-files' : ''} ${isLargeDeptCard ? 'case-card-large' : ''} ${stageClass}`;
            if (isDept && !isClosed) {
                card.setAttribute('onclick', `app.confirmS2Download('${s2Url}')`);
            }
            card.setAttribute('data-dept', c['主辦部門'] || '');

            // 訪客模式極簡卡片
            if (isGuest) {
                card.innerHTML = `
                    ${countdownHtml}
                    <div class="card-header">
                        <h4>${snLabel}${c['工程簡稱']}</h4>
                        <span class="badge ${isOverdue ? 'warning' : app.getBadgeClass(c['辦理狀態'])}">${app.formatStatus(c['辦理狀態'])}</span>
                    </div>
                    <div class="card-body">
                        <div class="info-row"><i class="fas fa-building"></i> ${c['主辦部門']}</div>
                        <div class="info-row"><i class="fas fa-hard-hat"></i> ${c['承攬商']}</div>
                        <div class="info-row" style="${isOverdue ? 'color:var(--warning);font-weight:700;' : ''}"><i class="fas fa-clock"></i> 限辦：${c['最晚應核章日期']}</div>
                    </div>
                `;
            } else {
                card.innerHTML = `
                    ${countdownHtml}
                    <div class="card-header" onclick="event.stopPropagation()">
                        <h4 class="report-clickable" onclick="app.openManage('${c.id}')">${snLabel}${c['工程簡稱']}</h4>
                        <span class="badge ${isOverdue ? 'warning' : app.getBadgeClass(c['辦理狀態'])}">${app.formatStatus(c['辦理狀態'])}</span>
                    </div>
                    <div class="card-body">
                        <div class="info-row"><i class="fas fa-building"></i> ${c['主辦部門']}</div>
                        <div class="info-row"><i class="fas fa-hard-hat"></i> ${c['承攬商']}</div>
                        <div class="info-row"><i class="fas fa-calendar-alt"></i> 查核：${c['查核日期']}</div>
                        <div class="info-row"><i class="fas fa-user-tie"></i> 人員：${c['查核人員'] || '無'}</div>
                        ${c['承辦人姓名'] ? `<div class="info-row"><i class="fas fa-user"></i> 承辦：${c['承辦人姓名']} <span style="font-size:0.75rem;">${c['承辦人電子信箱'] ? '('+c['承辦人電子信箱']+')' : ''}</span></div>` : ''}
                        ${c['承辦課長職稱'] ? `<div class="info-row"><i class="fas fa-user-shield"></i> 課長：${c['承辦課長職稱']} <span style="font-size:0.75rem;">${c['承辦課長電子信箱'] ? '('+c['承辦課長電子信箱']+')' : ''}</span></div>` : ''}
                        <div class="info-row" style="${isOverdue ? 'color:var(--warning);font-weight:700;' : ''}"><i class="fas fa-clock"></i> 限辦：${c['最晚應核章日期']}</div>
                        ${c['結案日期'] ? `<div class="info-row" style="color:var(--success);font-weight:700;"><i class="fas fa-check"></i> 結案：${new Date(c['結案日期']).toISOString().split('T')[0]}</div>` : ''}
                        
                        ${app.state.systemMode === 'progress' ? app.getProgressHtml(c) : `
                            <div style="margin-top:10px; padding:10px; background:rgba(0,0,0,0.03); border-radius:10px; font-size:0.8rem;">
                                <i class="fas fa-exclamation-circle" style="color:var(--primary);"></i> 缺失數：${app.state.deficiencies.filter(d => String(d.caseId) === String(c.id)).length}
                            </div>
                        `}
                    </div>
                    <div class="card-footer" onclick="event.stopPropagation()">
                        <button class="btn btn-primary" onclick="app.openManage('${c.id}')">管理</button>
                        <button class="btn btn-outline" onclick="app.viewHistory('${c.id}')"><i class="fas fa-history"></i></button>
                        ${isAdmin ? `<button class="btn btn-outline" style="color:var(--warning); border-color:var(--warning);" onclick="app.deleteCase('${c.id}')"><i class="fas fa-trash"></i></button>` : ''}
                    </div>
                    <div style="padding: 12px 24px; border-top: 1px dashed var(--border);">
                        ${app.getFileStatusHtml(c)}
                    </div>
                `;
            }
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
            const hasFiles = !!(c['\u7b2c2\u968e\u6bb5\u9023\u7d50-\u54e1\u5de5'] || c['\u7b2c2\u968e\u6bb5\u9023\u7d50-\u5ee0\u5546'] || c['\u7b2c3\u968e\u6bb5\u9023\u7d50'] || c['\u7b2c4\u968e\u6bb5\u9023\u7d50-\u54e1\u5de5'] || c['\u7b2c4\u968e\u6bb5\u9023\u7d50-\u627f\u652c\u5546']);
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
                <td><span class="badge ${app.getBadgeClass(c['辦理狀態'])}">${app.formatStatus(c['辦理狀態'])}</span></td>
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
    fetchDeficiencies: async () => { try { const res = await api.getDeficiencies(); app.state.deficiencies = res.data; if(app.state.currentView === 'deficiencies') app.renderDeficiencies(); app.updateStats(); } catch(e){} },
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
                    <td>
                        <div style="display:flex; gap:5px;">
                            ${(isAdmin || app.state.user.role === 'SafetyUploader') ? `<button class="btn btn-outline" style="padding:8px 12px;" onclick="app.openEditProjModal('${p.serial}')"><i class="fas fa-edit"></i></button>` : ''}
                            ${isAdmin ? `<button class="btn btn-outline" style="color:var(--warning); border-color:var(--warning); padding:8px 12px;" onclick="app.deleteProject('${p.serial}')"><i class="fas fa-trash"></i></button>` : ''}
                        </div>
                    </td>
                </tr>
            `;
        });
    },
    openNewProjectModal: () => {
        app.openModal('新增工程項目', `
            <div style="display:flex;flex-direction:column;gap:15px; max-height:70vh; overflow-y:auto; padding:5px;">
                <input type="text" id="pAbbr" placeholder="工程簡稱 (例如: 台中電纜)">
                <input type="text" id="pName" placeholder="工程全名">
                <input type="text" id="pContractor" placeholder="承攬商名稱">
                <input type="text" id="pDept" placeholder="主辦部門">

                <h4 style="margin:10px 0 0 0; color:var(--primary); border-bottom:1px solid var(--border); padding-bottom:5px;">預設承辦聯絡資訊</h4>
                <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
                    <div><label>承辦人姓名</label><input type="text" id="pContractorName" placeholder="例如：李四"></div>
                    <div><label>信箱</label><input type="email" id="pContractorEmail" placeholder="li@example.com"></div>
                    <div><label>承辦課長職稱</label><input type="text" id="pManagerTitle" placeholder="例如：張課長"></div>
                    <div><label>信箱</label><input type="email" id="pManagerEmail" placeholder="zhang@example.com"></div>
                </div>

                <button class="btn btn-primary" onclick="app.submitNewProject()">確認新增</button>
            </div>
        `);
    },
    openEditProjModal: (serial) => {
        const p = app.state.projects.find(proj => proj.serial == serial);
        if (!p) {
            app.showToast("找不到工程項目資料", "error");
            return;
        }
        app.openModal('編輯工程項目與承辦資訊', `
            <div style="display:flex;flex-direction:column;gap:15px; max-height:70vh; overflow-y:auto; padding:5px;">
                <input type="text" id="editPAbbr" value="${p.abbr}" placeholder="工程簡稱">
                <input type="text" id="editPName" value="${p.name}" placeholder="工程全名">
                <input type="text" id="editPContractor" value="${p.contractor}" placeholder="承攬商名稱">
                <input type="text" id="editPDept" value="${p.department}" placeholder="主辦部門">
                
                <h4 style="margin:10px 0 0 0; color:var(--primary); border-bottom:1px solid var(--border); padding-bottom:5px;">承辦聯絡資訊 (帶入案件之預設值)</h4>
                <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
                    <div><label>承辦人姓名</label><input type="text" id="editPContractorName" value="${p.contractorName || ''}" placeholder="例如：李四"></div>
                    <div><label>信箱</label><input type="email" id="editPContractorEmail" value="${p.contractorEmail || ''}" placeholder="li@example.com"></div>
                    <div><label>承辦課長職稱</label><input type="text" id="editPManagerTitle" value="${p.contractorManagerTitle || ''}" placeholder="例如：張課長"></div>
                    <div><label>信箱</label><input type="email" id="editPManagerEmail" value="${p.contractorManagerEmail || ''}" placeholder="zhang@example.com"></div>
                </div>

                <button class="btn btn-primary" onclick="app.submitEditProject('${p.serial}')">儲存變更</button>
            </div>
        `);
    },
    submitEditProject: async (serial) => {
        const payload = {
            serial,
            abbr: document.getElementById('editPAbbr').value.trim(),
            name: document.getElementById('editPName').value.trim(),
            contractor: document.getElementById('editPContractor').value.trim(),
            department: document.getElementById('editPDept').value.trim(),
            contractorName: document.getElementById('editPContractorName').value.trim(),
            contractorEmail: document.getElementById('editPContractorEmail').value.trim(),
            contractorManagerTitle: document.getElementById('editPManagerTitle').value.trim(),
            contractorManagerEmail: document.getElementById('editPManagerEmail').value.trim()
        };

        if (!payload.abbr || !payload.name) return app.showToast("簡稱與全名必填", "error");

        app.setModalLoading(true);
        try {
            const res = await api.updateProject(payload);
            app.state.projects = res.projects;
            app.renderProjects();
            app.closeModal();
            app.showToast("✅ 工程儲存成功");
        } catch(e) { app.showToast(e.message, "error"); } finally { app.setModalLoading(false); }
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
        let cases = app.state.cases;
        const isDept = app.state.user && app.state.user.role === 'DepartmentUploader';
        
        let yearFilter = new Date().getFullYear().toString();
        const yearSelect = document.getElementById('filterYear');
        if (yearSelect) yearFilter = yearSelect.value;
        
        if (yearFilter) {
            cases = cases.filter(c => c['查核日期'] && c['查核日期'].startsWith(yearFilter));
        }

        if (isDept) {
            // renderDeptStats 需要從原始 state 拿到部門的所有案件，再由自身依 year 過濾統計
            const allDeptCases = app.state.cases.filter(c => c['主辦部門'] === app.state.user.department);
            cases = allDeptCases.filter(c => !yearFilter || (c['查核日期'] && c['查核日期'].startsWith(yearFilter)));
            app.renderDeptStats(allDeptCases, yearFilter);
        }

        document.getElementById('stat-total').innerText = cases.length;
        document.getElementById('stat-active').innerText = cases.filter(c => c['辦理狀態'] !== '第4階段-已結案').length;
        document.getElementById('stat-closed').innerText = cases.filter(c => c['辦理狀態'] === '第4階段-已結案').length;
        document.getElementById('stat-overdue').innerText = cases.filter(c => c['辦理狀態'] !== '第4階段-已結案' && c['最晚應核章日期'] < todayStr).length;

        // 更新卡片標題 (角色區分)
        const titles = document.querySelectorAll('.stat-info .title');
        if (titles.length >= 4) {
            if (isDept) {
                titles[0].innerText = '本隊查核數';
                titles[1].innerText = '本隊進行中';
                titles[2].innerText = '本隊已逾期';
                titles[3].innerText = '本隊已完成';
            } else {
                titles[0].innerText = '查核總數';
                titles[1].innerText = '進行中';
                titles[2].innerText = '已逾期';
                titles[3].innerText = '已結案';
            }
        }
    },

    formatStatus: (status) => {
        if (app.state.user && app.state.user.role === 'DepartmentUploader') {
            if (status === '第3階段-工作隊版已處理') return '已上傳 (待結案)';
            if (status === '第2階段-改善單已上傳') return '待處理 (步驟2)';
            if (status === '第4階段-已結案') return '已結案';
        }
        return status;
    },

    renderDeptStats: (deptCases, currentYear) => {
        const projStatsTable = document.getElementById('projectStatsTable');
        const caseDatesTable = document.getElementById('caseDatesTable');
        if (!projStatsTable || !caseDatesTable) return;

        // 依年份篩選（currentYear 空則為全部年份）
        const yearDeptCases = currentYear
            ? deptCases.filter(c => c['查核日期'] && c['查核日期'].startsWith(currentYear))
            : deptCases;

        // 工程統計：按工程分群，每次查核列出缺失數，再加總
        const projGroups = {};
        yearDeptCases.forEach(c => {
            const p = c['工程簡稱'];
            if (!projGroups[p]) projGroups[p] = [];
            const cid = String(c.id);
            const defCount = app.state.deficiencies.filter(d => String(d.caseId) === cid).length;
            projGroups[p].push({ date: c['查核日期'], defs: defCount });
        });

        let projHtml = '<table style="width:100%; border-collapse:collapse; text-align:center;">';
        projHtml += '<tr style="border-bottom:2px solid var(--border); color:var(--text-muted);"><th style="text-align:left; padding:6px;">工程</th><th style="width:80px;">查核日</th><th style="width:50px;">缺失</th></tr>';
        let totalChecks = 0, totalDefs = 0;
        Object.entries(projGroups).sort((a,b) => b[1].length - a[1].length).forEach(([proj, items]) => {
            const subDefs = items.reduce((s, i) => s + i.defs, 0);
            totalChecks += items.length;
            totalDefs += subDefs;
            items.forEach((item, idx) => {
                projHtml += `<tr style="border-bottom:1px solid var(--border);">
                    <td style="text-align:left; padding:4px 6px;">${idx === 0 ? '<b>' + proj + '</b>' : ''}</td>
                    <td style="font-size:0.8rem;">${item.date}</td>
                    <td>${item.defs}</td>
                </tr>`;
            });
            // 小計行
            projHtml += `<tr style="border-bottom:2px solid var(--primary); background:rgba(99,102,241,0.04);">
                <td style="text-align:left; padding:4px 6px; font-weight:700; color:var(--primary);">${proj} 小計</td>
                <td style="font-weight:700;">${items.length} 次</td>
                <td style="font-weight:700; color:var(--warning);">${subDefs}</td>
            </tr>`;
        });
        // 總計行
        projHtml += `<tr style="background:rgba(99,102,241,0.08);">
            <td style="text-align:left; padding:6px; font-weight:800; font-size:0.95rem; color:var(--primary);">合計</td>
            <td style="font-weight:800;">${totalChecks} 次</td>
            <td style="font-weight:800; color:var(--danger);">${totalDefs}</td>
        </tr>`;
        projHtml += '</table>';
        projStatsTable.innerHTML = projHtml;

        // 查核日 vs 結案日（所選年度全部，依查核日期排序）
        let dateHtml = '<table style="width:100%; border-collapse:collapse;">';
        dateHtml += '<tr style="border-bottom:1px solid var(--border); color:var(--text-muted);"><th>查核日</th><th>工程</th><th>狀態</th><th>結案日</th></tr>';
        [...yearDeptCases].sort((a,b) => (a['查核日期'] || '').localeCompare(b['查核日期'] || '')).forEach(c => {
            const isClosed = c['辦理狀態'] === '第4階段-已結案';
            const closeDateStr = c['結案日期'] ? new Date(c['結案日期']).toISOString().split('T')[0] : '-';
            dateHtml += `<tr style="border-bottom:1px solid var(--border);">
                <td style="padding:6px 0;">${c['查核日期']}</td>
                <td>${c['工程簡稱']}</td>
                <td style="color:${isClosed ? 'var(--success)' : 'var(--warning)'}; font-weight:700;">${isClosed ? '已結案' : '進行中'}</td>
                <td style="color:var(--success);">${closeDateStr}</td>
            </tr>`;
        });
        dateHtml += '</table>';
        caseDatesTable.innerHTML = dateHtml;

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

                    <div class="deadline-hero">
                        <div class="label"><i class="fas fa-clock"></i> 最晚應結案日期 (核章期限)</div>
                        <div class="value">${c['最晚應核章日期']}</div>
                        <div class="note">⚠️ 註：結案日期之判定依據為「受查單位經辦人員之核章日期」，請務必準時辦理。</div>
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

                    <div class="lite-step-card ${c['第2階段連結'] ? 'active' : ''}">
                        <span class="step-badge">步驟 2</span>
                        <h4><i class="fas fa-file-pdf"></i> 上傳第 3 階段核章版</h4>
                        <p>完成現場改善並核章後，請將掃描後的 PDF 檔案在此回傳。系統將自動通知工安組結案。</p>
                        ${c['第2階段連結'] && !c['第3階段連結'] ? 
                            app.getUploadSection(id, 'stage3', '立即上傳核章版 PDF', '#fbbf24', '', false) : ''}
                        ${c['第3階段連結'] ? 
                            `<a href="${c['第3階段連結']}" target="_blank" class="btn btn-primary" style="width:100%; height:50px; justify-content:center; border-radius:15px; margin-bottom:10px;"><i class="fas fa-download"></i> 下載已上傳的 S3 核章版</a>
                             ` + app.getUploadSection(id, 'stage3', '已上傳 - 點選可更換檔案', '#fbbf24', '', true).replace('padding:20px;', 'padding:10px;') : ''}
                        ${!c['第2階段連結'] ? 
                            `<div style="color:var(--text-muted); background:rgba(0,0,0,0.03); padding:15px; border-radius:12px; text-align:center; border:1px dashed var(--border);"><i class="fas fa-lock"></i> 目前尚未開放上傳 (需先完成步驟 1)</div>` : ''}
                    </div>

                    ${archivingHtml}
                    <button class="btn btn-outline" style="width:100%; margin-top:10px; height:50px; justify-content:center;" onclick="app.viewHistory('${id}')"><i class="fas fa-history"></i> 查看此案件歷史紀錄</button>
                </div>
            `;
            app.openModal('經辦人員作業程序', liteHtml);
        } else {
            // Admin / SafetyUploader 的完整管理介面
            const allMembers = new Set();
            app.state.cases.forEach(caseItem => {
                if (caseItem['查核成員']) {
                    caseItem['查核成員'].split(/[,、]/).forEach(m => allMembers.add(m.trim()));
                }
            });
            const datalistHtml = `<datalist id="auditMembersList">${Array.from(allMembers).filter(Boolean).map(m => `<option value="${m}"></option>`).join('')}</datalist>`;

            let html = `
                ${datalistHtml}
                <div class="tabs-container">
                    <div class="tabs-header">
                        <button class="tab-btn active" onclick="app.switchTab(event, 'tabFiles')"><i class="fas fa-folder-open"></i> 檔案管理</button>
                        <button class="tab-btn" onclick="app.switchTab(event, 'tabDefs')"><i class="fas fa-list-ul"></i> 缺失項目</button>
                        <button class="tab-btn" onclick="app.switchTab(event, 'tabInfo')"><i class="fas fa-edit"></i> 編輯案件資料</button>
                    </div>

                    <!-- 分頁一：檔案管理 -->
                    <div id="tabFiles" class="tab-content active">
                        <div style="margin-bottom:15px; padding:12px; background:rgba(99,102,241,0.05); border-radius:12px; font-size:0.8rem; border:1px solid rgba(99,102,241,0.1);">
                            <i class="fas fa-info-circle"></i> 當前狀態：<b style="color:var(--primary);">${c['辦理狀態']}</b>
                        </div>
                                                                        <div class="manage-grid">
                            ${app.getUploadSection(id, 'stage2e', 'S2 員工', 'var(--danger)', '請僅上傳「員工版本」之 S2 檔案。', !!c['S2員工查核檔案位置'])}
                            ${app.getUploadSection(id, 'stage2c', 'S2 廠商', 'var(--danger)', '注意：請僅上傳「承攬商版本」之 S2 檔案。', !!c['S2廠商查核檔案位置'])}
                            ${app.getUploadSection(id, 'stage3', 'S3 廠商及員工改善後', '#fbbf24', '受查部門核章版資料', !!c['S3廠商及員工改善後核章檔案位置'])}
                            ${app.getUploadSection(id, 'stage4c', 'S4 廠商結案', 'var(--success)', '廠商完成改善結案', !!c['S4結案檔案位置'])}
                        </div>
                        
                        ${(c['第2階段連結-員工'] || c['第2階段連結-廠商'] || c['第3階段連結'] || c['第4階段連結-員工'] || c['第4階段連結-承攬商']) ? `
                        <div style="margin-top:15px; padding:12px; background:rgba(0,0,0,0.03); border-radius:12px; border:1px solid var(--border);">
                            <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(80px, 1fr)); gap:8px;">
                                ${c['第2階段連結-員工'] ? `<a href="${c['第2階段連結-員工']}" target="_blank" class="btn btn-outline" style="font-size:0.7rem; justify-content:center;"><i class="fas fa-user"></i> S2員</a>` : ''}
                                ${c['第2階段連結-廠商'] ? `<a href="${c['第2階段連結-廠商']}" target="_blank" class="btn btn-outline" style="font-size:0.7rem; justify-content:center;"><i class="fas fa-industry"></i> S2廠</a>` : ''}
                                ${c['第3階段連結'] ? `<a href="${c['第3階段連結']}" target="_blank" class="btn btn-outline" style="font-size:0.7rem; justify-content:center;"><i class="fas fa-stamp"></i> S3</a>` : ''}
                                ${c['第4階段連結-員工'] ? `<a href="${c['第4階段連結-員工']}" target="_blank" class="btn btn-outline" style="font-size:0.7rem; justify-content:center;"><i class="fas fa-user-check"></i> S4員</a>` : ''}
                                ${c['第4階段連結-承攬商'] ? `<a href="${c['第4階段連結-承攬商']}" target="_blank" class="btn btn-outline" style="font-size:0.7rem; justify-content:center;"><i class="fas fa-building-circle-check"></i> S4廠</a>` : ''}
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
                    
                    <!-- 分頁三：編輯案件資料 -->
                    <div id="tabInfo" class="tab-content" style="max-height:60vh; overflow-y:auto; padding:5px;">
                        <h4 style="margin:0 0 10px 0; color:var(--primary); border-bottom:1px solid var(--border); padding-bottom:5px;">查核人員資訊</h4>
                        <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px; margin-bottom:15px;">
                            <div><label>填表人</label><input type="text" id="editInspector" value="${c['填表人']||c['查核人員']||''}"></div>
                            <div><label>查核領隊</label><input type="text" id="editAuditLeader" value="${c['查核領隊']||''}"></div>
                            <div style="grid-column:1/-1"><label>查核成員</label><input type="text" id="editAuditMembers" list="auditMembersList" value="${c['查核成員']||''}"></div>
                        </div>

                        <h4 style="margin:0 0 10px 0; color:var(--primary); border-bottom:1px solid var(--border); padding-bottom:5px;">承辦聯絡資訊</h4>
                        <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px; margin-bottom:15px;">
                            <div><label>承辦人員姓名</label><input type="text" id="editContractorName" value="${c['承辦人員姓名']||c['承辦人姓名']||''}"></div>
                            <div><label>承辦人Email</label><input type="text" id="editContractorEmail" value="${c['承辦人Email']||c['承辦人電子信箱']||''}"></div>
                            <div><label>承辦課長姓名</label><input type="text" id="editContractorManagerTitle" value="${c['承辦課長姓名']||c['承辦課長職稱']||''}"></div>
                            <div><label>課長Email</label><input type="text" id="editContractorManagerEmail" value="${c['課長Email']||c['承辦課長電子信箱']||''}"></div>
                        </div>

                        <h4 style="margin:0 0 10px 0; color:var(--danger); border-bottom:1px solid var(--border); padding-bottom:5px;">案件狀態資訊</h4>
                        <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px; margin-bottom:15px;">
                            <div><label>結案日期</label><input type="date" id="editCloseDate" value="${c['結案日期'] ? new Date(c['結案日期']).toISOString().split('T')[0] : ''}"></div>
                        </div>

                        <button class="btn btn-primary" style="width:100%; justify-content:center; margin-top:10px;" onclick="app.submitEditCaseInfo('${id}')">儲存資料變更</button>
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
    submitEditCaseInfo: async (caseId) => {
        const details = {
            填表人: document.getElementById('editInspector').value.trim(),
            auditLeader: document.getElementById('editAuditLeader').value.trim(),
            auditMembers: document.getElementById('editAuditMembers').value.trim(),
            承辦人員姓名: document.getElementById('editContractorName').value.trim(),
            承辦人Email: document.getElementById('editContractorEmail').value.trim(),
            承辦課長姓名: document.getElementById('editContractorManagerTitle').value.trim(),
            課長Email: document.getElementById('editContractorManagerEmail').value.trim(),
            closeDate: document.getElementById('editCloseDate').value || null
        };

        app.setModalLoading(true);
        try {
            const res = await api.updateCase(caseId, details, app.state.user.email);
            app.state.cases = res.records;
            app.updateStats();
            app.renderView();
            app.showToast("✅ 案件資料更新成功");
        } catch (e) {
            app.showToast(e.message, "error");
        } finally {
            app.setModalLoading(false);
        }
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
    getUploadSection: (id, stage, label, color, note = '', exists = false) => {
        const isS2 = stage === 'stage2';
        return `
        <div class="upload-section ${exists ? 'has-file' : ''}" style="border-left: 5px solid ${color || 'var(--border)'}; position:relative;">
            ${exists ? `<span style="position:absolute; top:8px; right:8px; font-size:0.6rem; color:var(--success); background:rgba(16,185,129,0.1); padding:2px 6px; border-radius:4px;"><i class="fas fa-check"></i> 已存在</span>` : ''}
            <div class="upload-header" style="color:${color || 'inherit'}">
                <i class="fas fa-cloud-upload-alt"></i> ${label}
            </div>
            
            ${isS2 ? `
            <div class="upload-note" style="color:#ef4444; font-weight:700; background:rgba(239,68,68,0.05); padding:6px; border-radius:6px; border:1px dashed #fca5a5; margin:4px 0;">
                <i class="fas fa-exclamation-triangle"></i> 注意：請僅上傳「承攬商版本」之 S2 檔案
            </div>
            ` : ''}
            
            ${note ? `<p class="upload-note">${note}</p>` : ''}
            <div class="upload-actions">
                <input type="file" id="file_${stage}" style="width:100%; margin-bottom:12px; font-size:0.8rem;" />
                <button class="btn" style="width:100%; justify-content:center; background:${color || 'var(--primary)'}; color:white;" onclick="app.submitFile('${id}', '${stage}', ${exists})">
                    ${exists ? '替換現有檔案' : '確認上傳存檔'}
                </button>
            </div>
        </div>
    `;
    },
    submitFile: async (id, stage, isReplace = false) => {
        const input = document.getElementById(`file_${stage}`);
        if(!input.files.length) return app.showToast("請先選擇檔案", "error");

        // S2 上傳完整性驗證
        if (stage === 'stage2e' || stage === 'stage2c') {
            const c = app.state.cases.find(x => x['案件ID'] === id);
            if (c) {
                const missing = [];
                if (!c['查核領隊']) missing.push('查核領隊');
                if (!c['承辦人員姓名'] && !c['承辦人姓名']) missing.push('承辦人員姓名');
                if (!c['承辦人Email'] && !c['承辦人電子信箱']) missing.push('承辦人Email');
                if (!c['最晚應核章日期']) missing.push('最晚應核章日期');
                if (missing.length > 0) {
                    alert(`請填報人完成資料填寫。

尚未填寫欄位：${missing.join('、')}

【提醒】結案日期需待 S3廠商 改善完成後，以其「第一個核章日期」為準方可填列。`);
                    return;
                }
            }
        }
        
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
    confirmS2Download: (url) => {
        if (!url) return app.showToast("原始單尚未準備好", "warning");
        if (confirm("⚠️ 確認要下載此案件的 S2 檔案嗎？\n\n下載後請依規定時程辦理改善及核章。")) {
            const link = document.createElement('a');
            link.href = url;
            link.target = "_blank";
            link.click();
        }
    },
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
        
        const allMembers = new Set();
        app.state.cases.forEach(c => {
            if (c['查核成員']) {
                c['查核成員'].split(/[,、]/).forEach(m => allMembers.add(m.trim()));
            }
        });
        const datalistHtml = `<datalist id="auditMembersList">${Array.from(allMembers).filter(Boolean).map(m => `<option value="${m}"></option>`).join('')}</datalist>`;
        
        const html = `
            ${datalistHtml}
            <div style="display:flex;flex-direction:column;gap:15px; max-height:65vh; overflow-y:auto; padding:5px;">
                <div><label>工程：</label><select id="newProj" style="width:100%">${options}</select></div>
                <div><label>日期：</label><input type="date" id="newDate" value="${new Date().toISOString().split('T')[0]}" style="width:100%"></div>
                
                <h4 style="margin:10px 0 0 0; color:var(--primary); border-bottom:1px solid var(--border); padding-bottom:5px;">查核人員資訊</h4>
                <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
                    <div><label>查核領隊</label><input type="text" id="newAuditLeader" value="${(app.state.user ? app.state.user.email.split('@')[0] : '')}" placeholder="例如：王小明"></div>
                    <div><label>查核成員</label><input type="text" id="newAuditMembers" list="auditMembersList" placeholder="例如：陳大毛"></div>
                </div>

                <h4 style="margin:10px 0 0 0; color:var(--primary); border-bottom:1px solid var(--border); padding-bottom:5px;">承辦聯絡資訊</h4>
                <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">
                    <div><label>承辦人姓名</label><input type="text" id="newContractorName" placeholder="例如：李四"></div>
                    <div><label>承辦人電子信箱</label><input type="email" id="newContractorEmail" placeholder="li@example.com"></div>
                    <div><label>承辦課長職稱</label><input type="text" id="newContractorManagerTitle" placeholder="例如：張課長"></div>
                    <div><label>承辦課長電子信箱</label><input type="email" id="newContractorManagerEmail" placeholder="zhang@example.com"></div>
                </div>

                <button class="btn btn-primary" style="margin-top:10px; justify-content:center;" onclick="app.submitNewCase()">確認登錄</button>
            </div>
        `;
        app.openModal('登錄查核案件', html);
    },
    submitNewCase: async () => {
        const pAbbr = document.getElementById('newProj').value; 
        const date = document.getElementById('newDate').value;
        const auditLeader = document.getElementById('newAuditLeader').value.trim();
        const auditMembers = document.getElementById('newAuditMembers').value.trim();
        const contractorName = document.getElementById('newContractorName').value.trim();
        const contractorEmail = document.getElementById('newContractorEmail').value.trim();
        const contractorManagerTitle = document.getElementById('newContractorManagerTitle').value.trim();
        const contractorManagerEmail = document.getElementById('newContractorManagerEmail').value.trim();

        if (!pAbbr || !date) return app.showToast("工程與日期為必填", "error");
        
        app.setModalLoading(true);
        const pInfo = app.state.projects.find(p => p.abbr === pAbbr);
        try { 
            const payload = { 
                ...pInfo, 
                auditDate: date,
                auditLeader,
                auditMembers,
                contractorName,
                contractorEmail,
                contractorManagerTitle,
                contractorManagerEmail
            };
            const res = await api.createCase(payload); 
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
        const reason = prompt("請先輸入案件刪除理由（記錄至歷程）：");
        if (reason === null) return;
        if (!reason.trim()) return app.showToast("刪除案件必須輸入理由", "error");

        if (!confirm(`您輸入的刪除理由為：「${reason.trim()}」\n確定要刪除此案件？此操作無法恢復！`)) return;

        app.showLoading(true);
        try {
            const res = await api.deleteCase(id, reason.trim());
            app.state.cases = res.records || [];
            app.updateStats();
            app.renderView();
            app.showToast("案件已刪除");
        } catch (e) {
            app.showToast(e.message, "error");
        } finally { app.showLoading(false); }
    },

    initSystem: async () => {
        app.showLoading(true);
        try {
            const res = await api.getSystemMetadata();
            const meta = res.data;
            
            app.openModal('系統初始化管理 (雙模式切換)', `
                <div style="background:rgba(219,234,254,0.5); padding:16px; border-radius:12px; margin-bottom:20px; border:1px solid #bfdbfe;">
                    <h4 style="margin:0 0 10px 0; color:var(--primary); font-size:1rem;"><i class="fas fa-info-circle"></i> 系統當前資訊</h4>
                    <div style="display:grid; grid-template-columns:1fr 1fr; gap:10px; font-size:0.85rem;">
                        <div><b>基準年度：</b> <span style="color:var(--danger);">${meta.baselineYear}</span></div>
                        <div><b>操作者：</b> ${app.state.user.email.split('@')[0]}</div>
                        <div><b>備份狀態：</b> ${meta.lastBackup}</div>
                        <div><b>影響範圍：</b> 案件 ${meta.caseCount} 筆 / 缺失 ${meta.deficiencyCount} 筆</div>
                    </div>
                </div>

                <div class="init-mode-selector" style="display:flex; flex-direction:column; gap:12px;">
                    <label style="border:2px solid var(--border); padding:16px; border-radius:12px; cursor:pointer; display:block; transition:0.2s;" id="labelModeA">
                        <input type="radio" name="initMode" value="column_sync" checked style="margin-right:10px;" onchange="document.getElementById('resetWarning').classList.add('hidden')">
                        <b style="font-size:1rem;">模式 A：欄位檢查 / 同步</b>
                        <p style="margin:5px 0 0 25px; font-size:0.8rem; color:var(--text-muted);">僅檢查並補齊各分頁的標準標題列，<span style="color:var(--success);">不會影響現有案件資料。</span></p>
                    </label>

                    <label style="border:2px solid var(--border); padding:16px; border-radius:12px; cursor:pointer; display:block; transition:0.2s;" id="labelModeB">
                        <input type="radio" name="initMode" value="reset" style="margin-right:10px;" onchange="document.getElementById('resetWarning').classList.remove('hidden')">
                        <b style="font-size:1rem;">模式 B：資料匯出與全重置</b>
                        <p style="margin:5px 0 0 25px; font-size:0.8rem; color:var(--text-muted);">執行【自動備份】後，<span style="color:var(--danger);">清空所有案件、缺失與歷程</span>，將系統過渡至新年度 (${meta.baselineYear})。</p>
                    </label>
                </div>

                <div id="resetWarning" class="hidden" style="margin-top:15px; padding:12px; background:rgba(239,68,68,0.1); border-radius:8px; border:1px solid var(--danger); font-size:0.8rem; color:var(--danger);">
                    <i class="fas fa-exclamation-triangle"></i> 模式 B 為破壞性操作。系統將先複製一份當前資料至「System_Backups」資料夾，隨後清空正式環境資料。
                </div>

                <button class="btn btn-primary" onclick="app.submitInitSystem()" style="width:100%; margin-top:20px; justify-content:center; height:50px; font-size:1.1rem;">
                    <i class="fas fa-play-circle"></i> 執行作業
                </button>
            `);
        } catch (e) {
            app.showToast("無法獲取系統資訊: " + e.message, "error");
        } finally {
            app.showLoading(false);
        }
    },

    submitInitSystem: async () => {
        const mode = document.querySelector('input[name="initMode"]:checked').value;
        const modeText = mode === 'reset' ? '【資料匯出與重置】' : '【欄位檢查/同步】';
        
        if (mode === 'reset') {
            if (!confirm(`警告：您選擇了「${modeText}」。\n這將備份並清空所有資料！確定繼續？`)) return;
            const input = prompt("請輸入「RESET」以確認執行全系統重置：");
            if (input !== "RESET") return app.showToast("取消操作", "error");
        } else {
            if (!confirm(`確定執行「${modeText}」作業？這將掃描全部分頁標題並修補缺失欄位。`)) return;
        }

        app.setModalLoading(true);
        try {
            const res = await api.setupSystem(mode, "115年度");
            app.showToast(res.message, "success");
            app.closeModal();
            // 如果是重置，刷新畫面
            if (mode === 'reset') {
                app.state.cases = [];
                app.state.deficiencies = [];
                app.renderView();
                app.updateStats();
            }
        } catch (e) {
            app.showToast("執行失敗: " + e.message, "error");
        } finally {
            app.setModalLoading(false);
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
        const departments = Array.from(new Set(app.state.cases.map(c => c['主辦部門']).filter(Boolean)));
        const deptOptions = departments.map(d => `<option value="${d}">${d}</option>`).join('');

        app.openModal('統計與分析', `
            <div style="display:flex; flex-direction:column; gap:15px;">
                <div class="report-field-group">
                    <div><label>起始日期：</label><input type="date" id="reportStart" value="" style="width:100%;"></div>
                    <div><label>結束日期：</label><input type="date" id="reportEnd" value="" style="width:100%;"></div>
                </div>
                <div class="report-field-group">
                    <div>
                        <label>主辦部門：</label>
                        <select id="reportDept" style="width:100%;">
                            <option value="">-- 全部部門 --</option>
                            ${deptOptions}
                        </select>
                    </div>
                    <div>
                        <label>篩選工程：</label>
                        <select id="reportProj" style="width:100%;">
                            <option value="">-- 全部工程 --</option>
                            ${projectOptions}
                        </select>
                    </div>
                </div>
                <div class="report-field-group">
                    <div>
                        <label>結案狀態：</label>
                        <select id="reportStatus" style="width:100%;">
                            <option value="">-- 全部狀態 --</option>
                            <option value="已結案">僅列已結案</option>
                            <option value="進行中">僅列進行中</option>
                        </select>
                    </div>
                    <div>
                        <label>缺失內容關鍵字查詢：</label>
                        <input type="text" id="reportKeyword" placeholder="關鍵字 (如: 漏電, 安全帽...)" style="width:100%;">
                    </div>
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
        const deptFilter = document.getElementById('reportDept').value;
        const projAbbr = document.getElementById('reportProj').value;
        const statusFilter = document.getElementById('reportStatus').value;
        const keyword = document.getElementById('reportKeyword').value.trim().toLowerCase();

        // 1. 先篩選案件
        let filteredCases = app.state.cases.filter(c => {
            const date = c['查核日期'];
            if (!date) return false;
            const isDateMatch = (!start || date >= start) && (!end || date <= end);
            const isDeptMatch = deptFilter === '' || c['主辦部門'] === deptFilter;
            const isProjMatch = projAbbr === '' || c['工程簡稱'] === projAbbr;
            
            let isStatusMatch = true;
            if (statusFilter === '已結案') isStatusMatch = c['辦理狀態'] === '第4階段-已結案';
            if (statusFilter === '進行中') isStatusMatch = c['辦理狀態'] !== '第4階段-已結案';

            return isDateMatch && isDeptMatch && isProjMatch && isStatusMatch;
        });

        // 2. 獲取篩選案件的缺失，並應用關鍵字過濾
        let filteredDefs = app.state.deficiencies.filter(d => filteredCases.find(c => c.id == d.caseId));
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
    },

    // ==================== 測試信件 ====================
    triggerTestEmail: async () => {
        if (!confirm('確認發送三封測試信件至後端設定的 TEST_EMAIL？')) return;
        try {
            app.showToast('發送測試信...');
            const res = await api.testSendEmail();
            app.showToast(res.message, 'success');
        } catch (e) {
            app.showToast(e.message, 'error');
        }
    },

    // ==================== 手動稽催 ====================
    triggerDailyReminder: async () => {
        if (!confirm('手動執行三階段稽催？將透過实際信箱發送通知。')) return;
        try {
            const btn = document.getElementById('btnRunReminder');
            if (btn) { btn.disabled = true; btn.innerText = '執行中...'; }
            const res = await api.runDailyReminder();
            app.showToast(res.message, 'success');
            if (res.errors && res.errors.length) {
                console.warn('稽催錯誤:', res.errors);
                app.showToast(`有 ${res.errors.length} 筆發送失敗，請查看 Console`, 'error');
            }
        } catch (e) {
            app.showToast(e.message, 'error');
        } finally {
            const btn = document.getElementById('btnRunReminder');
            if (btn) { btn.disabled = false; btn.innerHTML = '<i class="fas fa-bell"></i> 執行稽催'; }
        }
    },

    // ==================== 建立觸發器 ====================
    triggerSetupTrigger: async () => {
        if (!confirm('將在 GAS 建立「每日上午8:00」自動稽催觸發器（已存在則艥過）。確定執行？')) return;
        try {
            const res = await api.setupTrigger();
            app.showToast(res.message, 'success');
        } catch (e) {
            app.showToast(e.message, 'error');
        }
    },

    // ==================== 帳號管理 ====================
    renderDeptAccounts: async () => {
        const tbody = document.getElementById('deptAccountListBody');
        if (!tbody) return;
        tbody.innerHTML = '<tr><td colspan="8" style="text-align:center; padding:30px;"><i class="fas fa-spinner fa-spin"></i> 載入中...</td></tr>';
        try {
            const res = await api.getDeptAccounts();
            const accounts = res.data || [];
            if (!accounts.length) {
                tbody.innerHTML = '<tr><td colspan="8" style="text-align:center; padding:40px; color:var(--text-muted);">\u5c1a無帳號資料</td></tr>';
                return;
            }
            tbody.innerHTML = accounts.map(a => `
                <tr>
                    <td><strong>${a.deptName}</strong></td>
                    <td>${a.contractorName || '-'}</td>
                    <td><a href="mailto:${a.contractorEmail}" style="color:var(--primary);">${a.contractorEmail || '-'}</a></td>
                    <td>${a.managerName || '-'}</td>
                    <td><a href="mailto:${a.managerEmail}" style="color:var(--primary);">${a.managerEmail || '-'}</a></td>
                    <td style="font-size:0.8rem; color:var(--text-muted);">${a.note || ''}</td>
                    <td style="font-size:0.8rem;">${a.createdAt || ''}</td>
                    <td>
                        <button class="btn btn-outline" style="padding:4px 10px; font-size:0.8rem;"
                            onclick="app.openRegisterDeptAccountModal(${JSON.stringify(a).replace(/"/g, '&quot;')})">
                            <i class="fas fa-edit"></i> 編輯
                        </button>
                    </td>
                </tr>
            `).join('');
        } catch (e) {
            tbody.innerHTML = `<tr><td colspan="8" style="text-align:center; padding:30px; color:var(--danger);">讀取失敗: ${e.message}</td></tr>`;
        }
    },

    openRegisterDeptAccountModal: (existing) => {
        const e = existing || {};
        app.openModal('新增 / 更新主辦部門帳號', `
            <p style="font-size:0.85rem; color:var(--text-muted); margin-bottom:16px;">
                <i class="fas fa-info-circle"></i> 同一部門名稱偵測到時自動覆蓋旧資料。
            </p>
            <div style="display:grid; gap:12px;">
                <label>部門名稱 *
                    <input id="daName" class="form-input" value="${e.deptName || ''}" placeholder="例：工務部" required>
                </label>
                <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
                    <label>承辦人姓名
                        <input id="daCtName" class="form-input" value="${e.contractorName || ''}" placeholder="姓名">
                    </label>
                    <label>承辦人 Email *
                        <input id="daCtEmail" class="form-input" type="email" value="${e.contractorEmail || ''}" placeholder="example@company.com">
                    </label>
                </div>
                <div style="display:grid; grid-template-columns:1fr 1fr; gap:12px;">
                    <label>課長姓名
                        <input id="daMgrName" class="form-input" value="${e.managerName || ''}" placeholder="姓名">
                    </label>
                    <label>課長 Email
                        <input id="daMgrEmail" class="form-input" type="email" value="${e.managerEmail || ''}" placeholder="example@company.com">
                    </label>
                </div>
                <label>備註
                    <input id="daNote" class="form-input" value="${e.note || ''}" placeholder="選填">
                </label>
                <button class="btn btn-primary" onclick="app.submitRegisterDeptAccount()">
                    <i class="fas fa-save"></i> 儲存
                </button>
            </div>
        `);
    },

    submitRegisterDeptAccount: async () => {
        const deptName = document.getElementById('daName')?.value.trim();
        const contractorEmail = document.getElementById('daCtEmail')?.value.trim();
        if (!deptName || !contractorEmail) {
            app.showToast('部門名稱與承辦人 Email 為必填欄位', 'error'); return;
        }
        try {
            app.setModalLoading(true);
            const res = await api.registerDeptAccount({
                deptName,
                contractorName: document.getElementById('daCtName')?.value.trim(),
                contractorEmail,
                managerName: document.getElementById('daMgrName')?.value.trim(),
                managerEmail: document.getElementById('daMgrEmail')?.value.trim(),
                note: document.getElementById('daNote')?.value.trim()
            });
            app.showToast(res.message, 'success');
            app.closeModal();
            app.renderDeptAccounts();
        } catch (e) {
            app.showToast(e.message, 'error');
        } finally {
            app.setModalLoading(false);
        }
    }
};

window.onload = () => app.initAuth();
