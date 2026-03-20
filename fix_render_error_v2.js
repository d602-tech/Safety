const fs = require('fs');
const path = require('path');

const filePath = 'd:/AI/GI01SafetyWalk/frontend/js/app.js';
let content = fs.readFileSync(filePath, 'utf8');

// 1. Fix getProgressHtml
const progressHtml = `    getProgressHtml: (c) => {
        const s2e = !!c['第2階段連結-員工'];
        const s2c = !!c['第2階段連結-廠商'];
        const s3 = !!c['第3階段連結'];
        const s4e = !!c['第4階段連結-員工'];
        const s4c = !!c['第4階段連結-承攬商'];
        const isClosed = c['辦理狀態'] === '第4階段-已結案';

        return \`
            <div class="progress-container">
                <div class="progress-label">
                    <span>階段進度</span>
                    <span>\\\${isClosed ? '100%' : ((s4e && s4c) ? '80%' : (s4e || s4c ? '70%' : (s3 ? '50%' : (s2e || s2c ? '25%' : '0%'))))}</span>
                </div>
                <div class="progress-track">
                    <div class="progress-step step-s1 active" title="S1: 已登錄"></div>
                    <div class="progress-step step-s2 \\\${s2e || s2c || s3 || s4e || s4c ? 'active' : ''}" title="S2: 員/廠改善單"></div>
                    <div class="progress-step step-s3 \\\${s3 || s4e || s4c ? 'active' : ''}" title="S3: 廠商及員工"></div>
                    <div class="progress-step step-s4 \\\${s4e || s4c || isClosed ? 'active' : ''}" title="S4: 結案 (員/廠)"></div>
                </div>
            </div>
        \`;
    },`;

// 2. Fix getFileStatusHtml
const fileStatusHtml = `    getFileStatusHtml: (c) => {
        const canAccess = app.state.user && (
            app.state.user.role === 'Admin' || 
            app.state.user.role === 'SafetyUploader' || 
            (app.state.user.role === 'DepartmentUploader' && c['主辦部門'] === app.state.user.department)
        );

        const stages = [
            { key: '第2階段連結-員工', label: 'S2 員工', class: 's2', icon: 'fa-user' },
            { key: '第2階段連結-廠商', label: 'S2 廠商', class: 's2', icon: 'fa-industry' },
            { key: '第3階段連結', label: 'S3', class: 's3', icon: 'fa-stamp' },
            { key: '第4階段連結-員工', label: 'S4 結(員工)', class: 's4', icon: 'fa-user-check' },
            { key: '第4階段連結-承攬商', label: 'S4 結(廠商)', class: 's4', icon: 'fa-building-circle-check' }
        ];

        return \`
            <div class="file-status-icons">
                \\\${stages.map(s => {
                    const url = c[s.key];
                    if (url && canAccess) {
                        return \\\`<a href="\\\${url}" target="_blank" class="file-icon uploaded \\\${s.class}" title="已上傳 \\\${s.label}"><i class="fas \\\${s.icon}"></i> \\\${s.label}</a>\\\`;
                    } else if (url && !canAccess) {
                        return \\\`<div class="file-icon uploaded" title="已上傳，但您無權下載" style="color:var(--text-muted);"><i class="fas fa-lock"></i> \\\${s.label}</div>\\\`;
                    } else {
                        return \\\`<div class="file-icon missing" title="\\\${s.label} 尚未上傳"><i class="fas \\\${s.icon}"></i> \\\${s.label}</div>\\\`;
                    }
                }).join('')}
            </div>
        \`;
    },`;

// 3. Fix manage-grid
const manageGridHtml = `<div class="manage-grid">
                            \\\${app.getUploadSection(id, 'stage2e', 'S2 員工', 'var(--danger)', '請僅上傳「員工版本」之 S2 檔案。', !!c['第2階段連結-員工'])}
                            \\\${app.getUploadSection(id, 'stage2c', 'S2 廠商', 'var(--danger)', '注意：請僅上傳「承攬商版本」之 S2 檔案。', !!c['第2階段連結-廠商'])}
                            \\\${app.getUploadSection(id, 'stage3', 'S3 廠商及員工', '#fbbf24', '受查部門核章版資料', !!c['第3階段連結'])}
                            \\\${app.getUploadSection(id, 'stage4e', 'S4 結(員工)', 'var(--success)', '承辦人完成結案報告', !!c['第4階段連結-員工'])}
                            \\\${app.getUploadSection(id, 'stage4c', 'S4 結(廠商)', 'var(--success)', '廠商完成改善結案', !!c['第4階段連結-承攬商'])}
                        </div>`;

// 4. Fix buttons block
const buttonsBlockHtml = `\\\${(c['第2階段連結-員工'] || c['第2階段連結-廠商'] || c['第3階段連結'] || c['第4階段連結-員工'] || c['第4階段連結-承攬商']) ? \\\`
                        <div style="margin-top:15px; padding:12px; background:rgba(0,0,0,0.03); border-radius:12px; border:1px solid var(--border);">
                            <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(80px, 1fr)); gap:8px;">
                                \\\${c['第2階段連結-員工'] ? \\\`<a href="\\\${c['第2階段連結-員工']}" target="_blank" class="btn btn-outline" style="font-size:0.7rem; justify-content:center;"><i class="fas fa-user"></i> S2 員工</a>\\\` : ''}
                                \\\${c['第2階段連結-廠商'] ? \\\`<a href="\\\${c['第2階段連結-廠商']}" target="_blank" class="btn btn-outline" style="font-size:0.7rem; justify-content:center;"><i class="fas fa-industry"></i> S2 廠商</a>\\\` : ''}
                                \\\${c['第3階段連結'] ? \\\`<a href="\\\${c['第3階段連結']}" target="_blank" class="btn btn-outline" style="font-size:0.7rem; justify-content:center;"><i class="fas fa-stamp"></i> S3</a>\\\` : ''}
                                \\\${c['第4階段連結-員工'] ? \\\`<a href="\\\${c['第4階段連結-員工']}" target="_blank" class="btn btn-outline" style="font-size:0.7rem; justify-content:center;"><i class="fas fa-user-check"></i> S4 結(員)</a>\\\` : ''}
                                \\\${c['第4階段連結-承攬商'] ? \\\`<a href="\\\${c['第4階段連結-承攬商']}" target="_blank" class="btn btn-outline" style="font-size:0.7rem; justify-content:center;"><i class="fas fa-building-circle-check"></i> S4 結(廠)</a>\\\` : ''}
                            </div>
                        </div>\\\` : ''}`;

// Apply replacements (using proper JS regex flags)
content = content.replace(/getProgressHtml: \(c\) => \{[\s\S]*?\n    \},/s, progressHtml);
content = content.replace(/getFileStatusHtml: \(c\) => \{[\s\S]*?\n    \},/s, fileStatusHtml);
content = content.replace(/<div class="manage-grid">[\s\S]*?<\/div>/s, manageGridHtml);

const buttonsRegex = /\\\$\{\(c\['.*?'\] \|\| c\['.*?'\] \|\| c\['.*?'\] \|\| c\['.*?'\] \|\| c\['.*?'\]\).*?<\/div>\\\` : ''\}/s;
content = content.replace(buttonsRegex, buttonsBlockHtml);

fs.writeFileSync(filePath, content, 'utf8');
console.log('Successfully fixed app.js formatting and labels (take 2)');
