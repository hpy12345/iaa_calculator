/*
 * IAA休闲游戏成本收益测算 - 前端交互逻辑（修复 & 优化版）
 * 主要修复与优化点：
 * - 增加安全的 DOM 获取函数，避免 null 导致的错误
 * - collectAllData 对各种输入方式增加容错（当未选择 radio 时不会抛错）
 * - displayExcelPreview 中不再强制把标签显示为百分比（更通用）
 * - downloadChart 使用 append/remove 保证兼容性
 * - 移除无用语句，统一部分默认值与防御式编程
 * - 保持原有结构与功能，尽量减少侵入式重构，便于替换
 */

// ========================================
// 全局变量
// ========================================
let periodCounter = 0;
let roiData = { type: 'manual', points: [] };
let retentionData = { type: 'manual', points: [] };
let calculationResults = null;
let dauChart = null;
let financeChart = null;
document.addEventListener('input', (e) => {
    if (e.target.closest('input, textarea, select')) {
        markUnsaved();
    }
});

document.addEventListener('change', (e) => {
    if (e.target.closest('input, textarea, select')) {
        markUnsaved();
    }
});

// 小工具：安全获取元素
function $id(id) {
    const el = document.getElementById(id);
    if (!el) {
        // 仅在开发时输出警告，避免污染生产日志
        // console.warn(`Element #${id} not found`);
    }
    return el;
}

// 小工具：安全 querySelector
function $qs(selector, root = document) {
    try {
        return root.querySelector(selector);
    } catch (e) {
        return null;
    }
}

// ========================================
// 页面初始化
// ========================================
document.addEventListener('DOMContentLoaded', function() {
    // 可在此初始化一些事件（如果需要）
});

// ========================================
// 项目创建 / 加载 / 删除
// ========================================
function createNewProject() {
    const nameInput = $id('newProjectName');
    const errorElement = $id('newProjectNameError');
    const confirmBtn = $id('confirmNewProjectBtn');

    if (nameInput) nameInput.value = '';
    if (errorElement) {
        errorElement.style.display = 'none';
        errorElement.textContent = '';
    }
    if (confirmBtn) confirmBtn.disabled = false;

    const dialog = $id('newProjectDialog');
    if (dialog) dialog.style.display = 'flex';

    setTimeout(() => {
        if (nameInput) nameInput.focus();
    }, 100);
}

function closeNewProjectDialog() {
    const dialog = $id('newProjectDialog');
    if (dialog) dialog.style.display = 'none';
}

async function confirmNewProject() {
    const nameInput = $id('newProjectName');
    const errorElement = $id('newProjectNameError');
    const confirmBtn = $id('confirmNewProjectBtn');
    const projectName = nameInput ? nameInput.value.trim() : '';

    if (errorElement) {
        errorElement.style.display = 'none';
        errorElement.textContent = '';
    }

    if (!projectName) {
        if (errorElement) {
            errorElement.textContent = '请输入项目名称';
            errorElement.style.display = 'block';
        }
        if (nameInput) nameInput.focus();
        return;
    }

    if (projectName.length > 50) {
        if (errorElement) {
            errorElement.textContent = '项目名称不能超过50个字符';
            errorElement.style.display = 'block';
        }
        if (nameInput) nameInput.focus();
        return;
    }

    if (confirmBtn) {
        confirmBtn.disabled = true;
        confirmBtn.innerHTML = '<span class="icon">⏳</span> 验证中...';
    }

    try {
        const response = await fetch('/list_projects');
        const result = await response.json();

        if (result && result.success) {
            const projects = result.projects || [];
            const existingProject = projects.find(p => p.name === projectName);

            if (existingProject) {
                if (errorElement) {
                    errorElement.textContent = '该项目名称已存在，请使用其他名称';
                    errorElement.style.display = 'block';
                }
                if (nameInput) nameInput.focus();
                if (confirmBtn) {
                    confirmBtn.disabled = false;
                    confirmBtn.innerHTML = '<span class="icon">✓</span> 确认创建';
                }
                return;
            }
        }

        await initializeAndSaveProject(projectName);
    } catch (error) {
        if (errorElement) {
            errorElement.textContent = '验证失败：' + (error && error.message ? error.message : String(error));
            errorElement.style.display = 'block';
        }
        if (confirmBtn) {
            confirmBtn.disabled = false;
            confirmBtn.innerHTML = '<span class="icon">✓</span> 确认创建';
        }
    }
}

async function initializeAndSaveProject(projectName) {
    const confirmBtn = $id('confirmNewProjectBtn');

    if (confirmBtn) confirmBtn.innerHTML = '<span class="icon">⏳</span> 创建中...';

    closeNewProjectDialog();

    const welcome = $id('welcomeSection');
    const main = $id('mainContent');
    if (welcome) welcome.style.display = 'none';
    if (main) main.style.display = '';

    const projectNameInput = $id('projectName');
    if (projectNameInput) projectNameInput.value = projectName;

    const periodsContainer = $id('investmentPeriods');
    if (periodsContainer) periodsContainer.innerHTML = '';
    periodCounter = 0;

    // 重置全局变量
    roiData = { type: 'manual', points: [] };
    retentionData = { type: 'manual', points: [] };
    calculationResults = null;

    // ROI：手动
    const roiManualRadio = $qs('input[name="roiInputType"][value="manual"]');
    if (roiManualRadio) roiManualRadio.checked = true;
    toggleRoiInput('manual');
    const roiManualBody = $id('roiManualBody');
    if (roiManualBody) roiManualBody.innerHTML = '';
    const roiExcelPreview = $id('roiExcelPreview');
    if (roiExcelPreview) roiExcelPreview.innerHTML = '';
    const roiFileInput = $id('roiExcelFile');
    if (roiFileInput) roiFileInput.value = '';

    // 添加默认值
    addRoiRow(1, 0.5);
    addRoiRow(7, 3.5);
    addRoiRow(30, 15);
    addRoiRow(60, 30);
    addRoiRow(90, 45);

    // retention
    const retentionManualRadio = $qs('input[name="retentionInputType"][value="manual"]');
    if (retentionManualRadio) retentionManualRadio.checked = true;
    toggleRetentionInput('manual');
    const retentionManualBody = $id('retentionManualBody');
    if (retentionManualBody) retentionManualBody.innerHTML = '';
    const retentionExcelPreview = $id('retentionExcelPreview');
    if (retentionExcelPreview) retentionExcelPreview.innerHTML = '';
    const retentionFileInput = $id('retentionExcelFile');
    if (retentionFileInput) retentionFileInput.value = '';

    addRetentionRow(0, 100);
    addRetentionRow(1, 45);
    addRetentionRow(7, 25);
    addRetentionRow(30, 12);
    addRetentionRow(60, 8);
    addRetentionRow(90, 6);

    // 其他设置
    const repayEl = $id('repaymentMonths');
    const targetEl = $id('targetDau');
    if (repayEl) repayEl.value = 1;
    if (targetEl) targetEl.value = 500;

    const resultsSection = $id('resultsSection');
    if (resultsSection) resultsSection.style.display = 'none';
    const exportCsvBtn = $id('exportCsvBtn');
    if (exportCsvBtn) exportCsvBtn.disabled = true;

    addInvestmentPeriod(null, null, true);

    await saveProject();
    markSaved();

    if (confirmBtn) {
        confirmBtn.disabled = false;
        confirmBtn.innerHTML = '<span class="icon">✓</span> 确认创建';
    }
}

function backToWelcome() {
    if (hasUnsavedChanges()) {
        const ok = confirm(
            '当前有未保存的修改。\n\n返回首页将丢失这些修改，是否继续？'
        );
        if (!ok) return;
    }

    const welcome = $id('welcomeSection');
    const main = $id('mainContent');
    if (welcome) welcome.style.display = '';
    if (main) main.style.display = 'none';

    const results = $id('resultsSection');
    if (results) results.style.display = 'none';
}

// ========================================
// 保存 / 加载 / 删除 项目
// ========================================
async function saveProject(isAutoSave = false) {
    const projectNameEl = $id('projectName');
    const projectName = projectNameEl ? projectNameEl.value.trim() : '';
    if (!projectName) {
        alert('项目名称不能为空！');
        return;
    }

    const projectData = collectAllData();

    try {
        const response = await fetch('/save_project', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(projectData)
        });

        const result = await response.json();

        if (result && result.success && !isAutoSave) {
            alert(`项目 "${projectName}" 已保存成功！`);
        }

        if (result && result.success) {
            markSaved();   // 👈 新增
        }
    } catch (error) {
        alert(`保存失败：${error && error.message ? error.message : String(error)}`);
    }
}

async function showLoadDialog() {
    if (hasUnsavedChanges()) {
        const ok = confirm(
            '当前项目有未保存的修改。\n\n继续加载其他项目将丢失这些修改，是否继续？'
        );
        if (!ok) return;
    }

    const projectList = $id('projectList');
    if (projectList) projectList.innerHTML = '<li style="text-align: center; color: #666;">正在加载...</li>';
    const dialog = $id('loadProjectDialog');
    if (dialog) dialog.style.display = 'flex';

    try {
        const response = await fetch('/list_projects');
        const result = await response.json();

        if (!projectList) return;
        projectList.innerHTML = '';

        if (!result || !result.success) {
            projectList.innerHTML = `<li style="text-align: center; color: #f00;">加载失败：${result ? result.error : '未知错误'}</li>`;
            return;
        }

        const projects = result.projects || [];
        if (projects.length === 0) {
            projectList.innerHTML = '<li style="text-align: center; color: #666;">暂无已保存的项目</li>';
        } else {
            projects.forEach(project => {
                const li = document.createElement('li');
                const savedTime = project.savedAt ? new Date(project.savedAt).toLocaleString('zh-CN') : '';
                li.innerHTML = `
                    <div class="project-info">
                        <span class="project-name">${project.name}</span>
                        <span class="project-time">${savedTime}</span>
                    </div>
                    <div class="project-actions">
                        <button class="btn-delete">删除</button>
                    </div>
                `;

                // 点击加载
                li.addEventListener('click', () => loadProject(project.name));

                // 删除按钮
                const delBtn = li.querySelector('.btn-delete');
                if (delBtn) {
                    delBtn.addEventListener('click', (ev) => {
                        ev.stopPropagation();
                        deleteProject(project.name);
                    });
                }

                projectList.appendChild(li);
            });
        }
    } catch (error) {
        if (projectList) projectList.innerHTML = `<li style="text-align: center; color: #f00;">加载失败：${error && error.message ? error.message : String(error)}</li>`;
    }
}

function closeLoadDialog() {
    const dialog = $id('loadProjectDialog');
    if (dialog) dialog.style.display = 'none';
}

async function loadProject(projectName) {
    try {
        const response = await fetch('/load_project', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ project_name: projectName })
        });

        const result = await response.json();
        if (!result || !result.success) {
            alert(`加载失败：${result ? result.error : '未知错误'}`);
            return;
        }

        const projectData = result.data || {};

        const projectNameEl = $id('projectName');
        if (projectNameEl) projectNameEl.value = projectName;

        const periodsContainer = $id('investmentPeriods');
        if (periodsContainer) periodsContainer.innerHTML = '';
        periodCounter = 0;

        if (projectData.investment_periods && projectData.investment_periods.length > 0) {
            projectData.investment_periods.forEach(period => addInvestmentPeriod(period, null, true));
        } else {
            addInvestmentPeriod(null, null, true);
        }

        // ROI 数据
        if (projectData.roi_data) {
            const roiType = projectData.roi_data.type || 'manual';
            const roiRadio = $qs(`input[name="roiInputType"][value="${roiType}"]`);
            if (roiRadio) roiRadio.checked = true;
            toggleRoiInput(roiType);

            if (roiType === 'manual') {
                const manualBody = $id('roiManualBody');
                if (manualBody) manualBody.innerHTML = '';
                if (projectData.roi_data.points && projectData.roi_data.points.length > 0) {
                    projectData.roi_data.points.forEach(point => addRoiRow(point.day, point.value));
                }
                roiData.points = [];
                const preview = $id('roiExcelPreview');
                if (preview) preview.innerHTML = '';
                const fileInput = $id('roiExcelFile');
                if (fileInput) fileInput.value = '';
            } else {
                roiData.points = projectData.roi_data.points || [];
                updateRoiExcelPreview(roiData.points);
                const manualBody = $id('roiManualBody');
                if (manualBody) manualBody.innerHTML = '';
            }
        } else {
            roiData.points = [];
            const preview = $id('roiExcelPreview');
            if (preview) preview.innerHTML = '';
            const manualBody = $id('roiManualBody');
            if (manualBody) manualBody.innerHTML = '';
        }

        // retention 数据
        if (projectData.retention_data) {
            const retentionType = projectData.retention_data.type || 'manual';
            const retentionRadio = $qs(`input[name="retentionInputType"][value="${retentionType}"]`);
            if (retentionRadio) retentionRadio.checked = true;
            toggleRetentionInput(retentionType);

            if (retentionType === 'manual') {
                const manualBody = $id('retentionManualBody');
                if (manualBody) manualBody.innerHTML = '';
                if (projectData.retention_data.points && projectData.retention_data.points.length > 0) {
                    projectData.retention_data.points.forEach(point => addRetentionRow(point.day, point.value));
                }
                retentionData.points = [];
                const preview = $id('retentionExcelPreview');
                if (preview) preview.innerHTML = '';
                const fileInput = $id('retentionExcelFile');
                if (fileInput) fileInput.value = '';
            } else {
                retentionData.points = projectData.retention_data.points || [];
                updateRetentionExcelPreview(retentionData.points);
                const manualBody = $id('retentionManualBody');
                if (manualBody) manualBody.innerHTML = '';
            }
        } else {
            retentionData.points = [];
            const preview = $id('retentionExcelPreview');
            if (preview) preview.innerHTML = '';
            const manualBody = $id('retentionManualBody');
            if (manualBody) manualBody.innerHTML = '';
        }

        // 其他设置
        const repayEl = $id('repaymentMonths');
        const targetEl = $id('targetDau');
        if (repayEl && projectData.repayment_months) repayEl.value = projectData.repayment_months;
        if (targetEl && projectData.target_dau) targetEl.value = projectData.target_dau;

        closeLoadDialog();

        const welcome = $id('welcomeSection');
        const main = $id('mainContent');
        if (welcome) welcome.style.display = 'none';
        if (main) main.style.display = '';

        const resultsSection = $id('resultsSection');
        if (resultsSection) resultsSection.style.display = 'none';
        const exportCsvBtn = $id('exportCsvBtn');
        if (exportCsvBtn) exportCsvBtn.disabled = true;

        alert(`项目 "${projectName}" 已加载成功！`);
        markSaved();
    } catch (error) {
        alert(`加载失败：${error && error.message ? error.message : String(error)}`);
    }
}

async function deleteProject(projectName) {
    if (!confirm(`确定要删除项目 "${projectName}" 吗？`)) return;

    try {
        const response = await fetch('/delete_project', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ project_name: projectName })
        });
        const result = await response.json();
        if (result && result.success) await showLoadDialog();
    } catch (error) {
        alert(`删除失败：${error && error.message ? error.message : String(error)}`);
    }
}

// ========================================
// 投资时间段管理
// ========================================
function addInvestmentPeriod(data = null, insertBeforeElement = null, skipAutoSave = false) {
    periodCounter++;
    const container = $id('investmentPeriods');
    if (!container) return;

    const periodCard = document.createElement('div');
    periodCard.className = 'period-card';
    periodCard.id = `period-${periodCounter}`;

    let defaultStart;
    const existingPeriods = container.querySelectorAll('.period-card');

    if (data && data.start) {
        defaultStart = data.start;
    } else if (existingPeriods.length > 0) {
        const lastPeriod = existingPeriods[existingPeriods.length - 1];
        const lastEndDate = $qs('.period-end', lastPeriod)?.value || '';
        if (lastEndDate) {
            const nextDay = new Date(lastEndDate);
            nextDay.setDate(nextDay.getDate() + 1);
            defaultStart = nextDay.toISOString().split('T')[0];
        } else {
            defaultStart = new Date().toISOString().split('T')[0];
        }
    } else {
        defaultStart = new Date().toISOString().split('T')[0];
    }

    let defaultEnd;
    if (data && data.end) {
        defaultEnd = data.end;
    } else {
        const startDate = new Date(defaultStart);
        startDate.setMonth(startDate.getMonth() + 2);
        defaultEnd = startDate.toISOString().split('T')[0];
    }

    const costType = (data && data.cost_type) || 'fixed';
    const costValue = (data && data.cost_value) || 10;
    const costStart = (data && data.cost_start) || 5;
    const costEnd = (data && data.cost_end) || 15;

    periodCard.innerHTML = `
        <div class="period-header">
            <span class="period-title">时间段 ${periodCounter}</span>
            <div class="period-header-actions">
                <button type="button" class="btn btn-insert" data-action="insert-before">⬆ 在此之前插入</button>
                <button type="button" class="btn btn-remove" data-action="remove">删除</button>
            </div>
        </div>
        <div class="period-grid">
            <div class="input-group">
                <label>起始时间</label>
                <input type="date" class="period-start" value="${defaultStart}">
            </div>
            <div class="input-group">
                <label>结束时间</label>
                <input type="date" class="period-end" value="${defaultEnd}">
            </div>
            <div class="input-group">
                <label>DNU (万人)</label>
                <input type="number" class="period-dnu" value="${data?.dnu || 1}" min="0" step="0.1">
            </div>
            <div class="input-group">
                <label>团队规模 (人)</label>
                <input type="number" class="period-team-size" value="${data?.team_size || 10}" min="1" step="1">
            </div>
            <div class="input-group">
                <label>用工成本 (万/人/天)</label>
                <input type="number" class="period-labor-cost" value="${data?.labor_cost || 0.025}" min="0" step="0.01">
            </div>
            <div class="input-group">
                <label>其他运营成本 (万元/天)</label>
                <input type="number" class="period-other-cost" value="${data?.other_cost || 0.1}" min="0" step="0.1">
            </div>
            <div class="input-group cost-type-inline">
                <label>日消耗 (万元)</label>
                <div class="cost-type-wrapper">
                    <div class="cost-inputs-inline">
                        <input type="number" class="period-cost-value cost-fixed-inline" value="${costValue}" min="0" step="0.1" placeholder="日消耗值" ${costType === 'linear' ? 'style="display:none"' : ''}>
                        <input type="number" class="period-cost-start cost-linear-inline" value="${costStart}" min="0" step="0.1" placeholder="初始值" ${costType === 'fixed' ? 'style="display:none"' : ''}>
                        <input type="number" class="period-cost-end cost-linear-inline" value="${costEnd}" min="0" step="0.1" placeholder="最终值" ${costType === 'fixed' ? 'style="display:none"' : ''}>
                    </div>
                    <div class="radio-group-inline">
                        <label class="radio-label">
                            <input type="radio" name="costType-${periodCounter}" value="fixed" ${costType === 'fixed' ? 'checked' : ''}>
                            定值
                        </label>
                        <label class="radio-label">
                            <input type="radio" name="costType-${periodCounter}" value="linear" ${costType === 'linear' ? 'checked' : ''}>
                            线性变化
                        </label>
                    </div>
                </div>
            </div>
        </div>
    `;

    // 将节点插入到容器中
    if (insertBeforeElement && insertBeforeElement.parentNode === container) {
        container.insertBefore(periodCard, insertBeforeElement);
    } else {
        container.appendChild(periodCard);
    }

    // 事件代理：插入/删除按钮
    const insertBtn = periodCard.querySelector('[data-action="insert-before"]');
    const removeBtn = periodCard.querySelector('[data-action="remove"]');

    if (insertBtn) {
        insertBtn.addEventListener('click', () => insertPeriodBefore(periodCard));
    }
    if (removeBtn) {
        removeBtn.addEventListener('click', () => removePeriodByElement(periodCard));
    }

    // Radio 切换事件
    const radios = periodCard.querySelectorAll(`input[name="costType-${periodCounter}"]`);
    radios.forEach(r => {
        r.addEventListener('change', (ev) => toggleCostType(periodCard.id, ev.target.value));
    });

    renumberPeriods();

    if (!skipAutoSave) saveProject(true);
}

// 新版：通过元素插入新时间段（传入目标元素），避免依赖 id 数字
function insertPeriodBefore(targetPeriodElement) {
    if (!targetPeriodElement || !targetPeriodElement.parentNode) return;

    const targetStartDate = $qs('.period-start', targetPeriodElement)?.value || '';

    let newEndDate;
    let newStartDate;

    if (targetStartDate) {
        const endDate = new Date(targetStartDate);
        endDate.setDate(endDate.getDate() - 1);
        newEndDate = endDate.toISOString().split('T')[0];

        const startDate = new Date(endDate);
        startDate.setMonth(startDate.getMonth() - 2);
        newStartDate = startDate.toISOString().split('T')[0];
    } else {
        const today = new Date();
        newEndDate = today.toISOString().split('T')[0];
        const threeMonthsAgo = new Date(today);
        threeMonthsAgo.setMonth(threeMonthsAgo.getMonth() - 2);
        newStartDate = threeMonthsAgo.toISOString().split('T')[0];
    }

    const newPeriodData = {
        start: newStartDate,
        end: newEndDate,
        dnu: 1,
        team_size: 10,
        labor_cost: 0.025,
        other_cost: 0.1,
        cost_type: 'fixed',
        cost_value: 10,
        cost_start: 5,
        cost_end: 15
    };

    addInvestmentPeriod(newPeriodData, targetPeriodElement);
}

function renumberPeriods() {
    const periods = document.querySelectorAll('.period-card');
    periods.forEach((period, index) => {
        const titleElement = period.querySelector('.period-title');
        if (titleElement) titleElement.textContent = `时间段 ${index + 1}`;
    });
}

function removePeriodByElement(periodElement) {
    if (!periodElement) return;
    periodElement.remove();

    if (document.querySelectorAll('.period-card').length === 0) {
        addInvestmentPeriod();
    } else {
        renumberPeriods();
        saveProject(true);
    }
}

// 旧版兼容：按 id 删除
function removePeriod(id) {
    const period = $id(`period-${id}`);
    if (period) removePeriodByElement(period);
}

function toggleCostType(periodIdOrElement, type) {
    let periodCard = null;
    if (typeof periodIdOrElement === 'string') {
        periodCard = $id(periodIdOrElement);
    } else if (periodIdOrElement instanceof Element) {
        periodCard = periodIdOrElement;
    }
    if (!periodCard) return;

    const fixedInputs = periodCard.querySelectorAll('.cost-fixed-inline');
    const linearInputs = periodCard.querySelectorAll('.cost-linear-inline');

    if (type === 'fixed') {
        fixedInputs.forEach(el => el.style.display = '');
        linearInputs.forEach(el => el.style.display = 'none');
    } else {
        fixedInputs.forEach(el => el.style.display = 'none');
        linearInputs.forEach(el => el.style.display = '');
    }
}

// ========================================
// ROI 管理
// ========================================
function toggleRoiInput(type) {
    const manual = $id('roiManualInput');
    const excel = $id('roiExcelInput');
    if (manual) manual.style.display = type === 'manual' ? '' : 'none';
    if (excel) excel.style.display = type === 'excel' ? '' : 'none';
    roiData.type = type;
    markUnsaved();
}

function updateRoiExcelPreview(points) {
    const previewData = (points || []).slice(0, 30);
displayExcelPreview('roiExcelPreview', previewData, 'ROI (%)', points ? points.length : 0);
}

function addRoiRow(day = '', value = '') {
    const tbody = $id('roiManualBody');
    if (!tbody) return;
    const row = document.createElement('tr');
    row.innerHTML = `
        <td><input type="number" class="roi-day" value="${day}" min="1" step="1" placeholder="天数"></td>
        <td><input type="number" class="roi-value" value="${value}" min="0" step="0.1" placeholder="ROI值"></td>
        <td><button type="button" class="btn btn-remove">删除</button></td>
    `;
    const delBtn = row.querySelector('.btn-remove');
    if (delBtn) delBtn.addEventListener('click', () => {
                row.remove();
                markUnsaved();});
    tbody.appendChild(row);
    markUnsaved();
}

function handleRoiExcel(event) {
    const file = event?.target?.files ? event.target.files[0] : null;
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

            roiData.points = [];
            const previewData = [];

            for (let i = 0; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (row && row.length >= 2) {
                    const day = parseInt(row[0]);
                    const value = parseFloat(row[1]);
                    if (!isNaN(day) && !isNaN(value) && day > 0) {
                        roiData.points.push({ day, value });
                        if (previewData.length < 30) previewData.push({ day, value });
                    }
                }
            }

            displayExcelPreview('roiExcelPreview', previewData, 'ROI (%)', roiData.points.length);
            markUnsaved();
        } catch (error) {
            alert('Excel文件解析失败：' + (error && error.message ? error.message : String(error)));
        }
    };
    reader.readAsArrayBuffer(file);
}

// ========================================
// 留存率 管理
// ========================================
function toggleRetentionInput(type) {
    const manual = $id('retentionManualInput');
    const excel = $id('retentionExcelInput');
    if (manual) manual.style.display = type === 'manual' ? '' : 'none';
    if (excel) excel.style.display = type === 'excel' ? '' : 'none';
    retentionData.type = type;
    markUnsaved();
}

function updateRetentionExcelPreview(points) {
    const previewData = (points || []).slice(0, 30);
displayExcelPreview('retentionExcelPreview', previewData, '留存率 (%)', points ? points.length : 0);
}

function addRetentionRow(day = '', value = '') {
    const tbody = $id('retentionManualBody');
    if (!tbody) return;
    const row = document.createElement('tr');
    row.innerHTML = `
        <td><input type="number" class="retention-day" value="${day}" min="1" step="1" placeholder="天数"></td>
        <td><input type="number" class="retention-value" value="${value}" min="0" max="100" step="0.1" placeholder="留存率"></td>
        <td><button type="button" class="btn btn-remove">删除</button></td>
    `;
    const delBtn = row.querySelector('.btn-remove');
    if (delBtn) delBtn.addEventListener('click', () => {
                row.remove();
                markUnsaved();});
    tbody.appendChild(row);
    markUnsaved();
}

function handleRetentionExcel(event) {
    const file = event?.target?.files ? event.target.files[0] : null;
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

            retentionData.points = [];
            const previewData = [];

            for (let i = 0; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (row && row.length >= 2) {
                    const day = parseInt(row[0]);
                    const value = parseFloat(row[1]);
                    if (!isNaN(day) && !isNaN(value) && day > 0) {
                        retentionData.points.push({ day, value });
                        if (previewData.length < 30) previewData.push({ day, value });
                    }
                }
            }

        displayExcelPreview('retentionExcelPreview', previewData, '留存率 (%)', retentionData.points.length);
        markUnsaved();
        } catch (error) {
            alert('Excel文件解析失败：' + (error && error.message ? error.message : String(error)));
        }
    };
    reader.readAsArrayBuffer(file);
}

function displayExcelPreview(containerId, data, label, totalCount = null) {
    const container = $id(containerId);
    if (!container) return;

    if (!data || data.length === 0) {
        container.innerHTML = '<p style="padding: 10px; color: #666;">未找到有效数据</p>';
        return;
    }

    let html = `
        <table>
            <thead>
                <tr>
                    <th>天数</th>
                    <th>${label}</th>
                </tr>
            </thead>
            <tbody>
    `;

    data.forEach(item => {
        const v = (typeof item.value === 'number') ? item.value.toFixed(2) : item.value;
        html += `<tr><td>${item.day}</td><td>${v}</td></tr>`;
    });

    html += '</tbody></table>';

    const total = totalCount !== null ? totalCount : data.length;
    if (total > 30) {
        html += `<p style="padding: 10px; color: #666; font-size: 0.85rem;">共导入 ${total} 条数据，预览前 30 条</p>`;
    } else {
        html += `<p style="padding: 10px; color: #666; font-size: 0.85rem;">共导入 ${total} 条数据</p>`;
    }

    container.innerHTML = html;
}

// ========================================
// 收集数据
// ========================================
function collectAllData() {
    const data = {
        project_name: ($id('projectName')?.value || '').trim(),
        investment_periods: [],
        roi_data: { type: 'manual', points: [] },
        retention_data: { type: 'manual', points: [] },
        repayment_months: parseInt($id('repaymentMonths')?.value) || 1,
        target_dau: parseInt($id('targetDau')?.value) || 1000
    };

    document.querySelectorAll('.period-card').forEach(card => {
        const periodId = (card.id || '').split('-')[1] || '';
        const costTypeEl = document.querySelector(`input[name="costType-${periodId}"]:checked`);
        const costType = costTypeEl ? costTypeEl.value : 'fixed';

        const period = {
            start: $qs('.period-start', card)?.value || '',
            end: $qs('.period-end', card)?.value || '',
            cost_type: costType,
            cost_value: parseFloat($qs('.period-cost-value', card)?.value) || 0,
            cost_start: parseFloat($qs('.period-cost-start', card)?.value) || 0,
            cost_end: parseFloat($qs('.period-cost-end', card)?.value) || 0,
            dnu: parseFloat($qs('.period-dnu', card)?.value) || 0,
            team_size: parseInt($qs('.period-team-size', card)?.value) || 0,
            labor_cost: parseFloat($qs('.period-labor-cost', card)?.value) || 0,
            other_cost: parseFloat($qs('.period-other-cost', card)?.value) || 0
        };

        data.investment_periods.push(period);
    });

    const roiInputTypeEl = document.querySelector('input[name="roiInputType"]:checked');
    const roiInputType = roiInputTypeEl ? roiInputTypeEl.value : (roiData?.type || 'manual');
    data.roi_data.type = roiInputType;

    if (roiInputType === 'manual') {
        document.querySelectorAll('#roiManualBody tr').forEach(row => {
            const day = parseInt(row.querySelector('.roi-day')?.value);
            const value = parseFloat(row.querySelector('.roi-value')?.value);
            if (!isNaN(day) && !isNaN(value)) data.roi_data.points.push({ day, value });
        });
    } else {
        data.roi_data.points = (roiData && roiData.points) ? roiData.points : [];
    }

    const retentionInputTypeEl = document.querySelector('input[name="retentionInputType"]:checked');
    const retentionInputType = retentionInputTypeEl ? retentionInputTypeEl.value : (retentionData?.type || 'manual');
    data.retention_data.type = retentionInputType;

    if (retentionInputType === 'manual') {
        document.querySelectorAll('#retentionManualBody tr').forEach(row => {
            const day = parseInt(row.querySelector('.retention-day')?.value);
            const value = parseFloat(row.querySelector('.retention-value')?.value);
            if (!isNaN(day) && !isNaN(value)) data.retention_data.points.push({ day, value });
        });
    } else {
        data.retention_data.points = (retentionData && retentionData.points) ? retentionData.points : [];
    }

    return data;
}

// ========================================
// 计算与结果展示
// ========================================
async function startCalculation() {
    // 计算前强制自动保存一次
    await saveProject(true);

    const data = collectAllData();

    if (data.investment_periods.length === 0) {
        alert('请至少添加一个投资时间段');
        return;
    }

    if (!data.roi_data || data.roi_data.points.length === 0) {
        alert('请输入ROI数据');
        return;
    }

    if (!data.retention_data || data.retention_data.points.length === 0) {
        alert('请输入留存率数据');
        return;
    }

    const loading = $id('loadingOverlay');
    if (loading) loading.style.display = 'flex';

    try {
        const response = await fetch('/calculate', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });

        if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);

        const results = await response.json();
        if (results.error) throw new Error(results.error);

        calculationResults = results;
        displayResults(results);

        const exportCsvBtn = $id('exportCsvBtn');
        if (exportCsvBtn) exportCsvBtn.disabled = false;

        const resultsSection = $id('resultsSection');
        if (resultsSection) resultsSection.style.display = '';
        $id('resultsSection')?.scrollIntoView({ behavior: 'smooth' });
    } catch (error) {
        alert('计算失败：' + (error && error.message ? error.message : String(error)));
    } finally {
        if (loading) loading.style.display = 'none';
    }
}

function displayResults(results) {
    const metrics = results.key_metrics || {};

    $id('metricMaxCash1m').textContent = formatCash(metrics.max_cash_demand_1m);
    $id('metricMaxCash2m').textContent = formatCash(metrics.max_cash_demand_2m);
    $id('metricDynamicBreakeven').textContent = formatDays(metrics.dynamic_profit_breakeven_day);
    $id('metricCumulativeBreakeven').textContent = formatDays(metrics.cumulative_profit_breakeven_day);
    $id('metricCashFlow1mBreakeven').textContent = formatDays(metrics.cumulative_cash_flow_1m_breakeven_day);
    $id('metricCashFlow2mBreakeven').textContent = formatDays(metrics.cumulative_cash_flow_2m_breakeven_day);
    $id('metricDay10mDau').textContent = formatDays(metrics.day_to_10m_dau);
    $id('metricDay2mDau').textContent = formatDays(metrics.day_to_2m_dau);
    $id('metricDayTargetDau').textContent = formatDays(metrics.day_to_target_dau);
    $id('metricTargetDauValue').textContent = metrics.target_dau || '--';

    const repaymentMonths = metrics.repayment_months || 1;
    if (repaymentMonths > 2) {
        $id('metricMaxCashNmCard').style.display = '';
        $id('metricCashFlowNmBreakevenCard').style.display = '';
        $id('metricMaxCashNmLabel').textContent = `${repaymentMonths}个月回款最大现金需求`;
        $id('metricCashFlowNmBreakevenLabel').textContent = `${repaymentMonths}个月回款现金流打正天数`;
        $id('metricMaxCashNm').textContent = formatCash(metrics.max_cash_demand_nm);
        $id('metricCashFlowNmBreakeven').textContent = formatDays(metrics.cumulative_cash_flow_nm_breakeven_day);
    } else {
        $id('metricMaxCashNmCard').style.display = 'none';
        $id('metricCashFlowNmBreakevenCard').style.display = 'none';
    }

    updateCharts(results.charts || {});
    updateQuarterlyTable(results.quarterly_table_data || [], repaymentMonths);
}

function formatCash(value) {
    if (value === null || value === undefined) return '--';
    if (typeof value !== 'number') return String(value);
    return value.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatNumber(value) {
    if (value === null || value === undefined) return '--';
    if (typeof value !== 'number') return String(value);
    return value.toLocaleString('zh-CN', { maximumFractionDigits: 2 });
}

function formatDays(value) {
    if (value === null || value === undefined || value < 0) return '未达成';
    return String(value);
}

function updateCharts(chartsData) {
    chartsData = chartsData || {};

    // DAU
    const dauCtxEl = $id('dauChart');
    if (dauCtxEl) {
        const dauCtx = dauCtxEl.getContext('2d');
        if (dauChart) dauChart.destroy();

        dauChart = new Chart(dauCtx, {
            type: 'line',
            data: {
                labels: (chartsData.dau_quarterly && chartsData.dau_quarterly.labels) || [],
                datasets: [{
                    label: 'DAU (万人)',
                    data: (chartsData.dau_quarterly && chartsData.dau_quarterly.data) || [],
                    borderColor: '#667eea',
                    backgroundColor: 'rgba(102, 126, 234, 0.1)',
                    borderWidth: 3,
                    fill: true,
                    tension: 0.4,
                    pointRadius: 6,
                    pointBackgroundColor: '#667eea',
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    legend: { display: true, position: 'top' },
                    tooltip: { callbacks: { label: function(context) { return `DAU: ${context.parsed.y.toLocaleString()} 万人`; } } }
                },
                scales: { y: { beginAtZero: true, title: { display: true, text: 'DAU (万人)' } } }
            }
        });
    }

    // Finance
    const financeCtxEl = $id('financeChart');
    if (financeCtxEl) {
        const financeCtx = financeCtxEl.getContext('2d');
        if (financeChart) financeChart.destroy();

        const fData = chartsData.finance_quarterly || { labels: [], income: [], cost: [], cumulative_profit: [] };
        const costDataNegative = (fData.cost || []).map(v => -Math.abs(v));

        financeChart = new Chart(financeCtx, {
            type: 'bar',
            data: {
                labels: fData.labels || [],
                datasets: [
                    { label: '当期成本', data: costDataNegative, stack: 'stack1', order: 2,
                        backgroundColor: 'rgba(239, 68, 68, 0.8)', borderColor: '#ef4444', borderWidth: 1 },
                    { label: '当期收入', data: fData.income || [], stack: 'stack1', order: 2,
                        backgroundColor: 'rgba(59, 130, 246, 0.8)', borderColor: '#3b82f6', borderWidth: 1 },
                    { label: '累计利润', data: fData.cumulative_profit || [], type: 'line', order: 1,
                        borderColor: '#22c55e', backgroundColor: 'rgba(34,197,94,0.15)', borderWidth: 3, tension: 0.4,
                        pointRadius: 5, pointBackgroundColor: '#22c55e', pointBorderColor: '#ffffff', pointBorderWidth: 2 }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: { legend: { display: true, position: 'top' }, tooltip: { callbacks: { label: function(context) { let label = context.dataset.label || ''; let value = context.parsed.y; if (label === '当期成本') value = Math.abs(value); return `${label}: ${value.toLocaleString()} 万元`; } } } },
                scales: { x: { stacked: true }, y: { stacked: false, title: { display: true, text: '金额 (万元)' }, ticks: { callback: function(value) { return Math.abs(value).toLocaleString(); } } } }
            }
        });
    }
}

// ========================================
// 季度表格
// ========================================
function updateQuarterlyTable(tableData, repaymentMonths = 1) {
    const thead = $id('quarterlyTableHead');
    const tbody = $id('quarterlyTableBody');
    if (!thead || !tbody) return;
    thead.innerHTML = '';
    tbody.innerHTML = '';

    if (!tableData || tableData.length === 0) return;

    const rowConfigs = [
        { label: '季度末时间点', field: 'end_date', styled: false, indent: false },
        { label: '季度末累计天数', field: 'cumulative_days', styled: false, indent: false },
        { label: '累计收入', field: 'cumulative_revenue', styled: true, indent: false },
        { label: '累计成本', field: 'cumulative_cost', styled: true, indent: false },
        { label: ' -UA成本', field: 'ua_cost', styled: false, indent: true },
        { label: ' -人员成本', field: 'personnel_cost', styled: false, indent: true },
        { label: ' -其他成本', field: 'other_cost', styled: false, indent: true },
        { label: '累计利润', field: 'cumulative_profit', styled: true, indent: false },
        { label: '累计净现金流（1个月回款）', field: 'cumulative_cash_flow_1m', styled: true, indent: false },
        { label: '累计净现金流（2个月回款）', field: 'cumulative_cash_flow_2m', styled: true, indent: false }
    ];

    if (repaymentMonths > 2) rowConfigs.push({ label: `累计净现金流（${repaymentMonths}个月回款）`, field: 'cumulative_cash_flow_nm', styled: true, indent: false });

    rowConfigs.push(
        { label: '当期收入', field: 'current_revenue', styled: true, indent: false },
        { label: '当期成本', field: 'current_cost', styled: true, indent: false },
        { label: ' -UA成本', field: 'current_ua_cost', styled: false, indent: true },
        { label: ' -人员成本', field: 'current_personnel_cost', styled: false, indent: true },
        { label: ' -其他成本', field: 'current_other_cost', styled: false, indent: true },
        { label: '当期利润', field: 'current_profit', styled: true, indent: false },
        { label: '当期资金需求（1个月回款）', field: 'current_cash_demand_1m', styled: false, indent: false },
        { label: '当期资金需求（2个月回款）', field: 'current_cash_demand_2m', styled: false, indent: false }
    );

    if (repaymentMonths > 2) rowConfigs.push({ label: `当期资金需求（${repaymentMonths}个月回款）`, field: 'current_cash_demand_nm', styled: false, indent: false });

    rowConfigs.push({ label: 'DAU', field: 'dau', styled: false, indent: false });

    const headerRow = document.createElement('tr');
    headerRow.innerHTML = '<th class="indicator-header">指标</th>';
    tableData.forEach(row => {
        headerRow.innerHTML += `<th>${row.quarter || '--'}</th>`;
    });
    thead.appendChild(headerRow);

    rowConfigs.forEach(config => {
        const tr = document.createElement('tr');
        const labelClass = config.indent ? 'indent-label' : 'indicator-label';
        tr.innerHTML = `<td class="${labelClass}">${config.label}</td>`;

        tableData.forEach(row => {
            let value = row[config.field];
            let displayValue;
            let cellClass = '';

            if (config.field === 'quarter' || config.field === 'end_date') {
                displayValue = value || '--';
            } else if (config.field === 'cumulative_days') {
                displayValue = (value !== null && value !== undefined) ? value : '--';
            } else {
                if (value === null || value === undefined) {
                    displayValue = '--';
                } else {
                    displayValue = formatNumber(value);
                    if (config.styled) cellClass = (typeof value === 'number' && value >= 0) ? 'positive' : 'negative';
                }
            }

            tr.innerHTML += `<td class="${cellClass}">${displayValue}</td>`;
        });

        tbody.appendChild(tr);
    });
}

// ========================================
// 导出与下载
// ========================================
function exportCsv() {
    if (!calculationResults || !calculationResults.data_file_saved) {
        alert('没有可导出的数据，请先进行计算');
        return;
    }

    const projectName = ($id('projectName')?.value || 'IAA计算结果').trim();
    const downloadUrl = `/export_csv?project_name=${encodeURIComponent(projectName)}`;

    const link = document.createElement('a');
    link.href = downloadUrl;
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

function downloadChart(chartId, chartName) {
    const canvas = $id(chartId);
    if (!canvas) return;
    const link = document.createElement('a');
    const projectName = ($id('projectName')?.value || 'IAA计算').trim();
    link.download = `${projectName}_${chartName}.png`;
    link.href = canvas.toDataURL('image/png');
    // 保证兼容性：先 append 再 click
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

// downloadQuarterlyTable 与 downloadQuarterlyTableAsImage 保持不变（原实现合理）

function downloadQuarterlyTable() {
    if (!calculationResults || !calculationResults.quarterly_table_data) {
        alert('没有可导出的季度汇总数据，请先进行计算');
        return;
    }

    const tableData = calculationResults.quarterly_table_data;
    const projectName = ($id('projectName')?.value || 'IAA计算').trim();
    const repaymentMonths = calculationResults.key_metrics?.repayment_months || 1;

    const rowConfigs = [
        { label: '季度末时间点', field: 'end_date' },
        { label: '季度末累计天数', field: 'cumulative_days' },
        { label: '累计收入', field: 'cumulative_revenue' },
        { label: '累计成本', field: 'cumulative_cost' },
        { label: ' -UA成本', field: 'ua_cost' },
        { label: ' -人员成本', field: 'personnel_cost' },
        { label: ' -其他成本', field: 'other_cost' },
        { label: '累计利润', field: 'cumulative_profit' },
        { label: '累计净现金流（1个月回款）', field: 'cumulative_cash_flow_1m' },
        { label: '累计净现金流（2个月回款）', field: 'cumulative_cash_flow_2m' }
    ];

    if (repaymentMonths > 2) rowConfigs.push({ label: `累计净现金流（${repaymentMonths}个月回款）`, field: 'cumulative_cash_flow_nm' });

    rowConfigs.push(
        { label: '当期收入', field: 'current_revenue' },
        { label: '当期成本', field: 'current_cost' },
        { label: ' -UA成本', field: 'current_ua_cost' },
        { label: ' -人员成本', field: 'current_personnel_cost' },
        { label: ' -其他成本', field: 'current_other_cost' },
        { label: '当期利润', field: 'current_profit' },
        { label: '当期资金需求（1个月回款）', field: 'current_cash_demand_1m' },
        { label: '当期资金需求（2个月回款）', field: 'current_cash_demand_2m' }
    );

    if (repaymentMonths > 2) rowConfigs.push({ label: `当期资金需求（${repaymentMonths}个月回款）`, field: 'current_cash_demand_nm' });

    rowConfigs.push({ label: 'DAU', field: 'dau' });

    const excelData = [];
    const headerRow = ['指标'];
    tableData.forEach(row => headerRow.push(row.quarter));
    excelData.push(headerRow);

    rowConfigs.forEach(config => {
        const dataRow = [config.label];
        tableData.forEach(row => {
            const value = row[config.field];
            if (value === null || value === undefined) dataRow.push('--');
            else dataRow.push(value);
        });
        excelData.push(dataRow);
    });

    const ws = XLSX.utils.aoa_to_sheet(excelData);
    const colWidths = [{ wch: 30 }];
    tableData.forEach(() => colWidths.push({ wch: 15 }));
    ws['!cols'] = colWidths;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '季度汇总数据');
    XLSX.writeFile(wb, `${projectName}_季度汇总数据.xlsx`);
}

function downloadQuarterlyTableAsImage() {
    const table = $id('quarterlyTable');
    if (!table) {
        alert('未找到季度汇总表格');
        return;
    }

    if (!calculationResults || !calculationResults.quarterly_table_data) {
        alert('没有可导出的季度汇总数据，请先进行计算');
        return;
    }

    const projectName = ($id('projectName')?.value || 'IAA计算').trim();

    if (typeof html2canvas === 'undefined') {
        const script = document.createElement('script');
        script.src = 'https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js';
        script.onload = function() { captureTableAsImage(table, projectName); };
        script.onerror = function() { alert('加载图片生成库失败，请检查网络连接'); };
        document.head.appendChild(script);
    } else {
        captureTableAsImage(table, projectName);
    }
}

function captureTableAsImage(table, projectName) {
    const tableContainer = table.closest('.table-container');
    const originalStyles = { container: tableContainer ? { overflow: tableContainer.style.overflow, overflowX: tableContainer.style.overflowX, maxWidth: tableContainer.style.maxWidth, width: tableContainer.style.width, padding: tableContainer.style.padding } : null, table: { width: table.style.width, margin: table.style.margin } };

    if (tableContainer) {
        tableContainer.style.overflow = 'visible';
        tableContainer.style.overflowX = 'visible';
        tableContainer.style.maxWidth = 'none';
        tableContainer.style.width = 'auto';
        tableContainer.style.padding = '10px';
    }
    table.style.margin = '0 10px 10px 0';
    void table.offsetWidth;

    const fullWidth = table.offsetWidth + 20;
    const fullHeight = table.offsetHeight + 20;

    html2canvas(table, { backgroundColor: '#ffffff', scale: 2, useCORS: true, logging: false, scrollX: 0, scrollY: 0, width: fullWidth, height: fullHeight, windowWidth: fullWidth + 100, windowHeight: fullHeight + 100, x: -10, y: -10 }).then(canvas => {
        if (tableContainer && originalStyles.container) {
            tableContainer.style.overflow = originalStyles.container.overflow;
            tableContainer.style.overflowX = originalStyles.container.overflowX;
            tableContainer.style.maxWidth = originalStyles.container.maxWidth;
            tableContainer.style.width = originalStyles.container.width;
            tableContainer.style.padding = originalStyles.container.padding;
        }
        table.style.width = originalStyles.table.width;
        table.style.margin = originalStyles.table.margin;

        const link = document.createElement('a');
        link.download = `${projectName}_季度汇总数据.png`;
        link.href = canvas.toDataURL('image/png');
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }).catch(error => {
        if (tableContainer && originalStyles.container) {
            tableContainer.style.overflow = originalStyles.container.overflow;
            tableContainer.style.overflowX = originalStyles.container.overflowX;
            tableContainer.style.maxWidth = originalStyles.container.maxWidth;
            tableContainer.style.width = originalStyles.container.width;
            tableContainer.style.padding = originalStyles.container.padding;
        }
        table.style.width = originalStyles.table.width;
        table.style.margin = originalStyles.table.margin;
        alert('生成图片失败：' + (error && error.message ? error.message : String(error)));
    });
}

window.addEventListener('beforeunload', function () {
    try {
        // 使用 sendBeacon，保证关闭页面时也能发出去
        const data = collectAllData();
        const blob = new Blob([JSON.stringify(data)], {
            type: 'application/json'
        });
        navigator.sendBeacon('/save_project', blob);
    } catch (e) {
    }
});

function markUnsaved() {
    const el = document.getElementById('saveStatus');
    if (!el) return;
    el.textContent = '未保存';
    el.classList.remove('saved');
    el.classList.add('unsaved');
}

function markSaved() {
    const el = document.getElementById('saveStatus');
    if (!el) return;
    el.textContent = '已保存';
    el.classList.remove('unsaved');
    el.classList.add('saved');
}

function hasUnsavedChanges() {
    const el = document.getElementById('saveStatus');
    return el && el.classList.contains('unsaved');
}
