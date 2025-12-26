/**
 * IAA休闲游戏成本收益测算 - 前端交互逻辑
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

// ========================================
// 页面初始化
// ========================================
document.addEventListener('DOMContentLoaded', function() {
    // 页面加载时不自动添加内容，等待用户新建或加载项目
    // 初始化时不执行任何操作
});

// ========================================
// 项目管理功能
// ========================================

/**
 * 新建项目 - 显示新建项目对话框
 */
function createNewProject() {
    // 清空输入框和错误信息
    document.getElementById('newProjectName').value = '';
    document.getElementById('newProjectNameError').style.display = 'none';
    document.getElementById('newProjectNameError').textContent = '';
    document.getElementById('confirmNewProjectBtn').disabled = false;
    
    // 显示新建项目对话框
    document.getElementById('newProjectDialog').style.display = 'flex';
    
    // 聚焦到输入框
    setTimeout(() => {
        document.getElementById('newProjectName').focus();
    }, 100);
}

/**
 * 关闭新建项目对话框
 */
function closeNewProjectDialog() {
    document.getElementById('newProjectDialog').style.display = 'none';
}

/**
 * 确认新建项目 - 验证名称并创建
 */
async function confirmNewProject() {
    const nameInput = document.getElementById('newProjectName');
    const errorElement = document.getElementById('newProjectNameError');
    const confirmBtn = document.getElementById('confirmNewProjectBtn');
    const projectName = nameInput.value.trim();
    
    // 清除之前的错误信息
    errorElement.style.display = 'none';
    errorElement.textContent = '';
    
    // 验证项目名称不为空
    if (!projectName) {
        errorElement.textContent = '请输入项目名称';
        errorElement.style.display = 'block';
        nameInput.focus();
        return;
    }
    
    // 验证项目名称长度
    if (projectName.length > 50) {
        errorElement.textContent = '项目名称不能超过50个字符';
        errorElement.style.display = 'block';
        nameInput.focus();
        return;
    }
    
    // 禁用确认按钮，防止重复点击
    confirmBtn.disabled = true;
    confirmBtn.innerHTML = '<span class="icon">⏳</span> 验证中...';
    
    try {
        // 检查项目名称是否已存在
        const response = await fetch('/list_projects');
        const result = await response.json();
        
        if (result.success) {
            const projects = result.projects || [];
            const existingProject = projects.find(p => p.name === projectName);
            
            if (existingProject) {
                errorElement.textContent = '该项目名称已存在，请使用其他名称';
                errorElement.style.display = 'block';
                nameInput.focus();
                confirmBtn.disabled = false;
                confirmBtn.innerHTML = '<span class="icon">✓</span> 确认创建';
                return;
            }
        }
        
        // 初始化项目数据并保存
        await initializeAndSaveProject(projectName);
        
    } catch (error) {
        errorElement.textContent = '验证失败：' + error.message;
        errorElement.style.display = 'block';
        confirmBtn.disabled = false;
        confirmBtn.innerHTML = '<span class="icon">✓</span> 确认创建';
    }
}

/**
 * 初始化并保存新项目
 */
async function initializeAndSaveProject(projectName) {
    document.getElementById('newProjectNameError');
    const confirmBtn = document.getElementById('confirmNewProjectBtn');
    
    confirmBtn.innerHTML = '<span class="icon">⏳</span> 创建中...';
    
    // 关闭对话框
    closeNewProjectDialog();
    
    // 隐藏欢迎页面，显示主内容
    document.getElementById('welcomeSection').style.display = 'none';
    document.getElementById('mainContent').style.display = '';
    
    // 设置项目名称
    document.getElementById('projectName').value = projectName;

    // 清空并重建时间段
    document.getElementById('investmentPeriods').innerHTML = '';
    periodCounter = 0;
    
    // 添加默认的ROI和留存率数据点
    document.getElementById('roiManualBody').innerHTML = '';
    addRoiRow(1, 0.5);
    addRoiRow(7, 3.5);
    addRoiRow(30, 15);
    addRoiRow(60, 30);
    addRoiRow(90, 45);

    document.getElementById('retentionManualBody').innerHTML = '';
    addRetentionRow(0, 100);
    addRetentionRow(1, 45);
    addRetentionRow(7, 25);
    addRetentionRow(30, 12);
    addRetentionRow(60, 8);
    addRetentionRow(90, 6);

    // 重置其他设置
    document.getElementById('repaymentMonths').value = 1;
    document.getElementById('targetDau').value = 500;

    // 隐藏结果区域
    document.getElementById('resultsSection').style.display = 'none';
    document.getElementById('exportCsvBtn').disabled = true;

    // 重置全局变量
    roiData = { type: 'manual', points: [] };
    retentionData = { type: 'manual', points: [] };
    calculationResults = null;
    
    // 添加第一个时间段
    addInvestmentPeriod(null, null, true);

    // 保存项目
    await saveProject(true);
    
    // 恢复按钮状态
    confirmBtn.disabled = false;
    confirmBtn.innerHTML = '<span class="icon">✓</span> 确认创建';
}

/**
 * 返回欢迎页面
 */
function backToWelcome() {
    if (confirm('确定要返回首页吗？未保存的数据将丢失。')) {
        document.getElementById('welcomeSection').style.display = '';
        document.getElementById('mainContent').style.display = 'none';
        document.getElementById('resultsSection').style.display = 'none';
    }
}


/**
 * 保存当前项目
 */
async function saveProject(isAutoSave = false) {
    const projectName = document.getElementById('projectName').value.trim();
    if (!projectName) {
        alert('项目名称不能为空！');
        return;
    }

    const projectData = collectAllData();

    try {
        const response = await fetch('/save_project', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(projectData)
        });

        const result = await response.json();

        if (result.success && !isAutoSave) {
            alert(`项目 "${projectName}" 已保存成功！`);
        }
    } catch (error) {
        alert(`保存失败：${error.message}`);
    }
}

/**
 * 显示加载项目对话框
 */
async function showLoadDialog() {
    const projectList = document.getElementById('projectList');
    projectList.innerHTML = '<li style="text-align: center; color: #666;">正在加载...</li>';
    
    document.getElementById('loadProjectDialog').style.display = 'flex';

    try {
        const response = await fetch('/list_projects');
        const result = await response.json();

        projectList.innerHTML = '';

        if (!result.success) {
            projectList.innerHTML = `<li style="text-align: center; color: #f00;">加载失败：${result.error}</li>`;
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
                        <button class="btn-delete" onclick="deleteProject('${project.name}', event)">删除</button>
                    </div>
                `;
                li.onclick = () => loadProject(project.name);
                projectList.appendChild(li);
            });
        }
    } catch (error) {
        projectList.innerHTML = `<li style="text-align: center; color: #f00;">加载失败：${error.message}</li>`;
    }
}

/**
 * 关闭加载项目对话框
 */
function closeLoadDialog() {
    document.getElementById('loadProjectDialog').style.display = 'none';
}

/**
 * 加载指定项目
 */
async function loadProject(projectName) {
    try {
        const response = await fetch('/load_project', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ project_name: projectName })
        });

        const result = await response.json();

        if (!result.success) {
            alert(`加载失败：${result.error}`);
            return;
        }

        const projectData = result.data;

        // 填充项目名称
        document.getElementById('projectName').value = projectName;

    // 清空并重建时间段
    document.getElementById('investmentPeriods').innerHTML = '';
    periodCounter = 0;

    if (projectData.investment_periods && projectData.investment_periods.length > 0) {
        projectData.investment_periods.forEach(period => {
            addInvestmentPeriod(period, null);  // skipAutoSave=true，避免多次弹出保存成功提示
        });
    } else {
        addInvestmentPeriod(null, null);  // skipAutoSave=true
    }

    // 填充ROI数据
    if (projectData.roi_data) {
        const roiType = projectData.roi_data.type || 'manual';
        document.querySelector(`input[name="roiInputType"][value="${roiType}"]`).checked = true;
        toggleRoiInput(roiType);

        if (projectData.roi_data.points && projectData.roi_data.points.length > 0) {
            if (roiType === 'manual') {
                document.getElementById('roiManualBody').innerHTML = '';
                projectData.roi_data.points.forEach(point => {
                    addRoiRow(point.day, point.value);
                });
            } else {
                // Excel类型：加载到全局变量并更新预览
                roiData.points = projectData.roi_data.points;
                updateRoiExcelPreview(roiData.points);
            }
        }
    }

    // 填充留存率数据
    if (projectData.retention_data) {
        const retentionType = projectData.retention_data.type || 'manual';
        document.querySelector(`input[name="retentionInputType"][value="${retentionType}"]`).checked = true;
        toggleRetentionInput(retentionType);

        if (projectData.retention_data.points && projectData.retention_data.points.length > 0) {
            if (retentionType === 'manual') {
                document.getElementById('retentionManualBody').innerHTML = '';
                projectData.retention_data.points.forEach(point => {
                    addRetentionRow(point.day, point.value);
                });
            } else {
                // Excel类型：加载到全局变量并更新预览
                retentionData.points = projectData.retention_data.points;
                updateRetentionExcelPreview(retentionData.points);
            }
        }
    }

    // 填充其他设置
    if (projectData.repayment_months) {
        document.getElementById('repaymentMonths').value = projectData.repayment_months;
    }
    if (projectData.target_dau) {
        document.getElementById('targetDau').value = projectData.target_dau;
    }

    closeLoadDialog();

    // 隐藏欢迎页面，显示主内容
    document.getElementById('welcomeSection').style.display = 'none';
    document.getElementById('mainContent').style.display = '';

    // 隐藏结果区域（加载项目后需要重新计算）
    document.getElementById('resultsSection').style.display = 'none';
    document.getElementById('exportCsvBtn').disabled = true;

    alert(`项目 "${projectName}" 已加载成功！`);

    } catch (error) {
        alert(`加载失败：${error.message}`);
    }
}

/**
 * 删除指定项目
 */
async function deleteProject(projectName, event) {
    event.stopPropagation();
    if (!confirm(`确定要删除项目 "${projectName}" 吗？`)) {
        return;
    }

    try {
        const response = await fetch('/delete_project', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ project_name: projectName })
        });

        const result = await response.json();

        if (result.success)
            await showLoadDialog(); // 刷新列表
    } catch (error) {
        alert(`删除失败：${error.message}`);
    }
}

/**
 * 添加投资时间段
 * @param {Object} data - 时间段数据（可选）
 * @param {HTMLElement} insertBeforeElement - 在此元素之前插入（可选，不传则追加到末尾）
 * @param {boolean} skipAutoSave - 是否跳过自动保存（用于初始化时）
 */
function addInvestmentPeriod(data = null, insertBeforeElement = null) {
    periodCounter++;
    const container = document.getElementById('investmentPeriods');

    const periodCard = document.createElement('div');
    periodCard.className = 'period-card';
    periodCard.id = `period-${periodCounter}`;

    // 获取上一个时间段的结束时间+1天作为新时间段的开始时间
    let defaultStart;
    const existingPeriods = document.querySelectorAll('.period-card');
    if (data?.start) {
        defaultStart = data.start;
    } else if (existingPeriods.length > 0) {
        // 获取最后一个时间段的结束时间
        const lastPeriod = existingPeriods[existingPeriods.length - 1];
        const lastEndDate = lastPeriod.querySelector('.period-end').value;
        if (lastEndDate) {
            // 将结束日期加一天作为新的开始日期
            const nextDay = new Date(lastEndDate);
            nextDay.setDate(nextDay.getDate() + 1);
            defaultStart = nextDay.toISOString().split('T')[0];
        } else {
            defaultStart = new Date().toISOString().split('T')[0];
        }
    } else {
        defaultStart = new Date().toISOString().split('T')[0];
    }

    // 结束时间默认为开始时间+3个月
    let defaultEnd;
    if (data?.end) {
        defaultEnd = data.end;
    } else {
        const startDate = new Date(defaultStart);
        startDate.setMonth(startDate.getMonth() + 2);
        defaultEnd = startDate.toISOString().split('T')[0];
    }

    const costType = data?.cost_type || 'fixed';
    const costValue = data?.cost_value || 10;
    const costStart = data?.cost_start || 5;
    const costEnd = data?.cost_end || 15;

    periodCard.innerHTML = `
        <div class="period-header">
            <span class="period-title">时间段 ${periodCounter}</span>
            <div class="period-header-actions">
                <button type="button" class="btn btn-insert" onclick="insertPeriodBefore(${periodCounter})" title="在此时间段之前插入新时间段">⬆ 在此之前插入</button>
                <button type="button" class="btn btn-remove" onclick="removePeriod(${periodCounter})">删除</button>
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
                            <input type="radio" name="costType-${periodCounter}" value="fixed" ${costType === 'fixed' ? 'checked' : ''} onchange="toggleCostType(${periodCounter}, 'fixed')">
                            定值
                        </label>
                        <label class="radio-label">
                            <input type="radio" name="costType-${periodCounter}" value="linear" ${costType === 'linear' ? 'checked' : ''} onchange="toggleCostType(${periodCounter}, 'linear')">
                            线性变化
                        </label>
                    </div>
                </div>
            </div>
        </div>
    `;

    // 根据是否指定了插入位置，决定是插入还是追加
    if (insertBeforeElement) {
        container.insertBefore(periodCard, insertBeforeElement);
    } else {
        container.appendChild(periodCard);
    }

    // 重新编号所有时间段
    renumberPeriods();
    
    // 自动保存项目数据
    saveProject(true);

}

/**
 * 在指定时间段之前插入新的时间段
 * @param {number} beforePeriodId - 在此ID对应的时间段之前插入
 */
function insertPeriodBefore(beforePeriodId) {
    const targetPeriod = document.getElementById(`period-${beforePeriodId}`);
    if (!targetPeriod) {
        console.error(`未找到ID为 period-${beforePeriodId} 的时间段`);
        return;
    }

    // 获取目标时间段的开始时间，作为新时间段的结束时间参考
    const targetStartDate = targetPeriod.querySelector('.period-start').value;
    
    // 计算新时间段的日期
    let newEndDate;
    let newStartDate;
    
    if (targetStartDate) {
        // 新时间段的结束时间 = 目标时间段开始时间的前一天
        const endDate = new Date(targetStartDate);
        endDate.setDate(endDate.getDate() - 1);
        newEndDate = endDate.toISOString().split('T')[0];

        // 新时间段的开始时间 = 结束时间往前推3个月
        const startDate = new Date(endDate);
        startDate.setMonth(startDate.getMonth() - 3);
        newStartDate = startDate.toISOString().split('T')[0];
    } else {
        // 如果目标时间段没有开始时间，使用默认值
        const today = new Date();
        newEndDate = today.toISOString().split('T')[0];
        const threeMonthsAgo = new Date(today);
        threeMonthsAgo.setMonth(threeMonthsAgo.getMonth() - 3);
        newStartDate = threeMonthsAgo.toISOString().split('T')[0];
    }

    // 创建新时间段的数据
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

    // 在目标时间段之前插入新时间段
    addInvestmentPeriod(newPeriodData, targetPeriod);
}

/**
 * 重新编号所有时间段的显示标题
 */
function renumberPeriods() {
    const periods = document.querySelectorAll('.period-card');
    periods.forEach((period, index) => {
        const titleElement = period.querySelector('.period-title');
        if (titleElement) {
            titleElement.textContent = `时间段 ${index + 1}`;
        }
    });
}

/**
 * 删除时间段
 */
function removePeriod(id) {
    const period = document.getElementById(`period-${id}`);
    if (period) {
        period.remove();
    }

    // 确保至少有一个时间段
    if (document.querySelectorAll('.period-card').length === 0) {
        addInvestmentPeriod();
    } else {
        // 重新编号所有时间段
        renumberPeriods();
    }
}

/**
 * 切换日消耗类型
 */
function toggleCostType(periodId, type) {
    const periodCard = document.getElementById(`period-${periodId}`);
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
// ROI数据管理
// ========================================

/**
 * 切换ROI输入方式
 */
function toggleRoiInput(type) {
    document.getElementById('roiManualInput').style.display = type === 'manual' ? '' : 'none';
    document.getElementById('roiExcelInput').style.display = type === 'excel' ? '' : 'none';
    roiData.type = type;
}

/**
 * 更新ROI Excel预览（用于加载项目时）
 */
function updateRoiExcelPreview(points) {
    const previewData = points.slice(0, 30);  // 只预览前30天
    displayExcelPreview('roiExcelPreview', previewData, 'ROI', points.length);
}

/**
 * 添加ROI数据行
 */
function addRoiRow(day = '', value = '') {
    const tbody = document.getElementById('roiManualBody');
    const row = document.createElement('tr');
    row.innerHTML = `
        <td><input type="number" class="roi-day" value="${day}" min="1" step="1" placeholder="天数"></td>
        <td><input type="number" class="roi-value" value="${value}" min="0" step="0.1" placeholder="ROI值"></td>
        <td><button type="button" class="btn btn-remove" onclick="this.closest('tr').remove()">删除</button></td>
    `;
    tbody.appendChild(row);
}

/**
 * 处理ROI Excel文件
 */
function handleRoiExcel(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

            // 修改：导入所有数据，不限制30天
            roiData.points = [];
            const previewData = [];

            // 修改：从第0行开始读取（包含表头行也尝试解析）
            for (let i = 0; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (row && row.length >= 2) {
                    const day = parseInt(row[0]);
                    const value = parseFloat(row[1]);
                    // 跳过无效数据（如表头）
                    if (!isNaN(day) && !isNaN(value) && day > 0) {
                        roiData.points.push({ day, value });
                        // 只预览前30天
                        if (previewData.length < 30) {
                            previewData.push({ day, value });
                        }
                    }
                }
            }

            // 显示预览（只显示前30天）
            displayExcelPreview('roiExcelPreview', previewData, 'ROI', roiData.points.length);
        } catch (error) {
            alert('Excel文件解析失败：' + error.message);
        }
    };
    reader.readAsArrayBuffer(file);
}

// ========================================
// 留存率数据管理
// ========================================

/**
 * 切换留存率输入方式
 */
function toggleRetentionInput(type) {
    document.getElementById('retentionManualInput').style.display = type === 'manual' ? '' : 'none';
    document.getElementById('retentionExcelInput').style.display = type === 'excel' ? '' : 'none';
    retentionData.type = type;
}

/**
 * 更新留存率Excel预览（用于加载项目时）
 */
function updateRetentionExcelPreview(points) {
    const previewData = points.slice(0, 30);  // 只预览前30天
    displayExcelPreview('retentionExcelPreview', previewData, '留存率', points.length);
}

/**
 * 添加留存率数据行
 */
function addRetentionRow(day = '', value = '') {
    const tbody = document.getElementById('retentionManualBody');
    const row = document.createElement('tr');
    row.innerHTML = `
        <td><input type="number" class="retention-day" value="${day}" min="1" step="1" placeholder="天数"></td>
        <td><input type="number" class="retention-value" value="${value}" min="0" max="100" step="0.1" placeholder="留存率"></td>
        <td><button type="button" class="btn btn-remove" onclick="this.closest('tr').remove()">删除</button></td>
    `;
    tbody.appendChild(row);
}

/**
 * 处理留存率Excel文件
 */
function handleRetentionExcel(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

            // 修改：导入所有数据，不限制30天
            retentionData.points = [];
            const previewData = [];

            // 修改：从第0行开始读取
            for (let i = 0; i < jsonData.length; i++) {
                const row = jsonData[i];
                if (row && row.length >= 2) {
                    const day = parseInt(row[0]);
                    const value = parseFloat(row[1]);
                    // 跳过无效数据（如表头）
                    if (!isNaN(day) && !isNaN(value) && day > 0) {
                        retentionData.points.push({ day, value });
                        // 只预览前30天
                        if (previewData.length < 30) {
                            previewData.push({ day, value });
                        }
                    }
                }
            }

            displayExcelPreview('retentionExcelPreview', previewData, '留存率', retentionData.points.length);
        } catch (error) {
            alert('Excel文件解析失败：' + error.message);
        }
    };
    reader.readAsArrayBuffer(file);
}

/**
 * 显示Excel预览表格
 * @param {string} containerId - 容器ID
 * @param {Array} data - 预览数据（前30天）
 * @param {string} label - 标签名称
 * @param {number} totalCount - 总导入条数
 */
function displayExcelPreview(containerId, data, label, totalCount = null) {
    const container = document.getElementById(containerId);
    if (data.length === 0) {
        container.innerHTML = '<p style="padding: 10px; color: #666;">未找到有效数据</p>';
        return;
    }

    let html = `
        <table>
            <thead>
                <tr>
                    <th>天数</th>
                    <th>${label} (%)</th>
                </tr>
            </thead>
            <tbody>
    `;

    data.forEach(item => {
        html += `<tr><td>${item.day}</td><td>${item.value.toFixed(2)}</td></tr>`;
    });

    html += '</tbody></table>';

    // 修改：显示总导入条数和预览条数
    const total = totalCount || data.length;
    if (total > 30) {
        html += `<p style="padding: 10px; color: #666; font-size: 0.85rem;">共导入 ${total} 条数据，预览前 30 条</p>`;
    } else {
        html += `<p style="padding: 10px; color: #666; font-size: 0.85rem;">共导入 ${total} 条数据</p>`;
    }

    container.innerHTML = html;
}

// ========================================
// 数据收集
// ========================================

/**
 * 收集所有输入数据
 */
function collectAllData() {
    const data = {
        project_name: document.getElementById('projectName').value.trim(),
        investment_periods: [],
        roi_data: { type: 'manual', points: [] },
        retention_data: { type: 'manual', points: [] },
        repayment_months: parseInt(document.getElementById('repaymentMonths').value) || 1,
        target_dau: parseInt(document.getElementById('targetDau').value) || 1000
    };

    // 收集时间段数据
    document.querySelectorAll('.period-card').forEach(card => {
        const periodId = card.id.split('-')[1];
        const costType = document.querySelector(`input[name="costType-${periodId}"]:checked`)?.value || 'fixed';

        const period = {
            start: card.querySelector('.period-start').value,
            end: card.querySelector('.period-end').value,
            cost_type: costType,
            // 修改：适配新的class名称
            cost_value: parseFloat(card.querySelector('.period-cost-value')?.value) || 0,
            cost_start: parseFloat(card.querySelector('.period-cost-start')?.value) || 0,
            cost_end: parseFloat(card.querySelector('.period-cost-end')?.value) || 0,
            dnu: parseFloat(card.querySelector('.period-dnu').value) || 0,
            team_size: parseInt(card.querySelector('.period-team-size').value) || 0,
            labor_cost: parseFloat(card.querySelector('.period-labor-cost').value) || 0,
            other_cost: parseFloat(card.querySelector('.period-other-cost').value) || 0
        };

        data.investment_periods.push(period);
    });

    // 收集ROI数据
    const roiInputType = document.querySelector('input[name="roiInputType"]:checked').value;
    data.roi_data.type = roiInputType;

    if (roiInputType === 'manual') {
        document.querySelectorAll('#roiManualBody tr').forEach(row => {
            const day = parseInt(row.querySelector('.roi-day').value);
            const value = parseFloat(row.querySelector('.roi-value').value);
            if (!isNaN(day) && !isNaN(value)) {
                data.roi_data.points.push({ day, value });
            }
        });
    } else {
        data.roi_data.points = roiData.points;
    }

    // 收集留存率数据
    const retentionInputType = document.querySelector('input[name="retentionInputType"]:checked').value;
    data.retention_data.type = retentionInputType;

    if (retentionInputType === 'manual') {
        document.querySelectorAll('#retentionManualBody tr').forEach(row => {
            const day = parseInt(row.querySelector('.retention-day').value);
            const value = parseFloat(row.querySelector('.retention-value').value);
            if (!isNaN(day) && !isNaN(value)) {
                data.retention_data.points.push({ day, value });
            }
        });
    } else {
        data.retention_data.points = retentionData.points;
    }

    return data;
}

// ========================================
// 计算与结果展示
// ========================================

/**
 * 开始计算
 */
async function startCalculation() {
    const data = collectAllData();

    // 验证数据
    if (data.investment_periods.length === 0) {
        alert('请至少添加一个投资时间段');
        return;
    }

    if (data.roi_data.points.length === 0) {
        alert('请输入ROI数据');
        return;
    }

    if (data.retention_data.points.length === 0) {
        alert('请输入留存率数据');
        return;
    }

    // 显示加载动画
    document.getElementById('loadingOverlay').style.display = 'flex';

    try {
        const response = await fetch('/calculate', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(data)
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const results = await response.json();

        if (results.error) {
            throw new Error(results.error);
        }

        calculationResults = results;
        displayResults(results);

        // 启用导出按钮
        document.getElementById('exportCsvBtn').disabled = false;

        // 显示结果区域
        document.getElementById('resultsSection').style.display = '';

        // 滚动到结果区域
        document.getElementById('resultsSection').scrollIntoView({ behavior: 'smooth' });

    } catch (error) {
        alert('计算失败：' + error.message);
    } finally {
        document.getElementById('loadingOverlay').style.display = 'none';
    }
}

/**
 * 显示计算结果
 */
function displayResults(results) {
    // 更新关键指标
    const metrics = results.key_metrics;
    // 修改：现金相关指标使用formatCash函数保留2位小数
    document.getElementById('metricMaxCash1m').textContent = formatCash(metrics.max_cash_demand_1m);
    document.getElementById('metricMaxCash2m').textContent = formatCash(metrics.max_cash_demand_2m);
    document.getElementById('metricDynamicBreakeven').textContent = formatDays(metrics.dynamic_profit_breakeven_day);
    document.getElementById('metricCumulativeBreakeven').textContent = formatDays(metrics.cumulative_profit_breakeven_day);
    document.getElementById('metricCashFlow1mBreakeven').textContent = formatDays(metrics.cumulative_cash_flow_1m_breakeven_day);
    document.getElementById('metricCashFlow2mBreakeven').textContent = formatDays(metrics.cumulative_cash_flow_2m_breakeven_day);
    document.getElementById('metricDay10mDau').textContent = formatDays(metrics.day_to_10m_dau);
    document.getElementById('metricDay2mDau').textContent = formatDays(metrics.day_to_2m_dau);
    document.getElementById('metricDayTargetDau').textContent = formatDays(metrics.day_to_target_dau);
    document.getElementById('metricTargetDauValue').textContent = metrics.target_dau || '--';

    // 更新图表
    updateCharts(results.charts);

    // 更新表格
    updateQuarterlyTable(results.quarterly_table_data);
}

/**
 * 格式化现金数字 - 保留2位小数
 */
function formatCash(value) {
    if (value === null || value === undefined) return '--';
    return value.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

/**
 * 格式化数字
 */
function formatNumber(value) {
    if (value === null || value === undefined) return '--';
    return value.toLocaleString('zh-CN', { maximumFractionDigits: 2 });
}

/**
 * 格式化天数
 */
function formatDays(value) {
    if (value === null || value === undefined || value < 0) return '未达成';
    return value.toString();
}

/**
 * 更新图表
 */
function updateCharts(chartsData) {
    // DAU季度图表
    const dauCtx = document.getElementById('dauChart').getContext('2d');

    if (dauChart) {
        dauChart.destroy();
    }

    dauChart = new Chart(dauCtx, {
        type: 'line',
        data: {
            labels: chartsData.dau_quarterly.labels,
            datasets: [{
                label: 'DAU (万人)',
                data: chartsData.dau_quarterly.data,
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
                legend: {
                    display: true,
                    position: 'top'
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return `DAU: ${context.parsed.y.toLocaleString()} 万人`;
                        }
                    }
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'DAU (万人)'
                    }
                }
            }
        }
    });

    // 收入成本图表 - 修改为上下分布的柱状图
    const financeCtx = document.getElementById('financeChart').getContext('2d');

    if (financeChart) {
        financeChart.destroy();
    }

    // 修改：将成本转为负数用于显示在x轴下方
    const costDataNegative = chartsData.finance_quarterly.cost.map(v => -Math.abs(v));

    financeChart = new Chart(financeCtx, {
        type: 'bar',
        data: {
            labels: chartsData.finance_quarterly.labels,
            datasets: [
                {
                    label: '当期成本',
                    data: costDataNegative,  // 使用负数数据
                    backgroundColor: 'rgba(220, 53, 69, 0.8)',
                    borderColor: '#dc3545',
                    borderWidth: 1,
                    stack: 'stack1',  // 添加相同的stack
                    order: 2
                },
                {
                    label: '当期收入',
                    data: chartsData.finance_quarterly.income,
                    backgroundColor: 'rgba(40, 167, 69, 0.8)',
                    borderColor: '#28a745',
                    borderWidth: 1,
                    stack: 'stack1',  // 添加相同的stack
                    order: 2
                },
                {
                    label: '累计利润',
                    data: chartsData.finance_quarterly.cumulative_profit,
                    type: 'line',
                    borderColor: '#764ba2',
                    backgroundColor: 'transparent',
                    borderWidth: 3,
                    pointRadius: 6,
                    pointBackgroundColor: '#764ba2',
                    pointBorderColor: '#fff',
                    pointBorderWidth: 2,
                    tension: 0.4,
                    order: 1,
                    yAxisID: 'y'
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    display: true,
                    position: 'top'
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            let label = context.dataset.label || '';
                            let value = context.parsed.y;
                            // 成本显示为正数（绝对值）
                            if (label === '当期成本') {
                                value = Math.abs(value);
                            }
                            return `${label}: ${value.toLocaleString()} 万元`;
                        }
                    }
                }
            },
            scales: {
                x: {
                    stacked: true  // 启用堆叠
                },
                y: {
                    stacked: false,  // Y轴不堆叠，允许负值显示
                    title: {
                        display: true,
                        text: '金额 (万元)'
                    },
                    ticks: {
                        callback: function(value) {
                            // Y轴标签显示绝对值
                            return Math.abs(value).toLocaleString();
                        }
                    }
                }
            }
        }
    });


}

/**
 * 更新季度汇总表格 - 转置布局：季度在上行，指标在左列
 */
function updateQuarterlyTable(tableData) {
    const thead = document.getElementById('quarterlyTableHead');
    const tbody = document.getElementById('quarterlyTableBody');
    thead.innerHTML = '';
    tbody.innerHTML = '';

    if (!tableData || tableData.length === 0) return;

    // 定义指标行配置：[指标名称, 数据字段, 是否需要正负样式, 是否为子项（需要缩进）]
    // 注意：后端返回的成本已经是负数，不需要再取反
    // 注意：季度名称已在表头显示，无需在数据行中重复
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
        { label: '累计净现金流（2个月回款）', field: 'cumulative_cash_flow_2m', styled: true, indent: false },
        { label: '当期收入', field: 'current_revenue', styled: true, indent: false },
        { label: '当期成本', field: 'current_cost', styled: true, indent: false },
        { label: ' -UA成本', field: 'current_ua_cost', styled: false, indent: true },
        { label: ' -人员成本', field: 'current_personnel_cost', styled: false, indent: true },
        { label: ' -其他成本', field: 'current_other_cost', styled: false, indent: true },
        { label: '当期利润', field: 'current_profit', styled: true, indent: false },
        { label: '当期资金需求（1个月回款）', field: 'current_cash_demand_1m', styled: false, indent: false },
        { label: '当期资金需求（2个月回款）', field: 'current_cash_demand_2m', styled: false, indent: false },
        { label: 'DAU', field: 'dau', styled: false, indent: false }
    ];

    // 创建表头行（季度标题）
    const headerRow = document.createElement('tr');
    headerRow.innerHTML = '<th class="indicator-header">指标</th>';
    tableData.forEach(row => {
        headerRow.innerHTML += `<th>${row.quarter}</th>`;
    });
    thead.appendChild(headerRow);

    // 创建每个指标行
    rowConfigs.forEach(config => {
        const tr = document.createElement('tr');
        
        // 指标名称列
        const labelClass = config.indent ? 'indent-label' : 'indicator-label';
        tr.innerHTML = `<td class="${labelClass}">${config.label}</td>`;
        
        // 各季度数据列
        tableData.forEach(row => {
            let value = row[config.field];
            let displayValue;
            let cellClass = '';
            
            // 处理数值显示
            if (config.field === 'quarter' || config.field === 'end_date') {
                displayValue = value || '--';
            } else if (config.field === 'cumulative_days') {
                displayValue = value !== null && value !== undefined ? value : '--';
            } else {
                // 数值字段
                if (value === null || value === undefined) {
                    displayValue = '--';
                } else {
                    displayValue = formatNumber(value);
                    
                    // 添加正负样式
                    if (config.styled) {
                        cellClass = value >= 0 ? 'positive' : 'negative';
                    }
                }
            }
            
            tr.innerHTML += `<td class="${cellClass}">${displayValue}</td>`;
        });
        
        tbody.appendChild(tr);
    });
}

// ========================================
// 导出功能
// ========================================

/**
 * 导出Excel
 * 调用后端接口下载Excel文件
 */
function exportCsv() {
    if (!calculationResults || !calculationResults.data_file_saved) {
        alert('没有可导出的数据，请先进行计算');
        return;
    }

    const projectName = document.getElementById('projectName').value.trim() || 'IAA计算结果';
    
    // 通过后端接口下载Excel文件
    const downloadUrl = `/export_csv?project_name=${encodeURIComponent(projectName)}`;
    
    // 创建隐藏的链接进行下载
    const link = document.createElement('a');
    link.href = downloadUrl;
    link.style.visibility = 'hidden';
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

/**
 * 下载图表
 */
function downloadChart(chartId, chartName) {
    const canvas = document.getElementById(chartId);
    const link = document.createElement('a');

    const projectName = document.getElementById('projectName').value.trim() || 'IAA计算';
    link.download = `${projectName}_${chartName}.png`;
    link.href = canvas.toDataURL('image/png');
    link.click();
}

/**
 * 下载季度汇总表格为Excel文件
 */
function downloadQuarterlyTable() {
    if (!calculationResults || !calculationResults.quarterly_table_data) {
        alert('没有可导出的季度汇总数据，请先进行计算');
        return;
    }

    const tableData = calculationResults.quarterly_table_data;
    const projectName = document.getElementById('projectName').value.trim() || 'IAA计算';

    // 定义指标行配置（与updateQuarterlyTable中保持一致）
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
        { label: '累计净现金流（2个月回款）', field: 'cumulative_cash_flow_2m' },
        { label: '当期收入', field: 'current_revenue' },
        { label: '当期成本', field: 'current_cost' },
        { label: ' -UA成本', field: 'current_ua_cost' },
        { label: ' -人员成本', field: 'current_personnel_cost' },
        { label: ' -其他成本', field: 'current_other_cost' },
        { label: '当期利润', field: 'current_profit' },
        { label: '当期资金需求（1个月回款）', field: 'current_cash_demand_1m' },
        { label: '当期资金需求（2个月回款）', field: 'current_cash_demand_2m' },
        { label: 'DAU', field: 'dau' }
    ];

    // 构建Excel数据（转置格式：行为指标，列为季度）
    const excelData = [];

    // 表头行：指标 + 各季度名称
    const headerRow = ['指标'];
    tableData.forEach(row => {
        headerRow.push(row.quarter);
    });
    excelData.push(headerRow);

    // 数据行：每个指标一行
    rowConfigs.forEach(config => {
        const dataRow = [config.label];
        tableData.forEach(row => {
            let value = row[config.field];
            if (value === null || value === undefined) {
                dataRow.push('--');
            } else if (typeof value === 'number') {
                dataRow.push(value);
            } else {
                dataRow.push(value);
            }
        });
        excelData.push(dataRow);
    });

    // 使用SheetJS创建Excel文件
    const ws = XLSX.utils.aoa_to_sheet(excelData);
    
    // 设置列宽
    const colWidths = [{ wch: 30 }];  // 第一列（指标名称）宽度
    tableData.forEach(() => {
        colWidths.push({ wch: 15 });  // 其他列宽度
    });
    ws['!cols'] = colWidths;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '季度汇总数据');

    // 下载文件
    XLSX.writeFile(wb, `${projectName}_季度汇总数据.xlsx`);
}

/**
 * 下载季度汇总表格为图片
 */
function downloadQuarterlyTableAsImage() {
    const table = document.getElementById('quarterlyTable');
    if (!table) {
        alert('未找到季度汇总表格');
        return;
    }

    if (!calculationResults || !calculationResults.quarterly_table_data) {
        alert('没有可导出的季度汇总数据，请先进行计算');
        return;
    }

    const projectName = document.getElementById('projectName').value.trim() || 'IAA计算';

    // 使用html2canvas将表格转换为图片
    // 先检查是否已加载html2canvas
    if (typeof html2canvas === 'undefined') {
        // 动态加载html2canvas库
        const script = document.createElement('script');
        script.src = 'https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js';
        script.onload = function() {
            captureTableAsImage(table, projectName);
        };
        script.onerror = function() {
            alert('加载图片生成库失败，请检查网络连接');
        };
        document.head.appendChild(script);
    } else {
        captureTableAsImage(table, projectName);
    }
}

/**
 * 捕获表格为图片并下载
 */
function captureTableAsImage(table, projectName) {
    // 获取表格容器以包含完整的表格
    const tableContainer = table.closest('.table-container');

    // 保存原始样式，以便截图后恢复
    const originalStyles = {
        container: tableContainer ? {
            overflow: tableContainer.style.overflow,
            overflowX: tableContainer.style.overflowX,
            maxWidth: tableContainer.style.maxWidth,
            width: tableContainer.style.width,
            padding: tableContainer.style.padding
        } : null,
        table: {
            width: table.style.width,
            margin: table.style.margin
        }
    };

    // 临时修改样式以显示完整表格
    if (tableContainer) {
        tableContainer.style.overflow = 'visible';
        tableContainer.style.overflowX = 'visible';
        tableContainer.style.maxWidth = 'none';
        tableContainer.style.width = 'auto';
        tableContainer.style.padding = '10px';  // 添加内边距确保边框可见
    }

    // 给表格添加右边距，确保边框完整
    table.style.margin = '0 10px 10px 0';

    // 强制重绘以应用样式更改
    void table.offsetWidth;

    // 获取表格的实际完整尺寸（包含边框）
    // offsetWidth/offsetHeight 包含边框、内边距和滚动条
    const fullWidth = table.offsetWidth + 20;  // 额外边距
    const fullHeight = table.offsetHeight + 20;

    // 直接截取表格元素而非容器，以获得更精确的尺寸
    html2canvas(table, {
        backgroundColor: '#ffffff',
        scale: 2,  // 提高图片清晰度
        useCORS: true,
        logging: false,
        scrollX: 0,
        scrollY: 0,
        width: fullWidth,
        height: fullHeight,
        windowWidth: fullWidth + 100,
        windowHeight: fullHeight + 100,
        x: -10,  // 向左偏移以包含左边距
        y: -10   // 向上偏移以包含上边距
    }).then(canvas => {
        // 恢复原始样式
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
        link.click();
    }).catch(error => {
        // 恢复原始样式
        if (tableContainer && originalStyles.container) {
            tableContainer.style.overflow = originalStyles.container.overflow;
            tableContainer.style.overflowX = originalStyles.container.overflowX;
            tableContainer.style.maxWidth = originalStyles.container.maxWidth;
            tableContainer.style.width = originalStyles.container.width;
            tableContainer.style.padding = originalStyles.container.padding;
        }
        table.style.width = originalStyles.table.width;
        table.style.margin = originalStyles.table.margin;
        alert('生成图片失败：' + error.message);
    });
}