/**
 * IAA休闲游戏成本收益测算 - 前端交互逻辑
 *
 * 使用 IIFE 封装以避免污染全局命名空间，
 * 仅将 HTML 行内事件处理器需要的函数挂载到 window 上。
 */
;(function () {
'use strict';

// ========================================
// 模块状态 —— 用 let 而非 const，因为运行期会被多次重置
// ========================================
let periodCounter = 0;
let roiData = { type: 'manual', points: [], curveName: '' };
let retentionData = { type: 'manual', points: [], curveName: '' };
let calculationResults = null;
let dauChart = null;
let financeChart = null;

// 监听所有表单控件的输入变化，统一标记「未保存」状态，
// 使用事件代理减少绑定次数
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

// 安全获取元素 —— 避免 null 引用导致后续链式调用抛错
function $id(id) {
  return document.getElementById(id);
}

// 安全 querySelector —— 捕获无效选择器异常
function $qs(selector, root = document) {
  try {
    return root.querySelector(selector);
  } catch (_) {
    return null;
  }
}

// 统一提取错误消息，避免全文重复 error && error.message ? error.message : String(error)
function getErrorMessage(error) {
  return (error && error.message) ? error.message : String(error);
}

// ========================================
// 页面初始化
// ========================================
document.addEventListener('DOMContentLoaded', () => {
  // 预留：可在此绑定全局快捷键、初始化第三方组件等
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

    if (result?.success) {
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
      errorElement.textContent = '验证失败：' + getErrorMessage(error);
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
  roiData = { type: 'manual', points: [], curveName: '' };
  retentionData = { type: 'manual', points: [], curveName: '' };
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
  // 隐藏曲线radio选项并重置标签
  const roiCurveRadioLabel = $id('roiCurveRadioLabel');
  if (roiCurveRadioLabel) roiCurveRadioLabel.style.display = 'none';
  const roiCurveRadioText = $id('roiCurveRadioText');
  if (roiCurveRadioText) roiCurveRadioText.textContent = 'ROI';
  const roiCurvePreview = $id('roiCurvePreview');
  if (roiCurvePreview) roiCurvePreview.innerHTML = '';

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
  // 隐藏曲线radio选项并重置标签
  const retentionCurveRadioLabel = $id('retentionCurveRadioLabel');
  if (retentionCurveRadioLabel) retentionCurveRadioLabel.style.display = 'none';
  const retentionCurveRadioText = $id('retentionCurveRadioText');
  if (retentionCurveRadioText) retentionCurveRadioText.textContent = '留存率';
  const retentionCurvePreview = $id('retentionCurvePreview');
  if (retentionCurvePreview) retentionCurvePreview.innerHTML = '';

  addRetentionRow(1, 100);
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

  const projectData = await collectAllData();

  try {
    const response = await fetch('/save_project', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(projectData)
    });

    const result = await response.json();

    if (result?.success && !isAutoSave) {
      alert(`项目 "${projectName}" 已保存成功！`);
    }

    if (result?.success) {
      markSaved();
    }
  } catch (error) {
    alert(`保存失败：${getErrorMessage(error)}`);
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
  if (projectList) projectList.innerHTML = '<li class="list-status">正在加载...</li>';
  const dialog = $id('loadProjectDialog');
  if (dialog) dialog.style.display = 'flex';

  try {
    const response = await fetch('/list_projects');
    const result = await response.json();

    if (!projectList) return;
    projectList.innerHTML = '';

    if (!result?.success) {
      projectList.innerHTML = `<li class="list-status is-error">加载失败：${result?.error || '未知错误'}</li>`;
      return;
    }

    const projects = result.projects || [];
    if (projects.length === 0) {
      projectList.innerHTML = '<li class="list-status">暂无已保存的项目</li>';
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
    if (projectList) projectList.innerHTML = `<li class="list-status is-error">加载失败：${getErrorMessage(error)}</li>`;
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
    if (!result?.success) {
      alert(`加载失败：${result?.error || '未知错误'}`);
      return;
    }

    const projectData = result.data || {};

    const projectNameEl = $id('projectName');
    if (projectNameEl) projectNameEl.value = projectName;

    const periodsContainer = $id('investmentPeriods');
    if (periodsContainer) periodsContainer.innerHTML = '';
    periodCounter = 0;

    if (projectData.investment_periods && projectData.investment_periods.length > 0) {
      for (const period of projectData.investment_periods) {
        await addInvestmentPeriod(period, null, true);
      }
    } else {
      await addInvestmentPeriod(null, null, true);
    }

    // ROI 数据
    if (projectData.roi_data) {
      const roiType = projectData.roi_data.type || 'manual';
      if (roiType === 'curve') {
        // 曲线模式：需先显示曲线 radio 选项并设置标签
        const curveName = projectData.roi_data.curveName || 'ROI';
        const roiCurveRadioLabel = $id('roiCurveRadioLabel');
        const roiCurveRadioText = $id('roiCurveRadioText');
        if (roiCurveRadioLabel) roiCurveRadioLabel.style.display = '';
        if (roiCurveRadioText) roiCurveRadioText.textContent = `${curveName}`;
        const roiRadio = $qs('input[name="roiInputType"][value="curve"]');
        if (roiRadio) roiRadio.checked = true;
        toggleRoiInput('curve');
        roiData.points = projectData.roi_data.points || [];
        roiData.curveName = curveName;
        const previewData = roiData.points.slice(0, 30);
        displayExcelPreview('roiCurvePreview', previewData, 'ROI (%)', roiData.points.length);
        const manualBody = $id('roiManualBody');
        if (manualBody) manualBody.innerHTML = '';
      } else if (roiType === 'manual') {
        const roiRadio = $qs('input[name="roiInputType"][value="manual"]');
        if (roiRadio) roiRadio.checked = true;
        toggleRoiInput('manual');
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
        const roiRadio = $qs(`input[name="roiInputType"][value="${roiType}"]`);
        if (roiRadio) roiRadio.checked = true;
        toggleRoiInput(roiType);
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
      if (retentionType === 'curve') {
        // 曲线模式：需先显示曲线 radio 选项并设置标签
        const curveName = projectData.retention_data.curveName || '留存率';
        const retentionCurveRadioLabel = $id('retentionCurveRadioLabel');
        const retentionCurveRadioText = $id('retentionCurveRadioText');
        if (retentionCurveRadioLabel) retentionCurveRadioLabel.style.display = '';
        if (retentionCurveRadioText) retentionCurveRadioText.textContent = `${curveName}`;
        const retentionRadio = $qs('input[name="retentionInputType"][value="curve"]');
        if (retentionRadio) retentionRadio.checked = true;
        toggleRetentionInput('curve');
        retentionData.points = projectData.retention_data.points || [];
        retentionData.curveName = curveName;
        const previewData = retentionData.points.slice(0, 30);
        displayExcelPreview('retentionCurvePreview', previewData, '留存率 (%)', retentionData.points.length);
        const manualBody = $id('retentionManualBody');
        if (manualBody) manualBody.innerHTML = '';
      } else if (retentionType === 'manual') {
        const retentionRadio = $qs('input[name="retentionInputType"][value="manual"]');
        if (retentionRadio) retentionRadio.checked = true;
        toggleRetentionInput('manual');
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
        const retentionRadio = $qs(`input[name="retentionInputType"][value="${retentionType}"]`);
        if (retentionRadio) retentionRadio.checked = true;
        toggleRetentionInput(retentionType);
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
    alert(`加载失败：${getErrorMessage(error)}`);
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
    if (result?.success) await showLoadDialog();
  } catch (error) {
    alert(`删除失败：${getErrorMessage(error)}`);
  }
}

// ========================================
// 投资时间段管理
// ========================================
async function addInvestmentPeriod(data = null, insertBeforeElement = null, skipAutoSave = false) {
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

  // 构建曲线选择器的选项列表
  const [roiCurves, retentionCurves] = await Promise.all([
    loadCurvesByType('roi'),
    loadCurvesByType('retention'),
  ]);
  const selectedRoiCurveId = (data && data.roi_curve_id) || '';
  const selectedRetentionCurveId = (data && data.retention_curve_id) || '';

  function buildCurveOptions(curves, selectedId, defaultLabel) {
    let opts = `<option value="">${defaultLabel}</option>`;
    curves.forEach(c => {
      const sel = c.id === selectedId ? 'selected' : '';
      opts += `<option value="${c.id}" ${sel}>${c.name}</option>`;
    });
    return opts;
  }

  periodCard.innerHTML = `
    <div class="period-header">
      <span class="period-title">时间段 ${periodCounter}</span>
      <div class="period-header-actions">
        <button type="button" class="btn btn-insert" data-action="insert-before">⬆ 向上插入</button>
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
      <div class="input-group period-curve-select-group">
        <label>ROI曲线</label>
        <select class="period-roi-curve">
${buildCurveOptions(roiCurves, selectedRoiCurveId, '全局曲线')}
        </select>
      </div>
      <div class="input-group period-curve-select-group">
        <label>留存率曲线</label>
        <select class="period-retention-curve">
${buildCurveOptions(retentionCurves, selectedRetentionCurveId, '全局曲线')}
        </select>
      </div>
      <div class="period-params-group">
        <div class="input-group">
          <label>DNU (万人)</label>
          <input type="number" class="period-dnu" value="${data?.dnu || 1}" min="0" step="0.1">
        </div>
        <div class="input-group">
          <label>团队规模 (人)</label>
          <input type="number" class="period-team-size" value="${data?.team_size ?? 10}" min="0" step="1">
        </div>
        <div class="input-group">
          <label>用工成本 (万/人/天)</label>
          <input type="number" class="period-labor-cost" value="${data?.labor_cost ?? 0.025}" min="0" step="0.01">
        </div>
        <div class="input-group">
          <label>其他运营成本 (万元/天)</label>
          <input type="number" class="period-other-cost" value="${data?.other_cost ?? 0.1}" min="0" step="0.1">
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

// 保留旧版 API 以兼容可能存在的外部调用
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
  const curve = $id('roiCurveInput');
  if (manual) manual.style.display = type === 'manual' ? '' : 'none';
  if (excel) excel.style.display = type === 'excel' ? '' : 'none';
  if (curve) curve.style.display = type === 'curve' ? '' : 'none';
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
  if (delBtn) {
    delBtn.addEventListener('click', () => {
      row.remove();
      markUnsaved();
    });
  }
  tbody.appendChild(row);
  markUnsaved();
}

/**
 * 智能解析Excel单元格数值：自动识别百分比格式和小数格式
 * - 百分比格式（如15%）：XLSX读取为0.15，乘以100转为15
 * - 小数格式（如0.15）：值在[0,1]之间，乘以100转为15
 * - 普通数值（如15）：值大于1，直接使用
 * @param {object} sheet - XLSX工作表对象
 * @param {string} cellAddress - 单元格地址（如'B2'）
 * @param {number} rawValue - sheet_to_json读取的原始数值
 * @returns {number} 转换后的百分比数值
 */
function parsePercentCellValue(sheet, cellAddress, rawValue) {
  const cell = sheet[cellAddress];
  if (cell) {
    // 检查单元格格式是否为百分比格式
    const numFmt = (cell.z) || '';
    if (numFmt.includes('%')) {
      // 百分比格式：XLSX将15%存储为0.15，需乘以100
      return rawValue * 100;
    }
  }
  // 非百分比格式：判断是否为小数（0到1之间）
  if (rawValue > 0 && rawValue < 1) {
    return rawValue * 100;
  }
  return rawValue;
}

/**
 * 从 ArrayBuffer 解析 Excel 文件，返回 {day, value}[] 数据点数组。
 * 自动识别百分比格式与小数格式，供 handleRoiExcel / handleRetentionExcel 共用。
 * @param {ArrayBuffer} arrayBuffer - FileReader 读取的文件内容
 * @returns {{day: number, value: number}[]}
 */
function parseExcelToPoints(arrayBuffer) {
  const data = new Uint8Array(arrayBuffer);
  // 启用 cellNF 以获取单元格格式信息
  const workbook = XLSX.read(data, { type: 'array', cellNF: true });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
  const range = XLSX.utils.decode_range(firstSheet['!ref'] || 'A1');
  const points = [];

  for (let i = 0; i < jsonData.length; i++) {
    const row = jsonData[i];
    if (row && row.length >= 2) {
      const day = parseInt(row[0], 10);
      const rawValue = parseFloat(row[1]);
      if (!isNaN(day) && !isNaN(rawValue) && day > 0) {
        // 构建第2列单元格地址，用于识别百分比格式
        const cellAddress = XLSX.utils.encode_cell({ r: range.s.r + i, c: range.s.c + 1 });
        const value = parsePercentCellValue(firstSheet, cellAddress, rawValue);
        points.push({ day, value });
      }
    }
  }
  return points;
}

function handleRoiExcel(event) {
  handleExcelUpload(event, roiData, 'roiExcelPreview', 'ROI (%)');
}

// ========================================
// 留存率 管理
// ========================================
function toggleRetentionInput(type) {
  const manual = $id('retentionManualInput');
  const excel = $id('retentionExcelInput');
  const curve = $id('retentionCurveInput');
  if (manual) manual.style.display = type === 'manual' ? '' : 'none';
  if (excel) excel.style.display = type === 'excel' ? '' : 'none';
  if (curve) curve.style.display = type === 'curve' ? '' : 'none';
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
  if (delBtn) {
    delBtn.addEventListener('click', () => {
      row.remove();
      markUnsaved();
    });
  }
  tbody.appendChild(row);
  markUnsaved();
}

function handleRetentionExcel(event) {
  handleExcelUpload(event, retentionData, 'retentionExcelPreview', '留存率 (%)');
}

/**
 * Excel 文件上传公共处理函数，供 handleRoiExcel / handleRetentionExcel 共用。
 * @param {Event} event - input[type=file] 的 change 事件
 * @param {object} stateRef - 对应的模块状态对象（roiData 或 retentionData）
 * @param {string} previewId - 预览容器的元素 ID
 * @param {string} valueLabel - 预览表格第二列的列标题
 */
function handleExcelUpload(event, stateRef, previewId, valueLabel) {
  const file = event?.target?.files ? event.target.files[0] : null;
  if (!file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      stateRef.points = parseExcelToPoints(e.target.result);
      displayExcelPreview(previewId, stateRef.points.slice(0, 30), valueLabel, stateRef.points.length);
      markUnsaved();
    } catch (error) {
      alert('Excel文件解析失败：' + getErrorMessage(error));
    }
  };
  reader.readAsArrayBuffer(file);
}

function displayExcelPreview(containerId, data, label, totalCount = null) {
  const container = $id(containerId);
  if (!container) return;

  if (!data || data.length === 0) {
    container.innerHTML = '<p class="excel-preview-empty">未找到有效数据</p>';
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
  const noteText = total > 30
    ? `共导入 ${total} 条数据，预览前 30 条`
    : `共导入 ${total} 条数据`;
  html += `<p class="excel-preview-note">${noteText}</p>`;

  container.innerHTML = html;
}

// ========================================
// 收集数据
// ========================================

/**
 * 从手动输入表格中读取数据点，供 collectAllData / collectAllDataSync 共用。
 * @param {'roi'|'retention'} type - 数据类型
 * @returns {{day: number, value: number}[]}
 */
function collectManualDataPoints(type) {
  const isRoi = type === 'roi';
  const tbodyId = isRoi ? 'roiManualBody' : 'retentionManualBody';
  const dayClass = isRoi ? '.roi-day' : '.retention-day';
  const valueClass = isRoi ? '.roi-value' : '.retention-value';
  const points = [];
  document.querySelectorAll(`#${tbodyId} tr`).forEach(row => {
    const day = parseInt(row.querySelector(dayClass)?.value, 10);
    const value = parseFloat(row.querySelector(valueClass)?.value);
    if (!isNaN(day) && !isNaN(value)) points.push({ day, value });
  });
  return points;
}

/**
 * 收集指定类型的全局数据（ROI 或留存率），供 collectAllData / collectAllDataSync 共用。
 * @param {'roi'|'retention'} type - 数据类型
 * @param {object} stateRef - 对应的模块状态对象（roiData 或 retentionData）
 * @returns {{ type: string, points: {day: number, value: number}[] }}
 */
function collectGlobalDataByType(type, stateRef) {
  const isRoi = type === 'roi';
  const radioName = isRoi ? 'roiInputType' : 'retentionInputType';
  const inputTypeEl = document.querySelector(`input[name="${radioName}"]:checked`);
  const inputType = inputTypeEl ? inputTypeEl.value : (stateRef?.type || 'manual');
  const points = inputType === 'manual'
    ? collectManualDataPoints(type)
    : ((stateRef && stateRef.points) ? stateRef.points : []);
  const result = { type: inputType, points };
  // 若为曲线模式，保存曲线名称以便项目加载时恢复
  if (inputType === 'curve' && stateRef && stateRef.curveName) {
    result.curveName = stateRef.curveName;
  }
  return result;
}

async function collectAllData() {
  const data = {
    project_name: ($id('projectName')?.value || '').trim(),
    investment_periods: [],
    roi_data: { type: 'manual', points: [] },
    retention_data: { type: 'manual', points: [] },
    repayment_months: parseInt($id('repaymentMonths')?.value, 10) || 1,
    target_dau: parseInt($id('targetDau')?.value, 10) || 1000
  };

  const allCurves = await loadAllCurves();

  document.querySelectorAll('.period-card').forEach(card => {
    const periodId = (card.id || '').split('-')[1] || '';
    const costTypeEl = document.querySelector(`input[name="costType-${periodId}"]:checked`);
    const costType = costTypeEl ? costTypeEl.value : 'fixed';

    // 读取该时间段选择的曲线ID
    const roiCurveId = $qs('.period-roi-curve', card)?.value || '';
    const retentionCurveId = $qs('.period-retention-curve', card)?.value || '';

    // 根据曲线ID查找对应的曲线数据
    const roiCurveObj = roiCurveId ? allCurves.find(c => c.id === roiCurveId && c.type === 'roi') : null;
    const retentionCurveObj = retentionCurveId ? allCurves.find(c => c.id === retentionCurveId && c.type === 'retention') : null;

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
      other_cost: parseFloat($qs('.period-other-cost', card)?.value) || 0,
      // 曲线选择信息（用于保存/恢复）
      roi_curve_id: roiCurveId,
      retention_curve_id: retentionCurveId,
      // 曲线数据点（用于后端计算）
      roi_curve_data: roiCurveObj ? { type: 'manual', points: roiCurveObj.points } : null,
      retention_curve_data: retentionCurveObj ? { type: 'manual', points: retentionCurveObj.points } : null
    };

    data.investment_periods.push(period);
  });

  data.roi_data = collectGlobalDataByType('roi', roiData);
  data.retention_data = collectGlobalDataByType('retention', retentionData);

  return data;
}

// ========================================
// 计算与结果展示
// ========================================
async function startCalculation() {
  // 计算前强制自动保存一次
  await saveProject(true);

  const data = await collectAllData();

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
    alert('计算失败：' + getErrorMessage(error));
  } finally {
    if (loading) loading.style.display = 'none';
  }
}

/**
 * 安全设置指标卡片的文本内容，元素不存在时静默跳过。
 * @param {string} id - 元素 ID
 * @param {string} text - 要设置的文本
 */
function setMetricText(id, text) {
  const el = $id(id);
  if (el) el.textContent = text;
}

function displayResults(results) {
  const metrics = results.key_metrics || {};

  // 固定指标卡片
  setMetricText('metricMaxCash1m',          formatCash(metrics.max_cash_demand_1m));
  setMetricText('metricMaxCash2m',          formatCash(metrics.max_cash_demand_2m));
  setMetricText('metricDynamicBreakeven',   formatDays(metrics.dynamic_profit_breakeven_day));
  setMetricText('metricCumulativeBreakeven',formatDays(metrics.cumulative_profit_breakeven_day));
  setMetricText('metricCashFlow1mBreakeven',formatDays(metrics.cumulative_cash_flow_1m_breakeven_day));
  setMetricText('metricCashFlow2mBreakeven',formatDays(metrics.cumulative_cash_flow_2m_breakeven_day));
  setMetricText('metricDay10mDau',          formatDays(metrics.day_to_10m_dau));
  setMetricText('metricDay2mDau',           formatDays(metrics.day_to_2m_dau));
  setMetricText('metricDayTargetDau',       formatDays(metrics.day_to_target_dau));
  setMetricText('metricTargetDauValue',     String(metrics.target_dau ?? '--'));

  // N 个月回款动态卡片（仅在 repaymentMonths > 2 时显示）
  const repaymentMonths = metrics.repayment_months || 1;
  const showNm = repaymentMonths > 2;
  const nmCardDisplay = showNm ? '' : 'none';
  $id('metricMaxCashNmCard').style.display        = nmCardDisplay;
  $id('metricCashFlowNmBreakevenCard').style.display = nmCardDisplay;
  if (showNm) {
    setMetricText('metricMaxCashNmLabel',          `${repaymentMonths}个月回款最大现金需求`);
    setMetricText('metricCashFlowNmBreakevenLabel',`${repaymentMonths}个月回款现金流打正天数`);
    setMetricText('metricMaxCashNm',               formatCash(metrics.max_cash_demand_nm));
    setMetricText('metricCashFlowNmBreakeven',     formatDays(metrics.cumulative_cash_flow_nm_breakeven_day));
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
        labels: chartsData.dau_quarterly?.labels || [],
        datasets: [{
          label: 'DAU (万人)',
          data: chartsData.dau_quarterly?.data || [],
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
          tooltip: {
            callbacks: {
              label: (context) => `DAU: ${context.parsed.y.toLocaleString()} 万人`
            }
          }
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
        plugins: {
          legend: { display: true, position: 'top' },
          tooltip: {
            callbacks: {
              label(context) {
                let label = context.dataset.label || '';
                let value = context.parsed.y;
                if (label === '当期成本') value = Math.abs(value);
                return `${label}: ${value.toLocaleString()} 万元`;
              }
            }
          }
        },
        scales: {
          x: { stacked: true },
          y: {
            stacked: false,
            title: { display: true, text: '金额 (万元)' },
            ticks: {
              callback(value) {
                return Math.abs(value).toLocaleString();
              }
            }
          }
        }
      }
    });
  }
}

// ========================================
// 季度表格
// ========================================

/**
 * 构建季度表格的行配置列表（updateQuarterlyTable 和 downloadQuarterlyTable 共用）
 * @param {number} repaymentMonths - 回款月数
 * @returns {Array<{label: string, field: string, styled?: boolean, indent?: boolean}>}
 */
function buildQuarterlyRowConfigs(repaymentMonths) {
  const configs = [
    { label: '季度末时间点',           field: 'end_date',               styled: false, indent: false },
    { label: '季度末累计天数',          field: 'cumulative_days',         styled: false, indent: false },
    { label: '累计收入',               field: 'cumulative_revenue',      styled: true,  indent: false },
    { label: '累计成本',               field: 'cumulative_cost',         styled: true,  indent: false },
    { label: ' -UA成本',              field: 'ua_cost',                 styled: false, indent: true  },
    { label: ' -人员成本',             field: 'personnel_cost',          styled: false, indent: true  },
    { label: ' -其他成本',             field: 'other_cost',              styled: false, indent: true  },
    { label: '累计利润',               field: 'cumulative_profit',       styled: true,  indent: false },
    { label: '累计净现金流（1个月回款）', field: 'cumulative_cash_flow_1m', styled: true,  indent: false },
    { label: '累计净现金流（2个月回款）', field: 'cumulative_cash_flow_2m', styled: true,  indent: false },
  ];

  if (repaymentMonths > 2) {
    configs.push({ label: `累计净现金流（${repaymentMonths}个月回款）`, field: 'cumulative_cash_flow_nm', styled: true, indent: false });
  }

  configs.push(
    { label: '当期收入',               field: 'current_revenue',         styled: true,  indent: false },
    { label: '当期成本',               field: 'current_cost',            styled: true,  indent: false },
    { label: ' -UA成本',              field: 'current_ua_cost',         styled: false, indent: true  },
    { label: ' -人员成本',             field: 'current_personnel_cost',  styled: false, indent: true  },
    { label: ' -其他成本',             field: 'current_other_cost',      styled: false, indent: true  },
    { label: '当期利润',               field: 'current_profit',          styled: true,  indent: false },
    { label: '当期资金需求（1个月回款）', field: 'current_cash_demand_1m',  styled: false, indent: false },
    { label: '当期资金需求（2个月回款）', field: 'current_cash_demand_2m',  styled: false, indent: false }
  );

  if (repaymentMonths > 2) {
    configs.push({ label: `当期资金需求（${repaymentMonths}个月回款）`, field: 'current_cash_demand_nm', styled: false, indent: false });
  }

  configs.push({ label: 'DAU', field: 'dau', styled: false, indent: false });
  return configs;
}

function updateQuarterlyTable(tableData, repaymentMonths = 1) {
  const thead = $id('quarterlyTableHead');
  const tbody = $id('quarterlyTableBody');
  if (!thead || !tbody) return;
  thead.innerHTML = '';
  tbody.innerHTML = '';

  if (!tableData || tableData.length === 0) return;

  const rowConfigs = buildQuarterlyRowConfigs(repaymentMonths);

  const headerRow = document.createElement('tr');
  const indicatorTh = document.createElement('th');
  indicatorTh.className = 'indicator-header';
  indicatorTh.textContent = '指标';
  headerRow.appendChild(indicatorTh);
  tableData.forEach(row => {
    const th = document.createElement('th');
    th.textContent = row.quarter || '--';
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  rowConfigs.forEach(config => {
    const tr = document.createElement('tr');
    const labelClass = config.indent ? 'indent-label' : 'indicator-label';
    tr.innerHTML = `<td class="${labelClass}">${config.label}</td>`;

    tableData.forEach(row => {
      const value = row[config.field];
      let displayValue;
      let cellClass = '';

      if (config.field === 'quarter' || config.field === 'end_date') {
        // 日期/季度字段：直接显示字符串
        displayValue = value || '--';
      } else if (value === null || value === undefined) {
        displayValue = '--';
      } else if (config.field === 'cumulative_days') {
        // 天数字段：整数，不需要格式化
        displayValue = value;
      } else {
        displayValue = formatNumber(value);
        if (config.styled) cellClass = (typeof value === 'number' && value >= 0) ? 'positive' : 'negative';
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

function downloadQuarterlyTable() {
  if (!calculationResults || !calculationResults.quarterly_table_data) {
    alert('没有可导出的季度汇总数据，请先进行计算');
    return;
  }

  const tableData = calculationResults.quarterly_table_data;
  const projectName = ($id('projectName')?.value || 'IAA计算').trim();
  const repaymentMonths = calculationResults.key_metrics?.repayment_months || 1;

  const rowConfigs = buildQuarterlyRowConfigs(repaymentMonths);

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
  // 截图前需要临时展开容器以获取完整表格尺寸，完成后恢复
  const originalStyles = {
    container: tableContainer
      ? {
        overflow: tableContainer.style.overflow,
        overflowX: tableContainer.style.overflowX,
        maxWidth: tableContainer.style.maxWidth,
        width: tableContainer.style.width,
        padding: tableContainer.style.padding,
      }
      : null,
    table: {
      width: table.style.width,
      margin: table.style.margin,
    },
  };

  // 提取样式恢复逻辑，避免在 .then / .catch 中重复
  function restoreTableStyles() {
    if (tableContainer && originalStyles.container) {
      tableContainer.style.overflow = originalStyles.container.overflow;
      tableContainer.style.overflowX = originalStyles.container.overflowX;
      tableContainer.style.maxWidth = originalStyles.container.maxWidth;
      tableContainer.style.width = originalStyles.container.width;
      tableContainer.style.padding = originalStyles.container.padding;
    }
    table.style.width = originalStyles.table.width;
    table.style.margin = originalStyles.table.margin;
  }

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

  html2canvas(table, {
    backgroundColor: '#ffffff',
    scale: 2,
    useCORS: true,
    logging: false,
    scrollX: 0,
    scrollY: 0,
    width: fullWidth,
    height: fullHeight,
    windowWidth: fullWidth + 100,
    windowHeight: fullHeight + 100,
    x: -10,
    y: -10,
  }).then(canvas => {
    restoreTableStyles();
    const link = document.createElement('a');
    link.download = `${projectName}_季度汇总数据.png`;
    link.href = canvas.toDataURL('image/png');
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  }).catch(error => {
    restoreTableStyles();
    alert('生成图片失败：' + getErrorMessage(error));
  });
}

/**
 * 同步版数据收集，专用于 beforeunload 场景。
 * 不依赖服务端曲线查询（无法在卸载阶段 await），
 * 只保存曲线 ID，不保存曲线数据点（加载时通过 ID 重新关联）。
 */
function collectAllDataSync() {
  const data = {
    project_name: ($id('projectName')?.value || '').trim(),
    investment_periods: [],
    roi_data: { type: 'manual', points: [] },
    retention_data: { type: 'manual', points: [] },
    repayment_months: parseInt($id('repaymentMonths')?.value, 10) || 1,
    target_dau: parseInt($id('targetDau')?.value, 10) || 1000
  };

  document.querySelectorAll('.period-card').forEach(card => {
    const periodId = (card.id || '').split('-')[1] || '';
    const costTypeEl = document.querySelector(`input[name="costType-${periodId}"]:checked`);
    const costType = costTypeEl ? costTypeEl.value : 'fixed';

    const roiCurveId = $qs('.period-roi-curve', card)?.value || '';
    const retentionCurveId = $qs('.period-retention-curve', card)?.value || '';

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
      other_cost: parseFloat($qs('.period-other-cost', card)?.value) || 0,
      roi_curve_id: roiCurveId,
      retention_curve_id: retentionCurveId,
      // beforeunload 场景下不查询服务端，曲线数据点留空，加载时通过 ID 重新关联
      roi_curve_data: null,
      retention_curve_data: null
    };

    data.investment_periods.push(period);
  });

  data.roi_data = collectGlobalDataByType('roi', roiData);
  data.retention_data = collectGlobalDataByType('retention', retentionData);

  return data;
}

// 页面关闭前自动保存 —— sendBeacon 不受页面卸载中断影响
window.addEventListener('beforeunload', () => {
  try {
    // 使用同步版本，避免 async collectAllData() 在卸载阶段无法 await 导致曲线数据丢失
    const data = collectAllDataSync();
    const blob = new Blob([JSON.stringify(data)], {
      type: 'application/json'
    });
    navigator.sendBeacon('/save_project', blob);
  } catch (_) {
    // 静默失败：卸载阶段无法向用户反馈错误
  }
});

// ========================================
// 保存状态标记 —— 类名与 CSS 中 .is-saved / .is-unsaved 保持一致
// ========================================
function markUnsaved() {
  const el = $id('saveStatus');
  if (!el) return;
  el.textContent = '未保存';
  el.classList.remove('is-saved');
  el.classList.add('is-unsaved');
}

function markSaved() {
  const el = $id('saveStatus');
  if (!el) return;
  el.textContent = '已保存';
  el.classList.remove('is-unsaved');
  el.classList.add('is-saved');
}

function hasUnsavedChanges() {
  const el = $id('saveStatus');
  return el && el.classList.contains('is-unsaved');
}

// ========================================
// 数据曲线服务端存储管理
// 存储路径：roi_data/<id>.json  /  retention_data/<id>.json
// 数据结构：{ id, name, type, points, createdAt, updatedAt }
// ========================================

/** 从服务端读取所有已保存的曲线（全部类型） */
async function loadAllCurves() {
  try {
    const response = await fetch('/list_all_curves');
    const result = await response.json();
    return (result?.success && Array.isArray(result.curves)) ? result.curves : [];
  } catch (_) {
    return [];
  }
}

/** 从服务端读取指定类型的曲线 */
async function loadCurvesByType(type) {
  try {
    const response = await fetch(`/list_curves?type=${encodeURIComponent(type)}`);
    const result = await response.json();
    return (result?.success && Array.isArray(result.curves)) ? result.curves : [];
  } catch (_) {
    return [];
  }
}

/** 保存单条曲线到服务端 */
async function saveCurveToServer(curve) {
  const response = await fetch('/save_curve', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(curve),
  });
  return response.json();
}

/** 生成简单唯一 ID */
function genCurveId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

// 当前待保存的曲线类型（'roi' | 'retention'），由对话框打开时设置
let _pendingCurveType = 'roi';
// 当前加载对话框的曲线类型
let _loadingCurveType = 'roi';
// 全量曲线缓存（对话框打开时刷新）
let _curveListCache = [];

/** 收集当前指定类型的数据点 */
function collectCurvePoints(type) {
  if (type === 'roi') {
    const inputType = $qs('input[name="roiInputType"]:checked');
    const mode = inputType ? inputType.value : 'manual';
    if (mode === 'excel') {
      return roiData.points.slice();
    }
    const points = [];
    const rows = document.querySelectorAll('#roiManualBody tr');
    rows.forEach(row => {
      const day = parseFloat(row.querySelector('.roi-day')?.value);
      const value = parseFloat(row.querySelector('.roi-value')?.value);
      if (!isNaN(day) && !isNaN(value) && day > 0) {
        points.push({ day, value });
      }
    });
    return points;
  } else {
    const inputType = $qs('input[name="retentionInputType"]:checked');
    const mode = inputType ? inputType.value : 'manual';
    if (mode === 'excel') {
      return retentionData.points.slice();
    }
    const points = [];
    const rows = document.querySelectorAll('#retentionManualBody tr');
    rows.forEach(row => {
      const day = parseFloat(row.querySelector('.retention-day')?.value);
      const value = parseFloat(row.querySelector('.retention-value')?.value);
      if (!isNaN(day) && !isNaN(value) && day > 0) {
        points.push({ day, value });
      }
    });
    return points;
  }
}

/** 打开保存曲线对话框 */
function showCurveSaveDialog(type) {
  _pendingCurveType = type;
  const nameInput = $id('curveNameInput');
  const errEl = $id('curveNameError');
  if (nameInput) nameInput.value = '';
  if (errEl) errEl.style.display = 'none';
  const dialog = $id('curveSaveDialog');
  if (dialog) {
    dialog.style.display = 'flex';
    dialog.setAttribute('aria-hidden', 'false');
    setTimeout(() => nameInput && nameInput.focus(), 50);
  }
  // 回车确认
  if (nameInput) {
    nameInput.onkeydown = (e) => { if (e.key === 'Enter') confirmCurveSave(); };
  }
}

/** 关闭保存曲线对话框 */
function closeCurveSaveDialog() {
  const dialog = $id('curveSaveDialog');
  if (dialog) {
    dialog.style.display = 'none';
    dialog.setAttribute('aria-hidden', 'true');
  }
}

/** 确认保存曲线 */
async function confirmCurveSave() {
  const nameInput = $id('curveNameInput');
  const errEl = $id('curveNameError');
  const name = nameInput ? nameInput.value.trim() : '';

  // 校验：非空
  if (!name) {
    if (errEl) { errEl.textContent = '曲线名称不能为空'; errEl.style.display = ''; }
    nameInput && nameInput.focus();
    return;
  }

  const points = collectCurvePoints(_pendingCurveType);
  if (points.length === 0) {
    if (errEl) { errEl.textContent = '当前没有有效数据点，请先输入数据'; errEl.style.display = ''; }
    return;
  }

  // 从服务端加载同类型曲线，检查是否重名
  const curves = await loadCurvesByType(_pendingCurveType);
  const existing = curves.find(c => c.name === name);

  if (existing) {
    // 重名：询问是否覆盖
    const ok = confirm(`已存在同名曲线「${name}」，是否覆盖更新？`);
    if (!ok) { nameInput && nameInput.focus(); return; }
    // 覆盖更新：复用原有ID
    const updatedCurve = {
      id: existing.id,
      name,
      type: _pendingCurveType,
      points,
      createdAt: existing.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
    try {
      const result = await saveCurveToServer(updatedCurve);
      if (!result?.success) throw new Error(result?.error || '保存失败');
      closeCurveSaveDialog();
      showToast(`曲线「${name}」已更新`);
      await updateAllPeriodCurveSelectors();
    } catch (e) {
      if (errEl) { errEl.textContent = '保存失败：' + getErrorMessage(e); errEl.style.display = ''; }
    }
    return;
  }

  // 新建
  const now = new Date().toISOString();
  const newCurve = {
    id: genCurveId(),
    name,
    type: _pendingCurveType,
    points,
    createdAt: now,
    updatedAt: now,
  };
  try {
    const result = await saveCurveToServer(newCurve);
    if (!result?.success) throw new Error(result?.error || '保存失败');
    closeCurveSaveDialog();
    showToast(`曲线「${name}」已保存（${points.length} 个数据点）`);
    await updateAllPeriodCurveSelectors();
  } catch (e) {
    if (errEl) { errEl.textContent = '保存失败：' + getErrorMessage(e); errEl.style.display = ''; }
  }
}

/** 打开加载曲线对话框 */
async function showCurveLoadDialog(type) {
  _loadingCurveType = type;
  const searchInput = $id('curveSearchInput');
  if (searchInput) searchInput.value = '';
  // 先显示对话框，再异步加载列表
  const dialog = $id('curveLoadDialog');
  if (dialog) {
    const title = $id('curveLoadDialogTitle');
    if (title) title.textContent = `加载${type === 'roi' ? 'ROI' : '留存率'}数据曲线`;
    dialog.style.display = 'flex';
    dialog.setAttribute('aria-hidden', 'false');
  }
  const list = $id('curveList');
  if (list) list.innerHTML = '<li class="list-status">正在加载...</li>';
  _curveListCache = await loadCurvesByType(type);
  renderCurveList(_curveListCache);
}

/** 关闭加载曲线对话框 */
function closeCurveLoadDialog() {
  const dialog = $id('curveLoadDialog');
  if (dialog) {
    dialog.style.display = 'none';
    dialog.setAttribute('aria-hidden', 'true');
  }
}

/** 渲染曲线列表 */
function renderCurveList(curves) {
  const list = $id('curveList');
  if (!list) return;
  list.innerHTML = '';

  if (!curves || curves.length === 0) {
    list.innerHTML = '<li class="list-status">暂无已保存的曲线</li>';
    return;
  }

  curves.forEach(curve => {
    const li = document.createElement('li');
    const typeLabel = curve.type === 'roi' ? 'ROI' : '留存率';
    const updatedTime = curve.updatedAt ? new Date(curve.updatedAt).toLocaleString('zh-CN') : '';
    li.innerHTML = `
      <div class="project-info">
        <span class="project-name">
          ${curve.name}
          <span class="curve-type-badge ${curve.type}">${typeLabel}</span>
        </span>
        <span class="project-time curve-points">${curve.points.length} 个数据点 · ${updatedTime}</span>
      </div>
      <div class="project-actions">
        <button class="btn-load-curve">加载</button>
        <button class="btn-delete-curve">删除</button>
      </div>
    `;

    const loadBtn = li.querySelector('.btn-load-curve');
    if (loadBtn) {
      loadBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        applyCurveData(curve);
      });
    }

    const delBtn = li.querySelector('.btn-delete-curve');
    if (delBtn) {
      delBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        deleteCurve(curve.id, curve.name, curve.type);
      });
    }

    // 点击整行也触发加载
    li.addEventListener('click', () => applyCurveData(curve));
    list.appendChild(li);
  });
}

/** 搜索过滤曲线列表 */
function filterCurveList() {
  const keyword = ($id('curveSearchInput')?.value || '').trim().toLowerCase();
  if (!keyword) {
    renderCurveList(_curveListCache);
    return;
  }
  const filtered = _curveListCache.filter(c => c.name.toLowerCase().includes(keyword));
  renderCurveList(filtered);
}

/** 将曲线数据应用到对应输入区 */
function applyCurveData(curve) {
  if (curve.type === 'roi') {
    // 显示曲线 radio 选项并更新标签文字
    const radioLabel = $id('roiCurveRadioLabel');
    const radioText = $id('roiCurveRadioText');
    if (radioLabel) radioLabel.style.display = '';
    if (radioText) radioText.textContent = `${curve.name}`;
    // 切换到曲线模式
    const radio = $qs('input[name="roiInputType"][value="curve"]');
    if (radio) { radio.checked = true; toggleRoiInput('curve'); }
    // 写入数据点和曲线名称
    roiData.points = curve.points.map(p => ({ day: p.day, value: p.value }));
    roiData.curveName = curve.name;
    // 渲染曲线预览：前30条，格式与 Excel 导入一致
    const previewData = roiData.points.slice(0, 30);
    displayExcelPreview('roiCurvePreview', previewData, 'ROI (%)', roiData.points.length);
  } else {
    // 显示曲线 radio 选项并更新标签文字
    const radioLabel = $id('retentionCurveRadioLabel');
    const radioText = $id('retentionCurveRadioText');
    if (radioLabel) radioLabel.style.display = '';
    if (radioText) radioText.textContent = `${curve.name}`;
    // 切换到曲线模式
    const radio = $qs('input[name="retentionInputType"][value="curve"]');
    if (radio) { radio.checked = true; toggleRetentionInput('curve'); }
    retentionData.points = curve.points.map(p => ({ day: p.day, value: p.value }));
    retentionData.curveName = curve.name;
    const previewData = retentionData.points.slice(0, 30);
    displayExcelPreview('retentionCurvePreview', previewData, '留存率 (%)', retentionData.points.length);
  }
  markUnsaved();
  closeCurveLoadDialog();
  showToast(`已加载曲线「${curve.name}」（${curve.points.length} 个数据点）`);
}

/** 删除指定曲线 */
async function deleteCurve(id, name, type) {
  if (!confirm(`确定要删除曲线「${name}」吗？此操作不可撤销。`)) return;
  try {
    const response = await fetch('/delete_curve', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ id, type: type || _loadingCurveType }),
    });
    const result = await response.json();
    if (!result?.success) throw new Error(result?.error || '删除失败');
    // 刷新缓存和列表
    _curveListCache = await loadCurvesByType(_loadingCurveType);
    const keyword = ($id('curveSearchInput')?.value || '').trim().toLowerCase();
    const filtered = keyword ? _curveListCache.filter(c => c.name.toLowerCase().includes(keyword)) : _curveListCache;
    renderCurveList(filtered);
    showToast(`曲线「${name}」已删除`);
    // 刷新所有时间段的曲线选择器
    await updateAllPeriodCurveSelectors();
  } catch (e) {
    alert('删除失败：' + getErrorMessage(e));
  }
}

/**
 * 刷新所有时间段卡片中的曲线下拉选择器选项
 * 在曲线保存/删除后调用，确保选项列表与服务端存储同步
 */
async function updateAllPeriodCurveSelectors() {
  const [roiCurves, retentionCurves] = await Promise.all([
    loadCurvesByType('roi'),
    loadCurvesByType('retention'),
  ]);

  document.querySelectorAll('.period-card').forEach(card => {
    const roiSelect = $qs('.period-roi-curve', card);
    const retentionSelect = $qs('.period-retention-curve', card);

    if (roiSelect) {
      const currentVal = roiSelect.value;
      let opts = '<option value="">全局曲线</option>';
      roiCurves.forEach(c => {
        const sel = c.id === currentVal ? 'selected' : '';
        opts += `<option value="${c.id}" ${sel}>${c.name}</option>`;
      });
      roiSelect.innerHTML = opts;
    }

    if (retentionSelect) {
      const currentVal = retentionSelect.value;
      let opts = '<option value="">全局曲线</option>';
      retentionCurves.forEach(c => {
        const sel = c.id === currentVal ? 'selected' : '';
        opts += `<option value="${c.id}" ${sel}>${c.name}</option>`;
      });
      retentionSelect.innerHTML = opts;
    }
  });
}

/** 轻量 Toast 提示（2秒自动消失），样式由 CSS .toast / .toast.is-hidden 管理 */
function showToast(message) {
  let toast = $id('_curveToast');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = '_curveToast';
    toast.className = 'toast';
    document.body.appendChild(toast);
  }
  toast.textContent = message;
  toast.classList.remove('is-hidden');
  clearTimeout(toast._timer);
  toast._timer = setTimeout(() => { toast.classList.add('is-hidden'); }, 2000);
}

// ========================================
// 将 HTML 行内事件处理器需要的函数挂载到全局
// ========================================
Object.assign(window, {
  createNewProject,
  closeNewProjectDialog,
  confirmNewProject,
  showLoadDialog,
  closeLoadDialog,
  saveProject,
  backToWelcome,
  addInvestmentPeriod,
  addRoiRow,
  addRetentionRow,
  toggleRoiInput,
  toggleRetentionInput,
  handleRoiExcel,
  handleRetentionExcel,
  startCalculation,
  exportCsv,
  downloadChart,
  downloadQuarterlyTable,
  downloadQuarterlyTableAsImage,
  // 数据曲线管理
  showCurveSaveDialog,
  closeCurveSaveDialog,
  confirmCurveSave,
  showCurveLoadDialog,
  closeCurveLoadDialog,
  filterCurveList,
  updateAllPeriodCurveSelectors,
});

})(); // IIFE 结束
