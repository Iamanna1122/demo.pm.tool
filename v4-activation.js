// Design Spec V4 final activation.
// This file is loaded after app.js, so it becomes the final UI contract.

getAllProjectsForCleanUI = function() {
  return (db.projects || []).filter(Boolean);
};

getProjectDisplay = function(project) {
  const custom = PROJECT_DISPLAY_V4[project.id] || {};
  const inferredGate = inferGateFromProjectV4(project);
  return {
    id: project.id,
    name: custom.name || `${project.id} · ${textOr(project.name, project.type || '產品專案')}`,
    spec: custom.spec || textOr(project.name, project.type || '待確認規格'),
    version: project.cleanGateState?.version || custom.version || textOr(project.currentVersion, 'V1'),
    gate: Number(project.cleanGateState?.gate || custom.gate || inferredGate),
    status: project.cleanGateState?.status || custom.status || normalizeStatusV4(project.status),
    priority: Number(project.priorityV4 || project.priority || custom.priority || 3),
    owner: project.cleanGateState?.owner || custom.owner || textOr(project.owner, 'PM 曼玉'),
    tracker: custom.tracker || project.assignmentV3?.pm || '曼玉',
    ngReason: project.cleanGateState?.ngReason || custom.ngReason || '',
    resultDue: project.cleanGateState?.resultDue || custom.resultDue || '',
    nextAction: project.cleanGateState?.nextAction || custom.nextAction || '',
    startDate: custom.startDate || project.milestones?.[0]?.target || '2026-06-01',
    planoDate: custom.planoDate || '2026-06-10',
    powerDate: custom.powerDate || project.deadline || '2026-06-25'
  };
};

prepareCleanNavigation = function() {
  document.title = '產品開發 Gate 管理系統';
  const brand = document.querySelector('.brand-text');
  if (brand) brand.textContent = '專案管理';
  const footer = document.getElementById('server-status-text');
  if (footer) footer.textContent = '本機資料已載入';
  setMenuItem('nav-vp-milestones', '我的工作', 'vp-milestones', true);
  setMenuItem('nav-project-master', '專案總覽', 'project-master', true);
  setMenuItem('nav-action-items', '專案進度', 'action-items', true);
  setMenuItem('nav-risks-issues', '團隊成員', 'risks-issues', true);
  setMenuItem('nav-gantt-chart', '月曆甘特', 'action-items', false);
  setMenuItem('nav-version-history', '版本紀錄', 'version-history', false);
  setMenuItem('nav-admin-settings', '系統管理', 'admin-settings', false);
  document.querySelectorAll('.menu-item').forEach(item => {
    item.onclick = (event) => {
      event.preventDefault();
      const tab = item.dataset.tab;
      if (tab === 'project-master') currentProjectDetailId = null;
      switchTab(tab);
    };
  });
  updateRoleSelectLabelsV4();
  updateLoginHelperV4();
};

initApp = function() {
  ensureRoleUsersV3();
  prepareCleanNavigation();
  const displayName = document.getElementById('user-display-name');
  const displayDept = document.getElementById('user-display-dept');
  if (displayName) displayName.innerText = currentUser.name || currentUser.username || '使用者';
  if (displayDept) displayDept.innerText = `${currentUser.dept || ''} · ${ROLE_LABELS_V4[activePerspective] || activePerspective}`;
  ensureNotificationBell();
  seedGateNotifications();
  applyPermissions();
  switchTab('vp-milestones');
};

switchTab = function(tabId) {
  document.querySelectorAll('.menu-item').forEach(item => {
    item.classList.toggle('active', item.dataset.tab === tabId);
  });
  document.querySelectorAll('.tab-pane, .tab-content').forEach(pane => pane.classList.remove('active'));
  const targetPane = document.getElementById(`tab-${tabId}`);
  if (targetPane) targetPane.classList.add('active');
  if (tabId === 'vp-milestones') renderVPMilestones();
  else if (tabId === 'project-master') renderProjectMaster();
  else if (tabId === 'action-items') renderActionItems();
  else if (tabId === 'risks-issues') renderRisksIssues();
  else if (tabId === 'admin-settings' && typeof renderAdminSettings === 'function') renderAdminSettings();
};

renderVPMilestones = function() {
  const container = document.getElementById('tab-vp-milestones');
  if (!container) return;
  currentProjectDetailId = null;
  const projects = getAllProjectsForCleanUI();
  const priorityOne = projects.filter(p => getProjectDisplay(p).priority === 1);
  const ngProjects = projects.filter(p => getProjectDisplay(p).status.includes('NG'));
  const cards = priorityOne.concat(projects.filter(p => !priorityOne.includes(p)).slice(0, 2));
  container.innerHTML = `
    <section class="simple-page v4-page">
      <div class="page-heading v4-heading">
        <p class="eyebrow">My Work</p>
        <h1>我的工作</h1>
        <p>${escapeHtml(ROLE_LABELS_V4[activePerspective] || activePerspective)} 視角：先看要處理哪一個專案、下一步要按哪裡。</p>
      </div>
      <div class="simple-stat-grid v4-stat-grid">
        <button class="simple-stat-card danger clickable-card" onclick="openProjectGateCenter('${escapeAttribute(ngProjects[0]?.id || priorityOne[0]?.id || projects[0]?.id || '')}')"><span>高優先處理</span><strong>${ngProjects.length}</strong><small>點擊進入對應專案</small></button>
        <button class="simple-stat-card warning clickable-card" onclick="switchTab('action-items')"><span>今日需追蹤</span><strong>${priorityOne.length}</strong><small>查看月曆與里程碑</small></button>
        <button class="simple-stat-card success clickable-card" onclick="switchTab('project-master')"><span>進行中專案</span><strong>${projects.length}</strong><small>進入專案總覽</small></button>
      </div>
      <div class="section-title-row"><h2>現在要處理的專案</h2><button class="btn btn-primary" onclick="switchTab('project-master')">查看全部專案</button></div>
      <div class="work-card-list v4-work-list">${cards.map(p => renderMyWorkCard(p)).join('')}</div>
    </section>
  `;
};

renderMyWorkCard = function(project) {
  const d = getProjectDisplay(project);
  const gate = GATE_DEFS_V4[String(d.gate)] || GATE_DEFS_V4['1'];
  return `<article class="work-card v4-work-card" onclick="openProjectGateCenter('${escapeAttribute(project.id)}')">
    <div><span class="priority-rank">第 ${d.priority} 優先</span><h3>${escapeHtml(d.name)}</h3><p>目前：Gate ${d.gate} ${escapeHtml(gate.name)}</p></div>
    <div class="work-card-meta"><span class="status-pill ${statusClassV4(d.status)}">${escapeHtml(d.status)}</span><span>負責：${escapeHtml(d.owner)}</span><span>下一步：${escapeHtml(d.nextAction || '等待下一個 Gate')}</span></div>
  </article>`;
};

renderProjectMaster = function() {
  const container = document.getElementById('tab-project-master');
  if (!container) return;
  if (currentProjectDetailId) return renderProjectGateCenter(currentProjectDetailId);
  const projects = getAllProjectsForCleanUI().sort((a, b) => getProjectDisplay(a).priority - getProjectDisplay(b).priority);
  const groups = [1, 2, 3].map(priority => ({
    priority,
    title: priority === 1 ? '第一優先' : priority === 2 ? '第二優先' : '第三優先',
    projects: projects.filter(project => Math.min(3, getProjectDisplay(project).priority) === priority)
  }));
  container.innerHTML = `<section class="simple-page v4-page"><div class="page-heading v4-heading"><p class="eyebrow">Projects</p><h1>專案總覽</h1><p>依優先順序排列。點任一張專案卡片，就會進入該專案的 Gate 中心。</p></div>${groups.map(group => renderPriorityProjectSectionV4(group)).join('')}</section>`;
};

renderProjectGateCenter = function(projectId) {
  const container = document.getElementById('tab-project-master');
  const project = db.projects.find(p => p.id === projectId);
  if (!container || !project) return;
  ensureGateCenterV4(project);
  const d = getProjectDisplay(project);
  const currentPhase = getCurrentPhaseV4(project);
  const activePhaseId = getActivePhaseV4(project);
  const activePhase = GATE_PHASES_V4.find(p => p.id === activePhaseId) || GATE_PHASES_V4[0];
  const activeVersion = getActiveVersionV4(project, activePhaseId);
  const currentGate = GATE_DEFS_V4[String(d.gate)] || GATE_DEFS_V4['1'];
  container.innerHTML = `
    <section class="simple-page v4-page gate-center-page">
      <div class="page-action-row"><button class="back-button" onclick="currentProjectDetailId=null; renderProjectMaster();">← 回專案總覽</button><button class="btn btn-outline" onclick="exportProjectMeetingExcel('${escapeAttribute(project.id)}')">匯出會議 Excel</button></div>
      <div class="project-detail-heading v4-heading"><p class="eyebrow">Project Gate Center</p><h1>${escapeHtml(d.name)}</h1></div>
      <article class="gate-summary-card v4-summary ${statusClassV4(d.status)}"><div class="summary-main"><span class="status-pill ${statusClassV4(d.status)}">${escapeHtml(d.status)}</span><h2>目前進度：Gate ${d.gate} ${escapeHtml(currentGate.name)}</h2><p>${escapeHtml(d.ngReason || currentGate.description)}</p></div><div class="summary-grid"><div><span>目前大項目</span><strong>${escapeHtml(GATE_PHASES_V4.find(p => p.id === currentPhase)?.label || '')}</strong></div><div><span>目前版本</span><strong>${escapeHtml(d.version)}</strong></div><div><span>下一步</span><strong>${escapeHtml(d.nextAction || '等待下一個 Gate')}</strong></div><div><span>負責 / PM追蹤</span><strong>${escapeHtml(d.owner)} / ${escapeHtml(d.tracker)}</strong></div></div></article>
      <div class="phase-dashboard v4-phase-dashboard">${GATE_PHASES_V4.map(phase => renderPhaseDashboardCardV4(project, phase, currentPhase, activePhaseId)).join('')}</div>
      <div class="phase-tabs v4-phase-tabs">${GATE_PHASES_V4.map(phase => `<button class="${phase.id === activePhaseId ? 'active' : ''}" onclick="setProjectPhaseTab('${escapeAttribute(project.id)}', '${phase.id}')">${escapeHtml(phase.label)}</button>`).join('')}</div>
      <section class="phase-panel v4-phase-panel"><div class="phase-panel-head"><div><h2>${escapeHtml(activePhase.label)}</h2><p>${escapeHtml(activePhase.subtitle)}</p></div>${renderVersionTabsV4(project, activePhaseId, activeVersion)}</div><div class="horizontal-gate-list v4-gate-list">${activePhase.gateKeys.map(key => renderGateCardV4(project, key, activeVersion)).join('')}</div></section>
    </section>`;
};

setProjectPhaseTab = function(projectId, phaseId) {
  gateCenterV2State[`v4:${projectId}:phase`] = phaseId;
  renderProjectGateCenter(projectId);
};

setProjectPhaseVersion = function(projectId, phaseId, version) {
  gateCenterV2State[`v4:${projectId}:${phaseId}`] = version;
  renderProjectGateCenter(projectId);
};

addProjectPhaseVersion = function(projectId, phaseId) {
  const project = db.projects.find(p => p.id === projectId);
  if (!project) return;
  ensureGateCenterV4(project);
  const versions = project.phaseVersionsV4[phaseId];
  const latest = versions[versions.length - 1] || 'V1';
  const defaultName = nextVersionNameV4(latest, phaseId, versions.length);
  const name = prompt('請輸入新版本名稱（例如 V15、V1.2）', defaultName);
  if (!name) return;
  const reason = prompt('建立原因（例如：Gate 9 NG，退回 Gate 6）', '版本調整 / NG 後新增');
  if (!versions.includes(name)) versions.push(name);
  project.gateRecordsV4[phaseId][name] = {};
  GATE_PHASES_V4.find(p => p.id === phaseId).gateKeys.forEach(gateKey => {
    project.gateRecordsV4[phaseId][name][gateKey] = { gateKey, version: name, status: '待開始', owner: getOwnerByGateV4(project, gateKey), dueDate: '', handoffDate: '', trialDate: '', confirmedBy: '', confirmedAt: '', imageLink: '', imageNote: '', ng: null, createReason: reason || '' };
  });
  gateCenterV2State[`v4:${projectId}:${phaseId}`] = name;
  saveDatabase();
  renderProjectGateCenter(projectId);
};

renderActionItems = function() {
  const container = document.getElementById('tab-action-items');
  if (!container) return;
  currentProjectDetailId = null;
  const view = gateCenterV2State.progressView || 'gantt';
  container.innerHTML = `<section class="simple-page v4-page progress-page-v2"><div class="page-heading v4-heading"><p class="eyebrow">Progress</p><h1>專案進度</h1><p>開會用：左邊看規格/版本/目前進度，右邊用月曆看這個月每天在做什麼。</p></div><div class="phase-tabs progress-tabs v4-phase-tabs"><button class="${view === 'gantt' ? 'active' : ''}" onclick="setProgressViewV2('gantt')">甘特圖（月曆）</button><button class="${view === 'milestone' ? 'active' : ''}" onclick="setProgressViewV2('milestone')">里程碑</button></div>${view === 'milestone' ? renderMilestoneProgressV4() : renderCalendarGanttV4()}</section>`;
};

renderRisksIssues = function() {
  const container = document.getElementById('tab-risks-issues');
  if (!container) return;
  currentProjectDetailId = null;
  container.innerHTML = `<section class="simple-page v4-page team-page-v2"><div class="page-heading v4-heading"><p class="eyebrow">Team</p><h1>團隊成員</h1><p>依部門看主管與同仁；權限不放在畫面上干擾使用者，由系統背景判斷。</p></div><div class="team-dept-grid-v2 v4-team-grid">${TEAM_STRUCTURE_V4.map(team => `<article class="team-dept-card-v2 v4-team-card"><div class="team-dept-head"><div><h2>${escapeHtml(team.dept)}</h2><p>主管：${escapeHtml(team.manager)} ${escapeHtml(team.title)} · ${escapeHtml(team.extension)}</p></div></div><h3>同仁</h3><div class="member-list-v4">${team.members.map(member => `<div class="member-row-v4"><strong>${escapeHtml(member.name)}</strong><span>${escapeHtml(member.title)}</span><em>${escapeHtml(member.extension)}</em></div>`).join('')}</div></article>`).join('')}</div></section>`;
};

exportProjectMeetingExcel = function(projectId) {
  const project = db.projects.find(p => p.id === projectId);
  if (!project) return;
  ensureGateCenterV4(project);
  const display = getProjectDisplay(project);
  const rows = [];
  GATE_PHASES_V4.forEach(phase => {
    getPhaseVersionsV4(project, phase.id).forEach(version => {
      phase.gateKeys.forEach(gateKey => {
        const def = GATE_DEFS_V4[gateKey];
        const rec = getGateRecordV4(project, gateKey, version);
        rows.push({ '專案名稱': display.name, '規格': display.spec, '大項目': phase.label, '版本': version, 'Gate': gateKey, 'Gate 名稱': def.name, '狀態': getGateStateV4(project, gateKey, version), '負責部門': def.department, '負責人': rec.owner, '預計完成時間': rec.dueDate || '', '移交打樣時間': rec.handoffDate || '', '預計試戴時間': rec.trialDate || '', '確認人': rec.confirmedBy || '', '確認時間': rec.confirmedAt || '', 'NG 時間': rec.ng?.time || '', 'NG 原因': rec.ng?.reason || '', 'NG 備註': rec.ng?.note || '', '下一步': rec.ng?.nextAction || display.nextAction || '', '圖片連結': rec.imageLink || '', '圖片備註': rec.imageNote || '', 'PM 追蹤': display.tracker });
      });
    });
  });
  const fileName = `${project.id}_會議進度報告_${new Date().toISOString().slice(0, 10)}.xlsx`;
  if (typeof XLSX !== 'undefined') {
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '會議進度');
    XLSX.writeFile(wb, fileName);
  } else {
    exportRowsAsCsvV2(rows, fileName.replace('.xlsx', '.csv'));
  }
};
