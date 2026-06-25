// ==========================================================================
// 鏡片產品開發多專案追蹤系統 - 前端核心業務邏輯
// ==========================================================================

let db = null;
let currentUser = null;
let activePerspective = 'VP'; // 預設副總/主管視角
let isServerMode = false;
let currentUploadedData = null; // 暫存匯入的解析資料
// currentProjectDetailId 已由舊版專案明細功能宣告；新版 Gate 中心沿用同一個狀態。
let currentStageSelection = null;

const PRODUCT_PROCESS_STAGES = [
  { id: 1, phase: "需求", name: "需求與規格確認", shortName: "規格確認", owner: "PM", description: "確認產品需求、含水量、厚度、材質、BC、DIA 與目標市場。" },
  { id: 2, phase: "平光段", name: "模仁設計（平光／CNC 模仁）", shortName: "平光模仁設計", owner: "RD／模具", description: "建立平光模仁設計與 CNC 加工條件；測試 NG 時回到此階段。" },
  { id: 3, phase: "平光段", name: "PP 模", shortName: "平光 PP 模", owner: "模具／製程", description: "製作 PP 模並確認成形條件與基本尺寸。" },
  { id: 4, phase: "平光段", name: "設計確認", shortName: "平光設計確認", owner: "RD／品保", description: "快固片（光固）、移印與構型確認。" },
  { id: 5, phase: "平光段", name: "臨床：試戴片＋測試（RD）", shortName: "平光試戴測試", owner: "RD", description: "製作試戴片、執行測試並形成平光測試判定。" },
  { id: 6, phase: "全光度段", name: "模仁設計（全光度／CNC 模仁）", shortName: "全光度模仁設計", owner: "RD／模具", description: "依平光 OK 基準調整焦度設計；全光度測試 NG 時回到此階段。" },
  { id: 7, phase: "全光度段", name: "PP 模", shortName: "全光度 PP 模", owner: "模具／製程", description: "製作全光度 PP 模並確認成形與焦度條件。" },
  { id: 8, phase: "全光度段", name: "設計確認", shortName: "全光度設計確認", owner: "RD／品保", description: "快固片（光固）、移印、焦度與構型確認。" },
  { id: 9, phase: "全光度段", name: "臨床：試戴片＋測試（RD）", shortName: "全光度試戴測試", owner: "RD", description: "執行全光度試戴與測試，形成量產前測試判定。" },
  { id: 10, phase: "量產", name: "量產", shortName: "量產", owner: "生產／品保／PM", description: "完成量產條件確認、良率追蹤與正式放行。" }
];

// 甘特圖週別配置 (對應 Excel E-AM 欄位)
const GANTT_WEEKS = [
  { col: "E", month: "25/11", day: "3" }, { col: "F", month: "", day: "10" }, { col: "G", month: "", day: "17" }, { col: "H", month: "", day: "24" },
  { col: "I", month: "25/12", day: "1" }, { col: "J", month: "", day: "8" }, { col: "K", month: "", day: "15" }, { col: "L", month: "", day: "22" }, { col: "M", month: "", day: "29" },
  { col: "N", month: "26/1", day: "5" }, { col: "O", month: "", day: "12" }, { col: "P", month: "", day: "19" }, { col: "Q", month: "", day: "26" },
  { col: "R", month: "26/2", day: "2" }, { col: "S", month: "", day: "9" }, { col: "T", month: "", day: "16" }, { col: "U", month: "", day: "23" },
  { col: "V", month: "26/3", day: "2" }, { col: "W", month: "", day: "9" }, { col: "X", month: "", day: "16" }, { col: "Y", month: "", day: "23" }, { col: "Z", month: "", day: "30" },
  { col: "AA", month: "26/4", day: "6" }, { col: "AB", month: "", day: "13" }, { col: "AC", month: "", day: "20" }, { col: "AD", month: "", day: "27" },
  { col: "AE", month: "26/5", day: "4" }, { col: "AF", month: "", day: "11" }, { col: "AG", month: "", day: "18" }, { col: "AH", month: "", day: "25" },
  { col: "AI", month: "26/6", day: "1" }, { col: "AJ", month: "", day: "8" }, { col: "AK", month: "", day: "15" }, { col: "AL", month: "", day: "22" }, { col: "AM", month: "", day: "29" }
];

// 預設甘特圖儲存格填色 (模擬 Excel 甘特圖樣式)
const DEFAULT_GANTT_CELLS = {
  "SIH-145-CONV": {
    "V9.1": { cols: ["AI", "AJ", "AK", "AL", "AM"], color: "green", desc: "V9.1 已安排量產中" },
    "V8": { cols: ["AE", "AF", "AG", "AH"], color: "blue", desc: "V8 等待全光度與臨床試戴" },
    "v7 / v7.1（平光）": { cols: ["E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD"], color: "green", desc: "v7.1 平光試戴OK 可量產" }
  },
  "SIH-145-SIH": {
    "V8": { cols: ["V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM"], color: "orange", desc: "V8 平光完成，全光度開發中" },
    "V7": { cols: ["E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U"], color: "red", desc: "V7 彩片油墨誤用 NG 需重做" }
  },
  "SIH-142-SIH": {
    "V18.2": { cols: ["E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD"], color: "orange", desc: "V18.2 舒適度提升，待全光度臨床" },
    "V14": { cols: ["AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM"], color: "red", desc: "V14 版本混亂且有嚴重漏白" }
  },
  "SIH-142-UT": {
    "V14": { cols: ["AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM"], color: "red", desc: "V14 版本重疊，量產適用性確認中" }
  },
  "SIH-142-CONV": {
    "V4": { cols: ["E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM"], color: "orange", desc: "V4 全光度開發與試戴排程確認" }
  }
};

// ==========================================================================
// 1. 初始化與資料處理
// ==========================================================================

document.addEventListener('DOMContentLoaded', () => {
  initAppConnection();
});

async function initAppConnection() {
  try {
    // 試圖向本機 API 伺服器發送 GET 請求
    const response = await fetch('/api/data');
    if (response.ok) {
      db = await response.json();
      isServerMode = true;
      document.getElementById('server-status-dot').className = 'status-indicator online';
      document.getElementById('server-status-text').innerText = '伺服器連線中 (API 儲存)';
    } else {
      throw new Error("Server response not OK");
    }
  } catch (e) {
    // 連線失敗，切換至 LocalStorage 模式
    console.warn("無法連接本機 API 伺服器，切換至 LocalStorage 模式。原因:", e.message);
    isServerMode = false;
    document.getElementById('server-status-dot').className = 'status-indicator offline';
    document.getElementById('server-status-text').innerText = '本機模式 (LocalStorage)';
    
    // 讀取 LocalStorage 或使用 data.js 的預設資料
    const localData = localStorage.getItem('LENS_PDM_DB');
    if (localData) {
      db = JSON.parse(localData);
    } else {
      db = window.INITIAL_DATA;
      localStorage.setItem('LENS_PDM_DB', JSON.stringify(db));
    }
  }
  
  db.projects.forEach(ensureProjectProcessModel);
  checkSession();
}

function checkSession() {
  const sessionUser = sessionStorage.getItem('LENS_USER');
  if (sessionUser) {
    currentUser = JSON.parse(sessionUser);
    const storedPerspective = sessionStorage.getItem('LENS_PERSPECTIVE');
    activePerspective = currentUser.role === 'Admin' && ['VP', 'PM', 'RD', 'Admin'].includes(storedPerspective)
      ? storedPerspective
      : currentUser.role;
    
    // 設定下拉選單的值
    document.getElementById('role-perspective-select').value = activePerspective;
    
    showApp();
    initApp();
  } else {
    showLogin();
  }
}

function showLogin() {
  document.getElementById('login-overlay').classList.remove('hidden');
  document.getElementById('app-container').classList.add('hidden');
  document.getElementById('login-error').style.display = 'none';
  document.getElementById('username').value = '';
  document.getElementById('password').value = '';
}

function showApp() {
  document.getElementById('login-overlay').classList.add('hidden');
  document.getElementById('app-container').classList.remove('hidden');
}

function handleLogin() {
  const userVal = document.getElementById('username').value.trim();
  const passVal = document.getElementById('password').value;
  
  const foundUser = db.users.find(u => u.username === userVal && u.password === passVal);
  if (foundUser) {
    currentUser = foundUser;
    activePerspective = foundUser.role; // 登入者預設其所屬角色
    
    sessionStorage.setItem('LENS_USER', JSON.stringify(currentUser));
    sessionStorage.setItem('LENS_PERSPECTIVE', activePerspective);
    document.getElementById('role-perspective-select').value = activePerspective;
    
    document.getElementById('login-error').style.display = 'none';
    showApp();
    initApp();
  } else {
    document.getElementById('login-error').style.display = 'block';
  }
}

function handleLogout() {
  currentUser = null;
  sessionStorage.removeItem('LENS_USER');
  sessionStorage.removeItem('LENS_PERSPECTIVE');
  showLogin();
}

// 變更協作視角 (PM / RD / VP / Admin)
function changePerspective(role) {
  const allowedRoles = currentUser && currentUser.role === 'Admin'
    ? ['VP', 'PM', 'RD', 'Admin']
    : [currentUser ? currentUser.role : 'VP'];
  if (!allowedRoles.includes(role)) {
    activePerspective = currentUser ? currentUser.role : 'VP';
    document.getElementById('role-perspective-select').value = activePerspective;
    return;
  }
  activePerspective = role;
  sessionStorage.setItem('LENS_PERSPECTIVE', role);
  
  // 更新 UI 權限限制
  applyPermissions();
  
  // 重新渲染目前所在的分頁
  const activeTabItem = document.querySelector('.menu-item.active');
  if (activeTabItem) {
    const tabId = activeTabItem.getAttribute('data-tab');
    switchTab(tabId);
  }
}

function applyPermissions() {
  const perspectiveSelect = document.getElementById('role-perspective-select');
  const isAdminAccount = currentUser && currentUser.role === 'Admin';
  perspectiveSelect.disabled = !isAdminAccount;
  [...perspectiveSelect.options].forEach(option => {
    option.hidden = !isAdminAccount && option.value !== currentUser.role;
  });

  // 1. 管理員菜單選項顯示/隱藏
  const navAdmin = document.getElementById('nav-admin-settings');
  if (activePerspective === 'Admin') {
    navAdmin.classList.remove('hidden');
  } else {
    navAdmin.classList.add('hidden');
  }
  
  // 2. 新增按鈕權限控制
  const btnAddProj = document.getElementById('btn-add-project-main');
  const btnAddVer = document.getElementById('btn-add-version-main');
  const btnImport = document.getElementById('btn-import-report-main');
  const btnAddAct = document.getElementById('btn-add-action-main');
  const btnAddRisk = document.getElementById('btn-add-risk-main');
  
  // 新增專案 (只有 Admin 和 PM 可以)
  if (activePerspective === 'Admin' || activePerspective === 'PM') {
    btnAddProj.classList.remove('hidden');
  } else {
    btnAddProj.classList.add('hidden');
  }
  
  // 手動登錄版本 / 匯入試戴報告 (只有 RD 和 Admin 可以)
  if (activePerspective === 'Admin' || activePerspective === 'RD') {
    btnAddVer.classList.remove('hidden');
    btnImport.classList.remove('hidden');
  } else {
    btnAddVer.classList.add('hidden');
    btnImport.classList.add('hidden');
  }
  
  // 新增行動項目 (Admin, PM)
  if (activePerspective === 'Admin' || activePerspective === 'PM') {
    btnAddAct.classList.remove('hidden');
  } else {
    btnAddAct.classList.add('hidden');
  }
  
  // 新增風險 (Admin, PM)
  if (activePerspective === 'Admin' || activePerspective === 'PM') {
    btnAddRisk.classList.remove('hidden');
  } else {
    btnAddRisk.classList.add('hidden');
  }

  const detailEdit = document.getElementById('btn-detail-edit');
  const detailVersion = document.getElementById('btn-detail-version');
  if (detailEdit) detailEdit.classList.toggle('hidden', !(activePerspective === 'Admin' || activePerspective === 'PM'));
  if (detailVersion) detailVersion.classList.toggle('hidden', !(activePerspective === 'Admin' || activePerspective === 'RD'));
}

function initApp() {
  // 設定使用者資訊顯示
  document.getElementById('user-display-name').innerText = currentUser.name;
  document.getElementById('user-display-dept').innerText = currentUser.dept;
  
  // 初始化導覽切換事件
  const menuItems = document.querySelectorAll('.menu-item');
  menuItems.forEach(item => {
    item.onclick = (e) => {
      e.preventDefault();
      menuItems.forEach(m => m.classList.remove('active'));
      item.classList.add('active');
      const tabId = item.getAttribute('data-tab');
      switchTab(tabId);
    };
  });
  
  applyPermissions();
  
  // 預設進入總覽
  switchTab('vp-milestones');
}

function switchTab(tabId) {
  // 隱藏所有分頁
  document.querySelectorAll('.tab-pane').forEach(pane => pane.classList.remove('active'));
  
  // 顯示選中分頁
  const targetPane = document.getElementById(`tab-${tabId}`);
  if (targetPane) targetPane.classList.add('active');
  
  // 觸發各自渲染邏輯
  if (tabId === 'vp-milestones') renderVPMilestones();
  else if (tabId === 'gantt-chart') renderGanttChart();
  else if (tabId === 'project-master') {
    if (currentProjectDetailId) {
      openProjectDetail(currentProjectDetailId, false);
    } else {
      renderProjectMaster();
    }
  }
  else if (tabId === 'version-history') renderVersionHistory();
  else if (tabId === 'action-items') renderActionItems();
  else if (tabId === 'risks-issues') renderRisksIssues();
  else if (tabId === 'admin-settings') renderAdminSettings();
}

async function saveDatabase() {
  if (isServerMode) {
    try {
      const response = await fetch('/api/data', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(db)
      });
      if (!response.ok) throw new Error("Server write failed");
    } catch (e) {
      console.error("寫入伺服器失敗，暫存至本機 LocalStorage:", e.message);
      localStorage.setItem('LENS_PDM_DB', JSON.stringify(db));
    }
  } else {
    localStorage.setItem('LENS_PDM_DB', JSON.stringify(db));
  }
}

// 匯出 JSON 資料庫
function exportDataJSON() {
  const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(db, null, 2));
  const downloadAnchor = document.createElement('a');
  downloadAnchor.setAttribute("href", dataStr);
  downloadAnchor.setAttribute("download", `lens_pdm_database_${new Date().toISOString().slice(0, 10)}.json`);
  document.body.appendChild(downloadAnchor);
  downloadAnchor.click();
  downloadAnchor.remove();
}

// ==========================================================================
// 2. 分頁 1: 副總里程碑總表
// ==========================================================================

function renderVPMilestones() {
  // 1. KPI 卡片計算
  document.getElementById('kpi-total-projects').innerText = db.projects.length;
  document.getElementById('kpi-red-projects').innerText = db.projects.filter(p => p.status === '紅').length;
  document.getElementById('kpi-yellow-projects').innerText = db.projects.filter(p => p.status === '黃').length;
  
  const pendingActions = db.actions.filter(a => a.status !== '已完成');
  document.getElementById('kpi-pending-actions').innerText = pendingActions.length;
  document.getElementById('kpi-high-actions').innerText = pendingActions.filter(a => a.priority === '高').length;
  document.getElementById('kpi-open-risks').innerText = db.risks.filter(r => r.status !== '已關閉').length;
  
  renderWorkbenchHero();
  renderRoleTaskCards();
  renderPriorityProjectCards();
  renderDepartmentChecklist();

  // 2. 專案快速進度表格渲染
  const tbody = document.getElementById('milestone-table-body');
  tbody.innerHTML = '';
  
  getSortedProjectsForWorkbench().forEach(proj => {
    const tr = document.createElement('tr');
    ensureProjectProcessModel(proj);
    tr.className = 'project-row-clickable';
    tr.tabIndex = 0;
    tr.onclick = () => {
      currentProjectDetailId = proj.id;
      document.querySelectorAll('.menu-item').forEach(item => item.classList.toggle('active', item.getAttribute('data-tab') === 'project-master'));
      switchTab('project-master');
    };
    
    // 燈號 Class 判定
    let lightClass = 'badge-gray';
    if (proj.status === '綠') lightClass = 'badge-green';
    else if (proj.status === '黃') lightClass = 'badge-orange';
    else if (proj.status === '紅') lightClass = 'badge-red';
    
    const stage = PRODUCT_PROCESS_STAGES[proj.process.currentStep - 1];
    const guidance = getProjectGuidance(proj);
    const assignment = getStageAssignment(proj, proj.process.currentStep);
    
    tr.innerHTML = `
      <td><strong>${proj.name}</strong><br><span style="font-size: 11px; color: var(--text-muted);">${proj.id}</span></td>
      <td><span class="badge ${lightClass}">${proj.status}燈</span></td>
      <td>
        <span class="badge badge-blue">第 ${proj.process.currentStep} 階段</span>
        <div class="workbench-table-stage">${escapeHtml(stage.shortName)}</div>
      </td>
      <td>
        <div class="workbench-next-line ${guidance.tone}">
          <strong>${escapeHtml(guidance.title)}</strong>
          <span>${escapeHtml(guidance.nextText)}</span>
          <em>${escapeHtml(assignment.department)}｜${escapeHtml(assignment.person)}　PM追蹤：${escapeHtml(assignment.tracker)}</em>
        </div>
      </td>
    `;
    
    tbody.appendChild(tr);
  });
}

function switchTabFromWorkbench(tabId) {
  document.querySelectorAll('.menu-item').forEach(item => {
    item.classList.toggle('active', item.getAttribute('data-tab') === tabId);
  });
  switchTab(tabId);
}

function renderWorkbenchHero() {
  const roleLabel = {
    PM: 'PM 今日工作台',
    RD: 'RD 今日工作台',
    VP: '主管今日工作台',
    Admin: '管理員總覽'
  }[activePerspective] || '今日工作台';
  const title = {
    PM: `${currentUser.name}，先看紅燈與卡關專案`,
    RD: `${currentUser.name}，先處理測試與版本判定`,
    VP: `${currentUser.name}，先看需要拍板的專案`,
    Admin: '系統管理員，先看全系統狀態'
  }[activePerspective] || '請先處理需要你確認的專案';
  const subtitle = {
    PM: '卡片會把每個專案的下一步、責任人與期限放在前面，不需要再自己去表格裡找。',
    RD: '卡片會優先列出第 5 / 第 9 階段測試、NG 原因與需要登錄的版本紀錄。',
    VP: '卡片會優先顯示紅燈、待核決與會影響上市時程的專案。',
    Admin: '卡片會彙整所有角色工作與異常，方便檢查權限和資料完整性。'
  }[activePerspective] || '系統會依角色整理任務卡，點卡片即可進入正確專案或功能。';

  document.getElementById('workbench-role-label').innerText = roleLabel;
  document.getElementById('workbench-title').innerText = title;
  document.getElementById('workbench-subtitle').innerText = subtitle;
}

function renderRoleTaskCards() {
  const container = document.getElementById('role-task-cards');
  const cards = getRoleTaskCards();
  container.innerHTML = cards.map(card => `
    <article class="role-task-card ${card.tone}">
      <div class="role-task-topline">
        <span>${escapeHtml(card.label)}</span>
        <strong>${escapeHtml(card.count)}</strong>
      </div>
      <h3>${escapeHtml(card.title)}</h3>
      <p>${escapeHtml(card.desc)}</p>
      <button class="btn ${card.primary ? 'btn-primary' : 'btn-outline'} btn-sm" onclick="${card.action}">
        ${escapeHtml(card.cta)}
      </button>
    </article>
  `).join('');
}

function getRoleTaskCards() {
  const redProjects = db.projects.filter(p => p.status === '紅');
  const highActions = db.actions.filter(a => a.priority === '高' && a.status !== '已完成');
  const myActions = getMyActions();
  const testProjects = db.projects.filter(p => {
    ensureProjectProcessModel(p);
    return [5, 9].includes(p.process.currentStep);
  });
  const approvalProjects = db.projects.filter(p => ["G2", "G3", "G4", "G5"].includes(p.gate));

  if (activePerspective === 'RD') {
    return [
      {
        label: '你要處理',
        count: `${myActions.length}`,
        title: '我的 RD 任務',
        desc: myActions.length ? '先處理到期與高優先任務，完成後專案才會往下一階段走。' : '目前沒有直接指派給你的 RD 任務。',
        cta: '看行動項目',
        action: "switchTabFromWorkbench('action-items')",
        tone: myActions.length ? 'warning' : 'calm',
        primary: myActions.length > 0
      },
      {
        label: '測試判定',
        count: `${testProjects.length}`,
        title: '第 5 / 第 9 階段測試',
        desc: '需要 RD 填寫 OK/NG 與版本紀錄；NG 會由系統指引回第 2 或第 6 階段。',
        cta: '登錄版本／測試',
        action: "switchTabFromWorkbench('version-history')",
        tone: 'info',
        primary: true
      },
      {
        label: '紅燈',
        count: `${redProjects.length}`,
        title: '需要技術釐清的卡點',
        desc: '先看 NG 原因、模仁/試戴/構型問題，再回填修正建議。',
        cta: '看卡關專案',
        action: "switchTabFromWorkbench('project-master')",
        tone: redProjects.length ? 'danger' : 'calm',
        primary: redProjects.length > 0
      }
    ];
  }

  if (activePerspective === 'VP') {
    return [
      {
        label: '需決策',
        count: `${approvalProjects.length}`,
        title: '待核決專案',
        desc: '看最新版本、風險與下一步，確認是否放行或退回。',
        cta: '看待核決',
        action: "switchTabFromWorkbench('project-master')",
        tone: approvalProjects.length ? 'warning' : 'calm',
        primary: approvalProjects.length > 0
      },
      {
        label: '紅燈',
        count: `${redProjects.length}`,
        title: '影響時程的異常',
        desc: '優先看紅燈專案是否需要資源、決策或跨部門升級。',
        cta: '看紅燈專案',
        action: "switchTabFromWorkbench('project-master')",
        tone: redProjects.length ? 'danger' : 'calm',
        primary: redProjects.length > 0
      },
      {
        label: '高優先',
        count: `${highActions.length}`,
        title: '高優先未完任務',
        desc: '確認高優先項目是否逾期，必要時要求 Owner 補回覆。',
        cta: '看行動項目',
        action: "switchTabFromWorkbench('action-items')",
        tone: highActions.length ? 'warning' : 'calm',
        primary: false
      }
    ];
  }

  return [
    {
      label: 'PM 追蹤',
      count: `${redProjects.length}`,
      title: '紅燈與卡關專案',
      desc: '先處理 NG、逾期、版本身分不清，以及需要跨部門確認的專案。',
      cta: '處理卡關專案',
      action: "switchTabFromWorkbench('project-master')",
      tone: redProjects.length ? 'danger' : 'calm',
      primary: redProjects.length > 0
    },
    {
      label: '任務',
      count: `${highActions.length}`,
      title: '高優先行動項目',
      desc: '確認 Owner、期限與完成條件，避免會議決議沒有落地。',
      cta: '看行動項目',
      action: "switchTabFromWorkbench('action-items')",
      tone: highActions.length ? 'warning' : 'calm',
      primary: false
    },
    {
      label: '排程',
      count: `${db.projects.length}`,
      title: '甘特圖與專案排程',
      desc: '看每個版本是否撞期、延誤或缺下一階段安排。',
      cta: '打開甘特圖',
      action: "switchTabFromWorkbench('gantt-chart')",
      tone: 'info',
      primary: false
    }
  ];
}

function renderPriorityProjectCards() {
  const container = document.getElementById('priority-project-cards');
  const projects = getSortedProjectsForWorkbench().slice(0, 5);
  if (!projects.length) {
    container.innerHTML = '<p class="empty-stage-record">目前沒有需要優先處理的專案。</p>';
    return;
  }

  container.innerHTML = projects.map(proj => {
    ensureProjectProcessModel(proj);
    const stage = PRODUCT_PROCESS_STAGES[proj.process.currentStep - 1];
    const guidance = getProjectGuidance(proj);
    const assignment = getStageAssignment(proj, proj.process.currentStep);
    return `
      <article class="priority-project-card ${guidance.tone}">
        <div class="priority-card-head">
          <div>
            <strong>${escapeHtml(proj.name)}</strong>
            <span>${escapeHtml(proj.id)} · ${escapeHtml(proj.currentVersion)}</span>
          </div>
          <span class="badge ${getLightClass(proj.status)}">${proj.status}燈</span>
        </div>
        <div class="priority-stage-line">
          <span>目前：第 ${proj.process.currentStep} 階段</span>
          <strong>${escapeHtml(stage.shortName)}</strong>
        </div>
        <div class="assignment-title">責任分工</div>
        ${renderAssignmentPills(assignment)}
        <p>${escapeHtml(guidance.nextText)}</p>
        <div class="priority-card-actions">
          <button class="btn btn-primary btn-sm" onclick="openProjectDetailFromWorkbench('${proj.id}')">${escapeHtml(guidance.cta)}</button>
          <button class="btn btn-outline btn-sm" onclick="openProjectDetailFromWorkbench('${proj.id}')">看完整流程</button>
        </div>
      </article>
    `;
  }).join('');
}

function renderDepartmentChecklist() {
  const container = document.getElementById('department-checklist');
  const checklist = getDepartmentChecklist(activePerspective);
  container.innerHTML = checklist.map((item, index) => `
    <article class="checklist-item ${item.required ? 'required' : ''}">
      <span class="checklist-index">${index + 1}</span>
      <div>
        <strong>${escapeHtml(item.title)}</strong>
        <p>${escapeHtml(item.desc)}</p>
      </div>
    </article>
  `).join('');
}

function getDepartmentChecklist(role) {
  const lists = {
    PM: [
      { title: '確認每個卡片都有下一步', desc: '紅燈專案必須填 Owner、期限、退回階段或決策需求。', required: true },
      { title: '把 NG 轉成任務', desc: '第 5 階段 NG 回第 2；第 9 階段 NG 回第 6，並建立下一輪修正任務。', required: true },
      { title: '檢查跨部門缺口', desc: 'RD、模仁、品保、生產、法規與供應鏈若有缺資料，要補行動項目。', required: false }
    ],
    RD: [
      { title: '填寫測試判定', desc: '第 5 / 第 9 階段要輸入 OK、NG、待確認與技術原因。', required: true },
      { title: '登錄版本紀錄', desc: '每次模仁、PP 模、試戴或修正都要留版本、差異、結論。', required: true },
      { title: '提出修正建議', desc: 'NG 時說明要改設計、模仁、焦度、構型或製程條件。', required: false }
    ],
    VP: [
      { title: '先看紅燈與待核決', desc: '確認是否要放行、退回或要求跨部門補資料。', required: true },
      { title: '看上市時程影響', desc: '若卡點會影響量產或上市，需決定資源與優先順序。', required: true },
      { title: '確認風險已有人負責', desc: '紅燈不得只有描述，必須有 Owner 與期限。', required: false }
    ],
    Admin: [
      { title: '確認角色權限', desc: '非 Admin 不可切換他人角色；各部門只看到可操作的按鈕。', required: true },
      { title: '檢查資料完整性', desc: '專案、版本、行動項目與風險資料需能互相關聯。', required: true },
      { title: '維護帳號', desc: '新增同仁帳號或調整部門角色。', required: false }
    ]
  };
  return lists[role] || lists.PM;
}

// ==========================================================================
// 16. Tail V4 activation：放在檔案最尾端，最後一次覆蓋舊版入口
// ==========================================================================

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

// ==========================================================================
// 15. Final V4 activation：確保舊版同名函式不會覆蓋 Excel 規格版
// ==========================================================================

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
  const myRole = ROLE_LABELS_V4[activePerspective] || activePerspective;
  const cards = priorityOne.concat(projects.filter(p => !priorityOne.includes(p)).slice(0, 2));
  container.innerHTML = `
    <section class="simple-page v4-page">
      <div class="page-heading v4-heading">
        <p class="eyebrow">My Work</p>
        <h1>我的工作</h1>
        <p>${escapeHtml(myRole)} 視角：先看你現在要處理哪一個專案、下一步要按哪裡。</p>
      </div>
      <div class="simple-stat-grid v4-stat-grid">
        <button class="simple-stat-card danger clickable-card" onclick="openProjectGateCenter('${escapeAttribute(ngProjects[0]?.id || priorityOne[0]?.id || projects[0]?.id || '')}')">
          <span>高優先處理</span><strong>${ngProjects.length}</strong><small>點擊進入對應專案</small>
        </button>
        <button class="simple-stat-card warning clickable-card" onclick="switchTab('action-items')">
          <span>今日需追蹤</span><strong>${priorityOne.length}</strong><small>查看月曆與里程碑</small>
        </button>
        <button class="simple-stat-card success clickable-card" onclick="switchTab('project-master')">
          <span>進行中專案</span><strong>${projects.length}</strong><small>進入專案總覽</small>
        </button>
      </div>
      <div class="section-title-row">
        <h2>現在要處理的專案</h2>
        <button class="btn btn-primary" onclick="switchTab('project-master')">查看全部專案</button>
      </div>
      <div class="work-card-list v4-work-list">
        ${cards.map(p => renderMyWorkCard(p)).join('')}
      </div>
    </section>
  `;
};

renderMyWorkCard = function(project) {
  const d = getProjectDisplay(project);
  const gate = GATE_DEFS_V4[String(d.gate)] || GATE_DEFS_V4['1'];
  return `
    <article class="work-card v4-work-card" onclick="openProjectGateCenter('${escapeAttribute(project.id)}')">
      <div>
        <span class="priority-rank">第 ${d.priority} 優先</span>
        <h3>${escapeHtml(d.name)}</h3>
        <p>目前：Gate ${d.gate} ${escapeHtml(gate.name)}</p>
      </div>
      <div class="work-card-meta">
        <span class="status-pill ${statusClassV4(d.status)}">${escapeHtml(d.status)}</span>
        <span>負責：${escapeHtml(d.owner)}</span>
        <span>下一步：${escapeHtml(d.nextAction || '等待下一個 Gate')}</span>
      </div>
    </article>
  `;
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
  container.innerHTML = `
    <section class="simple-page v4-page">
      <div class="page-heading v4-heading">
        <p class="eyebrow">Projects</p>
        <h1>專案總覽</h1>
        <p>依優先順序排列。點任一張專案卡片，就會進入該專案的 Gate 中心。</p>
      </div>
      ${groups.map(group => renderPriorityProjectSectionV4(group)).join('')}
    </section>
  `;
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
      <div class="page-action-row">
        <button class="back-button" onclick="currentProjectDetailId=null; renderProjectMaster();">← 回專案總覽</button>
        <button class="btn btn-outline" onclick="exportProjectMeetingExcel('${escapeAttribute(project.id)}')">匯出會議 Excel</button>
      </div>
      <div class="project-detail-heading v4-heading">
        <p class="eyebrow">Project Gate Center</p>
        <h1>${escapeHtml(d.name)}</h1>
      </div>
      <article class="gate-summary-card v4-summary ${statusClassV4(d.status)}">
        <div class="summary-main">
          <span class="status-pill ${statusClassV4(d.status)}">${escapeHtml(d.status)}</span>
          <h2>目前進度：Gate ${d.gate} ${escapeHtml(currentGate.name)}</h2>
          <p>${escapeHtml(d.ngReason || currentGate.description)}</p>
        </div>
        <div class="summary-grid">
          <div><span>目前大項目</span><strong>${escapeHtml(GATE_PHASES_V4.find(p => p.id === currentPhase)?.label || '')}</strong></div>
          <div><span>目前版本</span><strong>${escapeHtml(d.version)}</strong></div>
          <div><span>下一步</span><strong>${escapeHtml(d.nextAction || '等待下一個 Gate')}</strong></div>
          <div><span>負責 / PM追蹤</span><strong>${escapeHtml(d.owner)} / ${escapeHtml(d.tracker)}</strong></div>
        </div>
      </article>
      <div class="phase-dashboard v4-phase-dashboard">
        ${GATE_PHASES_V4.map(phase => renderPhaseDashboardCardV4(project, phase, currentPhase, activePhaseId)).join('')}
      </div>
      <div class="phase-tabs v4-phase-tabs">
        ${GATE_PHASES_V4.map(phase => `<button class="${phase.id === activePhaseId ? 'active' : ''}" onclick="setProjectPhaseTab('${escapeAttribute(project.id)}', '${phase.id}')">${escapeHtml(phase.label)}</button>`).join('')}
      </div>
      <section class="phase-panel v4-phase-panel">
        <div class="phase-panel-head">
          <div><h2>${escapeHtml(activePhase.label)}</h2><p>${escapeHtml(activePhase.subtitle)}</p></div>
          ${renderVersionTabsV4(project, activePhaseId, activeVersion)}
        </div>
        <div class="horizontal-gate-list v4-gate-list">
          ${activePhase.gateKeys.map(key => renderGateCardV4(project, key, activeVersion)).join('')}
        </div>
      </section>
    </section>
  `;
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
    project.gateRecordsV4[phaseId][name][gateKey] = {
      gateKey,
      version: name,
      status: '待開始',
      owner: getOwnerByGateV4(project, gateKey),
      dueDate: '',
      handoffDate: '',
      trialDate: '',
      confirmedBy: '',
      confirmedAt: '',
      imageLink: '',
      imageNote: '',
      ng: null,
      createReason: reason || ''
    };
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
  container.innerHTML = `
    <section class="simple-page v4-page progress-page-v2">
      <div class="page-heading v4-heading">
        <p class="eyebrow">Progress</p>
        <h1>專案進度</h1>
        <p>開會用：左邊看規格/版本/目前進度，右邊用月曆看這個月每天在做什麼。</p>
      </div>
      <div class="phase-tabs progress-tabs v4-phase-tabs">
        <button class="${view === 'gantt' ? 'active' : ''}" onclick="setProgressViewV2('gantt')">甘特圖（月曆）</button>
        <button class="${view === 'milestone' ? 'active' : ''}" onclick="setProgressViewV2('milestone')">里程碑</button>
      </div>
      ${view === 'milestone' ? renderMilestoneProgressV4() : renderCalendarGanttV4()}
    </section>
  `;
};

renderRisksIssues = function() {
  const container = document.getElementById('tab-risks-issues');
  if (!container) return;
  currentProjectDetailId = null;
  container.innerHTML = `
    <section class="simple-page v4-page team-page-v2">
      <div class="page-heading v4-heading">
        <p class="eyebrow">Team</p>
        <h1>團隊成員</h1>
        <p>依部門看主管與同仁；權限不放在畫面上干擾使用者，由系統背景判斷。</p>
      </div>
      <div class="team-dept-grid-v2 v4-team-grid">
        ${TEAM_STRUCTURE_V4.map(team => `
          <article class="team-dept-card-v2 v4-team-card">
            <div class="team-dept-head">
              <div><h2>${escapeHtml(team.dept)}</h2><p>主管：${escapeHtml(team.manager)} ${escapeHtml(team.title)} · ${escapeHtml(team.extension)}</p></div>
            </div>
            <h3>同仁</h3>
            <div class="member-list-v4">
              ${team.members.map(member => `<div class="member-row-v4"><strong>${escapeHtml(member.name)}</strong><span>${escapeHtml(member.title)}</span><em>${escapeHtml(member.extension)}</em></div>`).join('')}
            </div>
          </article>
        `).join('')}
      </div>
    </section>
  `;
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
        rows.push({
          '專案名稱': display.name,
          '規格': display.spec,
          '大項目': phase.label,
          '版本': version,
          'Gate': gateKey,
          'Gate 名稱': def.name,
          '狀態': getGateStateV4(project, gateKey, version),
          '負責部門': def.department,
          '負責人': rec.owner,
          '預計完成時間': rec.dueDate || '',
          '移交打樣時間': rec.handoffDate || '',
          '預計試戴時間': rec.trialDate || '',
          '確認人': rec.confirmedBy || '',
          '確認時間': rec.confirmedAt || '',
          'NG 時間': rec.ng?.time || '',
          'NG 原因': rec.ng?.reason || '',
          'NG 備註': rec.ng?.note || '',
          '下一步': rec.ng?.nextAction || display.nextAction || '',
          '圖片連結': rec.imageLink || '',
          '圖片備註': rec.imageNote || '',
          'PM 追蹤': display.tracker
        });
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

// ==========================================================================
// 14. Design Spec V4：依使用者 Excel 回填的設計架構重整主要畫面
// ==========================================================================

const PDM_V4_CACHE = 'design-spec-v4-20260625b';

const ROLE_LABELS_V4 = {
  VP: '高階主管',
  PM: 'PM',
  RD: 'RD 研發',
  CNC: 'CNC / 模仁',
  Supervisor: '部門主管',
  Employee: '一般同仁',
  Admin: '系統管理員'
};

const TEAM_MEMBERS_V4 = {
  PM: ['曼玉', '媃安', '宣祠'],
  RD: ['齊政', '世偉', '其他 RD'],
  CNC: ['鍾民', '其他模仁人員'],
  Sales: ['業務窗口'],
  Production: ['生產窗口', '品保窗口']
};

const TEAM_STRUCTURE_V4 = [
  { dept: 'RD 研發部', manager: 'Andy', title: '副總', extension: '分機待補', members: [
    { name: '齊政', title: 'RD', extension: '分機待補' },
    { name: '世偉', title: 'RD', extension: '分機待補' },
    { name: '其他 RD', title: '研發同仁', extension: '分機待補' }
  ]},
  { dept: 'CNC / 模仁', manager: '鍾民', title: '主管', extension: '分機待補', members: [
    { name: '鍾民', title: 'CNC / 模仁', extension: '分機待補' },
    { name: '其他模仁人員', title: 'CNC / 模仁', extension: '分機待補' }
  ]},
  { dept: '業務', manager: '待補', title: '主管', extension: '分機待補', members: [
    { name: '業務窗口', title: '業務窗口', extension: '分機待補' }
  ]},
  { dept: 'PM', manager: '曼玉', title: '經理', extension: '分機待補', members: [
    { name: '媃安', title: 'PM', extension: '分機待補' },
    { name: '宣祠', title: 'PM', extension: '分機待補' }
  ]}
];

const GATE_PHASES_V4 = [
  { id: 'demand', label: '1. 需求', subtitle: '規格、需求與責任窗口確認', gateKeys: ['1'] },
  { id: 'plano', label: '2. 平光', subtitle: '平光模仁、PP 模、設計確認與試戴判定', gateKeys: ['2', '3', '4', '5'] },
  { id: 'power', label: '3. 全光', subtitle: '全光度模仁、PP 模、設計確認與試戴判定', gateKeys: ['6', '7', '8', '9'] },
  { id: 'production', label: '4. 量產', subtitle: 'Gate 9 OK 後的量產準備與放行', gateKeys: ['P1', 'P2', 'P3'] }
];

const GATE_DEFS_V4 = {
  '1': { name: '需求與規格確認', department: '業務 / PM', confirmRole: 'PM', ownerGroup: 'PM', description: '業務提出需求，PM 確認含水量、厚度、DIA、BC、材質與目標市場。', checklist: ['需求來源確認', '產品規格確認', 'PM 建立專案'] },
  '2': { name: '平光模仁設計', department: 'CNC', confirmRole: 'CNC', ownerGroup: 'CNC', description: 'CNC / 模仁負責平光模仁設計與加工條件。', checklist: ['平光模仁設計完成', '加工條件確認', '預計完成時間已填'] },
  '3': { name: '平光 PP 模', department: 'CNC', confirmRole: 'CNC', ownerGroup: 'CNC', description: 'CNC 製作平光 PP 模，確認模具與成形條件。', checklist: ['PP 模完成', '基本尺寸確認', '可交付 RD'] },
  '4': { name: '平光設計確認', department: 'RD', confirmRole: 'RD', ownerGroup: 'RD', description: 'RD 以圖片與備註記錄設計確認，並填寫移交打樣時間。', checklist: ['圖片/連結已上傳', '設計差異已記錄', '移交打樣時間已填'] },
  '5': { name: '平光試戴測試 / 判定', department: 'RD', confirmRole: 'RD', ownerGroup: 'RD', description: 'RD 安排平光試戴與測試；OK 進全光，NG 退回 Gate 2 並新增下一版。', checklist: ['預計試戴時間已填', '測試結果已判定', 'NG 原因需完整'] },
  '6': { name: '全光度模仁設計', department: 'CNC', confirmRole: 'CNC', ownerGroup: 'CNC', description: 'CNC / 模仁依全光度需求設計模仁；Gate 9 NG 會回到此 Gate。', checklist: ['全光度模仁設計完成', '加工條件確認', '預計完成時間已填'] },
  '7': { name: '全光度 PP 模', department: 'CNC', confirmRole: 'CNC', ownerGroup: 'CNC', description: 'CNC 製作全光度 PP 模，確認焦度與成形條件。', checklist: ['PP 模完成', '焦度條件確認', '可交付 RD'] },
  '8': { name: '全光度設計確認', department: 'RD', confirmRole: 'RD', ownerGroup: 'RD', description: 'RD 確認全光度設計，必要時可附圖片或連結。', checklist: ['設計確認完成', '必要圖片/備註已記錄', '可安排試戴'] },
  '9': { name: '全光度試戴測試 / 判定', department: 'RD', confirmRole: 'RD', ownerGroup: 'RD', description: 'RD 執行全光度試戴與測試；OK 進量產準備，NG 退回 Gate 6 並新增下一版。', checklist: ['預計試戴時間已填', '測試結果已判定', 'NG 原因需完整'] },
  P1: { name: '量產準備', department: 'PM / 生產', confirmRole: 'PM', ownerGroup: 'PM', description: '彙整 Gate 9 OK 後的量產條件與資料。', checklist: ['量產資料整理', '版本與規格確認', '風險清單確認'] },
  P2: { name: '量產放行審核', department: '主管', confirmRole: 'Supervisor', ownerGroup: 'PM', description: '主管確認是否可進入量產放行。', checklist: ['主管審核完成', '量產條件可控', '會議結論已記錄'] },
  P3: { name: '量產追蹤', department: '生產 / 品保', confirmRole: 'Supervisor', ownerGroup: 'Production', description: '追蹤量產初期狀態、異常與回報。', checklist: ['首批狀態確認', '異常回報', 'PM 會議匯出'] }
};

const PROJECT_DISPLAY_V4 = {
  'SIH-142-UT': {
    name: 'SIH-142-UT · SiH 14.2 超薄',
    spec: '14.2（水膠）',
    version: 'V14',
    gate: 9,
    status: 'NG / 等待結果',
    priority: 1,
    owner: 'RD 齊政',
    tracker: '曼玉',
    ngReason: '全光度試戴測試有問題，等待 RD 測試結果',
    resultDue: '預計 2026/06/26 有結果',
    nextAction: '啟動 V15，退回 Gate 6 全光度模仁設計',
    startDate: '2026-06-04',
    planoDate: '2026-06-05',
    powerDate: '2026-06-26'
  },
  'SIH-145-CONV': {
    name: 'SIH-145-CONV · 14.5 Conventional',
    spec: '14.5（水膠）',
    version: 'V9.1',
    gate: 9,
    status: '進行中',
    priority: 1,
    owner: 'RD 世偉',
    tracker: '曼玉',
    nextAction: '全光度試戴完成後進入量產準備',
    startDate: '2026-05-28',
    planoDate: '2026-06-29',
    powerDate: '2026-06-30'
  },
  'SIH-145-SIH': {
    name: 'SIH-145-SIH · SiH 14.5',
    spec: '14.5（矽水膠）',
    version: 'V8',
    gate: 6,
    status: '進行中',
    priority: 2,
    owner: 'RD 世偉',
    tracker: '曼玉',
    nextAction: 'CNC 進行全光度模仁設計',
    startDate: '2026-06-02',
    planoDate: '2026-06-16',
    powerDate: '2026-06-25'
  },
  'SIH-142-SIH': {
    name: 'SIH-142-SIH · SiH 14.2',
    spec: '14.2（矽水膠）',
    version: 'V18.2',
    gate: 9,
    status: '進行中',
    priority: 2,
    owner: 'RD 齊政',
    tracker: '曼玉',
    nextAction: '等待全光度試戴判定',
    startDate: '2026-06-08',
    planoDate: '2026-06-16',
    powerDate: '2026-06-30'
  },
  'SIH-142-CONV': {
    name: 'SIH-142-CONV · 14.2 Conventional',
    spec: '14.2（水膠）',
    version: 'V4',
    gate: 8,
    status: '進行中',
    priority: 3,
    owner: 'RD 齊政',
    tracker: '曼玉',
    nextAction: '確認全光度設計後安排試戴',
    startDate: '2026-06-05',
    planoDate: '2026-06-09',
    powerDate: '2026-06-26'
  }
};

const PROJECT_EVENTS_V4 = [
  { projectId: 'SIH-142-UT', date: '2026-06-26', label: 'V14 全光試戴 NG', type: 'ng', phase: 'power', version: 'V14' },
  { projectId: 'SIH-142-UT', date: '2026-06-30', label: 'V15 退回 Gate 6', type: 'plan', phase: 'power', version: 'V15' },
  { projectId: 'SIH-145-CONV', date: '2026-06-18', label: 'V25 全光載入', type: 'power', phase: 'power', version: 'V9.1' },
  { projectId: 'SIH-145-CONV', date: '2026-06-26', label: 'V9.1 平光 OK', type: 'ok', phase: 'plano', version: 'V9.1' },
  { projectId: 'SIH-145-SIH', date: '2026-06-16', label: 'V8 平光試戴', type: 'plano', phase: 'plano', version: 'V8' },
  { projectId: 'SIH-145-SIH', date: '2026-06-25', label: 'V8 全光模仁', type: 'power', phase: 'power', version: 'V8' },
  { projectId: 'SIH-142-SIH', date: '2026-06-23', label: 'V18.2 全光確認', type: 'power', phase: 'power', version: 'V18.2' },
  { projectId: 'SIH-142-SIH', date: '2026-06-25', label: 'V18.2 測試安排', type: 'trial', phase: 'power', version: 'V18.2' },
  { projectId: 'SIH-142-CONV', date: '2026-06-26', label: 'V4 全光確認', type: 'power', phase: 'power', version: 'V4' }
];

function textOr(value, fallback) {
  const text = String(value || '');
  return /�|\?{2,}|蝝|銝|撠|璅|摰|頛/.test(text) ? fallback : text;
}

function getAllProjectsForCleanUI() {
  return (db.projects || []).filter(Boolean);
}

function getProjectDisplay(project) {
  const custom = PROJECT_DISPLAY_V4[project.id] || {};
  const inferredGate = inferGateFromProjectV4(project);
  return {
    id: project.id,
    name: custom.name || `${project.id} · ${textOr(project.name, project.type || '產品專案')}`,
    spec: custom.spec || textOr(project.name, project.type || '待確認規格'),
    version: custom.version || textOr(project.currentVersion, 'V1'),
    gate: Number(custom.gate || inferredGate),
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
}

function inferGateFromProjectV4(project) {
  const gateText = String(project.gate || '');
  const match = gateText.match(/\d+/);
  if (match) {
    const n = Number(match[0]);
    if (n >= 5) return Math.min(9, n + 4);
    if (n >= 1) return Math.min(9, n + 1);
  }
  return 1;
}

function normalizeStatusV4(status) {
  const text = String(status || '');
  if (text.includes('NG') || text.includes('蝝')) return 'NG / 等待結果';
  if (text.includes('完成') || text.includes('OK')) return '已完成';
  if (text.includes('暫') || text.includes('停')) return '暫停';
  return '進行中';
}

function statusClassV4(status) {
  const text = String(status || '');
  if (text.includes('NG') || text.includes('逾期')) return 'danger';
  if (text.includes('完成') || text.includes('OK')) return 'success';
  if (text.includes('等待') || text.includes('暫')) return 'warning';
  if (text.includes('待開始')) return 'neutral';
  return 'info';
}

function getPhaseByGateV4(gateKey) {
  const key = String(gateKey);
  return GATE_PHASES_V4.find(phase => phase.gateKeys.includes(key)) || GATE_PHASES_V4[0];
}

function getCurrentPhaseV4(project) {
  const gate = getProjectDisplay(project).gate;
  if (gate >= 10) return 'production';
  if (gate >= 6) return 'power';
  if (gate >= 2) return 'plano';
  return 'demand';
}

function getGateStateV4(project, gateKey, versionName) {
  const display = getProjectDisplay(project);
  const rec = getGateRecordV4(project, gateKey, versionName);
  if (rec.status) return rec.status;
  const numeric = Number(gateKey);
  if (String(gateKey).startsWith('P')) return display.gate >= 10 ? '進行中' : '待開始';
  if (numeric < display.gate) return '已完成';
  if (numeric === display.gate) return display.status.includes('NG') ? 'NG' : '進行中';
  return '待開始';
}

function getOwnerByGateV4(project, gateKey) {
  const def = GATE_DEFS_V4[String(gateKey)];
  const assignment = project.assignmentV3 || {};
  if (!def) return 'PM 曼玉';
  if (def.ownerGroup === 'CNC') return assignment.cnc ? `CNC ${assignment.cnc}` : 'CNC 鍾民';
  if (def.ownerGroup === 'RD') return assignment.rd ? `RD ${assignment.rd}` : (project.id && project.id.includes('142') ? 'RD 齊政' : 'RD 世偉');
  if (def.ownerGroup === 'Production') return '生產 / 品保窗口';
  return assignment.pm ? `PM ${assignment.pm}` : 'PM 曼玉';
}

function getDefaultVersionsV4(project, phaseId) {
  const display = getProjectDisplay(project);
  if (phaseId === 'demand') return ['需求確認'];
  if (phaseId === 'plano') return project.id === 'SIH-142-UT' ? ['V13', 'V14'] : ['V1'];
  if (phaseId === 'power') return [display.version || 'V1'];
  if (phaseId === 'production') return ['量產準備'];
  return ['V1'];
}

function ensureGateCenterV4(project) {
  if (!project.phaseVersionsV4) project.phaseVersionsV4 = {};
  if (!project.gateRecordsV4) project.gateRecordsV4 = {};
  GATE_PHASES_V4.forEach(phase => {
    if (!project.phaseVersionsV4[phase.id]) {
      project.phaseVersionsV4[phase.id] = getDefaultVersionsV4(project, phase.id);
    }
    if (!project.gateRecordsV4[phase.id]) project.gateRecordsV4[phase.id] = {};
    project.phaseVersionsV4[phase.id].forEach(version => {
      if (!project.gateRecordsV4[phase.id][version]) project.gateRecordsV4[phase.id][version] = {};
      phase.gateKeys.forEach(gateKey => {
        if (!project.gateRecordsV4[phase.id][version][gateKey]) {
          project.gateRecordsV4[phase.id][version][gateKey] = {
            gateKey,
            version,
            status: '',
            owner: getOwnerByGateV4(project, gateKey),
            dueDate: '',
            handoffDate: '',
            trialDate: '',
            confirmedBy: '',
            confirmedAt: '',
            imageLink: '',
            imageNote: '',
            ng: null
          };
        }
      });
    });
  });
  return project.gateRecordsV4;
}

function getActivePhaseV4(project) {
  const key = `v4:${project.id}:phase`;
  return gateCenterV2State[key] || getCurrentPhaseV4(project);
}

function getPhaseVersionsV4(project, phaseId) {
  ensureGateCenterV4(project);
  return project.phaseVersionsV4[phaseId] || getDefaultVersionsV4(project, phaseId);
}

function getActiveVersionV4(project, phaseId) {
  const key = `v4:${project.id}:${phaseId}`;
  const versions = getPhaseVersionsV4(project, phaseId);
  if (!gateCenterV2State[key] || !versions.includes(gateCenterV2State[key])) {
    gateCenterV2State[key] = versions[versions.length - 1];
  }
  return gateCenterV2State[key];
}

function getGateRecordV4(project, gateKey, versionName) {
  ensureGateCenterV4(project);
  const phase = getPhaseByGateV4(gateKey);
  const version = versionName || getActiveVersionV4(project, phase.id);
  if (!project.gateRecordsV4[phase.id][version]) project.gateRecordsV4[phase.id][version] = {};
  if (!project.gateRecordsV4[phase.id][version][gateKey]) {
    project.gateRecordsV4[phase.id][version][gateKey] = {
      gateKey,
      version,
      status: '',
      owner: getOwnerByGateV4(project, gateKey),
      dueDate: '',
      handoffDate: '',
      trialDate: '',
      confirmedBy: '',
      confirmedAt: '',
      imageLink: '',
      imageNote: '',
      ng: null
    };
  }
  return project.gateRecordsV4[phase.id][version][gateKey];
}

function prepareCleanNavigation() {
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
}

function updateRoleSelectLabelsV4() {
  const select = document.getElementById('role-perspective-select');
  if (!select) return;
  const labels = {
    VP: '高階主管',
    PM: 'PM',
    RD: 'RD 研發',
    CNC: 'CNC / 模仁',
    Supervisor: '部門主管',
    Employee: '一般同仁',
    Admin: '系統管理員'
  };
  [...select.options].forEach(option => {
    if (labels[option.value]) option.textContent = labels[option.value];
  });
}

function updateLoginHelperV4() {
  const helper = document.querySelector('.login-helper');
  if (!helper) return;
  helper.innerHTML = `
    <p>測試帳號密碼皆為 <strong>123</strong></p>
    <ul>
      <li><strong>pm_user</strong>：PM 曼玉</li>
      <li><strong>rd_user</strong>：RD 世偉</li>
      <li><strong>cnc_user</strong>：CNC 鍾民</li>
      <li><strong>supervisor_user</strong>：主管視角</li>
      <li><strong>admin</strong>：系統管理員</li>
    </ul>
  `;
}

function initApp() {
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
}

function switchTab(tabId) {
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
}

function renderVPMilestones() {
  const container = document.getElementById('tab-vp-milestones');
  if (!container) return;
  currentProjectDetailId = null;
  const projects = getAllProjectsForCleanUI();
  const priorityOne = projects.filter(p => getProjectDisplay(p).priority === 1);
  const ngProjects = projects.filter(p => getProjectDisplay(p).status.includes('NG'));
  const myRole = ROLE_LABELS_V4[activePerspective] || activePerspective;
  const cards = priorityOne.concat(projects.filter(p => !priorityOne.includes(p)).slice(0, 2));
  container.innerHTML = `
    <section class="simple-page v4-page">
      <div class="page-heading v4-heading">
        <p class="eyebrow">My Work</p>
        <h1>我的工作</h1>
        <p>${escapeHtml(myRole)} 視角：先看你現在要處理哪一個專案、下一步要按哪裡。</p>
      </div>

      <div class="simple-stat-grid v4-stat-grid">
        <button class="simple-stat-card danger clickable-card" onclick="openProjectGateCenter('${escapeAttribute(ngProjects[0]?.id || priorityOne[0]?.id || projects[0]?.id || '')}')">
          <span>高優先處理</span><strong>${ngProjects.length}</strong><small>點擊進入對應專案</small>
        </button>
        <button class="simple-stat-card warning clickable-card" onclick="switchTab('action-items')">
          <span>今日需追蹤</span><strong>${priorityOne.length}</strong><small>查看月曆與里程碑</small>
        </button>
        <button class="simple-stat-card success clickable-card" onclick="switchTab('project-master')">
          <span>進行中專案</span><strong>${projects.length}</strong><small>進入專案總覽</small>
        </button>
      </div>

      <div class="section-title-row">
        <h2>現在要處理的專案</h2>
        <button class="btn btn-primary" onclick="switchTab('project-master')">查看全部專案</button>
      </div>
      <div class="work-card-list v4-work-list">
        ${cards.map(p => renderMyWorkCard(p)).join('')}
      </div>
    </section>
  `;
}

function renderMyWorkCard(project) {
  const d = getProjectDisplay(project);
  const gate = GATE_DEFS_V4[String(d.gate)] || GATE_DEFS_V4['1'];
  return `
    <article class="work-card v4-work-card" onclick="openProjectGateCenter('${escapeAttribute(project.id)}')">
      <div>
        <span class="priority-rank">第 ${d.priority} 優先</span>
        <h3>${escapeHtml(d.name)}</h3>
        <p>目前：Gate ${d.gate} ${escapeHtml(gate.name)}</p>
      </div>
      <div class="work-card-meta">
        <span class="status-pill ${statusClassV4(d.status)}">${escapeHtml(d.status)}</span>
        <span>負責：${escapeHtml(d.owner)}</span>
        <span>下一步：${escapeHtml(d.nextAction || '等待下一個 Gate')}</span>
      </div>
    </article>
  `;
}

function renderProjectMaster() {
  const container = document.getElementById('tab-project-master');
  if (!container) return;
  if (currentProjectDetailId) return renderProjectGateCenter(currentProjectDetailId);
  const projects = getAllProjectsForCleanUI().sort((a, b) => getProjectDisplay(a).priority - getProjectDisplay(b).priority);
  const groups = [1, 2, 3].map(priority => ({
    priority,
    title: priority === 1 ? '第一優先' : priority === 2 ? '第二優先' : '第三優先',
    projects: projects.filter(project => Math.min(3, getProjectDisplay(project).priority) === priority)
  }));
  container.innerHTML = `
    <section class="simple-page v4-page">
      <div class="page-heading v4-heading">
        <p class="eyebrow">Projects</p>
        <h1>專案總覽</h1>
        <p>依優先順序排列。點任一張專案卡片，就會進入該專案的 Gate 中心。</p>
      </div>
      ${groups.map(group => renderPriorityProjectSectionV4(group)).join('')}
    </section>
  `;
}

function renderPriorityProjectSectionV4(group) {
  return `
    <section class="project-status-section v4-priority-section">
      <div class="project-section-title">
        <span class="section-dot priority-${group.priority}"></span>
        <h2>${group.title} (${group.projects.length})</h2>
      </div>
      <div class="project-card-grid v4-project-grid">
        ${group.projects.length ? group.projects.map(project => renderProjectOverviewCard(project)).join('') : renderEmptyCleanState('目前沒有專案')}
      </div>
    </section>
  `;
}

function renderProjectOverviewCard(project) {
  const d = getProjectDisplay(project);
  const gate = GATE_DEFS_V4[String(d.gate)] || GATE_DEFS_V4['1'];
  const canAssign = typeof canAssignProjectsV3 === 'function' ? canAssignProjectsV3() : ['Admin', 'Supervisor', 'VP'].includes(activePerspective);
  return `
    <article class="simple-project-card v4-project-card priority-${d.priority}">
      <div class="project-card-top" onclick="openProjectGateCenter('${escapeAttribute(project.id)}')">
        <div>
          <span class="priority-rank">第 ${d.priority} 優先</span>
          <h3>${escapeHtml(d.name)}</h3>
        </div>
        <span class="status-pill ${statusClassV4(d.status)}">${escapeHtml(d.status)}</span>
      </div>
      <div class="project-card-body" onclick="openProjectGateCenter('${escapeAttribute(project.id)}')">
        <div class="gate-progress-mini"><span style="width:${Math.min(100, Math.round((d.gate / 9) * 100))}%"></span></div>
        <p class="project-current">Gate ${d.gate} · ${escapeHtml(gate.name)}</p>
        <div class="project-card-footer">
          <span>版本 ${escapeHtml(d.version)}</span>
          <span>${escapeHtml(d.owner)}</span>
        </div>
        ${d.ngReason ? `
          <div class="ng-mini-block">
            <strong>NG 原因</strong>
            <span>${escapeHtml(d.ngReason)}</span>
            <em>${escapeHtml(d.resultDue || '等待結果')}</em>
          </div>` : ''}
      </div>
      <div class="project-manage-row">
        <button class="btn btn-primary btn-xs" onclick="openProjectGateCenter('${escapeAttribute(project.id)}')">進入 Gate 中心</button>
        ${canAssign ? `<button class="btn btn-outline btn-xs" onclick="assignProjectV4('${escapeAttribute(project.id)}')">編輯分派</button>` : ''}
      </div>
    </article>
  `;
}

function assignProjectV4(projectId) {
  if (typeof canAssignProjectsV3 === 'function' && !canAssignProjectsV3()) return;
  const project = db.projects.find(p => p.id === projectId);
  if (!project) return;
  const d = getProjectDisplay(project);
  const priority = prompt('請輸入優先順序：1=第一優先，2=第二優先，3=第三優先', String(d.priority));
  if (priority === null) return;
  project.priority = Math.max(1, Math.min(3, Number(priority) || d.priority));
  project.priorityV4 = project.priority;
  addGateNotificationV4(projectId, '專案優先順序更新', `${project.id} 改為第 ${project.priority} 優先`);
  saveDatabase();
  renderProjectMaster();
  renderNotificationsBell();
}

function openProjectGateCenter(projectId) {
  if (!projectId) return;
  currentProjectDetailId = projectId;
  switchTab('project-master');
}

function setProjectPhaseTab(projectId, phaseId) {
  const project = db.projects.find(p => p.id === projectId);
  if (!project) return;
  gateCenterV2State[`v4:${projectId}:phase`] = phaseId;
  renderProjectGateCenter(projectId);
}

function setProjectPhaseVersion(projectId, phaseId, version) {
  gateCenterV2State[`v4:${projectId}:${phaseId}`] = version;
  renderProjectGateCenter(projectId);
}

function addProjectPhaseVersion(projectId, phaseId) {
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
    project.gateRecordsV4[phaseId][name][gateKey] = {
      gateKey,
      version: name,
      status: '待開始',
      owner: getOwnerByGateV4(project, gateKey),
      dueDate: '',
      handoffDate: '',
      trialDate: '',
      confirmedBy: '',
      confirmedAt: '',
      imageLink: '',
      imageNote: '',
      ng: null,
      createReason: reason || ''
    };
  });
  gateCenterV2State[`v4:${projectId}:${phaseId}`] = name;
  addGateNotificationV4(projectId, '新增版本', `${project.id} 新增 ${name}：${reason || '未填原因'}`);
  saveDatabase();
  renderProjectGateCenter(projectId);
  renderNotificationsBell();
}

function editProjectPhaseVersion(projectId, phaseId, oldName) {
  const project = db.projects.find(p => p.id === projectId);
  if (!project) return;
  ensureGateCenterV4(project);
  const nextName = prompt('請輸入新的版本名稱', oldName);
  if (!nextName || nextName === oldName) return;
  const versions = project.phaseVersionsV4[phaseId] || [];
  const idx = versions.indexOf(oldName);
  if (idx >= 0) versions[idx] = nextName;
  const recordSet = project.gateRecordsV4[phaseId][oldName] || {};
  delete project.gateRecordsV4[phaseId][oldName];
  project.gateRecordsV4[phaseId][nextName] = recordSet;
  Object.values(recordSet).forEach(rec => { rec.version = nextName; });
  gateCenterV2State[`v4:${projectId}:${phaseId}`] = nextName;
  saveDatabase();
  renderProjectGateCenter(projectId);
}

function nextVersionNameV4(latest, phaseId, count) {
  if (phaseId === 'production') return `量產版 ${count + 1}`;
  const match = String(latest).match(/V(\d+)(?:\.(\d+))?/i);
  if (!match) return `V${count + 1}`;
  return `V${Number(match[1]) + 1}`;
}

function renderProjectGateCenter(projectId) {
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
      <div class="page-action-row">
        <button class="back-button" onclick="currentProjectDetailId=null; renderProjectMaster();">← 回專案總覽</button>
        <button class="btn btn-outline" onclick="exportProjectMeetingExcel('${escapeAttribute(project.id)}')">匯出會議 Excel</button>
      </div>

      <div class="project-detail-heading v4-heading">
        <p class="eyebrow">Project Gate Center</p>
        <h1>${escapeHtml(d.name)}</h1>
      </div>

      <article class="gate-summary-card v4-summary ${statusClassV4(d.status)}">
        <div class="summary-main">
          <span class="status-pill ${statusClassV4(d.status)}">${escapeHtml(d.status)}</span>
          <h2>目前進度：Gate ${d.gate} ${escapeHtml(currentGate.name)}</h2>
          <p>${escapeHtml(d.ngReason || currentGate.description)}</p>
        </div>
        <div class="summary-grid">
          <div><span>目前大項目</span><strong>${escapeHtml(GATE_PHASES_V4.find(p => p.id === currentPhase)?.label || '')}</strong></div>
          <div><span>目前版本</span><strong>${escapeHtml(d.version)}</strong></div>
          <div><span>下一步</span><strong>${escapeHtml(d.nextAction || '等待下一個 Gate')}</strong></div>
          <div><span>負責 / PM追蹤</span><strong>${escapeHtml(d.owner)} / ${escapeHtml(d.tracker)}</strong></div>
        </div>
      </article>

      <div class="phase-dashboard v4-phase-dashboard">
        ${GATE_PHASES_V4.map(phase => renderPhaseDashboardCardV4(project, phase, currentPhase, activePhaseId)).join('')}
      </div>

      <div class="phase-tabs v4-phase-tabs">
        ${GATE_PHASES_V4.map(phase => `
          <button class="${phase.id === activePhaseId ? 'active' : ''}" onclick="setProjectPhaseTab('${escapeAttribute(project.id)}', '${phase.id}')">${escapeHtml(phase.label)}</button>
        `).join('')}
      </div>

      <section class="phase-panel v4-phase-panel">
        <div class="phase-panel-head">
          <div>
            <h2>${escapeHtml(activePhase.label)}</h2>
            <p>${escapeHtml(activePhase.subtitle)}</p>
          </div>
          ${renderVersionTabsV4(project, activePhaseId, activeVersion)}
        </div>
        <div class="horizontal-gate-list v4-gate-list">
          ${activePhase.gateKeys.map(key => renderGateCardV4(project, key, activeVersion)).join('')}
        </div>
      </section>
    </section>
  `;
}

function renderPhaseDashboardCardV4(project, phase, currentPhase, activePhaseId) {
  const activeVersion = getActiveVersionV4(project, phase.id);
  const statuses = phase.gateKeys.map(gateKey => getGateStateV4(project, gateKey, activeVersion));
  const done = statuses.filter(s => s === '已完成').length;
  const state = statuses.includes('NG') ? 'NG' : done === phase.gateKeys.length ? '已完成' : phase.id === currentPhase ? '進行中' : '待開始';
  return `
    <button class="phase-dash-card ${phase.id === activePhaseId ? 'active' : ''} ${statusClassV4(state)}" onclick="setProjectPhaseTab('${escapeAttribute(project.id)}', '${phase.id}')">
      <span>${escapeHtml(phase.label)}</span>
      <strong>${escapeHtml(state)}</strong>
      <small>${done}/${phase.gateKeys.length} Gate 完成</small>
    </button>
  `;
}

function renderVersionTabsV4(project, phaseId, activeVersion) {
  const versions = getPhaseVersionsV4(project, phaseId);
  const canAdd = phaseId !== 'demand';
  return `
    <div class="version-tabs-v2 v4-version-tabs">
      ${versions.map(v => `
        <button class="${v === activeVersion ? 'active' : ''}" onclick="setProjectPhaseVersion('${escapeAttribute(project.id)}', '${phaseId}', '${escapeAttribute(v)}')">${escapeHtml(v)}</button>
      `).join('')}
      ${activeVersion ? `<button class="add-version subtle" onclick="editProjectPhaseVersion('${escapeAttribute(project.id)}', '${phaseId}', '${escapeAttribute(activeVersion)}')">編輯版本</button>` : ''}
      ${canAdd ? `<button class="add-version" onclick="addProjectPhaseVersion('${escapeAttribute(project.id)}', '${phaseId}')">＋新增頁籤</button>` : ''}
    </div>
  `;
}

function renderGateCardV4(project, gateKey, activeVersion) {
  const key = String(gateKey);
  const def = GATE_DEFS_V4[key];
  if (!def) return '';
  const rec = getGateRecordV4(project, key, activeVersion);
  const state = getGateStateV4(project, key, activeVersion);
  const isDone = state === '已完成';
  const canEdit = canEditGateV4(def, rec);
  const canConfirm = canConfirmGateV4(def, rec) && !isDone;
  const safe = safeDomId(`v4-${project.id}-${key}-${activeVersion}`);
  const roleHint = isDone ? '已完成，僅供追溯' : canConfirm ? '你可以確認此 Gate' : `只有 ${def.department} / 主管可確認`;
  const ownerOptions = getOwnerOptionsV4(def.ownerGroup, rec.owner);
  return `
    <article class="gate-row-card v4-gate-card ${statusClassV4(state)} ${isDone ? 'readonly' : ''}">
      <div class="gate-row-main">
        <div class="gate-row-title">
          <span class="gate-number">${key.startsWith('P') ? key : `Gate ${key}`}</span>
          <div>
            <h3>${escapeHtml(def.name)}</h3>
            <p>${escapeHtml(def.description)}</p>
          </div>
        </div>
        <span class="status-pill ${statusClassV4(state)}">${escapeHtml(state)}</span>
      </div>

      <div class="gate-row-fields v4-gate-fields">
        <label><span>版本</span><strong>${escapeHtml(activeVersion)}</strong></label>
        <label><span>部門</span><strong>${escapeHtml(def.department)}</strong></label>
        <label><span>負責人</span>
          <select id="owner-${safe}" ${canEdit ? '' : 'disabled'}>
            ${ownerOptions}
          </select>
        </label>
        <label><span>預計完成時間</span><input id="due-${safe}" type="date" value="${escapeAttribute(rec.dueDate || '')}" ${canEdit ? '' : 'disabled'}></label>
        ${key === '4' ? `<label><span>移交打樣時間</span><input id="handoff-${safe}" type="date" value="${escapeAttribute(rec.handoffDate || '')}" ${canEdit ? '' : 'disabled'}></label>` : ''}
        ${key === '5' || key === '9' ? `<label><span>預計試戴時間</span><input id="trial-${safe}" type="date" value="${escapeAttribute(rec.trialDate || '')}" ${canEdit ? '' : 'disabled'}></label>` : ''}
        <label><span>確認人</span><strong>${escapeHtml(rec.confirmedBy || '尚未確認')}</strong></label>
        <label><span>確認時間</span><strong>${escapeHtml(rec.confirmedAt || '尚未確認')}</strong></label>
      </div>

      ${renderGateChecklistV4(def)}
      ${key === '4' || key === '8' ? renderImageRecordBlockV4(project, key, rec, safe, canEdit) : ''}
      ${key === '5' || key === '9' ? renderNgBlockV4(project, key, rec, safe, canConfirm) : ''}

      <div class="gate-row-actions">
        <span>${escapeHtml(roleHint)}</span>
        ${canEdit ? `<button class="btn btn-outline" onclick="saveGateFieldsV4('${escapeAttribute(project.id)}', '${key}', '${escapeAttribute(activeVersion)}', '${safe}')">儲存欄位</button>` : ''}
        ${key === '5' || key === '9' ? `<button class="btn btn-outline danger" ${canConfirm ? '' : 'disabled'} onclick="showNgFormV2('${safe}')">填寫 NG</button>` : ''}
        <button class="btn btn-primary" ${canConfirm ? '' : 'disabled'} onclick="confirmProjectGateV4('${escapeAttribute(project.id)}', '${key}', '${escapeAttribute(activeVersion)}')">確認完成</button>
      </div>
    </article>
  `;
}

function renderGateChecklistV4(def) {
  return `
    <ul class="gate-checklist v4-checklist">
      ${def.checklist.map(item => `<li>${escapeHtml(item)}</li>`).join('')}
    </ul>
  `;
}

function getOwnerOptionsV4(group, selected) {
  const names = TEAM_MEMBERS_V4[group] || TEAM_MEMBERS_V4.PM;
  const normalized = names.map(name => group === 'RD' ? `RD ${name}` : group === 'CNC' ? `CNC ${name}` : group === 'PM' ? `PM ${name}` : name);
  const all = Array.from(new Set([selected, ...normalized].filter(Boolean)));
  return all.map(name => `<option value="${escapeAttribute(name)}" ${name === selected ? 'selected' : ''}>${escapeHtml(name)}</option>`).join('');
}

function renderImageRecordBlockV4(project, key, rec, safe, canEdit) {
  return `
    <div class="gate-evidence-block v4-evidence">
      <div>
        <strong>圖片紀錄</strong>
        <p>研發可貼圖片連結，用圖說明目前設計狀態。</p>
      </div>
      <input id="image-link-${safe}" value="${escapeAttribute(rec.imageLink || '')}" placeholder="貼上圖片連結或檔案路徑" ${canEdit ? '' : 'disabled'}>
      <input id="image-note-${safe}" value="${escapeAttribute(rec.imageNote || '')}" placeholder="圖片備註，例如：邊緣、移印、構型差異" ${canEdit ? '' : 'disabled'}>
      ${rec.imageLink ? `<a class="image-open-link" href="${escapeAttribute(rec.imageLink)}" target="_blank" rel="noopener">另開圖片</a>` : ''}
    </div>
  `;
}

function renderNgBlockV4(project, key, rec, safe, canEdit) {
  const ng = rec.ng || {};
  return `
    <div class="ng-record-block ${rec.ng ? '' : 'collapsed'}" id="ng-form-${safe}">
      <div class="ng-record-title">
        <strong>NG 紀錄</strong>
        <span>填寫時間 / 原因 / 備註，系統會新增下一版並退回指定 Gate。</span>
      </div>
      <div class="ng-form-grid">
        <label>NG 時間
          <input id="ng-time-${safe}" type="datetime-local" value="${escapeAttribute(toDatetimeLocalValue(ng.time) || getNowDatetimeLocal())}" ${canEdit ? '' : 'disabled'}>
        </label>
        <label>預計有結果
          <input id="ng-due-${safe}" type="date" value="${escapeAttribute(ng.due || '')}" ${canEdit ? '' : 'disabled'}>
        </label>
        <label>NG 原因
          <input id="ng-reason-${safe}" value="${escapeAttribute(ng.reason || '')}" placeholder="請輸入原因，例如：全光度試戴不穩定" ${canEdit ? '' : 'disabled'}>
        </label>
        <label>備註
          <textarea id="ng-note-${safe}" rows="2" placeholder="補充測試現象、討論結論、需誰協助" ${canEdit ? '' : 'disabled'}>${escapeHtml(ng.note || '')}</textarea>
        </label>
      </div>
      <div class="ng-form-actions">
        <button class="btn btn-outline danger" ${canEdit ? '' : 'disabled'} onclick="saveGateNgV4('${escapeAttribute(project.id)}', '${key}', '${safe}')">儲存 NG 並新增下一版</button>
      </div>
    </div>
  `;
}

function canEditGateV4(def, rec) {
  if (['Admin', 'VP', 'Supervisor'].includes(activePerspective)) return true;
  if (activePerspective === def.confirmRole) return true;
  if (activePerspective === 'Employee' && rec && isAssignedToCurrentUserV3(rec)) return true;
  return false;
}

function canConfirmGateV4(def, rec) {
  if (['Admin', 'VP', 'Supervisor'].includes(activePerspective)) return true;
  if (activePerspective === def.confirmRole) return true;
  if (activePerspective === 'Employee' && rec && isAssignedToCurrentUserV3(rec)) return true;
  return false;
}

function saveGateFieldsV4(projectId, gateKey, version, safe) {
  const project = db.projects.find(p => p.id === projectId);
  if (!project) return;
  const rec = getGateRecordV4(project, gateKey, version);
  rec.owner = getInputValue(`owner-${safe}`) || rec.owner;
  rec.dueDate = getInputValue(`due-${safe}`) || rec.dueDate;
  rec.handoffDate = getInputValue(`handoff-${safe}`) || rec.handoffDate;
  rec.trialDate = getInputValue(`trial-${safe}`) || rec.trialDate;
  rec.imageLink = getInputValue(`image-link-${safe}`) || rec.imageLink;
  rec.imageNote = getInputValue(`image-note-${safe}`) || rec.imageNote;
  addGateNotificationV4(projectId, 'Gate 欄位更新', `${GATE_DEFS_V4[gateKey].name} 更新負責人/時間欄位`);
  saveDatabase();
  renderProjectGateCenter(projectId);
  renderNotificationsBell();
}

function confirmProjectGateV4(projectId, gateKey, version) {
  const project = db.projects.find(p => p.id === projectId);
  if (!project) return;
  const def = GATE_DEFS_V4[String(gateKey)];
  const rec = getGateRecordV4(project, gateKey, version);
  if (!canConfirmGateV4(def, rec)) return;
  rec.status = '已完成';
  rec.confirmedBy = currentUser.name || ROLE_LABELS_V4[activePerspective] || '使用者';
  rec.confirmedAt = formatNowMinuteV2();
  const next = getNextGateKeyV4(gateKey);
  if (next) {
    const nextPhase = getPhaseByGateV4(next);
    const nextVersion = getActiveVersionV4(project, nextPhase.id);
    const nextRec = getGateRecordV4(project, next, nextVersion);
    if (!nextRec.status || nextRec.status === '待開始') nextRec.status = '進行中';
    project.cleanGateState = {
      ...(project.cleanGateState || {}),
      gate: String(next).startsWith('P') ? 10 : Number(next),
      status: '進行中',
      nextAction: `進行 ${String(next).startsWith('P') ? next : `Gate ${next}`} ${GATE_DEFS_V4[next].name}`,
      owner: nextRec.owner
    };
  } else {
    project.cleanGateState = { ...(project.cleanGateState || {}), gate: 10, status: '已完成', nextAction: '量產追蹤完成' };
  }
  addGateNotificationV4(projectId, 'Gate 確認完成', `${def.name} 由 ${rec.confirmedBy} 確認完成，${rec.confirmedAt}`);
  saveDatabase();
  renderProjectGateCenter(projectId);
  renderNotificationsBell();
}

function getNextGateKeyV4(gateKey) {
  const flat = GATE_PHASES_V4.flatMap(p => p.gateKeys);
  const idx = flat.indexOf(String(gateKey));
  return idx >= 0 ? flat[idx + 1] : null;
}

function saveGateNgV4(projectId, gateKey, safe) {
  const project = db.projects.find(p => p.id === projectId);
  if (!project) return;
  const phase = getPhaseByGateV4(gateKey);
  const version = getActiveVersionV4(project, phase.id);
  const rec = getGateRecordV4(project, gateKey, version);
  const def = GATE_DEFS_V4[String(gateKey)];
  if (!canConfirmGateV4(def, rec)) return;
  const backGate = String(gateKey) === '9' ? '6' : '2';
  const backPhase = String(gateKey) === '9' ? 'power' : 'plano';
  const reason = getInputValue(`ng-reason-${safe}`) || '未填原因';
  const due = getInputValue(`ng-due-${safe}`);
  const ngTime = formatDatetimeLocalMinute(getInputValue(`ng-time-${safe}`) || getNowDatetimeLocal());
  rec.status = 'NG';
  rec.confirmedBy = currentUser.name || ROLE_LABELS_V4[activePerspective] || '使用者';
  rec.confirmedAt = ngTime;
  rec.ng = {
    time: ngTime,
    due,
    reason,
    note: getInputValue(`ng-note-${safe}`),
    backGate,
    nextAction: `新增下一版，退回 Gate ${backGate} ${GATE_DEFS_V4[backGate].name}`
  };
  const versions = getPhaseVersionsV4(project, backPhase);
  const nextName = nextVersionNameV4(version, backPhase, versions.length);
  if (!versions.includes(nextName)) {
    versions.push(nextName);
    project.gateRecordsV4[backPhase][nextName] = {};
    GATE_PHASES_V4.find(p => p.id === backPhase).gateKeys.forEach(key => {
      project.gateRecordsV4[backPhase][nextName][key] = {
        gateKey: key,
        version: nextName,
        status: key === backGate ? '進行中' : '待開始',
        owner: getOwnerByGateV4(project, key),
        dueDate: '',
        confirmedBy: '',
        confirmedAt: '',
        imageLink: '',
        imageNote: '',
        ng: null,
        createReason: rec.ng.nextAction
      };
    });
  }
  gateCenterV2State[`v4:${projectId}:phase`] = backPhase;
  gateCenterV2State[`v4:${projectId}:${backPhase}`] = nextName;
  project.cleanGateState = {
    status: 'NG / 等待結果',
    gate: Number(backGate),
    version: nextName,
    owner: getOwnerByGateV4(project, backGate),
    ngReason: reason,
    resultDue: due ? `預計 ${due} 有結果` : '等待負責人回覆',
    nextAction: rec.ng.nextAction
  };
  addGateNotificationV4(projectId, 'NG 與版本更新', `${def.name} NG：${reason}；已新增 ${nextName}`);
  saveDatabase();
  renderProjectGateCenter(projectId);
  renderNotificationsBell();
}

function addGateNotificationV4(projectId, title, message) {
  if (typeof addGateNotification === 'function') {
    addGateNotification({ projectId, title, message, actor: currentUser?.name || ROLE_LABELS_V4[activePerspective] || '系統' });
  }
}

function renderActionItems() {
  const container = document.getElementById('tab-action-items');
  if (!container) return;
  currentProjectDetailId = null;
  const view = gateCenterV2State.progressView || 'gantt';
  container.innerHTML = `
    <section class="simple-page v4-page progress-page-v2">
      <div class="page-heading v4-heading">
        <p class="eyebrow">Progress</p>
        <h1>專案進度</h1>
        <p>開會用：左邊看規格/版本/目前進度，右邊用月曆看這個月每天在做什麼。</p>
      </div>
      <div class="phase-tabs progress-tabs v4-phase-tabs">
        <button class="${view === 'gantt' ? 'active' : ''}" onclick="setProgressViewV2('gantt')">甘特圖（月曆）</button>
        <button class="${view === 'milestone' ? 'active' : ''}" onclick="setProgressViewV2('milestone')">里程碑</button>
      </div>
      ${view === 'milestone' ? renderMilestoneProgressV4() : renderCalendarGanttV4()}
    </section>
  `;
}

function setProgressViewV2(view) {
  gateCenterV2State.progressView = view;
  renderActionItems();
}

function renderCalendarGanttV4() {
  const projects = getAllProjectsForCleanUI();
  const days = makeJune2026DaysV4();
  return `
    <div class="calendar-gantt-v4">
      <div class="calendar-left-v4">
        <div class="cal-left-header">
          <span>DIA</span><span>版本</span><span>平光</span><span>全光</span><span>目前進度</span>
        </div>
        ${projects.map(project => {
          const d = getProjectDisplay(project);
          const gate = GATE_DEFS_V4[String(d.gate)] || GATE_DEFS_V4['1'];
          return `
            <div class="cal-left-row" onclick="openProjectGateCenter('${escapeAttribute(project.id)}')">
              <strong>${escapeHtml(d.spec)}</strong>
              <span>${escapeHtml(d.version)}</span>
              <span>${escapeHtml(d.planoDate ? `${d.planoDate.slice(5)} ${d.gate > 5 ? '已製' : '進行中'}` : '待排')}</span>
              <span>${escapeHtml(d.powerDate ? `${d.powerDate.slice(5)} ${d.gate >= 9 ? d.status.replace(' / ', '/') : '安排中'}` : '待排')}</span>
              <span>${escapeHtml(`Gate ${d.gate} ${gate.name}`)}</span>
            </div>
          `;
        }).join('')}
      </div>
      <div class="calendar-right-v4">
        <div class="calendar-title-v4">2026 年 6 月</div>
        <div class="calendar-weekdays-v4">${['日', '一', '二', '三', '四', '五', '六'].map(d => `<span>${d}</span>`).join('')}</div>
        <div class="calendar-grid-v4">
          ${days.map(day => renderCalendarDayV4(day)).join('')}
        </div>
      </div>
    </div>
  `;
}

function makeJune2026DaysV4() {
  const days = [];
  for (let i = 0; i < 42; i += 1) {
    const d = new Date(2026, 5, 1 - 1 + i);
    days.push({
      date: d,
      iso: `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`,
      inMonth: d.getMonth() === 5,
      day: d.getDate()
    });
  }
  return days;
}

function renderCalendarDayV4(day) {
  const events = PROJECT_EVENTS_V4.filter(event => event.date === day.iso);
  return `
    <div class="calendar-day-v4 ${day.inMonth ? '' : 'muted'} ${day.iso === '2026-06-25' ? 'today' : ''}">
      <span class="day-num">${day.day}</span>
      <div class="day-events-v4">
        ${events.map(event => `
          <button class="calendar-event-v4 ${event.type}" onclick="openProjectGateCenter('${escapeAttribute(event.projectId)}')">
            ${escapeHtml(event.label)}
          </button>
        `).join('')}
      </div>
    </div>
  `;
}

function renderMilestoneProgressV4() {
  const projects = getAllProjectsForCleanUI();
  return `
    <div class="progress-table-card v4-milestone-card">
      <div class="progress-table-head milestone-grid">
        <span>規格</span><span>版次</span><span>目前進度</span><span>下一個階段</span>
      </div>
      ${projects.map(project => {
        const d = getProjectDisplay(project);
        const gate = GATE_DEFS_V4[String(d.gate)] || GATE_DEFS_V4['1'];
        return `
          <div class="progress-table-row milestone-grid" onclick="openProjectGateCenter('${escapeAttribute(project.id)}')">
            <strong>${escapeHtml(d.spec)}</strong>
            <span>${escapeHtml(d.version)}</span>
            <span>Gate ${d.gate} ${escapeHtml(gate.name)}</span>
            <span>${escapeHtml(d.nextAction || '等待下一階段')}</span>
          </div>
        `;
      }).join('')}
    </div>
  `;
}

function renderRisksIssues() {
  const container = document.getElementById('tab-risks-issues');
  if (!container) return;
  currentProjectDetailId = null;
  container.innerHTML = `
    <section class="simple-page v4-page team-page-v2">
      <div class="page-heading v4-heading">
        <p class="eyebrow">Team</p>
        <h1>團隊成員</h1>
        <p>依部門看主管與同仁；權限不放在畫面上干擾使用者，由系統背景判斷。</p>
      </div>
      <div class="team-dept-grid-v2 v4-team-grid">
        ${TEAM_STRUCTURE_V4.map(team => `
          <article class="team-dept-card-v2 v4-team-card">
            <div class="team-dept-head">
              <div>
                <h2>${escapeHtml(team.dept)}</h2>
                <p>主管：${escapeHtml(team.manager)} ${escapeHtml(team.title)} · ${escapeHtml(team.extension)}</p>
              </div>
            </div>
            <h3>同仁</h3>
            <div class="member-list-v4">
              ${team.members.map(member => `
                <div class="member-row-v4">
                  <strong>${escapeHtml(member.name)}</strong>
                  <span>${escapeHtml(member.title)}</span>
                  <em>${escapeHtml(member.extension)}</em>
                </div>
              `).join('')}
            </div>
          </article>
        `).join('')}
      </div>
    </section>
  `;
}

function exportProjectMeetingExcel(projectId) {
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
        rows.push({
          '專案名稱': display.name,
          '規格': display.spec,
          '大項目': phase.label,
          '版本': version,
          'Gate': gateKey,
          'Gate 名稱': def.name,
          '狀態': getGateStateV4(project, gateKey, version),
          '負責部門': def.department,
          '負責人': rec.owner,
          '預計完成時間': rec.dueDate || '',
          '移交打樣時間': rec.handoffDate || '',
          '預計試戴時間': rec.trialDate || '',
          '確認人': rec.confirmedBy || '',
          '確認時間': rec.confirmedAt || '',
          'NG 時間': rec.ng?.time || '',
          'NG 原因': rec.ng?.reason || '',
          'NG 備註': rec.ng?.note || '',
          '下一步': rec.ng?.nextAction || display.nextAction || '',
          '圖片連結': rec.imageLink || '',
          '圖片備註': rec.imageNote || '',
          'PM 追蹤': display.tracker
        });
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
}

function showLogin() {
  document.getElementById('login-overlay').classList.remove('hidden');
  document.getElementById('app-container').classList.add('hidden');
  document.getElementById('login-error').style.display = 'none';
  document.getElementById('username').value = '';
  document.getElementById('password').value = '';
  updateLoginHelperV3();
}

function updateLoginHelperV3() {
  const helper = document.querySelector('.login-helper');
  if (!helper) return;
  helper.innerHTML = `
    <p>測試帳號密碼皆為 123</p>
    <ul>
      <li><strong>admin</strong>（系統管理員，可切換所有視角）</li>
      <li><strong>pm_user</strong>（產品 PM）</li>
      <li><strong>rd_user</strong>（RD 研發）</li>
      <li><strong>cnc_user</strong>（CNC / 模仁）</li>
      <li><strong>supervisor_user</strong>（主管：可編輯、分派、審核）</li>
      <li><strong>employee_user</strong>（一般員工：只處理被分派 Gate）</li>
      <li><strong>vp_user</strong>（副總 / 高階主管）</li>
    </ul>
  `;
}

function openProjectDetailFromWorkbench(projectId) {
  currentProjectDetailId = projectId;
  switchTabFromWorkbench('project-master');
}

function getSortedProjectsForWorkbench() {
  return [...db.projects].sort((a, b) => {
    ensureProjectProcessModel(a);
    ensureProjectProcessModel(b);
    const score = (p) => {
      const overdue = p.deadline && new Date(p.deadline) < new Date(new Date().toISOString().slice(0, 10));
      return (p.status === '紅' ? 100 : p.status === '黃' ? 50 : 0)
        + (overdue ? 35 : 0)
        + ([5, 9].includes(p.process.currentStep) ? 20 : 0)
        + (100 - (p.priority || 99));
    };
    return score(b) - score(a);
  });
}

function getMyActions() {
  if (!currentUser) return [];
  return db.actions.filter(action => action.status !== '已完成'
    && (`${action.owner || ''}${action.coowners || ''}`.includes(currentUser.name)
      || (activePerspective === 'PM' && action.owner === '曼玉')
      || activePerspective === 'Admin'));
}

function getLightClass(status) {
  if (status === '綠') return 'badge-green';
  if (status === '黃') return 'badge-orange';
  if (status === '紅') return 'badge-red';
  return 'badge-gray';
}

function getProjectPeople(proj) {
  const people = `${proj.owner || ''}、${proj.collaborators || ''}`
    .split(/[、,，/／\s]+/)
    .map(name => name.trim())
    .filter(Boolean);
  return [...new Set(people)];
}

function pickProjectPerson(proj, preferredNames, fallback = '待指派') {
  const people = getProjectPeople(proj);
  return preferredNames.find(name => people.includes(name)) || people.find(name => preferredNames.includes(name)) || fallback;
}

function getStageAssignment(proj, stageId) {
  const tracker = proj.owner || 'PM待指派';
  const rdPerson = pickProjectPerson(proj, ['齊政', '世偉', '宣義'], 'RD待指派');
  const moldPerson = pickProjectPerson(proj, ['鍾民', '世偉'], '模仁待指派');
  const qaPerson = pickProjectPerson(proj, ['宣義', '世偉', '齊政'], '品保待指派');
  const productionPerson = pickProjectPerson(proj, ['世偉', '鍾民'], '生產待指派');

  const map = {
    1: { department: 'PM／產品企劃', person: tracker, tracker },
    2: { department: 'RD／模仁', person: `${rdPerson}／${moldPerson}`, tracker },
    3: { department: '模仁／製程', person: `${moldPerson}／${productionPerson}`, tracker },
    4: { department: 'RD／品保', person: `${rdPerson}／${qaPerson}`, tracker },
    5: { department: 'RD／臨床測試', person: rdPerson, tracker },
    6: { department: 'RD／模仁', person: `${rdPerson}／${moldPerson}`, tracker },
    7: { department: '模仁／製程', person: `${moldPerson}／${productionPerson}`, tracker },
    8: { department: 'RD／品保', person: `${rdPerson}／${qaPerson}`, tracker },
    9: { department: 'RD／臨床測試', person: rdPerson, tracker },
    10: { department: '生產／品保／PM', person: `${productionPerson}／${qaPerson}／${tracker}`, tracker }
  };
  return map[stageId] || { department: '待確認部門', person: '待指派', tracker };
}

function renderAssignmentPills(assignment) {
  return `
    <div class="assignment-pills">
      <span><b>部門</b>${escapeHtml(assignment.department)}</span>
      <span><b>負責</b>${escapeHtml(assignment.person)}</span>
      <span><b>追蹤</b>${escapeHtml(assignment.tracker)}</span>
    </div>
  `;
}

function getProjectGuidance(proj) {
  ensureProjectProcessModel(proj);
  const stage = PRODUCT_PROCESS_STAGES[proj.process.currentStep - 1];
  if (proj.status === '紅' && proj.process.currentStep === 5) {
    return {
      tone: 'danger',
      title: '平光測試 NG',
      nextText: '系統判斷下一步應退回第 2 階段「平光模仁設計」，建立下一輪修正。',
      cta: '處理 NG：回第 2 階段',
      backTo: 2
    };
  }
  if (proj.status === '紅' && proj.process.currentStep === 9) {
    return {
      tone: 'danger',
      title: '全光度測試 NG',
      nextText: '系統判斷下一步應退回第 6 階段「全光度模仁設計」，建立下一輪修正。',
      cta: '處理 NG：回第 6 階段',
      backTo: 6
    };
  }
  if (proj.process.currentStep === 5) {
    return {
      tone: 'warning',
      title: '等待平光測試判定',
      nextText: 'RD 需登錄平光測試 OK/NG；OK 進第 6，NG 回第 2。',
      cta: '填寫測試判定',
      backTo: null
    };
  }
  if (proj.process.currentStep === 9) {
    return {
      tone: proj.status === '紅' ? 'danger' : 'warning',
      title: '等待全光度測試判定',
      nextText: 'RD 需登錄全光度測試 OK/NG；OK 進第 10 量產，NG 回第 6。',
      cta: '填寫測試判定',
      backTo: null
    };
  }
  if (proj.process.currentStep === 10) {
    return {
      tone: 'success',
      title: '量產準備／追蹤',
      nextText: '確認生產排程、良率、品保與標示包材是否完成。',
      cta: '看量產條件',
      backTo: null
    };
  }
  const nextStage = PRODUCT_PROCESS_STAGES[proj.process.currentStep];
  return {
    tone: proj.status === '紅' ? 'danger' : proj.status === '黃' ? 'warning' : 'info',
    title: `${stage.shortName}進行中`,
    nextText: proj.nextStep || `完成本階段確認後，送往第 ${nextStage ? nextStage.id : proj.process.currentStep} 階段${nextStage ? `「${nextStage.shortName}」` : ''}。`,
    cta: '查看並更新進度',
    backTo: null
  };
}

function getGateName(gate) {
  const names = {
    "G0": "需求確認",
    "G1": "模仁設計",
    "G2": "平光測試",
    "G3": "全光設計",
    "G4": "全光試戴",
    "G5": "試量產與品質準備",
    "G6": "量產放行"
  };
  return names[gate] || '';
}

function getNextGate(gate) {
  const gates = ["G0", "G1", "G2", "G3", "G4", "G5", "G6"];
  const idx = gates.indexOf(gate);
  if (idx !== -1 && idx < gates.length - 1) {
    return gates[idx + 1];
  }
  return gate;
}

// ==========================================================================
// 3. 分頁 2: 專案甘特圖
// ==========================================================================

function renderGanttChart() {
  const header = document.getElementById('gantt-header');
  const body = document.getElementById('gantt-body');
  
  // 1. 渲染甘特圖表頭 (兩列，第一列是月份，第二列是週別/日期)
  let rowMonth = '<tr><th rowspan="2" class="sticky-col-1" style="background-color: var(--bg-primary);">規格專案</th><th rowspan="2" class="sticky-col-2" style="background-color: var(--bg-primary);">版本狀態</th>';
  let rowDay = '<tr>';
  
  GANTT_WEEKS.forEach(wk => {
    if (wk.month) {
      // 算出月份跨了多少週 (直到下個有設定 wk.month 的週或到最後)
      let colSpan = 1;
      const startIdx = GANTT_WEEKS.indexOf(wk);
      for (let i = startIdx + 1; i < GANTT_WEEKS.length; i++) {
        if (GANTT_WEEKS[i].month === "") colSpan++;
        else break;
      }
      rowMonth += `<th colspan="${colSpan}" style="font-size:10px;">${wk.month}</th>`;
    }
    rowDay += `<th style="font-size:9px; min-width:32px;">${wk.day}</th>`;
  });
  
  rowMonth += '</tr>';
  rowDay += '</tr>';
  header.innerHTML = rowMonth + rowDay;
  
  // 2. 渲染甘特圖內容 (以專案 -> 版本兩層渲染)
  body.innerHTML = '';
  
  db.projects.forEach(proj => {
    // 找出該專案的所有版本
    const projVersions = db.versions.filter(v => v.projectId === proj.id);
    
    // 若沒有版本，至少渲染一條空列
    if (projVersions.length === 0) {
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td class="sticky-col-1"><strong>${proj.name}</strong></td>
        <td class="sticky-col-2 text-muted">無版本資料</td>
        ${GANTT_WEEKS.map(() => '<td></td>').join('')}
      `;
      body.appendChild(tr);
      return;
    }
    
    // 渲染每個版本的甘特列
    projVersions.forEach((ver, vIdx) => {
      const tr = document.createElement('tr');
      
      // 第一列要合併專案名稱
      let projColHtml = '';
      if (vIdx === 0) {
        projColHtml = `<td class="sticky-col-1" rowspan="${projVersions.length}"><strong>${proj.name}</strong><br><span style="font-size:10px; color:var(--text-muted);">${proj.id}</span></td>`;
      }
      
      let verColHtml = `<td class="sticky-col-2"><strong>${ver.version}</strong> <span style="font-size:10px; color:var(--text-muted);">${ver.category}</span></td>`;
      
      // 渲染 35 週的填色儲存格
      let weekCellsHtml = '';
      
      // 查找預設填色配置
      const pConfig = DEFAULT_GANTT_CELLS[proj.id];
      const vConfig = pConfig ? pConfig[ver.version] : null;
      const activeCols = vConfig ? vConfig.cols : [];
      const cellColor = vConfig ? vConfig.color : '';
      const cellTooltip = vConfig ? `${ver.version}: ${vConfig.desc}` : `${ver.version}: ${ver.change}`;
      
      GANTT_WEEKS.forEach(wk => {
        const isCellActive = activeCols.includes(wk.col);
        if (isCellActive) {
          weekCellsHtml += `<td class="color-${cellColor} gantt-active-cell" data-tooltip="${cellTooltip}"></td>`;
        } else {
          weekCellsHtml += '<td></td>';
        }
      });
      
      tr.innerHTML = projColHtml + verColHtml + weekCellsHtml;
      body.appendChild(tr);
    });
  });
}

// ==========================================================================
// 4. 分頁 3: 專案主檔
// ==========================================================================

function renderProjectMaster() {
  const tbody = document.getElementById('project-table-body');
  tbody.innerHTML = '';
  document.getElementById('project-list-view').classList.remove('hidden');
  document.getElementById('project-detail-view').classList.add('hidden');
  
  const searchVal = document.getElementById('project-search').value.toLowerCase().trim();
  const typeFilter = document.getElementById('project-filter-type').value;
  const statusFilter = document.getElementById('project-filter-status').value;
  
  db.projects.forEach(proj => {
    // 篩選過濾邏輯
    if (searchVal && !proj.id.toLowerCase().includes(searchVal) && 
        !proj.name.toLowerCase().includes(searchVal) && 
        !proj.owner.toLowerCase().includes(searchVal)) {
      return;
    }
    
    if (typeFilter && proj.type !== typeFilter) return;
    if (statusFilter && proj.status !== statusFilter) return;
    
    ensureProjectProcessModel(proj);
    const tr = document.createElement('tr');
    tr.className = 'project-row-clickable';
    tr.tabIndex = 0;
    tr.setAttribute('role', 'button');
    tr.setAttribute('aria-label', `查看 ${proj.name} 專案進度`);
    tr.onclick = () => openProjectDetail(proj.id);
    tr.onkeydown = (event) => {
      if (event.key === 'Enter' || event.key === ' ') {
        event.preventDefault();
        openProjectDetail(proj.id);
      }
    };
    
    // 燈號樣式
    let lightClass = 'badge-gray';
    if (proj.status === '綠') lightClass = 'badge-green';
    else if (proj.status === '黃') lightClass = 'badge-orange';
    else if (proj.status === '紅') lightClass = 'badge-red';
    
    // 協作權限按鈕 (編輯專案)
    let actionBtn = `<button class="btn btn-primary btn-xs" onclick="event.stopPropagation(); openProjectDetail('${proj.id}')">查看進度</button>`;
    if (activePerspective === 'Admin' || activePerspective === 'PM') {
      actionBtn += ` <button class="btn btn-outline btn-xs" onclick="event.stopPropagation(); showEditProjectModal('${proj.id}')">編輯</button>`;
    }

    const stage = PRODUCT_PROCESS_STAGES.find(item => item.id === proj.process.currentStep);
    const assignment = getStageAssignment(proj, proj.process.currentStep);
    const iterationLabel = `${proj.process.phase === '全光度段' ? '全光度' : proj.process.phase === '量產' ? '量產' : '平光'}第 ${proj.process.currentCycle} 輪`;
    const focusText = proj.bottleneck || proj.nextStep || '目前無登錄卡點';
    
    tr.innerHTML = `
      <td>
        <div class="project-cell-title">${escapeHtml(proj.name)}</div>
        <div class="project-cell-meta">${escapeHtml(proj.id)} · ${escapeHtml(proj.type)}</div>
      </td>
      <td>
        <div class="stage-cell"><span class="stage-number">${proj.process.currentStep}</span><strong>${escapeHtml(stage.shortName)}</strong></div>
        <div class="project-cell-meta">${escapeHtml(proj.process.phase)} · ${getProjectProgress(proj)}%</div>
      </td>
      <td><strong>${escapeHtml(proj.currentVersion)}</strong><div class="project-cell-meta">${iterationLabel}</div></td>
      <td><span class="badge ${lightClass}">${proj.status}燈</span></td>
      <td>
        <div class="owner-stack">
          <strong>${escapeHtml(assignment.person)}</strong>
          <span>${escapeHtml(assignment.department)}</span>
          <small>PM追蹤：${escapeHtml(assignment.tracker)}</small>
        </div>
      </td>
      <td>${proj.deadline || '無期限'}</td>
      <td><div class="project-focus-text">${escapeHtml(focusText)}</div><div class="project-cell-meta">下一步：${escapeHtml(proj.nextStep || '待確認')}</div></td>
      <td>${actionBtn}</td>
    `;
    
    tbody.appendChild(tr);
  });
}

function ensureProjectProcessModel(proj) {
  const inferredStep = inferCurrentProcessStep(proj);
  const versions = db && db.versions ? db.versions.filter(v => v.projectId === proj.id) : [];
  const flatVersions = versions.filter(v => `${v.category || ''}${v.testType || ''}`.includes('平光'));
  const fullVersions = versions.filter(v => /全光|試戴/.test(`${v.category || ''}${v.testType || ''}`) && !`${v.category || ''}${v.testType || ''}`.includes('平光'));
  const phase = inferredStep <= 5 ? '平光段' : inferredStep <= 9 ? '全光度段' : '量產';
  const inferredCycle = phase === '平光段'
    ? Math.max(1, flatVersions.length)
    : phase === '全光度段'
      ? Math.max(1, fullVersions.length)
      : 1;

  proj.process = {
    currentStep: inferredStep,
    phase,
    currentCycle: inferredCycle,
    flatCycles: proj.process?.flatCycles || Math.max(1, flatVersions.length),
    fullCycles: proj.process?.fullCycles || Math.max(inferredStep >= 6 ? 1 : 0, fullVersions.length),
    lastDecision: proj.process?.lastDecision || inferLastDecision(proj, versions),
    selectedStage: proj.process?.selectedStage || inferredStep
  };
  return proj.process;
}

function inferCurrentProcessStep(proj) {
  const productionText = `${proj.statusDesc || ''}${proj.currentVersion || ''}`;
  if (!/尚未量產/.test(productionText) && /(已安排生產|生產中|已量產|量產版|正式量產)/.test(productionText)) return 10;
  if (proj.gate === 'G6') return 10;
  const gateMap = { G0: 1, G1: 2, G2: 5, G3: 6, G4: 9, G5: 10 };
  return gateMap[proj.gate] || 1;
}

function inferLastDecision(proj, versions) {
  const latest = versions[0];
  if (proj.status === '紅' || latest?.conclusion === 'NG' || latest?.isProduction === '否') return 'NG';
  if (proj.process?.currentStep >= 6 || /可進|OK|量產/.test(`${latest?.conclusion || ''}${proj.statusDesc || ''}`)) return 'OK';
  return '待判定';
}

function getProjectProgress(proj) {
  return Math.round(Math.max(0, Math.min(10, proj.process.currentStep - (proj.process.currentStep === 10 ? 0 : 1))) / 10 * 100);
}

function openProjectDetail(projectId, scrollToTop = true) {
  const proj = db.projects.find(p => p.id === projectId);
  if (!proj) return;
  ensureProjectProcessModel(proj);
  currentProjectDetailId = projectId;
  currentStageSelection = proj.process.currentStep;

  document.getElementById('project-list-view').classList.add('hidden');
  document.getElementById('project-detail-view').classList.remove('hidden');
  document.getElementById('btn-detail-edit').onclick = () => showEditProjectModal(projectId);

  renderProjectDetailHeader(proj);
  renderProjectProcess(proj);
  renderProjectStageInspector(proj, currentStageSelection);
  renderProjectVersionTimeline(proj);
  renderProjectExecutionFocus(proj);
  applyPermissions();

  if (scrollToTop) document.querySelector('.content-body').scrollTo({ top: 0, behavior: 'smooth' });
}

function closeProjectDetail() {
  currentProjectDetailId = null;
  currentStageSelection = null;
  renderProjectMaster();
}

function renderProjectDetailHeader(proj) {
  const stage = PRODUCT_PROCESS_STAGES.find(item => item.id === proj.process.currentStep);
  const statusClass = proj.status === '紅' ? 'badge-red' : proj.status === '綠' ? 'badge-green' : 'badge-orange';
  const guidance = getProjectGuidance(proj);
  const assignment = getStageAssignment(proj, proj.process.currentStep);
  document.getElementById('project-detail-header').innerHTML = `
    <div class="detail-title-block">
      <div class="detail-title-row">
        <div>
          <div class="project-detail-id">${escapeHtml(proj.id)} · ${escapeHtml(proj.type)}</div>
          <h1>${escapeHtml(proj.name)}</h1>
        </div>
        <span class="badge ${statusClass} detail-status-badge">${proj.status}燈</span>
      </div>
      <div class="detail-current-stage">
        <span>目前進度</span>
        <strong>${proj.process.currentStep}. ${escapeHtml(stage.name)}</strong>
      </div>
      <div class="detail-next-guide ${guidance.tone}">
        <div>
          <span>系統建議下一步</span>
          <strong>${escapeHtml(guidance.title)}</strong>
          <p>${escapeHtml(guidance.nextText)}</p>
        </div>
        <button class="btn ${guidance.tone === 'danger' ? 'btn-primary' : 'btn-outline'} btn-sm" onclick="selectProjectStage(${guidance.backTo || proj.process.currentStep})">
          ${escapeHtml(guidance.cta)}
        </button>
      </div>
    </div>
    <div class="detail-metric-grid">
      <div><span>目前版本</span><strong>${escapeHtml(proj.currentVersion)}</strong></div>
      <div><span>階段／迭代</span><strong>${escapeHtml(proj.process.phase)} · 第 ${proj.process.currentCycle} 輪</strong></div>
      <div><span>負責部門</span><strong>${escapeHtml(assignment.department)}</strong></div>
      <div><span>階段負責人</span><strong>${escapeHtml(assignment.person)}</strong></div>
      <div><span>PM追蹤者</span><strong>${escapeHtml(assignment.tracker)}</strong></div>
      <div><span>下一期限</span><strong>${escapeHtml(proj.deadline || '待確認')}</strong></div>
    </div>
    <div class="detail-progress-track"><span style="width:${getProjectProgress(proj)}%"></span></div>
  `;
}

function renderProjectProcess(proj) {
  const flow = document.getElementById('project-process-flow');
  const groups = [
    { title: '需求確認', subtitle: '產品定義與開發起點', stages: [1] },
    { title: '平光段', subtitle: `多版本迭代 · 已記錄 ${proj.process.flatCycles} 輪`, stages: [2, 3, 4, 5], decision: { label: '平光測試判定', backTo: 2 } },
    { title: '全光度段', subtitle: `焦度設計與驗證 · 已記錄 ${proj.process.fullCycles} 輪`, stages: [6, 7, 8, 9], decision: { label: '全光度測試判定', backTo: 6 } },
    { title: '量產', subtitle: '良率確認與正式放行', stages: [10] }
  ];

  flow.innerHTML = groups.map(group => `
    <section class="process-phase process-phase-${group.stages[0]}">
      <div class="process-phase-label">
        <strong>${group.title}</strong>
        <span>${group.subtitle}</span>
      </div>
      <div class="process-stage-row">
        ${group.stages.map(stageId => renderProcessStageNode(proj, stageId)).join('<span class="process-arrow" aria-hidden="true">→</span>')}
      </div>
      ${group.decision ? `
        <div class="process-decision ${getDecisionClass(proj, group.stages[group.stages.length - 1])}">
          <span>${group.decision.label}</span>
          <strong>${getDecisionText(proj, group.stages[group.stages.length - 1])}</strong>
          <small>NG 回到第 ${group.decision.backTo} 階段，舊版本紀錄保留</small>
        </div>
      ` : ''}
    </section>
  `).join('');
}

function renderProcessStageNode(proj, stageId) {
  const stage = PRODUCT_PROCESS_STAGES.find(item => item.id === stageId);
  const guidance = getProjectGuidance(proj);
  const assignment = getStageAssignment(proj, stageId);
  let state = 'pending';
  if (stageId < proj.process.currentStep) state = 'done';
  if (stageId === proj.process.currentStep) state = proj.status === '紅' ? 'blocked' : 'current';
  if (proj.process.currentStep === 10 && stageId === 10 && proj.status === '綠') state = 'done';
  const selected = currentStageSelection === stageId ? ' selected' : '';
  const nextTarget = guidance.backTo === stageId ? ' next-target' : '';
  return `
    <button class="process-stage ${state}${selected}${nextTarget}" onclick="selectProjectStage(${stageId})">
      <span class="process-stage-number">${stageId}</span>
      <span class="process-stage-copy">
        <strong>${escapeHtml(stage.shortName)}</strong>
        <small>${escapeHtml(assignment.department)}</small>
        <small>${escapeHtml(assignment.person)}</small>
      </span>
      <span class="process-stage-state">${guidance.backTo === stageId ? '下一步回到這裡' : getStageStateLabel(state)}</span>
    </button>
  `;
}

function getStageStateLabel(state) {
  return { done: '完成', current: '進行中', blocked: 'NG／卡關', pending: '未開始' }[state];
}

function getDecisionClass(proj, finalStage) {
  if (proj.process.currentStep > finalStage) return 'decision-ok';
  if (proj.process.currentStep === finalStage && proj.status === '紅') return 'decision-ng';
  return 'decision-waiting';
}

function getDecisionText(proj, finalStage) {
  if (proj.process.currentStep > finalStage) return 'OK';
  if (proj.process.currentStep === finalStage && proj.status === '紅') return 'NG';
  return '待判定';
}

function selectProjectStage(stageId) {
  const proj = db.projects.find(p => p.id === currentProjectDetailId);
  if (!proj) return;
  currentStageSelection = stageId;
  renderProjectProcess(proj);
  renderProjectStageInspector(proj, stageId);
}

function renderProjectStageInspector(proj, stageId) {
  const stage = PRODUCT_PROCESS_STAGES.find(item => item.id === stageId);
  const versions = getVersionsForStage(proj.id, stageId);
  const assignment = getStageAssignment(proj, stageId);
  let state = '未開始';
  if (stageId < proj.process.currentStep) state = '已完成';
  if (stageId === proj.process.currentStep) state = proj.status === '紅' ? 'NG／卡關' : '進行中';
  if (proj.process.currentStep === 10 && stageId === 10 && proj.status === '綠') state = '已完成';

  document.getElementById('project-stage-inspector').innerHTML = `
    <div class="inspector-kicker">階段 ${stage.id}</div>
    <h2>${escapeHtml(stage.name)}</h2>
    <div class="inspector-state state-${state === '已完成' ? 'done' : state === '進行中' ? 'current' : state.includes('NG') ? 'blocked' : 'pending'}">${state}</div>
    <p class="inspector-description">${escapeHtml(stage.description)}</p>
    <dl class="inspector-facts">
      <div><dt>負責部門</dt><dd>${escapeHtml(assignment.department)}</dd></div>
      <div><dt>階段負責人</dt><dd>${escapeHtml(assignment.person)}</dd></div>
      <div><dt>PM追蹤者</dt><dd>${escapeHtml(assignment.tracker)}</dd></div>
      <div><dt>目前版本</dt><dd>${escapeHtml(proj.currentVersion)}</dd></div>
      <div><dt>相關紀錄</dt><dd>${versions.length} 筆</dd></div>
      <div><dt>最後更新</dt><dd>${escapeHtml(proj.lastUpdate || '待確認')}</dd></div>
    </dl>
    <div class="inspector-records">
      <h3>此階段紀錄</h3>
      ${versions.length ? versions.slice(0, 4).map(version => `
        <article>
          <strong>${escapeHtml(version.version)} · ${escapeHtml(version.category || '版本')}</strong>
          <span>${escapeHtml(version.date || '未填日期')}</span>
          <p>${escapeHtml(version.conclusion || version.change || '尚無結論')}</p>
        </article>
      `).join('') : '<p class="empty-stage-record">尚無版本或測試紀錄。</p>'}
    </div>
  `;
}

function getVersionsForStage(projectId, stageId) {
  const versions = db.versions.filter(version => version.projectId === projectId);
  if ([2, 3, 4, 5].includes(stageId)) {
    return versions.filter(version => /平光|PP|製程|臨床|試戴/.test(`${version.category || ''}${version.testType || ''}`));
  }
  if ([6, 7, 8, 9].includes(stageId)) {
    return versions.filter(version => /全光|焦度|臨床|試戴/.test(`${version.category || ''}${version.testType || ''}`) && !`${version.category || ''}${version.testType || ''}`.includes('平光'));
  }
  if (stageId === 10) return versions.filter(version => /量產|生產|切片/.test(`${version.category || ''}${version.testType || ''}`));
  return [];
}

function renderProjectVersionTimeline(proj) {
  const versions = db.versions.filter(version => version.projectId === proj.id);
  document.getElementById('detail-version-count').innerText = `${versions.length} 筆`;
  document.getElementById('project-version-timeline').innerHTML = versions.length ? versions.map((version, index) => {
    const resultClass = version.conclusion === 'NG' || version.isProduction === '否' ? 'ng' : /OK|可進|量產/.test(`${version.conclusion || ''}${version.isProduction || ''}`) ? 'ok' : 'waiting';
    return `
      <article class="version-timeline-item ${resultClass}">
        <div class="version-timeline-marker"></div>
        <div class="version-timeline-content">
          <div class="version-timeline-head">
            <strong>${escapeHtml(version.version)} · ${escapeHtml(version.category || '版本紀錄')}</strong>
            <span>${escapeHtml(version.date || '日期待確認')}</span>
          </div>
          <p>${escapeHtml(version.change || version.testType || '無變更說明')}</p>
          <div class="version-result-row">
            <span>${escapeHtml(version.conclusion || '待判定')}</span>
            <small>${escapeHtml(version.owner || '')}</small>
          </div>
        </div>
      </article>
    `;
  }).join('') : '<div class="empty-timeline">尚無版本紀錄。</div>';
}

function renderProjectExecutionFocus(proj) {
  const overdue = proj.deadline && new Date(proj.deadline) < new Date(new Date().toISOString().slice(0, 10));
  document.getElementById('project-execution-focus').innerHTML = `
    <div class="focus-block focus-risk">
      <span>目前卡點</span>
      <strong>${escapeHtml(proj.bottleneck || '目前無登錄卡點')}</strong>
    </div>
    <div class="focus-block focus-next">
      <span>下一步</span>
      <strong>${escapeHtml(proj.nextStep || '待 PM 確認')}</strong>
    </div>
    <div class="focus-meta-grid">
      <div><span>期限狀態</span><strong class="${overdue ? 'text-red' : ''}">${overdue ? '已逾期' : '追蹤中'}</strong></div>
      <div><span>協作人員</span><strong>${escapeHtml(proj.collaborators || '待指派')}</strong></div>
      <div><span>信心程度</span><strong>${escapeHtml(proj.confidence || '待確認')}</strong></div>
      <div><span>最近更新</span><strong>${escapeHtml(proj.lastUpdate || '待確認')}</strong></div>
    </div>
    ${proj.note ? `<div class="focus-note"><span>補充說明</span><p>${escapeHtml(proj.note)}</p></div>` : ''}
  `;
}

function openVersionForCurrentProject() {
  if (!currentProjectDetailId) return;
  renderVersionHistory();
  showAddVersionModal();
  document.getElementById('ver-project-id').value = currentProjectDetailId;
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

// 專案 Modal 控制
function showAddProjectModal() {
  document.getElementById('project-modal-title').innerText = '新增開發專案';
  document.getElementById('project-form').reset();
  document.getElementById('project-id').disabled = false;
  openModal('modal-project');
}

function showEditProjectModal(projectId) {
  const proj = db.projects.find(p => p.id === projectId);
  if (!proj) return;
  
  document.getElementById('project-modal-title').innerText = `編輯專案: ${proj.id}`;
  
  // 填寫欄位
  document.getElementById('project-id').value = proj.id;
  document.getElementById('project-id').disabled = true; // 不允許修改 ID
  document.getElementById('project-name').value = proj.name;
  document.getElementById('project-type').value = proj.type;
  document.getElementById('project-owner').value = proj.owner;
  document.getElementById('project-collaborators').value = proj.collaborators || '';
  document.getElementById('project-priority').value = proj.priority;
  document.getElementById('project-status').value = proj.status;
  ensureProjectProcessModel(proj);
  document.getElementById('project-gate').value = String(proj.process.currentStep);
  document.getElementById('project-deadline').value = proj.deadline || '';
  document.getElementById('project-status-desc').value = proj.statusDesc || '';
  document.getElementById('project-bottleneck').value = proj.bottleneck || '';
  document.getElementById('project-next-step').value = proj.nextStep || '';
  
  openModal('modal-project');
}

function submitProjectForm() {
  const id = document.getElementById('project-id').value.trim();
  const name = document.getElementById('project-name').value.trim();
  const type = document.getElementById('project-type').value;
  const owner = document.getElementById('project-owner').value.trim();
  const collaborators = document.getElementById('project-collaborators').value.trim();
  const priority = parseInt(document.getElementById('project-priority').value);
  const status = document.getElementById('project-status').value;
  const currentStep = parseInt(document.getElementById('project-gate').value);
  const gate = stepToLegacyGate(currentStep);
  const deadline = document.getElementById('project-deadline').value;
  const statusDesc = document.getElementById('project-status-desc').value.trim();
  const bottleneck = document.getElementById('project-bottleneck').value.trim();
  const nextStep = document.getElementById('project-next-step').value.trim();
  
  const existingIdx = db.projects.findIndex(p => p.id === id);
  
  if (existingIdx !== -1) {
    // 編輯舊專案
    db.projects[existingIdx] = {
      ...db.projects[existingIdx],
      name, type, owner, collaborators, priority, status, gate, deadline, statusDesc, bottleneck, nextStep,
      process: {
        ...db.projects[existingIdx].process,
        currentStep,
        phase: currentStep <= 5 ? '平光段' : currentStep <= 9 ? '全光度段' : '量產',
        selectedStage: currentStep
      },
      lastUpdate: new Date().toISOString().slice(0,10)
    };
  } else {
    // 建立新專案
    const newProj = {
      id, name, type, currentVersion: "V1（新立案）", priority, gate, status, bottleneck, nextStep, owner, collaborators,
      deadline, statusDesc, lastUpdate: new Date().toISOString().slice(0,10), confidence: "中", note: "",
      process: {
        currentStep,
        phase: currentStep <= 5 ? '平光段' : currentStep <= 9 ? '全光度段' : '量產',
        currentCycle: 1,
        flatCycles: 1,
        fullCycles: currentStep >= 6 ? 1 : 0,
        lastDecision: '待判定',
        selectedStage: currentStep
      },
      milestones: [
        { "name": "G0 規格凍結", "target": deadline, "actual": "", "status": "blue" },
        { "name": "G1 平光設計", "target": "", "actual": "", "status": "gray" },
        { "name": "G2 平光測試", "target": "", "actual": "", "status": "gray" },
        { "name": "G3 全光設計", "target": "", "actual": "", "status": "gray" },
        { "name": "G4 全光試戴", "target": "", "actual": "", "status": "gray" },
        { "name": "G5 試量產", "target": "", "actual": "", "status": "gray" },
        { "name": "G6 正式量產", "target": "", "actual": "", "status": "gray" }
      ]
    };
    db.projects.push(newProj);
    
    // 初始化新專案的甘特圖預設填色配置，防止甘特圖出錯
    DEFAULT_GANTT_CELLS[id] = {
      "V1（新立案）": { cols: ["E", "F", "G"], color: "orange", desc: "新立案起跑" }
    };
  }
  
  saveDatabase();
  closeModal('modal-project');
  renderProjectMaster();
}

function stepToLegacyGate(step) {
  if (step <= 1) return 'G0';
  if (step <= 4) return 'G1';
  if (step === 5) return 'G2';
  if (step <= 8) return 'G3';
  if (step === 9) return 'G4';
  return 'G5';
}

// ==========================================================================
// 5. 分頁 4: 版本履歷與試戴分析匯入
// ==========================================================================

function renderVersionHistory() {
  const tbody = document.getElementById('version-table-body');
  tbody.innerHTML = '';
  
  // 載入專案 ID 到篩選下拉選單與新增 version 下拉選單中
  const filterSelect = document.getElementById('version-filter-project');
  const formSelect = document.getElementById('ver-project-id');
  const importSelect = document.getElementById('import-assoc-project');
  const actProjSelect = document.getElementById('act-project-id');
  const riskProjSelect = document.getElementById('risk-project-id');
  
  const savedFilterVal = filterSelect.value;
  
  filterSelect.innerHTML = '<option value="">所有專案</option>';
  formSelect.innerHTML = '';
  importSelect.innerHTML = '';
  actProjSelect.innerHTML = '';
  riskProjSelect.innerHTML = '';
  
  db.projects.forEach(p => {
    filterSelect.innerHTML += `<option value="${p.id}">${p.name} (${p.id})</option>`;
    formSelect.innerHTML += `<option value="${p.id}">${p.name} (${p.id})</option>`;
    importSelect.innerHTML += `<option value="${p.id}">${p.name} (${p.id})</option>`;
    actProjSelect.innerHTML += `<option value="${p.id}">${p.name} (${p.id})</option>`;
    riskProjSelect.innerHTML += `<option value="${p.id}">${p.name} (${p.id})</option>`;
  });
  
  filterSelect.value = savedFilterVal;
  
  const projectFilter = filterSelect.value;
  const prodFilter = document.getElementById('version-filter-production').value;
  
  db.versions.forEach(ver => {
    if (projectFilter && ver.projectId !== projectFilter) return;
    if (prodFilter && ver.isProduction !== prodFilter) return;
    
    const tr = document.createElement('tr');
    
    // 可量產判定配色
    let prodBadgeClass = 'badge-gray';
    if (ver.isProduction === '是') prodBadgeClass = 'badge-green';
    else if (ver.isProduction === '否') prodBadgeClass = 'badge-red';
    else if (ver.isProduction === '條件式可用') prodBadgeClass = 'badge-orange';
    
    tr.innerHTML = `
      <td><strong>${ver.projectId}</strong></td>
      <td><strong>${ver.version}</strong></td>
      <td><span class="badge badge-gray">${ver.category}</span></td>
      <td>${ver.change}</td>
      <td>${ver.testType || '無'}</td>
      <td>${ver.sampleSize || '無'}</td>
      <td>${ver.porosity ? `${parseFloat(ver.porosity) * 100}%` : '無'}</td>
      <td>${ver.comfort || '無'}</td>
      <td>${ver.defect || '無'}</td>
      <td>${ver.conclusion || '無'}</td>
      <td>${ver.nextStep || '無'}</td>
      <td>${ver.owner}</td>
      <td><span class="badge ${prodBadgeClass}">${ver.isProduction}</span></td>
      <td><span style="font-size:11px; color:var(--text-muted);">${ver.source}</span></td>
    `;
    tbody.appendChild(tr);
  });
}

function showAddVersionModal() {
  document.getElementById('version-form').reset();
  openModal('modal-version');
}

function submitVersionForm() {
  const projectId = document.getElementById('ver-project-id').value;
  const version = document.getElementById('ver-name').value.trim();
  const category = document.getElementById('ver-category').value;
  const testType = document.getElementById('ver-test-type').value.trim();
  const sampleSize = document.getElementById('ver-sample-size').value;
  const porosityVal = document.getElementById('ver-porosity').value.trim();
  const comfort = document.getElementById('ver-comfort').value.trim();
  const owner = document.getElementById('ver-owner').value.trim();
  const date = document.getElementById('ver-date').value;
  const isProduction = document.getElementById('ver-production').value;
  const change = document.getElementById('ver-change').value.trim();
  const defect = document.getElementById('ver-defect').value.trim();
  const conclusion = document.getElementById('ver-conclusion').value.trim();
  const nextStep = document.getElementById('ver-next-step').value.trim();
  const source = document.getElementById('ver-source').value.trim();
  
  // 處理氣孔率百分比轉換
  let porosity = "";
  if (porosityVal) {
    if (porosityVal.includes('%')) {
      porosity = (parseFloat(porosityVal.replace('%','')) / 100).toString();
    } else {
      porosity = parseFloat(porosityVal).toString();
    }
  }
  
  const newVer = {
    projectId, version, category, change, testType, sampleSize, porosity, comfort, defect, conclusion,
    nextStep, owner, date, source, isProduction
  };
  
  db.versions.unshift(newVer); // 最新的排在最前面
  
  // 同步更新專案主檔的目前版本名稱
  const proj = db.projects.find(p => p.id === projectId);
  if (proj) {
    proj.currentVersion = version;
    proj.lastUpdate = new Date().toISOString().slice(0,10);
  }
  
  // 更新甘特圖配置 (模擬新增版本之甘特列，填上最右側的5週)
  if (!DEFAULT_GANTT_CELLS[projectId]) DEFAULT_GANTT_CELLS[projectId] = {};
  DEFAULT_GANTT_CELLS[projectId][version] = {
    cols: ["AI", "AJ", "AK", "AL", "AM"],
    color: isProduction === '是' ? 'green' : (isProduction === '否' ? 'red' : 'orange'),
    desc: `手動新增版本：${change}`
  };
  
  saveDatabase();
  closeModal('modal-version');
  renderVersionHistory();
}

// 試戴報告問卷拖放/匯入控制
function showImportModal() {
  document.getElementById('analysis-result-panel').classList.add('hidden');
  document.getElementById('import-file-input').value = '';
  currentUploadedData = null;
  openModal('modal-import');
}

// 檔案處理
function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;
  processFile(file);
}

// 拖放區事件設定
document.addEventListener('DOMContentLoaded', () => {
  const dropZone = document.getElementById('drop-zone');
  if (!dropZone) return;
  
  dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('dragover');
  });
  
  dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('dragover');
  });
  
  dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file) {
      processFile(file);
    }
  });
});

function processFile(file) {
  const reader = new FileReader();
  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    try {
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(sheet);
      
      if (json.length === 0) {
        alert("Excel 檔案為空，請檢查！");
        return;
      }
      
      analyzeTrialData(json);
    } catch (err) {
      console.error(err);
      alert("檔案讀取失敗！請確認格式是否正確，或下載本系統提供的範本檔案。");
    }
  };
  reader.readAsArrayBuffer(file);
}

// 試戴資料分析引擎
function analyzeTrialData(rows) {
  let sampleSize = rows.length;
  let totalComfort = 0;
  let comfortCount = 0;
  let okCentrationCount = 0;
  let centrationCount = 0;
  
  let defectsMap = {};
  
  rows.forEach(r => {
    // 1. 抓取舒適度 (對應 舒適度, comfort, score 等字眼)
    const comfortKey = Object.keys(r).find(k => k.includes('舒適') || k.toLowerCase().includes('comfort'));
    if (comfortKey) {
      const val = parseFloat(r[comfortKey]);
      if (!isNaN(val)) {
        totalComfort += val;
        comfortCount++;
      }
    }
    
    // 2. 抓取定位結果 (對應 定位, centration 等字眼)
    const centKey = Object.keys(r).find(k => k.includes('定位') || k.toLowerCase().includes('centration'));
    if (centKey) {
      centrationCount++;
      const val = r[centKey].toString().trim().toUpperCase();
      if (val === 'OK' || val === '是' || val === '良好' || val === '1' || val === 'TRUE') {
        okCentrationCount++;
      }
    }
    
    // 3. 抓取意見反饋與缺陷字詞
    const commentKey = Object.keys(r).find(k => k.includes('意見') || k.includes('反饋') || k.includes('備註') || k.toLowerCase().includes('comment') || k.toLowerCase().includes('feedback'));
    if (commentKey && r[commentKey]) {
      const txt = r[commentKey].toString();
      const keywords = ["漏白", "氣孔", "掉色", "硬痕", "印痕", "突出", "模糊", "位移"];
      keywords.forEach(kw => {
        if (txt.includes(kw)) {
          defectsMap[kw] = (defectsMap[kw] || 0) + 1;
        }
      });
    }
  });
  
  const avgComfort = comfortCount > 0 ? (totalComfort / comfortCount).toFixed(1) : "無數據";
  const centRate = centrationCount > 0 ? ((okCentrationCount / centrationCount) * 100).toFixed(0) : "無數據";
  
  // 4. 解析為分析物件
  currentUploadedData = {
    sampleSize: sampleSize,
    avgComfort: avgComfort,
    centRate: centRate,
    defects: defectsMap,
    rawPorosity: 0.045 // 模擬由問卷關聯或Excel中附帶的平均氣孔率
  };
  
  // 5. 渲染看板介面
  document.getElementById('stat-sample-size').innerText = sampleSize;
  document.getElementById('stat-avg-comfort').innerText = avgComfort;
  document.getElementById('stat-avg-centration').innerText = centRate === '無數據' ? '無' : `${centRate}%`;
  
  // 渲染缺陷標籤
  const tagsContainer = document.getElementById('defect-tags-container');
  tagsContainer.innerHTML = '';
  const defectKeys = Object.keys(defectsMap);
  if (defectKeys.length === 0) {
    tagsContainer.innerHTML = '<span class="text-muted" style="font-size:12px;">未檢出明顯缺陷字句。</span>';
  } else {
    defectKeys.forEach(k => {
      const pct = ((defectsMap[k] / sampleSize) * 100).toFixed(0);
      tagsContainer.innerHTML += `<span class="defect-tag">${k} (${pct}%)</span>`;
    });
  }
  
  // 6. Go/NG 決策演算法判定
  const aiBox = document.getElementById('ai-rec-box');
  const badge = document.getElementById('rec-badge-type');
  const reason = document.getElementById('rec-badge-reason');
  
  let isGo = true;
  let reasonText = [];
  
  if (parseFloat(avgComfort) < 4.0) {
    isGo = false;
    reasonText.push(`舒適度評分均值為 ${avgComfort}，低於品保量產門檻 (4.0)。`);
  }
  
  if (parseFloat(centRate) < 90) {
    isGo = false;
    reasonText.push(`定位合格率僅 ${centRate}%，未達要求門檻 (90%)。`);
  }
  
  if (defectsMap["掉色"]) {
    isGo = false;
    reasonText.push(`受試者意見中包含「掉色」反饋，此為臨床安全性紅色阻礙，必須釐清油墨。`);
  }
  
  if (defectsMap["突出"]) {
    isGo = false;
    reasonText.push(`有 ${defectsMap["突出"]} 筆回饋提到旋轉「突出」，可能引發配戴不適，應評估薄區厚度。`);
  }
  
  if (isGo) {
    aiBox.className = 'ai-recommendation rec-go';
    badge.innerText = '建議通過 (Go)';
    reason.innerText = `各項指標優良（舒適度均分 ${avgComfort}，定位 ${centRate}% 良好），且無檢出掉色、硬痕等嚴重安全性缺陷，建議可進入下個 Gate 階段審核。`;
    currentUploadedData.isProduction = '是';
    currentUploadedData.conclusion = '試戴合格，數據均達標';
  } else {
    aiBox.className = 'ai-recommendation rec-ng';
    badge.innerText = '建議退回 (NG)';
    reason.innerText = reasonText.join(" ");
    currentUploadedData.isProduction = '否';
    currentUploadedData.conclusion = `試戴不通過：${reasonText.join(' ')}`;
  }
  
  // 顯示右側分析結果
  document.getElementById('analysis-result-panel').classList.remove('hidden');
}

// 儲存匯入的分析結果
function saveImportedReport() {
  if (!currentUploadedData) return;
  
  const projectId = document.getElementById('import-assoc-project').value;
  const version = document.getElementById('import-assoc-version').value.trim();
  
  if (!version) {
    alert("請輸入版本代號！");
    return;
  }
  
  // 整理缺陷描述
  const defectKeys = Object.keys(currentUploadedData.defects);
  const defectStr = defectKeys.length > 0 
    ? defectKeys.map(k => `${k}(共${currentUploadedData.defects[k]}次)`).join(', ')
    : '無明顯缺陷';
  
  const newVer = {
    projectId: projectId,
    version: version,
    category: "試戴版",
    change: `試戴資料匯入 (樣本數: ${currentUploadedData.sampleSize})`,
    testType: "臨床試戴",
    sampleSize: currentUploadedData.sampleSize.toString(),
    porosity: currentUploadedData.rawPorosity.toString(),
    comfort: `舒適:${currentUploadedData.avgComfort} / 定位:${currentUploadedData.centRate}%`,
    defect: defectStr,
    conclusion: currentUploadedData.conclusion,
    nextStep: currentUploadedData.isProduction === '是' ? '申請 Gate 階段通過簽核' : '研發修改邊緣及厚度',
    owner: currentUser.name,
    date: new Date().toISOString().slice(0, 10),
    source: "自動解析 Excel 問卷資料",
    isProduction: currentUploadedData.isProduction
  };
  
  db.versions.unshift(newVer);
  
  // 更新專案主檔目前版本
  const proj = db.projects.find(p => p.id === projectId);
  if (proj) {
    proj.currentVersion = version;
    proj.lastUpdate = new Date().toISOString().slice(0,10);
  }
  
  // 更新甘特圖
  if (!DEFAULT_GANTT_CELLS[projectId]) DEFAULT_GANTT_CELLS[projectId] = {};
  DEFAULT_GANTT_CELLS[projectId][version] = {
    cols: ["AI", "AJ", "AK", "AL", "AM"],
    color: currentUploadedData.isProduction === '是' ? 'green' : 'red',
    desc: `匯入試戴報告：舒適度 ${currentUploadedData.avgComfort}`
  };
  
  saveDatabase();
  closeModal('modal-import');
  
  // 切換至版本履歷分頁以看見結果
  const menuHistory = document.getElementById('nav-version-history');
  menuHistory.click();
}

// ==========================================================================
// 6. 分頁 5: 行動項目 Kanban Board
// ==========================================================================

function renderActionItems() {
  const cardsTodo = document.getElementById('cards-todo');
  const cardsInprogress = document.getElementById('cards-inprogress');
  const cardsCompleted = document.getElementById('cards-completed');
  
  cardsTodo.innerHTML = '';
  cardsInprogress.innerHTML = '';
  cardsCompleted.innerHTML = '';
  
  let todoCount = 0;
  let ipCount = 0;
  let compCount = 0;
  
  db.actions.forEach(act => {
    const card = document.createElement('div');
    card.className = 'kanban-card';
    
    // 檢查是否逾期
    const isOverdue = act.deadline && new Date(act.deadline) < new Date() && act.status !== '已完成';
    const deadlineClass = isOverdue ? 'kanban-card-deadline overdue' : 'kanban-card-deadline';
    
    // 生成移動控制按鈕
    let moveButtons = '';
    
    if (activePerspective === 'Admin' || activePerspective === 'PM') {
      if (act.status === '未開始') {
        moveButtons = `<button class="btn btn-outline btn-xs" onclick="moveAction('${act.id}', '進行中')">開始執行 ➡️</button>`;
      } else if (act.status === '進行中') {
        moveButtons = `
          <button class="btn btn-outline btn-xs" onclick="moveAction('${act.id}', '未開始')">⬅️ 退回</button>
          <button class="btn btn-primary btn-xs ml-10" onclick="moveAction('${act.id}', '已完成')">完成 ➡️</button>
        `;
      } else if (act.status === '已完成') {
        moveButtons = `<button class="btn btn-outline btn-xs" onclick="moveAction('${act.id}', '進行中')">⬅️ 重啟任務</button>`;
      }
    } else if (activePerspective === 'RD' && (act.owner === currentUser.name || act.coowners.includes(currentUser.name))) {
      // 研發角色僅能更改自己的任務狀態
      if (act.status === '未開始') {
        moveButtons = `<button class="btn btn-outline btn-xs" onclick="moveAction('${act.id}', '進行中')">開始執行 ➡️</button>`;
      } else if (act.status === '進行中') {
        moveButtons = `<button class="btn btn-primary btn-xs" onclick="moveAction('${act.id}', '已完成')">完成 ➡️</button>`;
      }
    }
    
    card.innerHTML = `
      <div class="kanban-card-project">${act.projectId}</div>
      <div class="kanban-card-title">${act.title}</div>
      <div class="kanban-card-desc">${act.desc}</div>
      <div class="kanban-card-footer">
        <span class="kanban-card-owner">${act.owner}</span>
        <span class="${deadlineClass}">${act.deadline || '無期限'}</span>
      </div>
      <div class="kanban-card-actions">
        ${moveButtons}
      </div>
    `;
    
    if (act.status === '未開始') {
      cardsTodo.appendChild(card);
      todoCount++;
    } else if (act.status === '進行中') {
      cardsInprogress.appendChild(card);
      ipCount++;
    } else if (act.status === '已完成') {
      cardsCompleted.appendChild(card);
      compCount++;
    }
  });
  
  document.getElementById('count-todo').innerText = todoCount;
  document.getElementById('count-inprogress').innerText = ipCount;
  document.getElementById('count-completed').innerText = compCount;
}

function moveAction(actionId, targetStatus) {
  const act = db.actions.find(a => a.id === actionId);
  if (act) {
    act.status = targetStatus;
    saveDatabase();
    renderActionItems();
  }
}

function showAddActionModal() {
  document.getElementById('action-form').reset();
  openModal('modal-action');
}

function submitActionForm() {
  const projectId = document.getElementById('act-project-id').value;
  const title = document.getElementById('act-title').value.trim();
  const desc = document.getElementById('act-desc').value.trim();
  const owner = document.getElementById('act-owner').value.trim();
  const coowners = document.getElementById('act-coowners').value.trim();
  const priority = document.getElementById('act-priority').value;
  const deadline = document.getElementById('act-deadline').value;
  const notes = document.getElementById('act-notes').value.trim();
  
  // 計算流水編號
  const lastNum = db.actions.length > 0 
    ? parseInt(db.actions[db.actions.length - 1].id.split('-')[1])
    : 0;
  const newId = `A-${(lastNum + 1).toString().padStart(3, '0')}`;
  
  const newAct = {
    id: newId, projectId, title, desc, owner, coowners, priority, deadline, status: "未開始", notes
  };
  
  db.actions.push(newAct);
  saveDatabase();
  closeModal('modal-action');
  renderActionItems();
}

// ==========================================================================
// 7. 分頁 6: 風險與待確認
// ==========================================================================

function renderRisksIssues() {
  const tbody = document.getElementById('risk-table-body');
  tbody.innerHTML = '';
  
  db.risks.forEach(risk => {
    const tr = document.createElement('tr');
    
    let lightClass = 'badge-gray';
    if (risk.light === '綠') lightClass = 'badge-green';
    else if (risk.light === '黃') lightClass = 'badge-orange';
    else if (risk.light === '紅') lightClass = 'badge-red';
    
    let actionButtons = '';
    if (activePerspective === 'Admin' || activePerspective === 'PM') {
      actionButtons = `
        <button class="btn btn-outline btn-xs" onclick="closeRisk('${risk.id}')">關閉</button>
      `;
    } else {
      actionButtons = `<span style="font-size:11px; color:var(--text-muted);">無編輯權</span>`;
    }
    
    tr.innerHTML = `
      <td><strong>${risk.id}</strong></td>
      <td><span class="badge ${risk.type === 'Issue' ? 'badge-red' : 'badge-orange'}">${risk.type}</span></td>
      <td><strong>${risk.projectId}</strong></td>
      <td><strong>${risk.title}</strong><br><span style="font-size:11.5px; color:var(--text-muted);">${risk.desc}</span></td>
      <td>${risk.probability}</td>
      <td>${risk.impact}</td>
      <td><strong>${risk.score}</strong></td>
      <td><span class="badge ${lightClass}">${risk.light}燈</span></td>
      <td>${risk.mitigation || '待規劃'}</td>
      <td><strong>${risk.owner}</strong></td>
      <td>${risk.deadline || '無期限'}</td>
      <td><span class="badge badge-gray">${risk.status}</span></td>
      <td>${actionButtons}</td>
    `;
    
    tbody.appendChild(tr);
  });
}

function closeRisk(riskId) {
  const risk = db.risks.find(r => r.id === riskId);
  if (risk) {
    risk.status = '已關閉';
    saveDatabase();
    renderRisksIssues();
  }
}

function showAddRiskModal() {
  document.getElementById('risk-form').reset();
  document.getElementById('risk-score-display').value = '1';
  openModal('modal-risk');
}

function calculateRiskScore() {
  const p = parseInt(document.getElementById('risk-prob').value);
  const i = parseInt(document.getElementById('risk-impact').value);
  const score = p * i;
  
  document.getElementById('risk-score-display').value = score;
}

function submitRiskForm() {
  const type = document.getElementById('risk-type').value;
  const projectId = document.getElementById('risk-project-id').value;
  const title = document.getElementById('risk-title').value.trim();
  const desc = document.getElementById('risk-desc').value.trim();
  const probability = parseInt(document.getElementById('risk-prob').value);
  const impact = parseInt(document.getElementById('risk-impact').value);
  const score = probability * impact;
  const owner = document.getElementById('risk-owner').value.trim();
  const deadline = document.getElementById('risk-deadline').value;
  const status = document.getElementById('risk-status').value;
  const mitigation = document.getElementById('risk-mitigation').value.trim();
  
  // 計算燈號
  let light = "綠";
  if (score >= 15) light = "紅";
  else if (score >= 8) light = "黃";
  
  // 計算流水號
  const lastNum = db.risks.length > 0 
    ? parseInt(db.risks[db.risks.length - 1].id.split('-')[1])
    : 0;
  const newId = `RI-${(lastNum + 1).toString().padStart(3, '0')}`;
  
  const newRisk = {
    id: newId, type, projectId, title, desc, probability, impact, score, light, mitigation, owner, deadline, status
  };
  
  db.risks.push(newRisk);
  saveDatabase();
  closeModal('modal-risk');
  renderRisksIssues();
}

// ==========================================================================
// 8. 分頁 7: 系統管理員面板
// ==========================================================================

function renderAdminSettings() {
  const tbody = document.getElementById('user-table-body');
  tbody.innerHTML = '';
  
  db.users.forEach(u => {
    const tr = document.createElement('tr');
    
    // 生成角色編輯的下拉選單
    const options = `
      <select class="form-control-sm" onchange="updateUserRole('${u.username}', this.value)" style="width: 100px;">
        <option value="Admin" ${u.role === 'Admin' ? 'selected' : ''}>Admin</option>
        <option value="PM" ${u.role === 'PM' ? 'selected' : ''}>PM</option>
        <option value="RD" ${u.role === 'RD' ? 'selected' : ''}>RD</option>
        <option value="CNC" ${u.role === 'CNC' ? 'selected' : ''}>CNC</option>
        <option value="VP" ${u.role === 'VP' ? 'selected' : ''}>VP</option>
        <option value="Supervisor" ${u.role === 'Supervisor' ? 'selected' : ''}>Supervisor</option>
        <option value="Employee" ${u.role === 'Employee' ? 'selected' : ''}>Employee</option>
      </select>
    `;
    
    let permissionsDesc = "";
    if (u.role === 'Admin') permissionsDesc = "系統全權限（帳號、專案、簽核）";
    else if (u.role === 'PM') permissionsDesc = "專案主檔、時程甘特圖編輯與指派權限";
    else if (u.role === 'RD') permissionsDesc = "版本履歷登錄、試戴報告問卷解析匯入";
    else if (u.role === 'CNC') permissionsDesc = "模仁設計、PP 模與 CNC Gate 確認權限";
    else if (u.role === 'VP') permissionsDesc = "高階里程碑檢視與 Gate 簽核核准";
    else if (u.role === 'Supervisor') permissionsDesc = "主管權限：編輯專案、分派負責人、審核 Gate";
    else if (u.role === 'Employee') permissionsDesc = "一般員工：查看並更新被分派的 Gate";
    
    tr.innerHTML = `
      <td><strong>${u.name}</strong></td>
      <td><code>${u.username}</code></td>
      <td>${u.dept}</td>
      <td>${options}</td>
      <td class="text-muted" style="font-size:12px;">${permissionsDesc}</td>
      <td>
        <button class="btn btn-outline btn-xs" onclick="deleteUser('${u.username}')">刪除</button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function updateUserRole(username, newRole) {
  const u = db.users.find(usr => usr.username === username);
  if (u) {
    u.role = newRole;
    saveDatabase();
    
    // 若目前修改的是自己，強制更新 session 與視角
    if (username === currentUser.username) {
      currentUser.role = newRole;
      sessionStorage.setItem('LENS_USER', JSON.stringify(currentUser));
      changePerspective(newRole);
    } else {
      renderAdminSettings();
    }
  }
}

function deleteUser(username) {
  if (username === currentUser.username) {
    alert("您無法刪除自己正在登入的帳號！");
    return;
  }
  
  if (confirm(`確定要刪除帳號 ${username} 嗎？`)) {
    db.users = db.users.filter(u => u.username !== username);
    saveDatabase();
    renderAdminSettings();
  }
}

function showAddUserModal() {
  document.getElementById('user-form').reset();
  openModal('modal-user');
}

function submitUserForm() {
  const name = document.getElementById('usr-name').value.trim();
  const username = document.getElementById('usr-username').value.trim();
  const password = document.getElementById('usr-password').value;
  const dept = document.getElementById('usr-dept').value.trim();
  const role = document.getElementById('usr-role').value;
  
  const exists = db.users.some(u => u.username === username);
  if (exists) {
    alert("此登入帳號已被使用，請更換！");
    return;
  }
  
  const newUser = { name, username, password, dept, role };
  db.users.push(newUser);
  saveDatabase();
  closeModal('modal-user');
  renderAdminSettings();
}

// ==========================================================================
// 9. 模擬主管 Gate 審核與核決流程
// ==========================================================================

function showGateApprovalModal(projectId, targetGate) {
  const proj = db.projects.find(p => p.id === projectId);
  if (!proj) return;
  
  // 找出最近一筆版本履歷
  const latestVer = db.versions.filter(v => v.projectId === projectId)[0];
  
  document.getElementById('appr-project-id').value = projectId;
  document.getElementById('appr-target-gate').value = targetGate;
  
  document.getElementById('appr-project-name').innerText = proj.name;
  document.getElementById('appr-gate-badge').innerText = targetGate;
  document.getElementById('appr-project-version').innerText = latestVer ? `${latestVer.version} (${latestVer.category})` : '無';
  
  document.getElementById('appr-comments').value = '';
  
  // 預設重設為 APPROVE
  document.getElementById('appr-decision').value = 'APPROVE';
  document.getElementById('reject-gate-group').classList.add('hidden');
  
  openModal('modal-gate-approval');
}

function toggleRejectionField(decision) {
  const grp = document.getElementById('reject-gate-group');
  if (decision === 'REJECT') {
    grp.classList.remove('hidden');
  } else {
    grp.classList.add('hidden');
  }
}

function submitGateApproval() {
  const projectId = document.getElementById('appr-project-id').value;
  const targetGate = document.getElementById('appr-target-gate').value;
  const decision = document.getElementById('appr-decision').value;
  const comments = document.getElementById('appr-comments').value.trim();
  
  const proj = db.projects.find(p => p.id === projectId);
  if (!proj) return;
  
  const latestVer = db.versions.filter(v => v.projectId === projectId)[0];
  const currentVer = latestVer ? latestVer.version : '未知版本';
  
  if (decision === 'APPROVE') {
    // 1. 更新專案主檔之 Gate
    proj.gate = targetGate;
    proj.status = '綠'; // 通過後重設為綠燈
    ensureProjectProcessModel(proj);
    proj.process.currentStep = legacyGateToProcessStep(targetGate);
    proj.process.phase = proj.process.currentStep <= 5 ? '平光段' : proj.process.currentStep <= 9 ? '全光度段' : '量產';
    proj.process.lastDecision = 'OK';
    
    // 2. 更新里程碑為實際完成 (即將目前的 Gate 填上實際日期)
    // 尋找前一個 Gate（即被批准通過的 Gate）並寫入實際日期
    const prevGate = getPreviousGate(targetGate);
    const ms = proj.milestones.find(m => m.name.includes(prevGate));
    if (ms) {
      ms.actual = new Date().toISOString().slice(0, 10);
      ms.status = 'green';
    }
    // 把目前的 Target 改為藍色進行中
    const currentMs = proj.milestones.find(m => m.name.includes(targetGate));
    if (currentMs) {
      currentMs.status = 'blue';
    }
    
    // 3. 寫入一筆版本履歷簽核紀錄
    const newVerLog = {
      projectId: projectId,
      version: `${currentVer}-G`,
      category: "量產版",
      change: `Gate ${prevGate} 審查通過簽核。核決者：${currentUser.name}`,
      testType: "Gate 階段核決",
      sampleSize: "",
      porosity: "",
      comfort: "",
      defect: "無阻礙缺陷，已簽退",
      conclusion: `核准通過通往 ${targetGate}。評語：${comments}`,
      nextStep: `進入 ${targetGate} 工作階段`,
      owner: currentUser.name,
      date: new Date().toISOString().slice(0, 10),
      source: "主管審核系統核決",
      isProduction: "是"
    };
    db.versions.unshift(newVerLog);
    
  } else {
    // 駁回退回
    const rejectToGate = document.getElementById('appr-reject-to').value;
    proj.gate = rejectToGate;
    proj.status = '紅'; // 退回標記為紅燈
    ensureProjectProcessModel(proj);
    proj.process.currentStep = ['G0', 'G1', 'G2'].includes(rejectToGate) ? 2 : 6;
    proj.process.phase = proj.process.currentStep === 2 ? '平光段' : '全光度段';
    proj.process.currentCycle += 1;
    proj.process.lastDecision = 'NG';
    
    // 更新里程碑狀態：將退回到的 Gate 標為藍色(重新執行)，原目標的 Gate 標為灰色或紅色
    const targetMs = proj.milestones.find(m => m.name.includes(targetGate));
    if (targetMs) targetMs.status = 'red';
    
    const rejectMs = proj.milestones.find(m => m.name.includes(rejectToGate));
    if (rejectMs) {
      rejectMs.actual = ''; // 清除舊完成日期，要求重做
      rejectMs.status = 'blue';
    }
    
    // 寫入版本履歷駁回紀錄
    const newVerLog = {
      projectId: projectId,
      version: `${currentVer}-R`,
      category: "候選版",
      change: `Gate 審查未通過，退回至 ${rejectToGate}。核決者：${currentUser.name}`,
      testType: "Gate 階段核決",
      sampleSize: "",
      porosity: "",
      comfort: "",
      defect: `退回原因：${comments}`,
      conclusion: `退回至 ${rejectToGate} 重新開發設計`,
      nextStep: "針對退回意見修正模仁結構與油墨",
      owner: currentUser.name,
      date: new Date().toISOString().slice(0, 10),
      source: "主管審核系統核決",
      isProduction: "否"
    };
    db.versions.unshift(newVerLog);
  }
  
  saveDatabase();
  closeModal('modal-gate-approval');
  
  // 重新渲染畫面
  renderVPMilestones();
}

function getPreviousGate(gate) {
  const gates = ["G0", "G1", "G2", "G3", "G4", "G5", "G6"];
  const idx = gates.indexOf(gate);
  if (idx > 0) {
    return gates[idx - 1];
  }
  return gate;
}

function legacyGateToProcessStep(gate) {
  return { G0: 1, G1: 2, G2: 5, G3: 6, G4: 9, G5: 10, G6: 10 }[gate] || 1;
}

// ==========================================================================
// 10. Modal 控制輔助函式
// ==========================================================================

function openModal(modalId) {
  document.getElementById(modalId).classList.add('active');
}

function closeModal(modalId) {
  document.getElementById(modalId).classList.remove('active');
}

// ==========================================================================
// 11. 2026-06-24 簡潔版專案 Gate 中心
//     保留既有資料，重新整理使用路徑：我的工作 → 專案總覽 → 單一專案 Gate 中心
// ==========================================================================

const CLEAN_UI_VERSION = 'gate-center-20260624a';

const PERSON_BOOK = {
  pm: '曼玉',
  rdPrimary: '齊政',
  rdBackup: '世偉',
  cnc: '鍾民',
  business: '業務窗口待指派',
  molding: '模仁 鍾民'
};

const GATE_STAGES = [
  {
    id: 1,
    name: '需求與規格確認',
    department: '業務／產品PM',
    confirmRole: 'PM',
    owner: `${PERSON_BOOK.business}／PM ${PERSON_BOOK.pm}`,
    description: '業務提出產品需求，產品 PM 確認規格完整、可進入開發。',
    checklist: ['含水量、厚度、DIA/BC', '目標市場與產品定位', '需求凍結與開案紀錄'],
    passRule: '需求完整且 PM 確認後，才可進 Gate 2。'
  },
  {
    id: 2,
    name: '平光模仁設計',
    department: 'CNC',
    confirmRole: 'CNC',
    owner: PERSON_BOOK.molding,
    description: '依規格製作平光段模仁設計，確認可加工與模仁版本。',
    checklist: ['平光模仁圖', 'CNC 加工可行性', '模仁版本標示'],
    passRule: 'CNC 負責人確認模仁設計可製作。'
  },
  {
    id: 3,
    name: '平光 PP 模',
    department: 'RD',
    confirmRole: 'RD',
    owner: `RD ${PERSON_BOOK.rdPrimary}`,
    description: '完成平光 PP 模製作，建立對應版本與試作紀錄。',
    checklist: ['PP 模完成', '版本紀錄', '異常紀錄'],
    passRule: 'RD 確認 PP 模可進入平光設計確認。'
  },
  {
    id: 4,
    name: '平光設計確認',
    department: 'RD',
    confirmRole: 'RD',
    owner: `RD ${PERSON_BOOK.rdPrimary}`,
    description: '確認快固片、移印、構型與外觀設計是否符合需求。',
    checklist: ['快固片確認', '移印確認', '構型確認'],
    passRule: 'RD 確認設計符合需求後進入試戴測試。'
  },
  {
    id: 5,
    name: '平光試戴測試／判定',
    department: 'RD',
    confirmRole: 'RD',
    owner: `RD ${PERSON_BOOK.rdPrimary}`,
    description: '進行平光試戴與測試判定；若 NG，退回 Gate 2 重啟版本。',
    checklist: ['試戴紀錄', '測試數據', 'OK/NG 判定'],
    passRule: 'OK 進 Gate 6；NG 退回 Gate 2。'
  },
  {
    id: 6,
    name: '全光度模仁設計',
    department: 'CNC',
    confirmRole: 'CNC',
    owner: PERSON_BOOK.molding,
    description: '依平光結果進行全光度模仁設計，只改焦度段，速度較快。',
    checklist: ['全光度模仁圖', '焦度範圍', '版本啟動紀錄'],
    passRule: 'CNC 負責人確認全光度模仁可製作。'
  },
  {
    id: 7,
    name: '全光度 PP 模',
    department: 'RD',
    confirmRole: 'RD',
    owner: `RD ${PERSON_BOOK.rdPrimary}`,
    description: '完成全光度 PP 模，建立對應版本與試作紀錄。',
    checklist: ['PP 模完成', '全光度版本紀錄', '異常紀錄'],
    passRule: 'RD 確認 PP 模可進入設計確認。'
  },
  {
    id: 8,
    name: '全光度設計確認',
    department: 'RD',
    confirmRole: 'RD',
    owner: `RD ${PERSON_BOOK.rdPrimary}`,
    description: '確認快固片、移印、構型與全光度設計是否符合需求。',
    checklist: ['快固片確認', '移印確認', '構型確認'],
    passRule: 'RD 確認設計符合需求後進入全光度試戴測試。'
  },
  {
    id: 9,
    name: '全光度試戴測試／判定',
    department: 'RD',
    confirmRole: 'RD',
    owner: `RD ${PERSON_BOOK.rdPrimary}`,
    description: '進行全光度試戴與測試判定；若 NG，退回 Gate 6 啟動下一版。',
    checklist: ['試戴紀錄', '測試數據', 'OK/NG 判定'],
    passRule: 'OK 進量產準備；NG 退回 Gate 6。'
  }
];

const PROJECT_DISPLAY_OVERRIDES = {
  'SIH-145-CONV': { name: 'SIH-145-CONV · 14.5 Conventional', status: '進行中', version: 'V9.1', gate: 9, owner: 'RD 世偉', tracker: '曼玉' },
  'SIH-145-SIH': { name: 'SIH-145-SIH · SiH 14.5', status: '進行中', version: 'V8', gate: 6, owner: 'RD 世偉', tracker: '曼玉' },
  'SIH-142-SIH': { name: 'SIH-142-SIH · SiH 14.2', status: '進行中', version: 'V18.2', gate: 9, owner: 'RD 齊政', tracker: '曼玉' },
  'SIH-142-UT': {
    name: 'SIH-142-UT · SiH 14.2 超薄',
    status: 'NG / 等待結果',
    version: 'V14',
    gate: 9,
    owner: 'RD 齊政／模仁 鍾民',
    tracker: '曼玉',
    ngReason: '全光度試戴測試有問題，等待測試結果彙整。',
    nextAction: '啟動 V15，退回 Gate 6 全光度模仁設計',
    resultDue: '預計 2026-06-26 有結果'
  },
  'SIH-142-CONV': { name: 'SIH-142-CONV · 14.2 Conventional', status: '進行中', version: 'V4', gate: 9, owner: 'RD 齊政', tracker: '曼玉' }
};

let currentProjectDetailId = null;

function initApp() {
  const displayName = document.getElementById('user-display-name');
  const displayDept = document.getElementById('user-display-dept');
  if (displayName) displayName.innerText = currentUser.name || currentUser.username || '使用者';
  if (displayDept) displayDept.innerText = currentUser.dept || currentUser.role || '';
  prepareCleanNavigation();
  ensureNotificationBell();
  seedGateNotifications();
  applyPermissions();
  switchTab('vp-milestones');
}

function prepareCleanNavigation() {
  document.title = '專案管理 Gate 中心';
  const brand = document.querySelector('.brand-text');
  if (brand) brand.textContent = '專案管理';
  const footer = document.getElementById('server-status-text');
  if (footer) footer.textContent = '本機資料已載入';

  setMenuItem('nav-vp-milestones', '我的工作', 'vp-milestones', true);
  setMenuItem('nav-project-master', '專案總覽', 'project-master', true);
  setMenuItem('nav-action-items', '所有任務', 'action-items', true);
  setMenuItem('nav-risks-issues', '團隊成員', 'risks-issues', true);
  setMenuItem('nav-gantt-chart', '專案甘特圖', 'gantt-chart', false);
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
}

function setMenuItem(id, label, tab, visible) {
  const el = document.getElementById(id);
  if (!el) return;
  const svg = el.querySelector('svg') ? el.querySelector('svg').outerHTML : '';
  el.dataset.tab = tab;
  el.innerHTML = `${svg}<span>${label}</span>`;
  el.classList.toggle('hidden', !visible);
}

function applyPermissions() {
  const adminItem = document.getElementById('nav-admin-settings');
  if (adminItem) adminItem.classList.add('hidden');
  ['btn-add-project-main', 'btn-add-version-main', 'btn-import-report-main', 'btn-add-action-main', 'btn-add-risk-main'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.classList.toggle('hidden', !['PM', 'Admin'].includes(activePerspective));
  });
}

function renderVPMilestones() {
  const container = document.getElementById('tab-vp-milestones');
  if (!container) return;
  currentProjectDetailId = null;
  const projects = getAllProjectsForCleanUI();
  const urgent = projects.filter(p => getProjectDisplay(p).status.includes('NG')).length;
  const dueToday = projects.filter(p => (p.deadline || '') === '2026-06-24').length;
  const active = projects.filter(p => getProjectBucket(p) === 'active').length;
  const myCards = projects
    .filter(p => getProjectBucket(p) === 'active')
    .map(p => renderMyWorkCard(p))
    .join('');

  container.innerHTML = `
    <section class="simple-page">
      <div class="page-heading">
        <p class="eyebrow">My Work</p>
        <h1>我的工作</h1>
        <p>你被指派或需要追蹤的專案，一進來就知道今天要看哪裡。</p>
      </div>

      <div class="simple-stat-grid">
        <article class="simple-stat-card danger"><span>高優先級 / NG</span><strong>${urgent}</strong></article>
        <article class="simple-stat-card warning"><span>今天到期</span><strong>${dueToday}</strong></article>
        <article class="simple-stat-card success"><span>進行中專案</span><strong>${active}</strong></article>
      </div>

      <div class="section-title-row">
        <h2>進行中</h2>
        <button class="btn btn-primary" onclick="switchTab('project-master')">看全部專案</button>
      </div>
      <div class="work-card-list">${myCards || renderEmptyCleanState('目前沒有進行中的工作')}</div>
    </section>
  `;
}

function renderMyWorkCard(project) {
  const d = getProjectDisplay(project);
  const gate = GATE_STAGES[d.gate - 1] || GATE_STAGES[0];
  return `
    <article class="work-card" onclick="openProjectGateCenter('${escapeHtml(project.id)}')">
      <div>
        <span class="status-pill ${statusClass(d.status)}">${escapeHtml(d.status)}</span>
        <h3>${escapeHtml(d.name)}</h3>
        <p>目前：Gate ${gate.id} ${escapeHtml(gate.name)}</p>
      </div>
      <div class="work-card-meta">
        <span>負責：${escapeHtml(d.owner)}</span>
        <span>PM追蹤：${escapeHtml(d.tracker)}</span>
        <span>版本：${escapeHtml(d.version)}</span>
      </div>
    </article>
  `;
}

function renderProjectMaster() {
  const container = document.getElementById('tab-project-master');
  if (!container) return;
  if (currentProjectDetailId) {
    renderProjectGateCenter(currentProjectDetailId);
    return;
  }

  const projects = getAllProjectsForCleanUI();
  const active = projects.filter(p => getProjectBucket(p) === 'active');
  const paused = projects.filter(p => getProjectBucket(p) === 'paused');
  const done = projects.filter(p => getProjectBucket(p) === 'done');

  container.innerHTML = `
    <section class="simple-page">
      <div class="page-heading">
        <p class="eyebrow">Projects</p>
        <h1>專案總覽</h1>
        <p>不用一直點進去找資訊；每張卡片直接顯示 Gate、版本、狀態、NG 原因與下一步。</p>
      </div>
      ${renderProjectSection('進行中專案', active, 'activity')}
      ${renderProjectSection('暫停', paused, 'pause')}
      ${renderProjectSection('已完成', done, 'check')}
    </section>
  `;
}

function renderProjectSection(title, projects, iconName) {
  return `
    <section class="project-status-section">
      <div class="project-section-title">
        <span class="section-dot ${iconName}"></span>
        <h2>${title} (${projects.length})</h2>
      </div>
      <div class="project-card-grid">
        ${projects.length ? projects.map(p => renderProjectOverviewCard(p)).join('') : renderEmptyCleanState('目前沒有專案')}
      </div>
    </section>
  `;
}

function renderProjectOverviewCard(project) {
  const d = getProjectDisplay(project);
  const gate = GATE_STAGES[d.gate - 1] || GATE_STAGES[0];
  const ngBlock = d.ngReason ? `
    <div class="ng-mini-block">
      <strong>NG 原因</strong>
      <span>${escapeHtml(d.ngReason)}</span>
      <em>${escapeHtml(d.resultDue || '等待結果')}</em>
    </div>` : '';
  return `
    <article class="simple-project-card" onclick="openProjectGateCenter('${escapeHtml(project.id)}')">
      <div class="project-card-top">
        <h3>${escapeHtml(d.name)}</h3>
        <span class="status-pill ${statusClass(d.status)}">${escapeHtml(d.status)}</span>
      </div>
      <div class="gate-progress-mini">
        <span style="width:${Math.round((d.gate / 9) * 100)}%"></span>
      </div>
      <p class="project-current">Gate ${gate.id} · ${escapeHtml(gate.name)}</p>
      ${ngBlock}
      <div class="project-card-footer">
        <span>版本 ${escapeHtml(d.version)}</span>
        <span>${escapeHtml(d.owner)}</span>
      </div>
    </article>
  `;
}

function openProjectGateCenter(projectId) {
  currentProjectDetailId = projectId;
  switchTab('project-master');
}

function renderProjectGateCenter(projectId) {
  const container = document.getElementById('tab-project-master');
  const project = db.projects.find(p => p.id === projectId);
  if (!container || !project) return;
  const d = getProjectDisplay(project);
  const gate = GATE_STAGES[d.gate - 1] || GATE_STAGES[0];
  const guidance = getGateGuidance(project);

  container.innerHTML = `
    <section class="simple-page gate-center-page">
      <button class="back-button" onclick="currentProjectDetailId=null; renderProjectMaster();">← 回專案總覽</button>

      <div class="project-detail-heading">
        <p class="eyebrow">Project Gate Center</p>
        <h1>${escapeHtml(d.name)}</h1>
      </div>

      <article class="gate-summary-card ${statusClass(d.status)}">
        <div class="summary-main">
          <span class="status-pill ${statusClass(d.status)}">${escapeHtml(d.status)}</span>
          <h2>目前進度：Gate ${gate.id} ${escapeHtml(gate.name)}</h2>
          <p>${escapeHtml(d.ngReason || gate.description)}</p>
        </div>
        <div class="summary-grid">
          <div><span>目前版本</span><strong>${escapeHtml(d.version)}</strong></div>
          <div><span>下一步</span><strong>${escapeHtml(d.nextAction || guidance.nextAction)}</strong></div>
          <div><span>負責</span><strong>${escapeHtml(d.owner)}</strong></div>
          <div><span>PM追蹤</span><strong>${escapeHtml(d.tracker)}</strong></div>
        </div>
      </article>

      <div class="section-title-row">
        <h2>9 個 Gate</h2>
        <p>量產不放進 Gate；Gate 9 OK 後才進入「量產準備 / 量產放行」。</p>
      </div>

      <div class="gate-grid">
        ${GATE_STAGES.map(g => renderGateCard(project, g)).join('')}
      </div>
    </section>
  `;
}

function renderGateCard(project, gate) {
  const d = getProjectDisplay(project);
  const state = getGateState(d.gate, gate.id, d.status);
  const assignment = getGateAssignment(project, gate);
  const canConfirm = canConfirmGate(gate);
  const isDecisionGate = gate.id === 5 || gate.id === 9;
  const disabledAttr = canConfirm ? '' : 'disabled';
  const lockedText = canConfirm ? '你可以確認這一關' : `僅 ${gate.department} 負責人可確認`;

  return `
    <article class="gate-card ${state} ${canConfirm ? '' : 'locked'}">
      <div class="gate-card-head">
        <span class="gate-number">Gate ${gate.id}</span>
        <span class="gate-state ${state}">${gateStateText(state)}</span>
      </div>
      <h3>${escapeHtml(gate.name)}</h3>
      <p>${escapeHtml(gate.description)}</p>

      <div class="gate-owner-grid">
        <div><span>部門</span><strong>${escapeHtml(assignment.department)}</strong></div>
        <div><span>負責</span><strong>${escapeHtml(assignment.owner)}</strong></div>
        <div><span>追蹤</span><strong>PM ${PERSON_BOOK.pm}</strong></div>
      </div>

      <ul class="gate-checklist">
        ${gate.checklist.map(item => `<li>${escapeHtml(item)}</li>`).join('')}
      </ul>

      <div class="gate-action-row">
        <span>${lockedText}</span>
        ${isDecisionGate ? `
          <button class="btn btn-outline" ${disabledAttr} onclick="confirmProjectGate('${escapeHtml(project.id)}', ${gate.id}, 'NG')">NG 退回</button>
          <button class="btn btn-primary" ${disabledAttr} onclick="confirmProjectGate('${escapeHtml(project.id)}', ${gate.id}, 'OK')">OK 確認</button>
        ` : `
          <button class="btn btn-primary" ${disabledAttr} onclick="confirmProjectGate('${escapeHtml(project.id)}', ${gate.id}, 'OK')">確認完成</button>
        `}
      </div>
    </article>
  `;
}

function renderActionItems() {
  const container = document.getElementById('tab-action-items');
  if (!container) return;
  currentProjectDetailId = null;
  const tasks = getCleanTasks();
  container.innerHTML = `
    <section class="simple-page">
      <div class="page-heading">
        <p class="eyebrow">Tasks</p>
        <h1>所有任務</h1>
        <p>把每個 Gate 需要誰做、何時完成，整理成可搜尋的任務清單。</p>
      </div>
      <div class="task-list-clean">
        ${tasks.map(t => `
          <article class="task-row-clean">
            <div>
              <h3>${escapeHtml(t.title)}</h3>
              <p>${escapeHtml(t.project)} · ${escapeHtml(t.owner)}</p>
            </div>
            <div class="task-meta-clean">
              <span class="status-pill ${statusClass(t.status)}">${escapeHtml(t.status)}</span>
              <span>${escapeHtml(t.deadline)}</span>
            </div>
          </article>
        `).join('')}
      </div>
    </section>
  `;
}

function renderRisksIssues() {
  const container = document.getElementById('tab-risks-issues');
  if (!container) return;
  currentProjectDetailId = null;
  const members = [
    { dept: 'PM', name: '曼玉', role: '產品 PM / 專案追蹤', load: '5 個專案追蹤中' },
    { dept: 'RD', name: '齊政', role: '全光度設計與測試', load: '3 個 Gate 負責' },
    { dept: 'RD', name: '世偉', role: '平光設計與試作', load: '2 個 Gate 負責' },
    { dept: 'CNC', name: '鍾民', role: '模仁設計 / CNC', load: 'Gate 2、Gate 6 負責' },
    { dept: '業務', name: '業務窗口待指派', role: '需求提出', load: 'Gate 1 負責' }
  ];
  const groups = [...new Set(members.map(m => m.dept))];
  container.innerHTML = `
    <section class="simple-page">
      <div class="page-heading">
        <p class="eyebrow">Team</p>
        <h1>團隊成員</h1>
        <p>每個部門與人名都要清楚，避免 Gate 卡住時不知道找誰。</p>
      </div>
      ${groups.map(group => `
        <section class="team-group">
          <h2>${group} (${members.filter(m => m.dept === group).length})</h2>
          <div class="team-card-grid">
            ${members.filter(m => m.dept === group).map(m => `
              <article class="team-card">
                <div class="avatar-clean">${escapeHtml(m.name.slice(0, 1))}</div>
                <div>
                  <h3>${escapeHtml(m.name)}</h3>
                  <span>${escapeHtml(m.role)}</span>
                  <p>${escapeHtml(m.load)}</p>
                </div>
              </article>
            `).join('')}
          </div>
        </section>
      `).join('')}
    </section>
  `;
}

function getAllProjectsForCleanUI() {
  return db.projects || [];
}

function getProjectDisplay(project) {
  const override = { ...(PROJECT_DISPLAY_OVERRIDES[project.id] || {}), ...(project.cleanGateState || {}) };
  const inferredGate = override.gate || inferGateFromProject(project);
  return {
    name: override.name || `${project.id} · ${project.name || '未命名專案'}`,
    status: override.status || normalizeStatus(project.status),
    version: override.version || project.currentVersion || 'V待定',
    gate: inferredGate,
    owner: override.owner || inferOwner(project, inferredGate),
    tracker: override.tracker || PERSON_BOOK.pm,
    ngReason: override.ngReason || '',
    nextAction: override.nextAction || project.nextStep || '',
    resultDue: override.resultDue || ''
  };
}

function inferGateFromProject(project) {
  if (project.process && project.process.currentStep) return Math.max(1, Math.min(9, Number(project.process.currentStep)));
  return { G0: 1, G1: 2, G2: 5, G3: 6, G4: 9, G5: 9, G6: 9 }[project.gate] || 1;
}

function inferOwner(project, gateId) {
  if ([2, 6].includes(gateId)) return PERSON_BOOK.molding;
  if (gateId === 1) return `${PERSON_BOOK.business}／PM ${PERSON_BOOK.pm}`;
  return `RD ${project.id.includes('142') ? PERSON_BOOK.rdPrimary : PERSON_BOOK.rdBackup}`;
}

function normalizeStatus(status) {
  const text = String(status || '');
  if (text.includes('紅') || text.includes('NG') || text.includes('蝝')) return 'NG / 等待結果';
  if (text.includes('綠') || text.includes('完成')) return '順利';
  return '進行中';
}

function getProjectBucket(project) {
  const d = getProjectDisplay(project);
  if ((project.statusDesc || '').includes('暫停')) return 'paused';
  if ((project.gate === 'G6' || (project.process && Number(project.process.currentStep) >= 10)) && d.status === '順利') return 'done';
  return 'active';
}

function getGateGuidance(project) {
  const d = getProjectDisplay(project);
  if (d.gate === 9 && d.status.includes('NG')) return { nextAction: `啟動 ${nextVersion(d.version)}，退回 Gate 6 全光度模仁設計` };
  if (d.gate === 5 && d.status.includes('NG')) return { nextAction: `啟動 ${nextVersion(d.version)}，退回 Gate 2 平光模仁設計` };
  if (d.gate === 9) return { nextAction: 'Gate 9 OK 後，進入量產準備 / 量產放行' };
  const nextGate = GATE_STAGES[d.gate] || GATE_STAGES[d.gate - 1];
  return { nextAction: nextGate ? `進入 Gate ${nextGate.id} ${nextGate.name}` : '等待下一步指派' };
}

function getGateState(currentGate, gateId, projectStatus) {
  if (gateId < currentGate) return 'done';
  if (gateId > currentGate) return 'pending';
  if (String(projectStatus).includes('NG')) return 'blocked';
  return 'current';
}

function gateStateText(state) {
  return { done: '已完成', current: '進行中', blocked: 'NG', pending: '待開始' }[state] || state;
}

function getGateAssignment(project, gate) {
  return {
    department: gate.department,
    owner: inferOwner(project, gate.id)
  };
}

function canConfirmGate(gate) {
  if (activePerspective === 'Admin') return true;
  if (gate.confirmRole === 'PM') return activePerspective === 'PM';
  if (gate.confirmRole === 'RD') return activePerspective === 'RD';
  return false;
}

function confirmProjectGate(projectId, gateId, decision) {
  const project = db.projects.find(p => p.id === projectId);
  const gate = GATE_STAGES[gateId - 1];
  if (!project || !gate || !canConfirmGate(gate)) return;
  const today = new Date().toISOString().slice(0, 10);
  const display = getProjectDisplay(project);
  project.lastUpdate = today;

  if (decision === 'NG') {
    const backGate = gateId === 9 ? 6 : 2;
    project.status = '紅';
    project.gate = backGate === 6 ? 'G3' : 'G1';
    project.nextStep = `啟動 ${nextVersion(display.version)}，退回 Gate ${backGate} ${GATE_STAGES[backGate - 1].name}`;
    project.currentVersion = nextVersion(display.version);
    project.process = { ...(project.process || {}), currentStep: backGate, lastDecision: 'NG' };
    project.cleanGateState = {
      status: 'NG / 等待結果',
      gate: backGate,
      version: project.currentVersion,
      owner: inferOwner(project, backGate),
      nextAction: project.nextStep,
      ngReason: `${gate.name} 判定 NG，需退回前段重啟版本。`,
      resultDue: '等待負責人回填結果時間'
    };
  } else {
    const nextGate = Math.min(10, gateId + 1);
    project.status = gateId === 9 ? '綠' : '黃';
    project.gate = nextGate >= 10 ? 'G5' : stepToCleanLegacyGate(nextGate);
    project.nextStep = gateId === 9 ? '進入量產準備 / 量產放行' : `進入 Gate ${nextGate} ${GATE_STAGES[nextGate - 1].name}`;
    project.process = { ...(project.process || {}), currentStep: nextGate, lastDecision: 'OK' };
    project.cleanGateState = {
      status: gateId === 9 ? '順利' : '進行中',
      gate: Math.min(9, nextGate),
      version: project.currentVersion,
      owner: inferOwner(project, Math.min(9, nextGate)),
      nextAction: project.nextStep,
      ngReason: '',
      resultDue: ''
    };
  }

  addGateNotification({
    projectId,
    title: `${currentUser.name || activePerspective} 更新 ${projectId}`,
    message: `${gate.name} 已判定 ${decision}，下一步：${project.nextStep}`,
    actor: currentUser.name || activePerspective
  });
  saveDatabase();
  renderNotificationsBell();
  renderProjectGateCenter(projectId);
}

function stepToCleanLegacyGate(step) {
  return { 1: 'G0', 2: 'G1', 3: 'G1', 4: 'G2', 5: 'G2', 6: 'G3', 7: 'G3', 8: 'G4', 9: 'G4' }[step] || 'G0';
}

function nextVersion(version) {
  const text = String(version || 'V1');
  const match = text.match(/V(\d+)/i);
  if (!match) return 'V下一版';
  return `V${Number(match[1]) + 1}`;
}

function statusClass(status) {
  const text = String(status || '');
  if (text.includes('NG') || text.includes('紅') || text.includes('阻')) return 'danger';
  if (text.includes('等待') || text.includes('暫')) return 'warning';
  if (text.includes('完成') || text.includes('順利')) return 'success';
  return 'info';
}

function getCleanTasks() {
  const projects = getAllProjectsForCleanUI();
  const tasks = [];
  projects.forEach(project => {
    const d = getProjectDisplay(project);
    const gate = GATE_STAGES[d.gate - 1] || GATE_STAGES[0];
    tasks.push({
      project: d.name,
      title: `Gate ${gate.id} ${gate.name}`,
      owner: d.owner,
      deadline: project.deadline || '待排程',
      status: d.status
    });
    if (d.nextAction) {
      tasks.push({
        project: d.name,
        title: d.nextAction,
        owner: d.owner,
        deadline: d.resultDue || project.deadline || '待排程',
        status: d.status.includes('NG') ? '等待結果' : '進行中'
      });
    }
  });
  return tasks;
}

function ensureNotificationBell() {
  const headerRight = document.querySelector('.header-right');
  if (!headerRight || document.getElementById('notification-bell')) return;
  headerRight.insertAdjacentHTML('afterbegin', `
    <div class="notification-wrap">
      <button id="notification-bell" class="bell-button" onclick="toggleNotificationPanel()" title="通知">
        🔔 <span id="notification-count">0</span>
      </button>
      <div id="notification-panel" class="notification-panel hidden"></div>
    </div>
  `);
  renderNotificationsBell();
}

function seedGateNotifications() {
  if (!db.notifications) db.notifications = [];
  if (db.notifications.length) return;
  db.notifications = [
    {
      id: `N-${Date.now()}-1`,
      projectId: 'SIH-142-UT',
      title: 'SIH-142-UT 狀態更新',
      message: 'RD 齊政更新 Gate 9：全光度試戴測試 NG，等待結果；下一步預計啟動 V15 退回 Gate 6。',
      actor: '齊政',
      date: '2026-06-24',
      read: false
    },
    {
      id: `N-${Date.now()}-2`,
      projectId: 'SIH-142-SIH',
      title: 'SIH-142-SIH 進度更新',
      message: 'Gate 9 全光度試戴測試進行中，PM 曼玉追蹤結果。',
      actor: '曼玉',
      date: '2026-06-24',
      read: false
    }
  ];
  saveDatabase();
}

function addGateNotification(payload) {
  if (!db.notifications) db.notifications = [];
  db.notifications.unshift({
    id: `N-${Date.now()}`,
    projectId: payload.projectId,
    title: payload.title,
    message: payload.message,
    actor: payload.actor,
    date: new Date().toISOString().slice(0, 10),
    read: false
  });
}

function renderNotificationsBell() {
  const count = document.getElementById('notification-count');
  const panel = document.getElementById('notification-panel');
  if (!count || !panel) return;
  const notes = db.notifications || [];
  count.textContent = notes.filter(n => !n.read).length;
  panel.innerHTML = `
    <div class="notification-panel-head">
      <strong>通知</strong>
      <button onclick="markAllNotificationsRead()">全部已讀</button>
    </div>
    ${notes.length ? notes.slice(0, 8).map(n => `
      <article class="notification-item ${n.read ? 'read' : ''}" onclick="openProjectGateCenter('${escapeHtml(n.projectId)}')">
        <strong>${escapeHtml(n.title)}</strong>
        <p>${escapeHtml(n.message)}</p>
        <span>${escapeHtml(n.actor)} · ${escapeHtml(n.date)}</span>
      </article>
    `).join('') : '<p class="empty-text">目前沒有通知</p>'}
  `;
}

function toggleNotificationPanel() {
  const panel = document.getElementById('notification-panel');
  if (panel) panel.classList.toggle('hidden');
}

function markAllNotificationsRead() {
  (db.notifications || []).forEach(n => { n.read = true; });
  saveDatabase();
  renderNotificationsBell();
}

function renderEmptyCleanState(text) {
  return `<div class="empty-clean-state">${escapeHtml(text)}</div>`;
}

function switchTab(tabId) {
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
  else if (tabId === 'gantt-chart' && typeof renderGanttChart === 'function') renderGanttChart();
  else if (tabId === 'version-history' && typeof renderVersionHistory === 'function') renderVersionHistory();
  else if (tabId === 'admin-settings' && typeof renderAdminSettings === 'function') renderAdminSettings();
}

// ==========================================================================
// 12. Gate Center v2：亮色版 / 4 大分頁 / 橫向 Gate 卡片 / 會議匯出
// ==========================================================================

var gateCenterV2State = window.gateCenterV2State || {
  phaseByProject: {},
  versionByProjectPhase: {},
  progressView: 'gantt'
};
window.gateCenterV2State = gateCenterV2State;

const GATE_PHASES_V2 = [
  { id: 'demand', label: '需求', subtitle: '規格確認', gateKeys: ['1'] },
  { id: 'plano', label: '平光', subtitle: '模仁 / PP / 設計 / 試戴', gateKeys: ['2', '3', '4', '5'] },
  { id: 'power', label: '全光', subtitle: '全光度模仁到測試', gateKeys: ['6', '7', '8', '9'] },
  { id: 'production', label: '量產', subtitle: 'Gate 9 OK 後放行', gateKeys: ['P1', 'P2', 'P3', 'P4'] }
];

const PRODUCTION_GATES_V2 = [
  { id: 'P1', name: '量產準備', department: 'PM / 生產', confirmRole: 'PM', owner: 'PM 曼聿 / 生產主管', description: '確認量產前資料、規格、版本與放行條件。', checklist: ['量產資料完整', '版本確認', '風險清單關閉'] },
  { id: 'P2', name: '量產放行', department: '主管審核', confirmRole: 'VP', owner: '部門主管', description: '主管審核量產放行，確認可進入正式生產。', checklist: ['主管審核', '放行紀錄', '跨部門確認'] },
  { id: 'P3', name: '生產確認', department: '生產', confirmRole: 'VP', owner: '生產主管', description: '確認首批生產狀態與異常回報。', checklist: ['首批狀態', '異常回報', '品質確認'] },
  { id: 'P4', name: '上市 / 出貨前確認', department: 'PM / 業務', confirmRole: 'PM', owner: 'PM 曼聿 / 業務窗口', description: '確認上市、出貨與會議報告資料。', checklist: ['出貨確認', '上市資料', '會議報告'] }
];

const TEAM_STRUCTURE_V2 = [
  { dept: 'RD 研發部', manager: 'Andy', title: '副總', members: ['齊政', '世偉', '其他 RD'] },
  { dept: 'CNC / 模仁', manager: '某某', title: '主管', members: ['鍾民', '其他模仁人員'] },
  { dept: '業務', manager: '某某', title: '主管', members: ['業務窗口'] },
  { dept: 'PM', manager: '曼聿', title: '經理', members: ['媃安', '宣祠'] }
];

function prepareCleanNavigation() {
  document.title = '產品開發 Gate 管理';
  const brand = document.querySelector('.brand-text');
  if (brand) brand.textContent = '專案管理';
  const footer = document.getElementById('server-status-text');
  if (footer) footer.textContent = '本機資料已載入';

  setMenuItem('nav-vp-milestones', '我的工作', 'vp-milestones', true);
  setMenuItem('nav-project-master', '專案總覽', 'project-master', true);
  setMenuItem('nav-action-items', '專案進度', 'action-items', true);
  setMenuItem('nav-risks-issues', '團隊成員', 'risks-issues', true);
  setMenuItem('nav-gantt-chart', '專案甘特圖', 'gantt-chart', false);
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
}

function getGateDefinitionV2(gateKey) {
  const key = String(gateKey);
  if (key.startsWith('P')) return PRODUCTION_GATES_V2.find(g => g.id === key);
  return (GATE_STAGES || []).find(g => String(g.id) === key);
}

function getGatePhaseV2(gateKey) {
  const key = String(gateKey);
  return GATE_PHASES_V2.find(phase => phase.gateKeys.includes(key)) || GATE_PHASES_V2[0];
}

function ensureGateCenterV2(project) {
  if (!project.gateRecordsV2) project.gateRecordsV2 = {};
  const display = getProjectDisplay(project);
  const currentGate = String(display.gate || inferGateFromProject(project));
  const allKeys = GATE_PHASES_V2.flatMap(p => p.gateKeys);

  allKeys.forEach(key => {
    const def = getGateDefinitionV2(key);
    if (!def) return;
    if (!project.gateRecordsV2[key]) {
      project.gateRecordsV2[key] = {
        gateKey: key,
        status: getInitialGateStatusV2(key, currentGate, display.status),
        owner: inferOwnerV2(project, key),
        department: def.department,
        dueDate: inferDueDateV2(project, key),
        confirmedBy: '',
        confirmedAt: '',
        supervisor: inferSupervisorV2(key),
        supervisorStatus: '待審核',
        supervisorAt: '',
        note: '',
        imageLink: '',
        imageNote: '',
        ng: null
      };
    }
  });

  if (!project.phaseVersionsV2) {
    project.phaseVersionsV2 = {
      demand: ['需求確認'],
      plano: ['V1'],
      power: [display.version || project.currentVersion || 'V1'],
      production: ['量產準備']
    };
  }

  return project.gateRecordsV2;
}

function getInitialGateStatusV2(key, currentGate, projectStatus) {
  if (String(projectStatus || '').includes('NG') && key === currentGate) return 'NG';
  if (key.startsWith('P')) return Number(currentGate) >= 10 ? '進行中' : '待開始';
  const numericKey = Number(key);
  const numericCurrent = Number(currentGate);
  if (numericKey < numericCurrent) return '已完成';
  if (numericKey === numericCurrent) return String(projectStatus || '').includes('NG') ? 'NG' : '進行中';
  return '待開始';
}

function inferOwnerV2(project, gateKey) {
  const key = String(gateKey);
  if (key === '1') return '業務窗口 / PM 曼聿';
  if (['2', '6'].includes(key)) return 'CNC 鍾民';
  if (['3', '4', '5', '7', '8', '9'].includes(key)) return project.id && project.id.includes('142') ? 'RD 齊政' : 'RD 世偉';
  if (key === 'P1') return 'PM 曼聿 / 生產主管';
  if (key === 'P2') return '部門主管';
  if (key === 'P3') return '生產主管';
  return 'PM 曼聿 / 業務窗口';
}

function inferSupervisorV2(gateKey) {
  const key = String(gateKey);
  if (['3', '4', '5', '7', '8', '9'].includes(key)) return 'Andy 副總';
  if (['2', '6'].includes(key)) return 'CNC 主管';
  if (key.startsWith('P')) return '部門主管';
  return '曼聿 經理';
}

function inferDueDateV2(project, gateKey) {
  if (project.deadline) return project.deadline;
  return '';
}

function getCurrentPhaseV2(project) {
  const display = getProjectDisplay(project);
  if (display.gate >= 10) return 'production';
  if (display.gate >= 6) return 'power';
  if (display.gate >= 2) return 'plano';
  return 'demand';
}

function getActivePhaseV2(project) {
  return gateCenterV2State.phaseByProject[project.id] || getCurrentPhaseV2(project);
}

function getActiveVersionV2(project, phaseId) {
  const key = `${project.id}:${phaseId}`;
  const versions = getPhaseVersionsV2(project, phaseId);
  if (!gateCenterV2State.versionByProjectPhase[key]) {
    gateCenterV2State.versionByProjectPhase[key] = versions[versions.length - 1];
  }
  return gateCenterV2State.versionByProjectPhase[key];
}

function getPhaseVersionsV2(project, phaseId) {
  ensureGateCenterV2(project);
  return project.phaseVersionsV2[phaseId] || ['目前版本'];
}

function setProjectPhaseTab(projectId, phaseId) {
  const project = db.projects.find(p => p.id === projectId);
  if (!project) return;
  gateCenterV2State.phaseByProject[projectId] = phaseId;
  renderProjectGateCenter(projectId);
}

function setProjectPhaseVersion(projectId, phaseId, version) {
  gateCenterV2State.versionByProjectPhase[`${projectId}:${phaseId}`] = version;
  renderProjectGateCenter(projectId);
}

function addProjectPhaseVersion(projectId, phaseId) {
  const project = db.projects.find(p => p.id === projectId);
  if (!project) return;
  ensureGateCenterV2(project);
  const versions = project.phaseVersionsV2[phaseId] || ['V1'];
  let next = 'V2';
  const latest = versions[versions.length - 1] || 'V1';
  const match = String(latest).match(/V(\d+)/i);
  if (match) next = `V${Number(match[1]) + 1}`;
  else if (phaseId === 'production') next = `量產版 ${versions.length + 1}`;
  versions.push(next);
  project.phaseVersionsV2[phaseId] = versions;
  gateCenterV2State.versionByProjectPhase[`${projectId}:${phaseId}`] = next;
  saveDatabase();
  renderProjectGateCenter(projectId);
}

function renderProjectGateCenter(projectId) {
  const container = document.getElementById('tab-project-master');
  const project = db.projects.find(p => p.id === projectId);
  if (!container || !project) return;

  ensureGateCenterV2(project);
  const display = getProjectDisplay(project);
  const currentPhase = getCurrentPhaseV2(project);
  const activePhaseId = getActivePhaseV2(project);
  const activePhase = GATE_PHASES_V2.find(p => p.id === activePhaseId) || GATE_PHASES_V2[0];
  const activeVersion = getActiveVersionV2(project, activePhaseId);
  const currentGateDef = getGateDefinitionV2(String(display.gate)) || getGateDefinitionV2(activePhase.gateKeys[0]);
  const guidance = getGateGuidanceV2(project);

  container.innerHTML = `
    <section class="simple-page gate-center-page gate-center-v2">
      <div class="page-action-row">
        <button class="back-button" onclick="currentProjectDetailId=null; renderProjectMaster();">← 回專案總覽</button>
        <button class="btn btn-primary" onclick="exportProjectMeetingExcel('${escapeHtml(project.id)}')">匯出會議 Excel</button>
      </div>

      <div class="project-detail-heading">
        <p class="eyebrow">Project Gate Center</p>
        <h1>${escapeHtml(display.name)}</h1>
      </div>

      <article class="gate-summary-card light-summary ${statusClass(display.status)}">
        <div class="summary-main">
          <span class="status-pill ${statusClass(display.status)}">${escapeHtml(display.status)}</span>
          <h2>目前：${escapeHtml(getPhaseLabelV2(currentPhase))} / ${escapeHtml(currentGateDef ? currentGateDef.name : '待確認')}</h2>
          <p>${escapeHtml(display.ngReason || guidance.reason || '依 Gate 紀錄追蹤目前進度、負責人、確認時間與下一步。')}</p>
        </div>
        <div class="summary-grid">
          <div><span>目前大項目</span><strong>${escapeHtml(getPhaseLabelV2(currentPhase))}</strong></div>
          <div><span>目前版本</span><strong>${escapeHtml(display.version || activeVersion)}</strong></div>
          <div><span>下一步</span><strong>${escapeHtml(display.nextAction || guidance.nextAction)}</strong></div>
          <div><span>PM追蹤</span><strong>${escapeHtml(display.tracker || '曼聿')}</strong></div>
        </div>
      </article>

      <div class="phase-dashboard">
        ${GATE_PHASES_V2.map(phase => renderPhaseDashboardCardV2(project, phase, currentPhase, activePhaseId)).join('')}
      </div>

      <div class="phase-tabs">
        ${GATE_PHASES_V2.map(phase => `
          <button class="${phase.id === activePhaseId ? 'active' : ''}" onclick="setProjectPhaseTab('${escapeHtml(project.id)}', '${phase.id}')">
            ${escapeHtml(phase.label)}
          </button>
        `).join('')}
      </div>

      <section class="phase-panel">
        <div class="phase-panel-head">
          <div>
            <h2>${escapeHtml(activePhase.label)}</h2>
            <p>${escapeHtml(activePhase.subtitle)}</p>
          </div>
          ${renderVersionTabsV2(project, activePhaseId, activeVersion)}
        </div>

        <div class="horizontal-gate-list">
          ${activePhase.gateKeys.map(key => renderGateCardV2(project, key, activeVersion)).join('')}
        </div>
      </section>
    </section>
  `;
}

function renderPhaseDashboardCardV2(project, phase, currentPhase, activePhaseId) {
  const records = ensureGateCenterV2(project);
  const statuses = phase.gateKeys.map(key => records[key] ? records[key].status : '待開始');
  const completeCount = statuses.filter(s => s === '已完成').length;
  const blocked = statuses.some(s => s === 'NG');
  const inProgress = phase.id === currentPhase || statuses.some(s => s === '進行中');
  const state = blocked ? 'NG' : completeCount === phase.gateKeys.length ? '已完成' : inProgress ? '進行中' : '待開始';
  return `
    <button class="phase-dash-card ${phase.id === activePhaseId ? 'active' : ''} ${stateClassV2(state)}" onclick="setProjectPhaseTab('${escapeHtml(project.id)}', '${phase.id}')">
      <span>${escapeHtml(phase.label)}</span>
      <strong>${escapeHtml(state)}</strong>
      <small>${completeCount}/${phase.gateKeys.length} Gate 完成</small>
    </button>
  `;
}

function renderVersionTabsV2(project, phaseId, activeVersion) {
  const versions = getPhaseVersionsV2(project, phaseId);
  const canAdd = ['plano', 'power', 'production'].includes(phaseId);
  return `
    <div class="version-tabs-v2">
      ${versions.map(v => `<button class="${v === activeVersion ? 'active' : ''}" onclick="setProjectPhaseVersion('${escapeHtml(project.id)}', '${phaseId}', '${escapeHtml(v)}')">${escapeHtml(v)}</button>`).join('')}
      ${canAdd ? `<button class="add-version" onclick="addProjectPhaseVersion('${escapeHtml(project.id)}', '${phaseId}')">＋新增頁籤</button>` : ''}
    </div>
  `;
}

function renderGateCardV2(project, gateKey, activeVersion) {
  const key = String(gateKey);
  const def = getGateDefinitionV2(key);
  const records = ensureGateCenterV2(project);
  const rec = records[key];
  if (!def || !rec) return '';
  const canConfirm = canConfirmGateV2(def);
  const isDone = rec.status === '已完成';
  const isDecisionGate = key === '5' || key === '9';
  const safe = safeDomId(`${project.id}-${key}`);
  const confirmDisabled = !canConfirm || isDone ? 'disabled' : '';
  const lockedText = isDone ? '已完成，不可編輯' : canConfirm ? '你可以確認此 Gate' : `僅 ${def.department} / 主管可確認`;
  const imageBlock = key === '4' ? renderImageRecordBlockV2(project, key, rec, safe, isDone) : '';
  const ngBlock = isDecisionGate ? renderNgBlockV2(project, key, rec, safe, canConfirm && !isDone) : '';

  return `
    <article class="gate-row-card ${stateClassV2(rec.status)} ${isDone ? 'readonly' : ''}">
      <div class="gate-row-main">
        <div class="gate-row-title">
          <span class="gate-number">${key.startsWith('P') ? key : `Gate ${key}`}</span>
          <div>
            <h3>${escapeHtml(def.name)}</h3>
            <p>${escapeHtml(def.description || '')}</p>
          </div>
        </div>
        <span class="status-pill ${stateClassV2(rec.status)}">${escapeHtml(rec.status)}</span>
      </div>

      <div class="gate-row-fields">
        <div><span>版本</span><strong>${escapeHtml(activeVersion)}</strong></div>
        <div><span>部門</span><strong>${escapeHtml(def.department)}</strong></div>
        <div><span>負責人</span><strong>${escapeHtml(rec.owner)}</strong></div>
        <div><span>預計完成</span><strong>${escapeHtml(rec.dueDate || '待排程')}</strong></div>
        <div><span>確認人</span><strong>${escapeHtml(rec.confirmedBy || '尚未確認')}</strong></div>
        <div><span>確認時間</span><strong>${escapeHtml(rec.confirmedAt || '尚未確認')}</strong></div>
        <div><span>主管審核</span><strong>${escapeHtml(rec.supervisor)} / ${escapeHtml(rec.supervisorStatus)}</strong></div>
      </div>

      ${imageBlock}
      ${ngBlock}

      <div class="gate-row-actions">
        <span>${escapeHtml(lockedText)}</span>
        <button class="btn btn-primary" ${confirmDisabled} onclick="confirmProjectGateV2('${escapeHtml(project.id)}', '${key}', 'OK')">確認完成</button>
        ${isDecisionGate ? `<button class="btn btn-outline danger" ${confirmDisabled} onclick="showNgFormV2('${safe}')">NG 填寫原因</button>` : ''}
      </div>
    </article>
  `;
}

function renderImageRecordBlockV2(project, key, rec, safe, isDone) {
  const disabled = isDone ? 'disabled' : '';
  return `
    <div class="gate-evidence-block">
      <div>
        <strong>圖片紀錄</strong>
        <p>平光設計確認可放圖片連結，點開後另開圖檔。</p>
      </div>
      <input id="image-link-${safe}" ${disabled} value="${escapeAttribute(rec.imageLink || '')}" placeholder="貼上圖片連結或公司資料夾路徑">
      <input id="image-note-${safe}" ${disabled} value="${escapeAttribute(rec.imageNote || '')}" placeholder="圖片備註，例如：構型調整點">
      <button class="btn btn-outline" ${disabled} onclick="saveGateImageRecordV2('${escapeHtml(project.id)}', '${key}', '${safe}')">儲存圖片紀錄</button>
      ${rec.imageLink ? `<a class="image-open-link" href="${escapeAttribute(rec.imageLink)}" target="_blank" rel="noopener">另開圖片</a>` : ''}
    </div>
  `;
}

function renderNgBlockV2(project, key, rec, safe, canEdit) {
  const ng = rec.ng || {};
  const disabled = canEdit ? '' : 'disabled';
  return `
    <div class="ng-record-block ${rec.ng ? '' : 'collapsed'}" id="ng-form-${safe}">
      <div class="ng-record-title">
        <strong>NG 紀錄</strong>
        <span>${rec.ng ? '已填寫' : '負責人需填時間 / 原因 / 備註'}</span>
      </div>
      <div class="ng-form-grid">
        <label>NG 時間
          <input id="ng-time-${safe}" ${disabled} type="datetime-local" value="${escapeAttribute(toDatetimeLocalValue(ng.time) || getNowDatetimeLocal())}">
        </label>
        <label>預計重新完成
          <input id="ng-due-${safe}" ${disabled} type="date" value="${escapeAttribute(ng.due || '')}">
        </label>
        <label>NG 原因
          <input id="ng-reason-${safe}" ${disabled} value="${escapeAttribute(ng.reason || '')}" placeholder="例如：臨床試戴不通過 / 全光度測試異常">
        </label>
        <label>備註
          <textarea id="ng-note-${safe}" ${disabled} rows="2" placeholder="補充原因、處理方向、需協作部門">${escapeHtml(ng.note || '')}</textarea>
        </label>
      </div>
      <div class="ng-form-actions">
        <button class="btn btn-outline danger" ${disabled} onclick="saveGateNgV2('${escapeHtml(project.id)}', '${key}', '${safe}')">儲存 NG 並新增下一版</button>
      </div>
    </div>
  `;
}

function showNgFormV2(safeId) {
  const el = document.getElementById(`ng-form-${safeId}`);
  if (el) el.classList.remove('collapsed');
}

function confirmProjectGateV2(projectId, gateKey, decision) {
  const project = db.projects.find(p => p.id === projectId);
  const key = String(gateKey);
  const def = getGateDefinitionV2(key);
  if (!project || !def || !canConfirmGateV2(def)) return;
  ensureGateCenterV2(project);
  const rec = project.gateRecordsV2[key];
  const now = formatNowMinuteV2();
  rec.status = '已完成';
  rec.confirmedBy = currentUser.name || activePerspective || '使用者';
  rec.confirmedAt = now;
  rec.supervisorStatus = activePerspective === 'VP' || activePerspective === 'Admin' ? '已審核' : rec.supervisorStatus;
  if (rec.supervisorStatus === '已審核') rec.supervisorAt = now;

  const phase = getGatePhaseV2(key);
  const idx = phase.gateKeys.indexOf(key);
  const nextKey = phase.gateKeys[idx + 1];
  if (nextKey && project.gateRecordsV2[nextKey] && project.gateRecordsV2[nextKey].status === '待開始') {
    project.gateRecordsV2[nextKey].status = '進行中';
  }
  updateProjectProgressFromRecordsV2(project);
  addGateNotification({
    projectId,
    title: `${projectId} Gate 更新`,
    message: `${def.name} 已由 ${rec.confirmedBy} 確認完成，時間 ${now}`,
    actor: rec.confirmedBy
  });
  saveDatabase();
  renderNotificationsBell();
  renderProjectGateCenter(projectId);
}

function saveGateNgV2(projectId, gateKey, safeId) {
  const project = db.projects.find(p => p.id === projectId);
  const key = String(gateKey);
  const def = getGateDefinitionV2(key);
  if (!project || !def || !canConfirmGateV2(def)) return;
  ensureGateCenterV2(project);
  const rec = project.gateRecordsV2[key];
  const timeRaw = getInputValue(`ng-time-${safeId}`) || getNowDatetimeLocal();
  const ngTime = formatDatetimeLocalMinute(timeRaw);
  const reason = getInputValue(`ng-reason-${safeId}`) || '未填原因';
  const note = getInputValue(`ng-note-${safeId}`);
  const due = getInputValue(`ng-due-${safeId}`);
  const backGate = key === '9' ? '6' : '2';
  const backPhase = key === '9' ? 'power' : 'plano';
  rec.status = 'NG';
  rec.confirmedBy = currentUser.name || activePerspective || '使用者';
  rec.confirmedAt = ngTime;
  rec.ng = {
    time: ngTime,
    reason,
    note,
    due,
    backGate,
    nextAction: `新增下一版，退回 Gate ${backGate} ${getGateDefinitionV2(backGate).name}`
  };

  project.status = '紅';
  project.cleanGateState = {
    status: 'NG / 等待結果',
    gate: Number(backGate),
    version: getActiveVersionV2(project, backPhase),
    owner: inferOwnerV2(project, backGate),
    nextAction: rec.ng.nextAction,
    ngReason: reason,
    resultDue: due ? `預計 ${due} 有結果` : '等待負責人回填結果時間'
  };
  project.gate = backGate === '6' ? 'G3' : 'G1';
  project.process = { ...(project.process || {}), currentStep: Number(backGate), lastDecision: 'NG' };
  addProjectPhaseVersion(projectId, backPhase);
  if (project.gateRecordsV2[backGate]) {
    project.gateRecordsV2[backGate].status = '進行中';
  }
  addGateNotification({
    projectId,
    title: `${projectId} NG 紀錄`,
    message: `${def.name} NG：${reason}；${rec.ng.nextAction}`,
    actor: rec.confirmedBy
  });
  saveDatabase();
  renderNotificationsBell();
  gateCenterV2State.phaseByProject[projectId] = backPhase;
  renderProjectGateCenter(projectId);
}

function saveGateImageRecordV2(projectId, gateKey, safeId) {
  const project = db.projects.find(p => p.id === projectId);
  if (!project) return;
  ensureGateCenterV2(project);
  const rec = project.gateRecordsV2[String(gateKey)];
  if (!rec || rec.status === '已完成') return;
  rec.imageLink = getInputValue(`image-link-${safeId}`);
  rec.imageNote = getInputValue(`image-note-${safeId}`);
  saveDatabase();
  renderProjectGateCenter(projectId);
}

function updateProjectProgressFromRecordsV2(project) {
  const records = ensureGateCenterV2(project);
  const active = Object.values(records).find(r => r.status === 'NG') || Object.values(records).find(r => r.status === '進行中');
  if (!active) return;
  const key = String(active.gateKey);
  if (!key.startsWith('P')) {
    project.cleanGateState = { ...(project.cleanGateState || {}), gate: Number(key), status: active.status === 'NG' ? 'NG / 等待結果' : '進行中' };
  } else {
    project.cleanGateState = { ...(project.cleanGateState || {}), gate: 10, status: active.status === 'NG' ? 'NG / 等待結果' : '量產中' };
  }
}

function canConfirmGateV2(def) {
  if (!def) return false;
  if (activePerspective === 'Admin' || activePerspective === 'VP') return true;
  if (def.confirmRole === 'PM') return activePerspective === 'PM';
  if (def.confirmRole === 'RD') return activePerspective === 'RD';
  return false;
}

function getGateGuidanceV2(project) {
  const records = ensureGateCenterV2(project);
  const ngRecord = Object.values(records).find(r => r.status === 'NG' && r.ng);
  if (ngRecord) return { reason: ngRecord.ng.reason, nextAction: ngRecord.ng.nextAction };
  const display = getProjectDisplay(project);
  if (display.gate >= 10) return { nextAction: '進入量產分頁，完成量產準備與放行' };
  const next = getGateDefinitionV2(String(display.gate));
  return { nextAction: next ? `完成 ${next.name} 後進入下一 Gate` : '等待下一步指派' };
}

function getPhaseLabelV2(phaseId) {
  return (GATE_PHASES_V2.find(p => p.id === phaseId) || {}).label || phaseId;
}

function stateClassV2(state) {
  const text = String(state || '');
  if (text.includes('NG')) return 'danger';
  if (text.includes('完成') || text.includes('已審核')) return 'success';
  if (text.includes('進行') || text.includes('量產')) return 'info';
  if (text.includes('待')) return 'warning';
  return 'neutral';
}

function safeDomId(text) {
  return String(text).replace(/[^a-zA-Z0-9_-]/g, '-');
}

function getInputValue(id) {
  const el = document.getElementById(id);
  return el ? el.value.trim() : '';
}

function formatNowMinuteV2() {
  const d = new Date();
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}/${pad(d.getMonth() + 1)}/${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function getNowDatetimeLocal() {
  const d = new Date();
  const pad = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function toDatetimeLocalValue(value) {
  if (!value) return '';
  const text = String(value).replace(/\//g, '-').replace(' ', 'T');
  return text.length >= 16 ? text.slice(0, 16) : '';
}

function formatDatetimeLocalMinute(value) {
  if (!value) return formatNowMinuteV2();
  const [date, time] = String(value).split('T');
  return `${date.replace(/-/g, '/')}${time ? ` ${time.slice(0, 5)}` : ''}`;
}

function escapeAttribute(value) {
  return escapeHtml(String(value || '')).replace(/"/g, '&quot;');
}

function renderActionItems() {
  const container = document.getElementById('tab-action-items');
  if (!container) return;
  currentProjectDetailId = null;
  const projects = getAllProjectsForCleanUI();
  container.innerHTML = `
    <section class="simple-page progress-page-v2">
      <div class="page-heading">
        <h1>專案進度</h1>
        <p>給會議快速看：規格、版次、目前進度、時間軸與下一階段。</p>
      </div>
      <div class="phase-tabs progress-tabs">
        <button class="${gateCenterV2State.progressView === 'gantt' ? 'active' : ''}" onclick="setProgressViewV2('gantt')">甘特圖</button>
        <button class="${gateCenterV2State.progressView === 'milestone' ? 'active' : ''}" onclick="setProgressViewV2('milestone')">里程碑</button>
      </div>
      ${gateCenterV2State.progressView === 'milestone' ? renderMilestoneProgressV2(projects) : renderGanttProgressV2(projects)}
    </section>
  `;
}

function setProgressViewV2(view) {
  gateCenterV2State.progressView = view;
  renderActionItems();
}

function renderGanttProgressV2(projects) {
  return `
    <div class="progress-table-card">
      <div class="progress-table-head gantt-grid">
        <span>規格</span><span>版次</span><span>目前進度</span><span>時間軸</span>
      </div>
      ${projects.map(project => {
        const d = getProjectDisplay(project);
        const phase = getCurrentPhaseV2(project);
        const percent = phase === 'demand' ? 15 : phase === 'plano' ? 45 : phase === 'power' ? 75 : 95;
        return `
          <div class="progress-table-row gantt-grid" onclick="openProjectGateCenter('${escapeHtml(project.id)}')">
            <strong>${escapeHtml(d.name)}</strong>
            <span>${escapeHtml(d.version)}</span>
            <span>${escapeHtml(getPhaseLabelV2(phase))} / Gate ${escapeHtml(String(d.gate))}</span>
            <div class="gantt-line"><span style="width:${percent}%"></span></div>
          </div>
        `;
      }).join('')}
    </div>
  `;
}

function renderMilestoneProgressV2(projects) {
  return `
    <div class="progress-table-card">
      <div class="progress-table-head milestone-grid">
        <span>規格</span><span>版次</span><span>目前進度</span><span>下一個階段</span>
      </div>
      ${projects.map(project => {
        const d = getProjectDisplay(project);
        const guidance = getGateGuidanceV2(project);
        return `
          <div class="progress-table-row milestone-grid" onclick="openProjectGateCenter('${escapeHtml(project.id)}')">
            <strong>${escapeHtml(d.name)}</strong>
            <span>${escapeHtml(d.version)}</span>
            <span>${escapeHtml(getPhaseLabelV2(getCurrentPhaseV2(project)))}</span>
            <span>${escapeHtml(d.nextAction || guidance.nextAction)}</span>
          </div>
        `;
      }).join('')}
    </div>
  `;
}

function renderRisksIssues() {
  const container = document.getElementById('tab-risks-issues');
  if (!container) return;
  currentProjectDetailId = null;
  container.innerHTML = `
    <section class="simple-page team-page-v2">
      <div class="page-heading">
        <h1>團隊成員</h1>
        <p>依部門呈現主管與同仁；主管具備分派、編輯與審核權。</p>
      </div>
      <div class="team-dept-grid-v2">
        ${TEAM_STRUCTURE_V2.map(team => `
          <article class="team-dept-card-v2">
            <div class="team-dept-head">
              <div>
                <h2>${escapeHtml(team.dept)}</h2>
                <p>主管：${escapeHtml(team.manager)} ${escapeHtml(team.title)}</p>
              </div>
              <span class="manager-badge">主管權限</span>
            </div>
            <div class="manager-permissions">
              <span>可分派 / 編輯專案</span>
              <span>審核權</span>
              <span>查看部門 Gate</span>
            </div>
            <h3>同仁</h3>
            <div class="member-list-v2">
              ${team.members.map(member => `<span>${escapeHtml(member)}</span>`).join('')}
            </div>
          </article>
        `).join('')}
      </div>
    </section>
  `;
}

function exportProjectMeetingExcel(projectId) {
  const project = db.projects.find(p => p.id === projectId);
  if (!project) return;
  ensureGateCenterV2(project);
  const display = getProjectDisplay(project);
  const rows = Object.values(project.gateRecordsV2).map(rec => {
    const def = getGateDefinitionV2(rec.gateKey);
    const phase = getGatePhaseV2(rec.gateKey);
    return {
      '專案名稱': display.name,
      '規格': project.type || '',
      '大項目': phase.label,
      '版次': getActiveVersionV2(project, phase.id),
      'Gate': rec.gateKey,
      'Gate 名稱': def ? def.name : '',
      '狀態': rec.status,
      '負責部門': def ? def.department : rec.department,
      '負責人': rec.owner,
      '確認人': rec.confirmedBy,
      '確認時間': rec.confirmedAt,
      '主管審核': `${rec.supervisor} / ${rec.supervisorStatus}`,
      'NG 時間': rec.ng ? rec.ng.time : '',
      'NG 原因': rec.ng ? rec.ng.reason : '',
      'NG 備註': rec.ng ? rec.ng.note : '',
      '下一步': rec.ng ? rec.ng.nextAction : (display.nextAction || ''),
      '預計完成時間': rec.dueDate || (rec.ng ? rec.ng.due : ''),
      'PM追蹤': display.tracker || '曼聿',
      '圖片 / 附件連結': rec.imageLink || '',
      '備註': rec.note || rec.imageNote || ''
    };
  });

  const fileName = `${project.id}_會議進度報告_${new Date().toISOString().slice(0, 10)}.xlsx`;
  if (typeof XLSX !== 'undefined') {
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Gate進度');
    XLSX.writeFile(wb, fileName);
  } else {
    exportRowsAsCsvV2(rows, fileName.replace('.xlsx', '.csv'));
  }
}

function exportRowsAsCsvV2(rows, fileName) {
  if (!rows.length) return;
  const headers = Object.keys(rows[0]);
  const csv = [headers.join(',')]
    .concat(rows.map(row => headers.map(h => `"${String(row[h] || '').replace(/"/g, '""')}"`).join(',')))
    .join('\n');
  const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(a.href);
}

// ==========================================================================
// 13. Role & Permission v3：CNC / 主管 / 一般員工
// ==========================================================================

const ROLE_OPTIONS_V3 = ['VP', 'PM', 'RD', 'CNC', 'Supervisor', 'Employee', 'Admin'];

const ROLE_LABELS_V3 = {
  VP: '副總/主管',
  PM: '產品 PM',
  RD: '研發 RD',
  CNC: 'CNC / 模仁',
  Supervisor: '主管',
  Employee: '一般員工',
  Admin: '系統管理員'
};

const ROLE_DEMO_USERS_V3 = [
  { username: 'cnc_user', password: '123', role: 'CNC', name: '鍾民', dept: 'CNC / 模仁' },
  { username: 'supervisor_user', password: '123', role: 'Supervisor', name: 'Andy', dept: 'RD 研發部 / 主管' },
  { username: 'employee_user', password: '123', role: 'Employee', name: '媃安', dept: 'PM / 一般員工' }
];

function ensureRoleUsersV3() {
  if (!db || !Array.isArray(db.users)) return;
  ROLE_DEMO_USERS_V3.forEach(user => {
    const existing = db.users.find(u => u.username === user.username);
    if (!existing) db.users.push({ ...user });
  });
}

function ensureRoleSelectV3() {
  const select = document.getElementById('role-perspective-select');
  if (!select) return;
  ROLE_OPTIONS_V3.forEach(role => {
    if (![...select.options].some(option => option.value === role)) {
      const option = document.createElement('option');
      option.value = role;
      option.textContent = `${ROLE_LABELS_V3[role]} 視角`;
      select.appendChild(option);
    }
  });
}

function getAllowedPerspectivesV3() {
  const role = currentUser ? currentUser.role : 'VP';
  if (role === 'Admin') return ROLE_OPTIONS_V3;
  if (role === 'Supervisor' || role === 'VP') return ['Supervisor', 'Employee', role].filter((value, index, arr) => arr.indexOf(value) === index);
  return [role];
}

function checkSession() {
  const sessionUser = sessionStorage.getItem('LENS_USER');
  ensureRoleUsersV3();
  ensureRoleSelectV3();
  if (sessionUser) {
    currentUser = JSON.parse(sessionUser);
    const storedPerspective = sessionStorage.getItem('LENS_PERSPECTIVE');
    const allowed = getAllowedPerspectivesV3();
    activePerspective = allowed.includes(storedPerspective) ? storedPerspective : currentUser.role;
    const select = document.getElementById('role-perspective-select');
    if (select) select.value = activePerspective;
    showApp();
    initApp();
  } else {
    showLogin();
  }
}

function handleLogin() {
  ensureRoleUsersV3();
  ensureRoleSelectV3();
  const userVal = document.getElementById('username').value.trim();
  const passVal = document.getElementById('password').value;
  const foundUser = db.users.find(u => u.username === userVal && u.password === passVal);
  if (foundUser) {
    currentUser = foundUser;
    activePerspective = foundUser.role;
    sessionStorage.setItem('LENS_USER', JSON.stringify(currentUser));
    sessionStorage.setItem('LENS_PERSPECTIVE', activePerspective);
    document.getElementById('role-perspective-select').value = activePerspective;
    document.getElementById('login-error').style.display = 'none';
    saveDatabase();
    showApp();
    initApp();
  } else {
    document.getElementById('login-error').style.display = 'block';
  }
}

function changePerspective(role) {
  ensureRoleSelectV3();
  const allowedRoles = getAllowedPerspectivesV3();
  if (!allowedRoles.includes(role)) {
    activePerspective = currentUser ? currentUser.role : 'VP';
    document.getElementById('role-perspective-select').value = activePerspective;
    return;
  }
  activePerspective = role;
  sessionStorage.setItem('LENS_PERSPECTIVE', role);
  applyPermissions();
  const activeTabItem = document.querySelector('.menu-item.active');
  if (activeTabItem) switchTab(activeTabItem.getAttribute('data-tab'));
}

function initApp() {
  ensureRoleUsersV3();
  ensureRoleSelectV3();
  normalizeGateDefinitionsV3();
  const displayName = document.getElementById('user-display-name');
  const displayDept = document.getElementById('user-display-dept');
  if (displayName) displayName.innerText = currentUser.name || currentUser.username || '使用者';
  if (displayDept) displayDept.innerText = `${currentUser.dept || currentUser.role || ''} · ${ROLE_LABELS_V3[activePerspective] || activePerspective}`;
  prepareCleanNavigation();
  ensureNotificationBell();
  seedGateNotifications();
  applyPermissions();
  switchTab('vp-milestones');
}

function normalizeGateDefinitionsV3() {
  const setGate = (id, patch) => {
    const gate = GATE_STAGES.find(g => String(g.id) === String(id));
    if (gate) Object.assign(gate, patch);
  };
  setGate(2, { department: 'CNC', confirmRole: 'CNC', owner: 'CNC 鍾民' });
  setGate(3, { department: 'CNC', confirmRole: 'CNC', owner: 'CNC 鍾民', description: 'CNC / 模仁完成平光 PP 模製作，建立版本與試作紀錄。' });
  setGate(6, { department: 'CNC', confirmRole: 'CNC', owner: 'CNC 鍾民' });
  setGate(7, { department: 'CNC', confirmRole: 'CNC', owner: 'CNC 鍾民', description: 'CNC / 模仁完成全光度 PP 模製作，建立版本與試作紀錄。' });
}

function applyPermissions() {
  ensureRoleSelectV3();
  const perspectiveSelect = document.getElementById('role-perspective-select');
  if (perspectiveSelect) {
    const allowed = getAllowedPerspectivesV3();
    perspectiveSelect.disabled = currentUser && currentUser.role !== 'Admin' && currentUser.role !== 'Supervisor' && currentUser.role !== 'VP';
    [...perspectiveSelect.options].forEach(option => {
      option.hidden = !allowed.includes(option.value);
    });
  }
  const adminItem = document.getElementById('nav-admin-settings');
  if (adminItem) adminItem.classList.toggle('hidden', activePerspective !== 'Admin');
  const canManage = canManageProjectsV3();
  ['btn-add-project-main', 'btn-add-action-main', 'btn-add-risk-main'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.classList.toggle('hidden', !canManage);
  });
  ['btn-add-version-main', 'btn-import-report-main'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.classList.toggle('hidden', !['RD', 'CNC', 'Supervisor', 'VP', 'Admin'].includes(activePerspective));
  });
}

function canManageProjectsV3() {
  return ['Admin', 'Supervisor', 'VP', 'PM'].includes(activePerspective);
}

function canAssignProjectsV3() {
  return ['Admin', 'Supervisor', 'VP'].includes(activePerspective);
}

function canApproveGateV3() {
  return ['Admin', 'Supervisor', 'VP'].includes(activePerspective);
}

function isAssignedToCurrentUserV3(rec) {
  if (!currentUser || !rec) return false;
  const haystack = `${rec.owner || ''} ${rec.assignedTo || ''}`;
  return haystack.includes(currentUser.name) || haystack.includes(currentUser.username);
}

function ensureGateCenterV2(project) {
  if (!project.gateRecordsV2) project.gateRecordsV2 = {};
  const display = getProjectDisplay(project);
  const currentGate = String(display.gate || inferGateFromProject(project));
  const allKeys = GATE_PHASES_V2.flatMap(p => p.gateKeys);
  allKeys.forEach(key => {
    const def = getGateDefinitionV2(key);
    if (!def) return;
    if (!project.gateRecordsV2[key]) {
      project.gateRecordsV2[key] = {
        gateKey: key,
        status: getInitialGateStatusV2(key, currentGate, display.status),
        owner: inferOwnerV2(project, key),
        assignedTo: inferOwnerV2(project, key),
        department: def.department,
        dueDate: inferDueDateV2(project, key),
        confirmedBy: '',
        confirmedAt: '',
        supervisor: inferSupervisorV2(key),
        supervisorStatus: '待審核',
        supervisorAt: '',
        note: '',
        imageLink: '',
        imageNote: '',
        ng: null
      };
    } else {
      project.gateRecordsV2[key].department = def.department;
      project.gateRecordsV2[key].supervisor = project.gateRecordsV2[key].supervisor || inferSupervisorV2(key);
      if (!project.gateRecordsV2[key].assignedTo) project.gateRecordsV2[key].assignedTo = project.gateRecordsV2[key].owner || inferOwnerV2(project, key);
      if (!project.gateRecordsV2[key].confirmedBy && !project.gateRecordsV2[key].manualOwner) {
        project.gateRecordsV2[key].owner = inferOwnerV2(project, key);
        project.gateRecordsV2[key].assignedTo = inferOwnerV2(project, key);
      }
    }
  });
  if (!project.phaseVersionsV2) {
    project.phaseVersionsV2 = {
      demand: ['需求確認'],
      plano: ['V1'],
      power: [display.version || project.currentVersion || 'V1'],
      production: ['量產準備']
    };
  }
  return project.gateRecordsV2;
}

function inferOwnerV2(project, gateKey) {
  const key = String(gateKey);
  if (key === '1') return '業務窗口 / PM 曼聿';
  if (['2', '3', '6', '7'].includes(key)) return 'CNC 鍾民';
  if (['4', '5', '8', '9'].includes(key)) return project.id && project.id.includes('142') ? 'RD 齊政' : 'RD 世偉';
  if (key === 'P1') return 'PM 曼聿 / 生產主管';
  if (key === 'P2') return '部門主管';
  if (key === 'P3') return '生產主管';
  return 'PM 曼聿 / 業務窗口';
}

function canConfirmGateRecordV3(project, def, rec) {
  if (!def || !rec) return false;
  if (activePerspective === 'Admin') return true;
  if (activePerspective === 'Employee') return isAssignedToCurrentUserV3(rec);
  if (canApproveGateV3()) return true;
  if (def.confirmRole === 'PM') return activePerspective === 'PM';
  if (def.confirmRole === 'RD') return activePerspective === 'RD';
  if (def.confirmRole === 'CNC') return activePerspective === 'CNC';
  return false;
}

function renderProjectOverviewCard(project) {
  const d = getProjectDisplay(project);
  const gate = GATE_STAGES[d.gate - 1] || GATE_STAGES[0];
  const canManage = canManageProjectsV3();
  const canAssign = canAssignProjectsV3();
  const ngBlock = d.ngReason ? `
    <div class="ng-mini-block">
      <strong>NG 原因</strong>
      <span>${escapeHtml(d.ngReason)}</span>
      <em>${escapeHtml(d.resultDue || '等待結果')}</em>
    </div>` : '';
  return `
    <article class="simple-project-card">
      <div class="project-card-top" onclick="openProjectGateCenter('${escapeHtml(project.id)}')">
        <h3>${escapeHtml(d.name)}</h3>
        <span class="status-pill ${statusClass(d.status)}">${escapeHtml(d.status)}</span>
      </div>
      <div onclick="openProjectGateCenter('${escapeHtml(project.id)}')">
        <div class="gate-progress-mini"><span style="width:${Math.round((d.gate / 9) * 100)}%"></span></div>
        <p class="project-current">Gate ${gate.id} · ${escapeHtml(gate.name)}</p>
        ${ngBlock}
        <div class="project-card-footer">
          <span>版本 ${escapeHtml(d.version)}</span>
          <span>${escapeHtml(d.owner)}</span>
        </div>
      </div>
      ${canManage ? `
        <div class="project-manage-row">
          <button class="btn btn-outline btn-xs" onclick="openProjectGateCenter('${escapeHtml(project.id)}')">查看 / 編輯</button>
          ${canAssign ? `<button class="btn btn-primary btn-xs" onclick="assignProjectV3('${escapeHtml(project.id)}')">分派專案</button>` : ''}
        </div>
      ` : ''}
    </article>
  `;
}

function assignProjectV3(projectId) {
  if (!canAssignProjectsV3()) return;
  const project = db.projects.find(p => p.id === projectId);
  if (!project) return;
  const current = project.assignmentV3 || {};
  const pm = prompt('分派 PM / 追蹤者', current.pm || '曼聿');
  if (pm === null) return;
  const rd = prompt('分派 RD 負責人', current.rd || (project.id.includes('142') ? '齊政' : '世偉'));
  if (rd === null) return;
  const cnc = prompt('分派 CNC / 模仁負責人', current.cnc || '鍾民');
  if (cnc === null) return;
  const priority = prompt('專案優先順序（1=第一優先，2=第二優先，3=第三優先）', String(project.priority || 2));
  if (priority === null) return;
  project.assignmentV3 = { pm, rd, cnc };
  project.owner = pm;
  project.priority = Math.max(1, Math.min(3, Number(priority) || 2));
  ensureGateCenterV2(project);
  Object.entries(project.gateRecordsV2).forEach(([key, rec]) => {
    if (['2', '3', '6', '7'].includes(String(key))) rec.owner = rec.assignedTo = `CNC ${cnc}`;
    if (['4', '5', '8', '9'].includes(String(key))) rec.owner = rec.assignedTo = `RD ${rd}`;
    if (String(key) === '1') rec.owner = rec.assignedTo = `業務窗口 / PM ${pm}`;
    rec.manualOwner = true;
  });
  addGateNotification({
    projectId,
    title: `${projectId} 已重新分派`,
    message: `PM：${pm}；RD：${rd}；CNC：${cnc}；優先順序：${project.priority}`,
    actor: currentUser.name || activePerspective
  });
  saveDatabase();
  renderProjectMaster();
  renderNotificationsBell();
}

function renderGateCardV2(project, gateKey, activeVersion) {
  const key = String(gateKey);
  const def = getGateDefinitionV2(key);
  const records = ensureGateCenterV2(project);
  const rec = records[key];
  if (!def || !rec) return '';
  const canConfirm = canConfirmGateRecordV3(project, def, rec);
  const canAssign = canAssignProjectsV3();
  const canApprove = canApproveGateV3();
  const isDone = rec.status === '已完成';
  const isDecisionGate = key === '5' || key === '9';
  const safe = safeDomId(`${project.id}-${key}`);
  const confirmDisabled = !canConfirm || isDone ? 'disabled' : '';
  const lockedText = isDone ? '已完成，不可編輯' : canConfirm ? '你可以確認此 Gate' : `僅 ${def.department} 負責人 / 被分派者可確認`;
  const imageBlock = key === '4' ? renderImageRecordBlockV2(project, key, rec, safe, isDone) : '';
  const ngBlock = isDecisionGate ? renderNgBlockV2(project, key, rec, safe, canConfirm && !isDone) : '';
  const managementBlock = (canAssign || canApprove) ? renderGateManagementV3(project, key, rec, safe, canAssign, canApprove) : '';

  return `
    <article class="gate-row-card ${stateClassV2(rec.status)} ${isDone ? 'readonly' : ''}">
      <div class="gate-row-main">
        <div class="gate-row-title">
          <span class="gate-number">${key.startsWith('P') ? key : `Gate ${key}`}</span>
          <div>
            <h3>${escapeHtml(def.name)}</h3>
            <p>${escapeHtml(def.description || '')}</p>
          </div>
        </div>
        <span class="status-pill ${stateClassV2(rec.status)}">${escapeHtml(rec.status)}</span>
      </div>

      <div class="gate-row-fields">
        <div><span>版本</span><strong>${escapeHtml(activeVersion)}</strong></div>
        <div><span>部門</span><strong>${escapeHtml(def.department)}</strong></div>
        <div><span>負責人</span><strong>${escapeHtml(rec.owner)}</strong></div>
        <div><span>預計完成</span><strong>${escapeHtml(rec.dueDate || '待排程')}</strong></div>
        <div><span>確認人</span><strong>${escapeHtml(rec.confirmedBy || '尚未確認')}</strong></div>
        <div><span>確認時間</span><strong>${escapeHtml(rec.confirmedAt || '尚未確認')}</strong></div>
        <div><span>主管審核</span><strong>${escapeHtml(rec.supervisor)} / ${escapeHtml(rec.supervisorStatus)}</strong></div>
      </div>

      ${managementBlock}
      ${imageBlock}
      ${ngBlock}

      <div class="gate-row-actions">
        <span>${escapeHtml(lockedText)}</span>
        <button class="btn btn-primary" ${confirmDisabled} onclick="confirmProjectGateV2('${escapeHtml(project.id)}', '${key}', 'OK')">確認完成</button>
        ${isDecisionGate ? `<button class="btn btn-outline danger" ${confirmDisabled} onclick="showNgFormV2('${safe}')">NG 填寫原因</button>` : ''}
      </div>
    </article>
  `;
}

function renderGateManagementV3(project, key, rec, safe, canAssign, canApprove) {
  return `
    <div class="gate-management-block">
      ${canAssign ? `
        <label>分派負責人
          <input id="assign-owner-${safe}" value="${escapeAttribute(rec.owner || '')}" placeholder="例如：CNC 鍾民 / RD 齊政">
        </label>
        <label>預計完成時間
          <input id="assign-due-${safe}" type="date" value="${escapeAttribute(rec.dueDate || '')}">
        </label>
        <button class="btn btn-outline" onclick="saveGateAssignmentV3('${escapeHtml(project.id)}', '${escapeHtml(key)}', '${safe}')">儲存分派</button>
      ` : ''}
      ${canApprove ? `
        <button class="btn btn-primary" onclick="approveGateV3('${escapeHtml(project.id)}', '${escapeHtml(key)}')">主管審核通過</button>
      ` : ''}
    </div>
  `;
}

function saveGateAssignmentV3(projectId, gateKey, safeId) {
  if (!canAssignProjectsV3()) return;
  const project = db.projects.find(p => p.id === projectId);
  if (!project) return;
  ensureGateCenterV2(project);
  const rec = project.gateRecordsV2[String(gateKey)];
  if (!rec) return;
  const owner = getInputValue(`assign-owner-${safeId}`);
  const due = getInputValue(`assign-due-${safeId}`);
  if (owner) {
    rec.owner = owner;
    rec.assignedTo = owner;
    rec.manualOwner = true;
  }
  if (due) rec.dueDate = due;
  addGateNotification({
    projectId,
    title: `${projectId} Gate 已分派`,
    message: `${gateKey} 分派給 ${rec.owner}，預計 ${rec.dueDate || '待排程'}`,
    actor: currentUser.name || activePerspective
  });
  saveDatabase();
  renderNotificationsBell();
  renderProjectGateCenter(projectId);
}

function approveGateV3(projectId, gateKey) {
  if (!canApproveGateV3()) return;
  const project = db.projects.find(p => p.id === projectId);
  if (!project) return;
  ensureGateCenterV2(project);
  const rec = project.gateRecordsV2[String(gateKey)];
  if (!rec) return;
  rec.supervisorStatus = '已審核';
  rec.supervisorAt = formatNowMinuteV2();
  rec.supervisorBy = currentUser.name || activePerspective;
  addGateNotification({
    projectId,
    title: `${projectId} Gate 主管審核`,
    message: `${gateKey} 已由 ${rec.supervisorBy} 審核通過`,
    actor: rec.supervisorBy
  });
  saveDatabase();
  renderNotificationsBell();
  renderProjectGateCenter(projectId);
}

function confirmProjectGateV2(projectId, gateKey, decision) {
  const project = db.projects.find(p => p.id === projectId);
  const key = String(gateKey);
  const def = getGateDefinitionV2(key);
  if (!project || !def) return;
  ensureGateCenterV2(project);
  const rec = project.gateRecordsV2[key];
  if (!canConfirmGateRecordV3(project, def, rec)) return;
  const now = formatNowMinuteV2();
  rec.status = '已完成';
  rec.confirmedBy = currentUser.name || activePerspective || '使用者';
  rec.confirmedAt = now;
  rec.supervisorStatus = canApproveGateV3() ? '已審核' : rec.supervisorStatus;
  if (rec.supervisorStatus === '已審核') rec.supervisorAt = now;
  const phase = getGatePhaseV2(key);
  const idx = phase.gateKeys.indexOf(key);
  const nextKey = phase.gateKeys[idx + 1];
  if (nextKey && project.gateRecordsV2[nextKey] && project.gateRecordsV2[nextKey].status === '待開始') {
    project.gateRecordsV2[nextKey].status = '進行中';
  }
  updateProjectProgressFromRecordsV2(project);
  addGateNotification({
    projectId,
    title: `${projectId} Gate 更新`,
    message: `${def.name} 已由 ${rec.confirmedBy} 確認完成，時間 ${now}`,
    actor: rec.confirmedBy
  });
  saveDatabase();
  renderNotificationsBell();
  renderProjectGateCenter(projectId);
}

function saveGateNgV2(projectId, gateKey, safeId) {
  const project = db.projects.find(p => p.id === projectId);
  const key = String(gateKey);
  const def = getGateDefinitionV2(key);
  if (!project || !def) return;
  ensureGateCenterV2(project);
  const rec = project.gateRecordsV2[key];
  if (!canConfirmGateRecordV3(project, def, rec)) return;
  const timeRaw = getInputValue(`ng-time-${safeId}`) || getNowDatetimeLocal();
  const ngTime = formatDatetimeLocalMinute(timeRaw);
  const reason = getInputValue(`ng-reason-${safeId}`) || '未填原因';
  const note = getInputValue(`ng-note-${safeId}`);
  const due = getInputValue(`ng-due-${safeId}`);
  const backGate = key === '9' ? '6' : '2';
  const backPhase = key === '9' ? 'power' : 'plano';
  rec.status = 'NG';
  rec.confirmedBy = currentUser.name || activePerspective || '使用者';
  rec.confirmedAt = ngTime;
  rec.ng = {
    time: ngTime,
    reason,
    note,
    due,
    backGate,
    nextAction: `新增下一版，退回 Gate ${backGate} ${getGateDefinitionV2(backGate).name}`
  };
  project.status = '紅';
  project.cleanGateState = {
    status: 'NG / 等待結果',
    gate: Number(backGate),
    version: getActiveVersionV2(project, backPhase),
    owner: inferOwnerV2(project, backGate),
    nextAction: rec.ng.nextAction,
    ngReason: reason,
    resultDue: due ? `預計 ${due} 有結果` : '等待負責人回填結果時間'
  };
  project.gate = backGate === '6' ? 'G3' : 'G1';
  project.process = { ...(project.process || {}), currentStep: Number(backGate), lastDecision: 'NG' };
  addProjectPhaseVersion(projectId, backPhase);
  if (project.gateRecordsV2[backGate]) project.gateRecordsV2[backGate].status = '進行中';
  addGateNotification({
    projectId,
    title: `${projectId} NG 紀錄`,
    message: `${def.name} NG：${reason}；${rec.ng.nextAction}`,
    actor: rec.confirmedBy
  });
  saveDatabase();
  renderNotificationsBell();
  gateCenterV2State.phaseByProject[projectId] = backPhase;
  renderProjectGateCenter(projectId);
}

function getRoleTaskCards() {
  const redProjects = db.projects.filter(p => normalizeStatus(p.status).includes('NG') || String(p.status).includes('紅'));
  const myAssignedGates = getMyAssignedGatesV3();
  if (activePerspective === 'CNC') {
    return [
      { label: 'CNC 待確認 Gate', count: myAssignedGates.length || '看全部', tone: 'blue', title: '模仁 / PP 模確認', desc: '優先處理 Gate 2、3、6、7。', cta: '看專案總覽', action: "switchTab('project-master')" },
      { label: 'NG / 待處理', count: redProjects.length, tone: 'red', title: '異常專案', desc: '確認是否需要重新設計或重啟版本。', cta: '看紅燈', action: "switchTab('project-master')" }
    ];
  }
  if (activePerspective === 'Supervisor') {
    return [
      { label: '待審核 / 分派', count: myAssignedGates.length || db.projects.length, tone: 'purple', title: '主管工作台', desc: '可編輯專案、分派 Gate 負責人並審核。', cta: '看專案', action: "switchTab('project-master')" },
      { label: 'NG 決策', count: redProjects.length, tone: 'red', title: '需要主管拍板', desc: '確認退回 Gate、下一版與負責人。', cta: '看異常', action: "switchTab('project-master')" }
    ];
  }
  if (activePerspective === 'Employee') {
    return [
      { label: '我的 Gate', count: myAssignedGates.length, tone: 'green', title: '被分派的工作', desc: '只能更新自己負責的 Gate。', cta: '看我的專案', action: "switchTab('project-master')" }
    ];
  }
  return [
    { label: '進行中專案', count: db.projects.length, tone: 'blue', title: '專案總覽', desc: '查看所有專案 Gate 進度。', cta: '看專案', action: "switchTab('project-master')" },
    { label: 'NG / 風險', count: redProjects.length, tone: 'red', title: '異常追蹤', desc: '優先處理卡關專案。', cta: '看異常', action: "switchTab('project-master')" }
  ];
}

function getMyAssignedGatesV3() {
  if (!currentUser) return [];
  const list = [];
  (db.projects || []).forEach(project => {
    ensureGateCenterV2(project);
    Object.values(project.gateRecordsV2 || {}).forEach(rec => {
      if (rec.status !== '已完成' && isAssignedToCurrentUserV3(rec)) {
        list.push({ project, rec });
      }
    });
  });
  return list;
}

function getDepartmentChecklist(role) {
  const lists = {
    PM: [
      { title: '確認每個專案都有下一步', desc: 'PM 應確認 Gate、負責人、期限與會議報告資訊完整。', required: true },
      { title: '匯出會議 Excel', desc: '開會前輸出專案進度。', required: false }
    ],
    RD: [
      { title: '更新 RD Gate', desc: 'Gate 4、5、8、9 需要 RD 確認設計、試戴與 NG 原因。', required: true }
    ],
    CNC: [
      { title: '更新 CNC Gate', desc: 'Gate 2、3、6、7 由 CNC / 模仁確認模仁與 PP 模進度。', required: true }
    ],
    Supervisor: [
      { title: '分派專案與 Gate', desc: '主管可分派負責人、調整期限與審核 Gate。', required: true },
      { title: '審核 NG 下一步', desc: 'NG 後確認退回 Gate、下一版與負責人。', required: true }
    ],
    Employee: [
      { title: '完成被分派 Gate', desc: '一般員工只需處理自己被分派的 Gate。', required: true }
    ],
    VP: [
      { title: '審核關鍵 Gate', desc: '確認卡關、量產與跨部門決策。', required: true }
    ]
  };
  return lists[role] || lists.PM;
}
