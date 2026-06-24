// ==========================================================================
// 鏡片產品開發多專案追蹤系統 - 前端核心業務邏輯
// ==========================================================================

let db = null;
let currentUser = null;
let activePerspective = 'VP'; // 預設副總/主管視角
let isServerMode = false;
let currentUploadedData = null; // 暫存匯入的解析資料
let currentProjectDetailId = null;
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
      <td>${proj.owner}</td>
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
      <div><span>主要 Owner</span><strong>${escapeHtml(proj.owner)}</strong></div>
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
        <small>${escapeHtml(stage.owner)}</small>
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
      <div><dt>主要負責</dt><dd>${escapeHtml(stage.owner)}</dd></div>
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
        <option value="VP" ${u.role === 'VP' ? 'selected' : ''}>VP</option>
      </select>
    `;
    
    let permissionsDesc = "";
    if (u.role === 'Admin') permissionsDesc = "系統全權限（帳號、專案、簽核）";
    else if (u.role === 'PM') permissionsDesc = "專案主檔、時程甘特圖編輯與指派權限";
    else if (u.role === 'RD') permissionsDesc = "版本履歷登錄、試戴報告問卷解析匯入";
    else if (u.role === 'VP') permissionsDesc = "高階里程碑檢視與 Gate 簽核核准";
    
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
