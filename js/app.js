// ============================================================
//  Raimak LMS — App Logic & UI
// ============================================================

// ── App State ─────────────────────────────────────────────────
const State = {
  leads:       [],
  contractors: [],
  activityLog: [],
  currentView: "dashboard",
  filters: { status: "all", search: "", assignedTo: "all" },
  editingLeadId: null,
  loading: false,
};

// ── Boot ──────────────────────────────────────────────────────
window.addEventListener("DOMContentLoaded", async () => {
  try {
    const redirectResult = await Auth.init();

    if (!Auth.isSignedIn()) {
      showLoginScreen();
      return;
    }

    if (redirectResult) {
      // Came back from MSAL redirect — clean the URL
      window.history.replaceState({}, document.title, window.location.pathname);
    }

    showAppShell();
    await loadAllData();
    renderDashboard();
  } catch (err) {
    console.error("Boot error:", err);
    showLoginScreen();
  }
});

// ── Load Data ─────────────────────────────────────────────────
async function loadAllData() {
  setLoading(true);
  try {
    const [rawLeads, contractors, activityLog] = await Promise.all([
      Graph.getLeads(),
      Graph.getContractors(),
      Graph.getActivityLog(),
    ]);

    State.contractors = contractors;
    State.leads       = Graph.applyBusinessRules(rawLeads, contractors);
    State.activityLog = activityLog;
  } catch (err) {
    UI.showToast("Failed to load data from SharePoint: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

// ============================================================
//  SCREENS
// ============================================================

function showLoginScreen() {
  document.getElementById("app").innerHTML = `
    <div class="login-screen">
      <div class="login-card">
        <div class="login-logo">
          <svg width="48" height="48" viewBox="0 0 48 48" fill="none">
            <rect width="48" height="48" rx="12" fill="#F6A623"/>
            <path d="M12 36V12h8l8 12 8-12h8v24h-8V22l-8 12-8-12v14z" fill="#0D0F14"/>
          </svg>
          <span>Raimak</span>
        </div>
        <h1>Lead Management</h1>
        <p>Sign in with your Raimak Microsoft account to access the system.</p>
        <button class="btn-primary btn-lg" onclick="Auth.signIn()">
          <svg width="20" height="20" viewBox="0 0 21 21" fill="none" style="margin-right:8px">
            <path d="M10 1H1v9h9V1zM20 1h-9v9h9V1zM10 11H1v9h9v-9zM20 11h-9v9h9v-9z" fill="currentColor"/>
          </svg>
          Sign in with Microsoft
        </button>
        <p class="login-version">v${Config.rules.appVersion} · Raimak Leadship</p>
      </div>
    </div>
  `;
}

function showAppShell() {
  const user = Auth.getUser();
  document.getElementById("app").innerHTML = `
    <div class="app-shell">

      <!-- Sidebar -->
      <aside class="sidebar" id="sidebar">
        <div class="sidebar-brand">
          <svg width="36" height="36" viewBox="0 0 48 48" fill="none">
            <rect width="48" height="48" rx="10" fill="#F6A623"/>
            <path d="M12 36V12h8l8 12 8-12h8v24h-8V22l-8 12-8-12v14z" fill="#0D0F14"/>
          </svg>
          <span>Raimak</span>
        </div>

        <nav class="sidebar-nav">
          <a class="nav-item active" data-view="dashboard" onclick="navigate('dashboard')">
            <svg width="18" height="18" fill="none" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1" fill="currentColor"/><rect x="14" y="3" width="7" height="7" rx="1" fill="currentColor"/><rect x="3" y="14" width="7" height="7" rx="1" fill="currentColor"/><rect x="14" y="14" width="7" height="7" rx="1" fill="currentColor"/></svg>
            Dashboard
          </a>
          <a class="nav-item" data-view="leads" onclick="navigate('leads')">
            <svg width="18" height="18" fill="none" viewBox="0 0 24 24"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><circle cx="9" cy="7" r="4" stroke="currentColor" stroke-width="2"/><path d="M23 21v-2a4 4 0 0 0-3-3.87" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><path d="M16 3.13a4 4 0 0 1 0 7.75" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
            Leads
            <span class="badge" id="badge-leads"></span>
          </a>
          <a class="nav-item" data-view="contractors" onclick="navigate('contractors')">
            <svg width="18" height="18" fill="none" viewBox="0 0 24 24"><rect x="2" y="7" width="20" height="14" rx="2" stroke="currentColor" stroke-width="2"/><path d="M16 7V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v2" stroke="currentColor" stroke-width="2"/></svg>
            Contractors
          </a>
          <a class="nav-item" data-view="activity" onclick="navigate('activity')">
            <svg width="18" height="18" fill="none" viewBox="0 0 24 24"><polyline points="22,12 18,12 15,21 9,3 6,12 2,12" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>
            Activity Log
          </a>
        </nav>

        <div class="sidebar-footer">
          <div class="user-info">
            <div class="user-avatar">${(user?.name || "U")[0].toUpperCase()}</div>
            <div class="user-meta">
              <span class="user-name">${user?.name || "User"}</span>
              <span class="user-email">${user?.email || ""}</span>
            </div>
          </div>
          <button class="btn-ghost" onclick="Auth.signOut()" title="Sign Out">
            <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><polyline points="16,17 21,12 16,7" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><line x1="21" y1="12" x2="9" y2="12" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
          </button>
        </div>
      </aside>

      <!-- Main -->
      <main class="main-content" id="main-content">
        <div class="loading-overlay" id="loading-overlay" style="display:none">
          <div class="spinner"></div>
        </div>
      </main>
    </div>

    <!-- Toast -->
    <div id="toast-container"></div>

    <!-- Modal -->
    <div id="modal-overlay" class="modal-overlay" style="display:none" onclick="closeModal(event)">
      <div class="modal" id="modal"></div>
    </div>
  `;
}

// ============================================================
//  NAVIGATION
// ============================================================

function navigate(view) {
  State.currentView = view;
  document.querySelectorAll(".nav-item").forEach(el => el.classList.remove("active"));
  document.querySelector(`[data-view="${view}"]`)?.classList.add("active");

  switch (view) {
    case "dashboard":   renderDashboard();   break;
    case "leads":       renderLeads();       break;
    case "contractors": renderContractors(); break;
    case "activity":    renderActivity();    break;
  }
}

// ============================================================
//  DASHBOARD
// ============================================================

function renderDashboard() {
  const { leads, contractors } = State;
  const { maxLeadsPerAgent, recycleAfterDays } = Config.rules;

  const total       = leads.length;
  const active      = leads.filter(l => !["Won","Lost","Recycled"].includes(l.status)).length;
  const won         = leads.filter(l => l.status === "Won").length;
  const needRecycle = leads.filter(l => l.flags?.includes("needs_recycle")).length;
  const overloaded  = leads.filter(l => l.flags?.includes("agent_overloaded")).length;
  const coolOff     = leads.filter(l => l.flags?.includes("cool_off")).length;

  const statusCounts = {};
  Config.leadStatuses.forEach(s => {
    statusCounts[s] = leads.filter(l => l.status === s).length;
  });

  // Recent leads (last 5)
  const recentLeads = [...leads]
    .sort((a,b) => new Date(b.createdAt) - new Date(a.createdAt))
    .slice(0, 5);

  // Agent load
  const agentLoad = contractors.map(c => {
    const count = leads.filter(l => l.assignedTo === c.name && !["Won","Lost","Recycled"].includes(l.status)).length;
    const pct   = Math.min(100, Math.round((count / maxLeadsPerAgent) * 100));
    return { ...c, count, pct };
  }).sort((a,b) => b.count - a.count).slice(0, 6);

  const main = document.getElementById("main-content");
  main.innerHTML = `
    <div class="view-header">
      <div>
        <h1 class="view-title">Dashboard</h1>
        <span class="view-subtitle">Raimak Lead Management · v${Config.rules.appVersion}</span>
      </div>
      <button class="btn-primary" onclick="openAddLeadModal()">
        <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"/><line x1="5" y1="12" x2="19" y2="12" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"/></svg>
        Add Lead
      </button>
    </div>

    <!-- KPI Cards -->
    <div class="kpi-grid">
      <div class="kpi-card kpi-primary">
        <span class="kpi-label">Total Leads</span>
        <span class="kpi-value">${total}</span>
        <span class="kpi-sub">${active} active</span>
      </div>
      <div class="kpi-card kpi-success">
        <span class="kpi-label">Won</span>
        <span class="kpi-value">${won}</span>
        <span class="kpi-sub">${total ? Math.round((won/total)*100) : 0}% conversion</span>
      </div>
      <div class="kpi-card ${needRecycle > 0 ? 'kpi-warn' : 'kpi-neutral'}">
        <span class="kpi-label">Needs Recycle</span>
        <span class="kpi-value">${needRecycle}</span>
        <span class="kpi-sub">>${recycleAfterDays} days inactive</span>
      </div>
      <div class="kpi-card ${coolOff > 0 ? 'kpi-info' : 'kpi-neutral'}">
        <span class="kpi-label">In Cool-Off</span>
        <span class="kpi-value">${coolOff}</span>
        <span class="kpi-sub">${Config.rules.coolOffDays}-day rule</span>
      </div>
    </div>

    ${overloaded > 0 ? `
    <div class="alert alert-warn">
      <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z" stroke="currentColor" stroke-width="2"/><line x1="12" y1="9" x2="12" y2="13" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><line x1="12" y1="17" x2="12.01" y2="17" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
      <strong>${overloaded} lead(s)</strong> are assigned to agents already at the ${maxLeadsPerAgent}-lead limit.
    </div>` : ""}

    <div class="two-col">
      <!-- Status breakdown -->
      <div class="card">
        <div class="card-header">
          <h2 class="card-title">Pipeline Status</h2>
        </div>
        <div class="status-breakdown">
          ${Config.leadStatuses.map(s => `
            <div class="status-row">
              <span class="status-badge status-${s.toLowerCase().replace(/\s+/g,'-')}">${s}</span>
              <div class="status-bar-wrap">
                <div class="status-bar" style="width:${total ? (statusCounts[s]/total)*100 : 0}%"></div>
              </div>
              <span class="status-count">${statusCounts[s]}</span>
            </div>
          `).join("")}
        </div>
      </div>

      <!-- Agent load -->
      <div class="card">
        <div class="card-header">
          <h2 class="card-title">Agent Load</h2>
          <span class="card-meta">Max ${maxLeadsPerAgent} leads/agent</span>
        </div>
        <div class="agent-load-list">
          ${agentLoad.length ? agentLoad.map(a => `
            <div class="agent-row">
              <div class="agent-avatar">${a.name[0].toUpperCase()}</div>
              <div class="agent-info">
                <span class="agent-name">${a.name}</span>
                <div class="load-bar-wrap">
                  <div class="load-bar ${a.pct >= 100 ? 'load-full' : a.pct >= 80 ? 'load-high' : ''}"
                       style="width:${a.pct}%"></div>
                </div>
              </div>
              <span class="agent-count ${a.count >= maxLeadsPerAgent ? 'text-danger' : ''}">${a.count}/${maxLeadsPerAgent}</span>
            </div>
          `).join("") : '<p class="empty-state">No contractors found.</p>'}
        </div>
      </div>
    </div>

    <!-- Recent Leads -->
    <div class="card">
      <div class="card-header">
        <h2 class="card-title">Recent Leads</h2>
        <button class="btn-ghost-sm" onclick="navigate('leads')">View all →</button>
      </div>
      ${renderLeadsTable(recentLeads, true)}
    </div>
  `;

  updateBadges();
}

// ============================================================
//  LEADS VIEW
// ============================================================

function renderLeads() {
  const main = document.getElementById("main-content");
  const contractors = State.contractors.map(c => c.name);

  main.innerHTML = `
    <div class="view-header">
      <div>
        <h1 class="view-title">Leads</h1>
        <span class="view-subtitle">${State.leads.length} total leads</span>
      </div>
      <div class="header-actions">
        <button class="btn-ghost" onclick="refreshData()" title="Refresh">
          <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><polyline points="23,4 23,10 17,10" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><polyline points="1,20 1,14 7,14" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>
          Refresh
        </button>
        <button class="btn-ghost" onclick="exportCSV()">
          <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><polyline points="7,10 12,15 17,10" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><line x1="12" y1="15" x2="12" y2="3" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
          Export CSV
        </button>
        <button class="btn-primary" onclick="openAddLeadModal()">
          <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"/><line x1="5" y1="12" x2="19" y2="12" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"/></svg>
          Add Lead
        </button>
      </div>
    </div>

    <!-- Filters -->
    <div class="filters-bar">
      <div class="search-wrap">
        <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8" stroke="currentColor" stroke-width="2"/><line x1="21" y1="21" x2="16.65" y2="16.65" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
        <input type="text" id="search-input" class="search-input" placeholder="Search by name, company, email…"
               value="${State.filters.search}" oninput="applyFilters()">
      </div>
      <select class="filter-select" id="filter-status" onchange="applyFilters()">
        <option value="all">All Statuses</option>
        ${Config.leadStatuses.map(s => `<option value="${s}" ${State.filters.status===s?'selected':''}>${s}</option>`).join("")}
      </select>
      <select class="filter-select" id="filter-agent" onchange="applyFilters()">
        <option value="all">All Agents</option>
        ${contractors.map(c => `<option value="${c}" ${State.filters.assignedTo===c?'selected':''}>${c}</option>`).join("")}
      </select>
    </div>

    <div class="card" id="leads-table-wrap">
      ${renderLeadsTable(getFilteredLeads())}
    </div>
  `;
}

function getFilteredLeads() {
  let leads = [...State.leads];
  const { status, search, assignedTo } = State.filters;

  if (status !== "all")      leads = leads.filter(l => l.status === status);
  if (assignedTo !== "all")  leads = leads.filter(l => l.assignedTo === assignedTo);
  if (search.trim()) {
    const q = search.toLowerCase();
    leads = leads.filter(l =>
      l.name.toLowerCase().includes(q)    ||
      l.company.toLowerCase().includes(q) ||
      l.email.toLowerCase().includes(q)
    );
  }
  return leads;
}

function applyFilters() {
  State.filters.search     = document.getElementById("search-input")?.value || "";
  State.filters.status     = document.getElementById("filter-status")?.value || "all";
  State.filters.assignedTo = document.getElementById("filter-agent")?.value  || "all";

  const filtered = getFilteredLeads();
  const wrap = document.getElementById("leads-table-wrap");
  if (wrap) wrap.innerHTML = renderLeadsTable(filtered);
}

function renderLeadsTable(leads, compact = false) {
  if (!leads.length) {
    return `<div class="empty-state"><p>No leads found.</p></div>`;
  }

  return `
    <div class="table-wrap">
      <table class="data-table">
        <thead>
          <tr>
            <th>Name</th>
            <th>Company</th>
            ${compact ? "" : "<th>Email</th>"}
            <th>Status</th>
            <th>Assigned To</th>
            <th>Last Contacted</th>
            ${compact ? "" : "<th>Flags</th><th></th>"}
          </tr>
        </thead>
        <tbody>
          ${leads.map(lead => `
            <tr class="lead-row ${lead.flags?.includes('needs_recycle') ? 'row-warn' : ''}"
                onclick="openEditLeadModal('${lead.id}')">
              <td class="td-name">
                <span class="lead-name">${escHtml(lead.name)}</span>
                ${lead.source ? `<span class="lead-source">${escHtml(lead.source)}</span>` : ""}
              </td>
              <td>${escHtml(lead.company)}</td>
              ${compact ? "" : `<td class="td-email">${escHtml(lead.email)}</td>`}
              <td><span class="status-badge status-${lead.status.toLowerCase().replace(/\s+/g,'-')}">${lead.status}</span></td>
              <td>${escHtml(lead.assignedTo || "—")}</td>
              <td>${formatDate(lead.lastContacted) || formatDate(lead.createdAt) || "—"}</td>
              ${compact ? "" : `
              <td class="td-flags">
                ${(lead.flags || []).map(f => `<span class="flag flag-${f}">${flagLabel(f)}</span>`).join("")}
              </td>
              <td class="td-actions">
                <button class="btn-icon" onclick="event.stopPropagation(); openEditLeadModal('${lead.id}')" title="Edit">
                  <svg width="14" height="14" fill="none" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
                </button>
                <button class="btn-icon btn-danger" onclick="event.stopPropagation(); deleteLead('${lead.id}')" title="Delete">
                  <svg width="14" height="14" fill="none" viewBox="0 0 24 24"><polyline points="3,6 5,6 21,6" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
                </button>
              </td>`}
            </tr>
          `).join("")}
        </tbody>
      </table>
    </div>
  `;
}

// ============================================================
//  CONTRACTORS VIEW
// ============================================================

function renderContractors() {
  const main = document.getElementById("main-content");
  const { contractors, leads } = State;
  const { maxLeadsPerAgent } = Config.rules;

  main.innerHTML = `
    <div class="view-header">
      <div>
        <h1 class="view-title">Contractors</h1>
        <span class="view-subtitle">${contractors.length} agents</span>
      </div>
    </div>
    <div class="contractor-grid">
      ${contractors.length ? contractors.map(c => {
        const count = leads.filter(l => l.assignedTo === c.name && !["Won","Lost","Recycled"].includes(l.status)).length;
        const pct   = Math.min(100, Math.round((count / maxLeadsPerAgent) * 100));
        return `
          <div class="contractor-card">
            <div class="contractor-header">
              <div class="contractor-avatar">${c.name[0].toUpperCase()}</div>
              <div>
                <div class="contractor-name">${escHtml(c.name)}</div>
                <div class="contractor-role">${escHtml(c.role)}</div>
              </div>
              <span class="status-dot ${c.active ? 'dot-active' : 'dot-inactive'}"></span>
            </div>
            <div class="contractor-email">${escHtml(c.email || "No email")}</div>
            <div class="contractor-stats">
              <div class="load-label">
                <span>Lead Load</span>
                <span class="${count >= maxLeadsPerAgent ? 'text-danger' : ''}">${count}/${maxLeadsPerAgent}</span>
              </div>
              <div class="load-bar-wrap">
                <div class="load-bar ${pct >= 100 ? 'load-full' : pct >= 80 ? 'load-high' : ''}"
                     style="width:${pct}%"></div>
              </div>
            </div>
          </div>
        `;
      }).join("") : '<p class="empty-state">No contractors found in SharePoint.</p>'}
    </div>
  `;
}

// ============================================================
//  ACTIVITY LOG VIEW
// ============================================================

function renderActivity() {
  const main = document.getElementById("main-content");
  const { activityLog } = State;

  main.innerHTML = `
    <div class="view-header">
      <h1 class="view-title">Activity Log</h1>
      <span class="view-subtitle">${activityLog.length} entries</span>
    </div>
    <div class="card">
      <div class="table-wrap">
        <table class="data-table">
          <thead>
            <tr>
              <th>Time</th>
              <th>Lead</th>
              <th>Action</th>
              <th>Agent</th>
              <th>Notes</th>
            </tr>
          </thead>
          <tbody>
            ${activityLog.length ? activityLog.map(entry => `
              <tr>
                <td class="td-mono">${formatDateTime(entry.timestamp)}</td>
                <td>${escHtml(entry.leadName || entry.leadId || "—")}</td>
                <td><span class="action-badge">${escHtml(entry.action || "—")}</span></td>
                <td>${escHtml(entry.agent || "—")}</td>
                <td class="td-notes">${escHtml(entry.notes || "")}</td>
              </tr>
            `).join("") : `<tr><td colspan="5" class="empty-state">No activity logged yet.</td></tr>`}
          </tbody>
        </table>
      </div>
    </div>
  `;
}

// ============================================================
//  LEAD MODAL (Add / Edit)
// ============================================================

function openAddLeadModal() {
  State.editingLeadId = null;
  renderLeadModal(null);
}

function openEditLeadModal(id) {
  const lead = State.leads.find(l => l.id === id);
  if (!lead) return;
  State.editingLeadId = id;
  renderLeadModal(lead);
}

function renderLeadModal(lead) {
  const isEdit = !!lead;
  const contractors = State.contractors.map(c => c.name);
  const canContact  = !lead || !Graph.isInCoolOff(lead);

  document.getElementById("modal").innerHTML = `
    <div class="modal-header">
      <h2>${isEdit ? "Edit Lead" : "Add New Lead"}</h2>
      <button class="btn-icon" onclick="closeModal()">
        <svg width="18" height="18" fill="none" viewBox="0 0 24 24"><line x1="18" y1="6" x2="6" y2="18" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><line x1="6" y1="6" x2="18" y2="18" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
      </button>
    </div>

    ${isEdit && !canContact ? `
      <div class="alert alert-info" style="margin:0 0 16px">
        <svg width="14" height="14" fill="none" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" stroke="currentColor" stroke-width="2"/><line x1="12" y1="8" x2="12" y2="12" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><line x1="12" y1="16" x2="12.01" y2="16" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
        Cool-off period active — ${Config.rules.coolOffDays} days since last contact.
      </div>
    ` : ""}

    <div class="modal-form">
      <div class="form-row">
        <div class="form-group">
          <label>Full Name *</label>
          <input type="text" id="f-name" class="form-input" value="${escHtml(lead?.name || '')}" placeholder="Jane Smith" required>
        </div>
        <div class="form-group">
          <label>Company</label>
          <input type="text" id="f-company" class="form-input" value="${escHtml(lead?.company || '')}" placeholder="Acme Corp">
        </div>
      </div>
      <div class="form-row">
        <div class="form-group">
          <label>Email</label>
          <input type="email" id="f-email" class="form-input" value="${escHtml(lead?.email || '')}" placeholder="jane@acme.com">
        </div>
        <div class="form-group">
          <label>Phone</label>
          <input type="tel" id="f-phone" class="form-input" value="${escHtml(lead?.phone || '')}" placeholder="+1 555 0100">
        </div>
      </div>
      <div class="form-row">
        <div class="form-group">
          <label>Status</label>
          <select id="f-status" class="form-input">
            ${Config.leadStatuses.map(s => `<option value="${s}" ${(lead?.status||'New')===s?'selected':''}>${s}</option>`).join("")}
          </select>
        </div>
        <div class="form-group">
          <label>Source</label>
          <select id="f-source" class="form-input">
            <option value="">— Select —</option>
            ${Config.leadSources.map(s => `<option value="${s}" ${lead?.source===s?'selected':''}>${s}</option>`).join("")}
          </select>
        </div>
      </div>
      <div class="form-row">
        <div class="form-group">
          <label>Assigned To</label>
          <select id="f-assigned" class="form-input">
            <option value="">— Unassigned —</option>
            ${contractors.map(c => {
              const canTake = Graph.canAgentTakeLead(c, State.leads.filter(l => l.id !== lead?.id));
              return `<option value="${c}" ${lead?.assignedTo===c?'selected':''} ${!canTake && lead?.assignedTo!==c?'disabled':''}>
                ${c}${!canTake && lead?.assignedTo!==c ? ' (full)' : ''}
              </option>`;
            }).join("")}
          </select>
        </div>
        <div class="form-group">
          <label>Deal Value (£)</label>
          <input type="number" id="f-value" class="form-input" value="${lead?.value || ''}" placeholder="0.00">
        </div>
      </div>
      <div class="form-row">
        <div class="form-group">
          <label>Last Contacted</label>
          <input type="date" id="f-lastcontacted" class="form-input" value="${lead?.lastContacted ? lead.lastContacted.split('T')[0] : ''}">
        </div>
      </div>
      <div class="form-group form-group-full">
        <label>Notes</label>
        <textarea id="f-notes" class="form-input form-textarea" placeholder="Any notes about this lead…">${escHtml(lead?.notes || '')}</textarea>
      </div>
    </div>

    <div class="modal-footer">
      <button class="btn-ghost" onclick="closeModal()">Cancel</button>
      <button class="btn-primary" onclick="${isEdit ? 'submitEditLead()' : 'submitAddLead()'}">
        ${isEdit ? "Save Changes" : "Add Lead"}
      </button>
    </div>
  `;

  document.getElementById("modal-overlay").style.display = "flex";
}

async function submitAddLead() {
  const fields = collectLeadForm();
  if (!fields) return;

  const assignedTo = fields.AssignedTo;
  if (assignedTo && !Graph.canAgentTakeLead(assignedTo, State.leads)) {
    UI.showToast(`${assignedTo} is already at the ${Config.rules.maxLeadsPerAgent}-lead limit.`, "error");
    return;
  }

  setLoading(true);
  try {
    const newLead = await Graph.addLead(fields);
    await Graph.logActivity({
      LeadId: newLead.id, LeadName: fields.Title,
      Action: "Lead Created", Agent: Auth.getUser()?.name || "",
    });
    await refreshData();
    closeModal();
    UI.showToast("Lead added successfully!", "success");
  } catch (err) {
    UI.showToast("Failed to add lead: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

async function submitEditLead() {
  const fields = collectLeadForm();
  if (!fields) return;

  setLoading(true);
  try {
    await Graph.updateLead(State.editingLeadId, fields);
    await Graph.logActivity({
      LeadId: State.editingLeadId, LeadName: fields.Title,
      Action: "Lead Updated", Agent: Auth.getUser()?.name || "",
    });
    await refreshData();
    closeModal();
    UI.showToast("Lead updated successfully!", "success");
  } catch (err) {
    UI.showToast("Failed to update lead: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

function collectLeadForm() {
  const name = document.getElementById("f-name")?.value?.trim();
  if (!name) { UI.showToast("Name is required.", "error"); return null; }
  return {
    Title:         name,
    Company:       document.getElementById("f-company")?.value?.trim() || "",
    Email:         document.getElementById("f-email")?.value?.trim() || "",
    Phone:         document.getElementById("f-phone")?.value?.trim() || "",
    Status:        document.getElementById("f-status")?.value || "New",
    LeadSource:    document.getElementById("f-source")?.value || "",
    AssignedTo:    document.getElementById("f-assigned")?.value || "",
    DealValue:     document.getElementById("f-value")?.value || "",
    LastContacted: document.getElementById("f-lastcontacted")?.value || "",
    Notes:         document.getElementById("f-notes")?.value?.trim() || "",
  };
}

async function deleteLead(id) {
  const lead = State.leads.find(l => l.id === id);
  if (!confirm(`Delete lead "${lead?.name}"? This cannot be undone.`)) return;

  setLoading(true);
  try {
    await Graph.deleteLead(id);
    await refreshData();
    UI.showToast("Lead deleted.", "success");
  } catch (err) {
    UI.showToast("Failed to delete lead: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

function closeModal(event) {
  if (event && event.target !== document.getElementById("modal-overlay")) return;
  document.getElementById("modal-overlay").style.display = "none";
}

// ============================================================
//  UTILITIES
// ============================================================

async function refreshData() {
  await loadAllData();
  navigate(State.currentView);
}

function setLoading(on) {
  State.loading = on;
  const overlay = document.getElementById("loading-overlay");
  if (overlay) overlay.style.display = on ? "flex" : "none";
}

function updateBadges() {
  const needsAttention = State.leads.filter(l =>
    l.flags?.includes("needs_recycle") || l.flags?.includes("agent_overloaded")
  ).length;
  const badge = document.getElementById("badge-leads");
  if (badge) {
    badge.textContent = needsAttention > 0 ? needsAttention : "";
    badge.style.display = needsAttention > 0 ? "inline-flex" : "none";
  }
}

function exportCSV() {
  const leads = getFilteredLeads();
  const headers = ["Name","Company","Email","Phone","Status","Source","Assigned To","Value","Last Contacted","Created","Notes"];
  const rows = leads.map(l => [
    l.name, l.company, l.email, l.phone, l.status, l.source,
    l.assignedTo, l.value, l.lastContacted, l.createdAt, l.notes,
  ].map(v => `"${String(v || "").replace(/"/g, '""')}"`));

  const csv = [headers.join(","), ...rows.map(r => r.join(","))].join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `raimak-leads-${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
  UI.showToast("Exported CSV successfully!", "success");
}

function flagLabel(flag) {
  const map = {
    cool_off:         "⏱ Cool-off",
    needs_recycle:    "♻ Recycle",
    agent_overloaded: "⚠ Overloaded",
  };
  return map[flag] || flag;
}

function formatDate(d) {
  if (!d) return "";
  return new Date(d).toLocaleDateString("en-GB", { day:"2-digit", month:"short", year:"numeric" });
}

function formatDateTime(d) {
  if (!d) return "";
  return new Date(d).toLocaleString("en-GB", { day:"2-digit", month:"short", year:"numeric", hour:"2-digit", minute:"2-digit" });
}

function escHtml(str) {
  return String(str || "")
    .replace(/&/g,"&amp;")
    .replace(/</g,"&lt;")
    .replace(/>/g,"&gt;")
    .replace(/"/g,"&quot;");
}

// ── UI Helpers ─────────────────────────────────────────────────
const UI = {
  showToast(msg, type = "info") {
    const container = document.getElementById("toast-container");
    if (!container) return;
    const toast = document.createElement("div");
    toast.className = `toast toast-${type}`;
    toast.textContent = msg;
    container.appendChild(toast);
    setTimeout(() => toast.classList.add("show"), 10);
    setTimeout(() => { toast.classList.remove("show"); setTimeout(() => toast.remove(), 300); }, 4000);
  },
};
