// Raimak LMS - App Logic v2.0

// ── State ──────────────────────────────────────────────────────
const State = {
  leads:         [],
  contractors:   [],
  activityLog:   [],
  todaySales:    [],
  currentView:   "dashboard",
  filters:       { status: "all", search: "", assignedTo: "all" },
  editingLeadId: null,
  loading:       false,
  role:          "agent",       // "admin" or "agent"
  currentUser:   null,
  salesFeedTimer: null,
  currentAgentLead: null,       // the one lead shown to the agent in feed mode
};

// ── Role Helpers ───────────────────────────────────────────────
function isAdmin() { return State.role === "admin"; }

function detectRole(user) {
  if (!user) return "agent";
  const email = (user.email || "").toLowerCase();
  const admins = (Config.roles.admins || []).map(function(a) { return a.toLowerCase(); });
  return admins.includes(email) ? "admin" : "agent";
}

// ── Boot ──────────────────────────────────────────────────────
window.addEventListener("DOMContentLoaded", async function() {
  try {
    const redirectResult = await Auth.init();
    if (!Auth.isSignedIn()) { showLoginScreen(); return; }
    if (redirectResult) {
      window.history.replaceState({}, document.title, window.location.pathname);
    }
    State.currentUser = Auth.getUser();
    State.role        = detectRole(State.currentUser);
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
    const [rawLeads, contractors, activityLog, todaySales] = await Promise.all([
      Graph.getLeads(),
      Graph.getContractors(),
      Graph.getActivityLog(),
      Graph.getTodaySales(),
    ]);
    State.contractors = contractors;
    State.leads       = Graph.applyBusinessRules(rawLeads, contractors);
    State.activityLog = activityLog;
    State.todaySales  = todaySales;
  } catch (err) {
    UI.showToast("Failed to load data: " + err.message, "error");
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
        <div class="login-logo"><img src="Raimak.png" alt="Raimak"></div>
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
  const user = State.currentUser;
  const adminNav = isAdmin() ? `
    <a class="nav-item" data-view="assign" onclick="navigate('assign')">
      <svg width="18" height="18" fill="none" viewBox="0 0 24 24"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><circle cx="9" cy="7" r="4" stroke="currentColor" stroke-width="2"/><line x1="19" y1="8" x2="19" y2="14" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><line x1="22" y1="11" x2="16" y2="11" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
      Assign Leads
    </a>
    <a class="nav-item" data-view="report" onclick="navigate('report')">
      <svg width="18" height="18" fill="none" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="currentColor" stroke-width="2"/><polyline points="14,2 14,8 20,8" stroke="currentColor" stroke-width="2"/><line x1="16" y1="13" x2="8" y2="13" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><line x1="16" y1="17" x2="8" y2="17" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
      Daily Report
    </a>` : "";

  const agentNav = !isAdmin() ? `
    <a class="nav-item" data-view="myleads" onclick="navigate('myleads')">
      <svg width="18" height="18" fill="none" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" stroke="currentColor" stroke-width="2"/><polyline points="12,6 12,12 16,14" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
      My Leads
    </a>` : "";

  document.getElementById("app").innerHTML = `
    <div class="app-shell">
      <aside class="sidebar" id="sidebar">
        <div class="sidebar-brand"><img src="Raimak.png" alt="Raimak"></div>
        <nav class="sidebar-nav">
          <a class="nav-item active" data-view="dashboard" onclick="navigate('dashboard')">
            <svg width="18" height="18" fill="none" viewBox="0 0 24 24"><rect x="3" y="3" width="7" height="7" rx="1" fill="currentColor"/><rect x="14" y="3" width="7" height="7" rx="1" fill="currentColor"/><rect x="3" y="14" width="7" height="7" rx="1" fill="currentColor"/><rect x="14" y="14" width="7" height="7" rx="1" fill="currentColor"/></svg>
            Dashboard
          </a>
          ${agentNav}
          ${adminNav}
          <a class="nav-item" data-view="leads" onclick="navigate('leads')">
            <svg width="18" height="18" fill="none" viewBox="0 0 24 24"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><circle cx="9" cy="7" r="4" stroke="currentColor" stroke-width="2"/></svg>
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
            <div class="user-avatar">${((user && user.name) || "U")[0].toUpperCase()}</div>
            <div class="user-meta">
              <span class="user-name">${(user && user.name) || "User"}</span>
              <span class="user-email">${(user && user.email) || ""}</span>
            </div>
          </div>
          <button class="btn-ghost" onclick="Auth.signOut()" title="Sign Out">
            <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><polyline points="16,17 21,12 16,7" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><line x1="21" y1="12" x2="9" y2="12" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
          </button>
        </div>
      </aside>
      <main class="main-content" id="main-content">
        <div class="loading-overlay" id="loading-overlay" style="display:none"><div class="spinner"></div></div>
      </main>
    </div>
    <div id="toast-container"></div>
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
  document.querySelectorAll(".nav-item").forEach(function(el) { el.classList.remove("active"); });
  const navEl = document.querySelector("[data-view='" + view + "']");
  if (navEl) navEl.classList.add("active");
  switch (view) {
    case "dashboard":   renderDashboard();   break;
    case "leads":       renderLeads();       break;
    case "myleads":     renderMyLeads();     break;
    case "assign":      renderAssignLeads(); break;
    case "contractors": renderContractors(); break;
    case "activity":    renderActivity();    break;
    case "report":      renderDailyReport(); break;
  }
}

// ============================================================
//  DASHBOARD
// ============================================================

function renderDashboard() {
  const leads       = State.leads;
  const todaySales  = State.todaySales;
  const { maxLeadsPerAgent, recycleAfterDays } = Config.rules;

  const total       = leads.length;
  const active      = leads.filter(function(l) { return !Config.terminalStatuses.includes(l.status); }).length;
  const sold        = leads.filter(function(l) { return l.status === "Sold"; }).length;
  const needRecycle = leads.filter(function(l) { return l.flags && l.flags.includes("needs_recycle"); }).length;
  const coolOff     = leads.filter(function(l) { return l.flags && l.flags.includes("cool_off"); }).length;

  // Status counts
  const statusCounts = {};
  Config.leadStatuses.forEach(function(s) {
    statusCounts[s] = leads.filter(function(l) { return l.status === s; }).length;
  });

  // Top 5 agents by sales today
  const agentSales = {};
  todaySales.forEach(function(l) {
    if (l.assignedTo) agentSales[l.assignedTo] = (agentSales[l.assignedTo] || 0) + 1;
  });
  const top5 = Object.entries(agentSales)
    .sort(function(a, b) { return b[1] - a[1]; })
    .slice(0, 5);

  const recentLeads = leads.slice().sort(function(a, b) {
    return new Date(b.createdAt) - new Date(a.createdAt);
  }).slice(0, 5);

  const main = document.getElementById("main-content");
  main.innerHTML = `
    <div class="view-header">
      <div>
        <h1 class="view-title">Dashboard</h1>
        <span class="view-subtitle">${isAdmin() ? "Admin View" : "Agent View"} · v${Config.rules.appVersion}</span>
      </div>
      ${isAdmin() ? `<button class="btn-primary" onclick="openAddLeadModal()">
        <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"/><line x1="5" y1="12" x2="19" y2="12" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"/></svg>
        Add Lead
      </button>` : ""}
    </div>

    <div class="kpi-grid">
      <div class="kpi-card kpi-primary">
        <span class="kpi-label">Total Leads</span>
        <span class="kpi-value">${total}</span>
        <span class="kpi-sub">${active} active</span>
      </div>
      <div class="kpi-card kpi-success">
        <span class="kpi-label">Sold Today</span>
        <span class="kpi-value">${todaySales.length}</span>
        <span class="kpi-sub">${total ? Math.round((sold/total)*100) : 0}% all-time rate</span>
      </div>
      <div class="kpi-card ${needRecycle > 0 ? "kpi-warn" : "kpi-neutral"}">
        <span class="kpi-label">Needs Recycle</span>
        <span class="kpi-value">${needRecycle}</span>
        <span class="kpi-sub">&gt;${recycleAfterDays} days inactive</span>
      </div>
      <div class="kpi-card ${coolOff > 0 ? "kpi-info" : "kpi-neutral"}">
        <span class="kpi-label">In Cool-Off</span>
        <span class="kpi-value">${coolOff}</span>
        <span class="kpi-sub">${Config.rules.coolOffDays}-day rule</span>
      </div>
    </div>

    <div class="two-col">
      <!-- Pipeline -->
      <div class="card">
        <div class="card-header"><h2 class="card-title">Pipeline Status</h2></div>
        <div class="status-breakdown">
          ${Config.leadStatuses.map(function(s) { return `
            <div class="status-row">
              <span class="status-badge status-${s.toLowerCase().replace(/\s+/g,"-").replace(/[^a-z0-9-]/g,"")}">${s}</span>
              <div class="status-bar-wrap"><div class="status-bar" style="width:${total ? (statusCounts[s]/total)*100 : 0}%"></div></div>
              <span class="status-count">${statusCounts[s]}</span>
            </div>
          `; }).join("")}
        </div>
      </div>

      <!-- Live Sales Feed + Top 5 -->
      <div class="card">
        <div class="card-header">
          <h2 class="card-title">
            <span class="live-dot"></span> Live Sales Feed
          </h2>
          <span class="card-meta" id="sales-feed-time">Today</span>
        </div>
        <div class="sales-feed" id="sales-feed">
          ${todaySales.length ? todaySales.slice(0,8).map(function(l) { return `
            <div class="sale-entry">
              <div class="sale-icon">&#127881;</div>
              <div class="sale-info">
                <span class="sale-name">${escHtml(l.name)}</span>
                <span class="sale-agent">${escHtml(l.assignedTo || "Unassigned")}</span>
              </div>
              <span class="sale-time">${formatTime(l.modified)}</span>
            </div>
          `; }).join("") : `<p class="empty-state" style="padding:24px">No sales yet today. Go get them!</p>`}
        </div>

        <div class="card-header" style="border-top:1px solid var(--border);border-bottom:none;margin-top:4px">
          <h2 class="card-title">Top 5 Agents Today</h2>
        </div>
        <div class="top5-list">
          ${top5.length ? top5.map(function(entry, i) { return `
            <div class="top5-row">
              <span class="top5-rank rank-${i+1}">${i+1}</span>
              <span class="top5-name">${escHtml(entry[0])}</span>
              <span class="top5-count">${entry[1]} sale${entry[1] !== 1 ? "s" : ""}</span>
            </div>
          `; }).join("") : `<p class="empty-state" style="padding:16px 20px;font-size:12px">No sales recorded today yet.</p>`}
        </div>
      </div>
    </div>

    <!-- Recent Leads -->
    <div class="card">
      <div class="card-header">
        <h2 class="card-title">Recent Leads</h2>
        <button class="btn-ghost-sm" onclick="navigate('leads')">View all</button>
      </div>
      ${renderLeadsTable(recentLeads, true)}
    </div>
  `;

  updateBadges();
  startSalesFeedPolling();
}

// ── Live sales feed polling ─────────────────────────────────
function startSalesFeedPolling() {
  if (State.salesFeedTimer) clearInterval(State.salesFeedTimer);
  State.salesFeedTimer = setInterval(async function() {
    if (State.currentView !== "dashboard") return;
    try {
      State.todaySales = await Graph.getTodaySales();
      const feed = document.getElementById("sales-feed");
      if (!feed) return;
      const time = document.getElementById("sales-feed-time");
      if (time) time.textContent = "Updated " + formatTime(new Date().toISOString());
      if (!State.todaySales.length) return;
      feed.innerHTML = State.todaySales.slice(0, 8).map(function(l) { return `
        <div class="sale-entry">
          <div class="sale-icon">&#127881;</div>
          <div class="sale-info">
            <span class="sale-name">${escHtml(l.name)}</span>
            <span class="sale-agent">${escHtml(l.assignedTo || "Unassigned")}</span>
          </div>
          <span class="sale-time">${formatTime(l.modified)}</span>
        </div>
      `; }).join("");
    } catch (e) { /* silent */ }
  }, Config.salesFeedInterval);
}

// ============================================================
//  AGENT - MY LEADS (one-at-a-time feed)
// ============================================================

function renderMyLeads() {
  const user    = State.currentUser;
  const myLeads = State.leads.filter(function(l) {
    return l.assignedTo === (user && user.name) && !Config.terminalStatuses.includes(l.status);
  });
  const contactsToday = Graph.agentContactsToday((user && user.name) || "", State.activityLog);
  const atLimit       = contactsToday >= Config.rules.maxContactsPerDay;

  const main = document.getElementById("main-content");
  main.innerHTML = `
    <div class="view-header">
      <div>
        <h1 class="view-title">My Leads</h1>
        <span class="view-subtitle">${myLeads.length} assigned to you</span>
      </div>
      <div class="contacts-today-badge ${atLimit ? "badge-full" : ""}">
        <svg width="14" height="14" fill="none" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" stroke="currentColor" stroke-width="2"/><polyline points="12,6 12,12 16,14" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
        ${contactsToday}/${Config.rules.maxContactsPerDay} contacts today
      </div>
    </div>

    ${atLimit ? `
    <div class="alert alert-warn">
      <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z" stroke="currentColor" stroke-width="2"/><line x1="12" y1="9" x2="12" y2="13" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
      You have reached your daily contact limit of ${Config.rules.maxContactsPerDay}. Great work today!
    </div>` : ""}

    <div id="lead-feed-wrap">
      ${renderLeadFeedCard(myLeads, contactsToday)}
    </div>

    <div class="card" style="margin-top:20px">
      <div class="card-header"><h2 class="card-title">All My Assigned Leads</h2></div>
      ${renderLeadsTable(myLeads, false, true)}
    </div>
  `;
}

function renderLeadFeedCard(myLeads, contactsToday) {
  // Show the first lead that is not in cool-off
  const lead = myLeads.find(function(l) { return !Graph.isInCoolOff(l); });

  if (!lead) {
    return `
      <div class="feed-card feed-card-empty">
        <div class="feed-empty-icon">&#10003;</div>
        <h3>You are all caught up!</h3>
        <p>${myLeads.length > 0 ? "Your remaining leads are in the cool-off period. Check back soon." : "No leads currently assigned to you. Ask your manager to assign some."}</p>
      </div>
    `;
  }

  const inCoolOff = Graph.isInCoolOff(lead);
  const atLimit   = contactsToday >= Config.rules.maxContactsPerDay;

  return `
    <div class="feed-card">
      <div class="feed-card-header">
        <span class="feed-label">NEXT LEAD</span>
        ${lead.flags && lead.flags.includes("cool_off") ? `<span class="flag flag-cool_off">Cool-off</span>` : ""}
      </div>
      <h2 class="feed-name">${escHtml(lead.name)}</h2>
      <div class="feed-meta-row">
        ${lead.company ? `<span class="feed-meta"><svg width="13" height="13" fill="none" viewBox="0 0 24 24"><rect x="2" y="7" width="20" height="14" rx="2" stroke="currentColor" stroke-width="2"/><path d="M16 7V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v2" stroke="currentColor" stroke-width="2"/></svg>${escHtml(lead.company)}</span>` : ""}
        ${lead.phone ? `<span class="feed-meta"><svg width="13" height="13" fill="none" viewBox="0 0 24 24"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07A19.5 19.5 0 0 1 4.69 12 19.79 19.79 0 0 1 1.61 3.38 2 2 0 0 1 3.6 1.22h3a2 2 0 0 1 2 1.72c.127.96.361 1.903.7 2.81a2 2 0 0 1-.45 2.11L7.91 8.96a16 16 0 0 0 6 6l.92-.92a2 2 0 0 1 2.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0 1 21.73 16.92z" stroke="currentColor" stroke-width="2"/></svg>${escHtml(lead.phone)}</span>` : ""}
        ${lead.email ? `<span class="feed-meta"><svg width="13" height="13" fill="none" viewBox="0 0 24 24"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z" stroke="currentColor" stroke-width="2"/><polyline points="22,6 12,13 2,6" stroke="currentColor" stroke-width="2"/></svg>${escHtml(lead.email)}</span>` : ""}
      </div>
      ${lead.notes ? `<p class="feed-notes">${escHtml(lead.notes)}</p>` : ""}

      <div class="feed-status-row">
        <span class="feed-label">UPDATE STATUS</span>
        <div class="feed-status-buttons">
          ${Config.leadStatuses.filter(function(s) { return s !== "New"; }).map(function(s) { return `
            <button class="status-btn status-btn-${s.toLowerCase().replace(/\s+/g,"-").replace(/[^a-z0-9-]/g,"")}"
                    onclick="agentUpdateStatus('${lead.id}', '${s}')"
                    ${atLimit && !Config.terminalStatuses.includes(s) ? "disabled title='Daily limit reached'" : ""}>
              ${s}
            </button>
          `; }).join("")}
        </div>
      </div>
      <div class="feed-note-row">
        <textarea id="feed-notes" class="form-input form-textarea" placeholder="Add a note about this contact..." style="min-height:60px"></textarea>
        <button class="btn-primary" onclick="agentSaveNote('${lead.id}')">Save Note</button>
      </div>
    </div>
  `;
}

// Agent updates status from feed
async function agentUpdateStatus(leadId, newStatus) {
  const user = State.currentUser;
  const lead = State.leads.find(function(l) { return l.id === leadId; });
  if (!lead) return;

  // Enforce cool-off
  if (Graph.isInCoolOff(lead) && !Config.terminalStatuses.includes(newStatus)) {
    UI.showToast("This lead is in the " + Config.rules.coolOffDays + "-day cool-off period.", "error");
    return;
  }

  setLoading(true);
  try {
    const today = new Date().toISOString().split("T")[0];
    await Graph.updateLead(leadId, { Status: newStatus, LastContacted: today });
    await Graph.logActivity({
      LeadId:     leadId,
      LeadName:   lead.name,
      Action:     "Status: " + newStatus,
      Agent:      (user && user.name) || "",
      AgentEmail: (user && user.email) || "",
      Notes:      document.getElementById("feed-notes") ? document.getElementById("feed-notes").value : "",
    });
    UI.showToast("Lead marked as " + newStatus, "success");
    if (newStatus === "Sold") {
      UI.showConfetti();
    }
    await loadAllData();
    renderMyLeads();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

async function agentSaveNote(leadId) {
  const notes = document.getElementById("feed-notes");
  if (!notes || !notes.value.trim()) { UI.showToast("Please add a note first.", "error"); return; }
  const lead = State.leads.find(function(l) { return l.id === leadId; });
  if (!lead) return;
  setLoading(true);
  try {
    await Graph.updateLead(leadId, { Notes: notes.value.trim() });
    await Graph.logActivity({
      LeadId:     leadId,
      LeadName:   lead.name,
      Action:     "Note Added",
      Agent:      (State.currentUser && State.currentUser.name) || "",
      AgentEmail: (State.currentUser && State.currentUser.email) || "",
      Notes:      notes.value.trim(),
    });
    UI.showToast("Note saved!", "success");
    await loadAllData();
    renderMyLeads();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

// ============================================================
//  ADMIN - ASSIGN LEADS
// ============================================================

function renderAssignLeads() {
  const { leads, contractors } = State;
  const unassigned = leads.filter(function(l) { return !l.assignedTo && !Config.terminalStatuses.includes(l.status); });
  const { maxLeadsPerAgent } = Config.rules;

  const main = document.getElementById("main-content");
  main.innerHTML = `
    <div class="view-header">
      <div>
        <h1 class="view-title">Assign Leads</h1>
        <span class="view-subtitle">${unassigned.length} unassigned leads</span>
      </div>
      <button class="btn-primary" onclick="autoAssignLeads()">
        <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><polyline points="23,4 23,10 17,10" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><polyline points="1,20 1,14 7,14" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
        Auto-Assign All
      </button>
    </div>

    <!-- Agent load overview -->
    <div class="assign-agent-grid">
      ${contractors.map(function(c) {
        const count = leads.filter(function(l) { return l.assignedTo === c.name && !Config.terminalStatuses.includes(l.status); }).length;
        const pct   = Math.min(100, Math.round((count / maxLeadsPerAgent) * 100));
        const full  = count >= maxLeadsPerAgent;
        return `
          <div class="assign-agent-card ${full ? "agent-full" : ""}">
            <div class="contractor-avatar">${c.name[0].toUpperCase()}</div>
            <div class="assign-agent-info">
              <span class="contractor-name">${escHtml(c.name)}</span>
              <div class="load-bar-wrap">
                <div class="load-bar ${pct >= 100 ? "load-full" : pct >= 80 ? "load-high" : ""}" style="width:${pct}%"></div>
              </div>
              <span class="assign-count ${full ? "text-danger" : ""}">${count}/${maxLeadsPerAgent} ${full ? "(FULL)" : ""}</span>
            </div>
          </div>
        `;
      }).join("")}
    </div>

    <!-- Unassigned leads table -->
    <div class="card">
      <div class="card-header">
        <h2 class="card-title">Unassigned Leads (${unassigned.length})</h2>
      </div>
      <div class="table-wrap">
        <table class="data-table">
          <thead><tr>
            <th>Name</th><th>Company</th><th>Phone</th><th>Status</th><th>Source</th><th>Assign To</th>
          </tr></thead>
          <tbody>
            ${unassigned.length ? unassigned.map(function(lead) { return `
              <tr>
                <td><span class="lead-name">${escHtml(lead.name)}</span></td>
                <td>${escHtml(lead.company)}</td>
                <td class="td-mono">${escHtml(lead.phone)}</td>
                <td><span class="status-badge status-${lead.status.toLowerCase().replace(/\s+/g,"-").replace(/[^a-z0-9-]/g,"")}">${lead.status}</span></td>
                <td>${escHtml(lead.source)}</td>
                <td>
                  <div class="assign-select-row">
                    <select class="filter-select assign-select" id="assign-${lead.id}">
                      <option value="">Select agent</option>
                      ${contractors.map(function(c) {
                        const cnt = leads.filter(function(l) { return l.assignedTo === c.name && !Config.terminalStatuses.includes(l.status); }).length;
                        const full = cnt >= maxLeadsPerAgent;
                        return `<option value="${escHtml(c.name)}" ${full ? "disabled" : ""}>${escHtml(c.name)} (${cnt}/${maxLeadsPerAgent}${full ? " FULL" : ""})</option>`;
                      }).join("")}
                    </select>
                    <button class="btn-primary" style="padding:6px 12px;font-size:12px" onclick="assignLead('${lead.id}')">Assign</button>
                  </div>
                </td>
              </tr>
            `; }).join("") : `<tr><td colspan="6" class="empty-state">All leads are assigned!</td></tr>`}
          </tbody>
        </table>
      </div>
    </div>
  `;
}

async function assignLead(leadId) {
  const select = document.getElementById("assign-" + leadId);
  const agent  = select && select.value;
  if (!agent) { UI.showToast("Please select an agent first.", "error"); return; }

  if (!Graph.canAgentTakeLead(agent, State.leads)) {
    UI.showToast(agent + " is already at the " + Config.rules.maxLeadsPerAgent + "-lead limit.", "error");
    return;
  }

  const lead = State.leads.find(function(l) { return l.id === leadId; });
  setLoading(true);
  try {
    await Graph.updateLead(leadId, { AssignedTo: agent });
    await Graph.logActivity({
      LeadId:   leadId,
      LeadName: lead ? lead.name : "",
      Action:   "Assigned",
      Agent:    agent,
      Notes:    "Assigned by " + ((State.currentUser && State.currentUser.name) || "Admin"),
    });
    UI.showToast("Lead assigned to " + agent, "success");
    await loadAllData();
    renderAssignLeads();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

async function autoAssignLeads() {
  const { leads, contractors } = State;
  const unassigned = leads.filter(function(l) { return !l.assignedTo && !Config.terminalStatuses.includes(l.status); });
  if (!unassigned.length) { UI.showToast("No unassigned leads to assign.", "info"); return; }

  if (!confirm("Auto-assign " + unassigned.length + " leads evenly across available agents?")) return;

  setLoading(true);
  try {
    // Build available agent slots
    const slots = [];
    contractors.forEach(function(c) {
      const current = leads.filter(function(l) { return l.assignedTo === c.name && !Config.terminalStatuses.includes(l.status); }).length;
      const available = Config.rules.maxLeadsPerAgent - current;
      for (let i = 0; i < available; i++) slots.push(c.name);
    });

    let assigned = 0;
    for (let i = 0; i < Math.min(unassigned.length, slots.length); i++) {
      await Graph.updateLead(unassigned[i].id, { AssignedTo: slots[i] });
      assigned++;
    }

    UI.showToast("Assigned " + assigned + " leads successfully!", "success");
    await loadAllData();
    renderAssignLeads();
  } catch (err) {
    UI.showToast("Auto-assign failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

// ============================================================
//  LEADS VIEW (Admin full view)
// ============================================================

function renderLeads() {
  const contractors = State.contractors.map(function(c) { return c.name; });
  const main = document.getElementById("main-content");

  main.innerHTML = `
    <div class="view-header">
      <div>
        <h1 class="view-title">All Leads</h1>
        <span class="view-subtitle">${State.leads.length} total</span>
      </div>
      <div class="header-actions">
        <button class="btn-ghost" onclick="refreshData()">
          <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><polyline points="23,4 23,10 17,10" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><polyline points="1,20 1,14 7,14" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
          Refresh
        </button>
        <button class="btn-ghost" onclick="exportCSV()">
          <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><polyline points="7,10 12,15 17,10" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/><line x1="12" y1="15" x2="12" y2="3" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
          Export CSV
        </button>
        ${isAdmin() ? `<button class="btn-primary" onclick="openAddLeadModal()">
          <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"/><line x1="5" y1="12" x2="19" y2="12" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"/></svg>
          Add Lead
        </button>` : ""}
      </div>
    </div>
    <div class="filters-bar">
      <div class="search-wrap">
        <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><circle cx="11" cy="11" r="8" stroke="currentColor" stroke-width="2"/><line x1="21" y1="21" x2="16.65" y2="16.65" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
        <input type="text" id="search-input" class="search-input" placeholder="Search name, company, email..." value="${State.filters.search}" oninput="applyFilters()">
      </div>
      <select class="filter-select" id="filter-status" onchange="applyFilters()">
        <option value="all">All Statuses</option>
        ${Config.leadStatuses.map(function(s) { return `<option value="${s}" ${State.filters.status===s?"selected":""}>${s}</option>`; }).join("")}
      </select>
      <select class="filter-select" id="filter-agent" onchange="applyFilters()">
        <option value="all">All Agents</option>
        ${contractors.map(function(c) { return `<option value="${c}" ${State.filters.assignedTo===c?"selected":""}>${c}</option>`; }).join("")}
      </select>
    </div>
    <div class="card" id="leads-table-wrap">
      ${renderLeadsTable(getFilteredLeads())}
    </div>
  `;
}

function getFilteredLeads() {
  let leads = State.leads.slice();
  const { status, search, assignedTo } = State.filters;
  if (status !== "all")     leads = leads.filter(function(l) { return l.status === status; });
  if (assignedTo !== "all") leads = leads.filter(function(l) { return l.assignedTo === assignedTo; });
  if (search.trim()) {
    const q = search.toLowerCase();
    leads = leads.filter(function(l) {
      return l.name.toLowerCase().includes(q) || l.company.toLowerCase().includes(q) || l.email.toLowerCase().includes(q);
    });
  }
  return leads;
}

function applyFilters() {
  State.filters.search     = (document.getElementById("search-input")    || {}).value || "";
  State.filters.status     = (document.getElementById("filter-status")   || {}).value || "all";
  State.filters.assignedTo = (document.getElementById("filter-agent")    || {}).value || "all";
  const wrap = document.getElementById("leads-table-wrap");
  if (wrap) wrap.innerHTML = renderLeadsTable(getFilteredLeads());
}

function renderLeadsTable(leads, compact, agentView) {
  if (!leads.length) return `<div class="empty-state"><p>No leads found.</p></div>`;
  return `
    <div class="table-wrap">
      <table class="data-table">
        <thead><tr>
          <th>Name</th><th>Company</th>
          ${compact ? "" : "<th>Phone</th>"}
          <th>Status</th><th>Assigned To</th><th>Last Contacted</th>
          ${compact ? "" : "<th>Flags</th><th></th>"}
        </tr></thead>
        <tbody>
          ${leads.map(function(lead) { return `
            <tr class="lead-row ${lead.flags && lead.flags.includes("needs_recycle") ? "row-warn" : ""}"
                onclick="${isAdmin() ? "openEditLeadModal('" + lead.id + "')" : ""}">
              <td><span class="lead-name">${escHtml(lead.name)}</span>${lead.source ? `<span class="lead-source">${escHtml(lead.source)}</span>` : ""}</td>
              <td>${escHtml(lead.company)}</td>
              ${compact ? "" : `<td class="td-mono">${escHtml(lead.phone)}</td>`}
              <td><span class="status-badge status-${lead.status.toLowerCase().replace(/\s+/g,"-").replace(/[^a-z0-9-]/g,"")}">${lead.status}</span></td>
              <td>${escHtml(lead.assignedTo || "—")}</td>
              <td>${formatDate(lead.lastContacted) || formatDate(lead.createdAt) || "—"}</td>
              ${compact ? "" : `
              <td class="td-flags">${(lead.flags||[]).map(function(f) { return `<span class="flag flag-${f}">${flagLabel(f)}</span>`; }).join("")}</td>
              <td class="td-actions">
                ${isAdmin() ? `
                  <button class="btn-icon" onclick="event.stopPropagation();openEditLeadModal('${lead.id}')" title="Edit">
                    <svg width="14" height="14" fill="none" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
                  </button>
                  <button class="btn-icon btn-danger" onclick="event.stopPropagation();deleteLead('${lead.id}')" title="Delete">
                    <svg width="14" height="14" fill="none" viewBox="0 0 24 24"><polyline points="3,6 5,6 21,6" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
                  </button>
                ` : ""}
              </td>`}
            </tr>
          `; }).join("")}
        </tbody>
      </table>
    </div>
  `;
}

// ============================================================
//  DAILY REPORT (Admin only)
// ============================================================

async function renderDailyReport() {
  const main = document.getElementById("main-content");
  main.innerHTML = `
    <div class="view-header">
      <h1 class="view-title">Daily Report</h1>
      <button class="btn-ghost" onclick="exportReportCSV()">
        <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><polyline points="7,10 12,15 17,10" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><line x1="12" y1="15" x2="12" y2="3" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
        Export Report CSV
      </button>
    </div>
    <div class="card"><div class="empty-state" style="padding:40px">Loading report...</div></div>
  `;

  try {
    const stats = await Graph.getDailyStats();
    const today = new Date().toLocaleDateString("en-GB", { weekday:"long", year:"numeric", month:"long", day:"numeric" });

    main.innerHTML = `
      <div class="view-header">
        <div>
          <h1 class="view-title">Daily Report</h1>
          <span class="view-subtitle">${today}</span>
        </div>
        <button class="btn-ghost" onclick="exportReportCSV()">
          <svg width="16" height="16" fill="none" viewBox="0 0 24 24"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><polyline points="7,10 12,15 17,10" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><line x1="12" y1="15" x2="12" y2="3" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
          Export Report CSV
        </button>
      </div>

      <div class="kpi-grid">
        <div class="kpi-card kpi-primary">
          <span class="kpi-label">Total Contacts Today</span>
          <span class="kpi-value">${stats.reduce(function(s,a) { return s + a.contacts; }, 0)}</span>
        </div>
        <div class="kpi-card kpi-success">
          <span class="kpi-label">Total Sales Today</span>
          <span class="kpi-value">${State.todaySales.length}</span>
        </div>
        <div class="kpi-card kpi-info">
          <span class="kpi-label">Active Agents Today</span>
          <span class="kpi-value">${stats.length}</span>
        </div>
        <div class="kpi-card kpi-neutral">
          <span class="kpi-label">Avg Contacts/Agent</span>
          <span class="kpi-value">${stats.length ? Math.round(stats.reduce(function(s,a) { return s + a.contacts; }, 0) / stats.length) : 0}</span>
        </div>
      </div>

      <div class="card">
        <div class="card-header"><h2 class="card-title">Agent Breakdown</h2></div>
        <div class="table-wrap">
          <table class="data-table">
            <thead><tr>
              <th>Agent</th><th>Contacts Today</th><th>Sales Today</th><th>Contact Limit</th><th>Last Action</th>
            </tr></thead>
            <tbody>
              ${stats.length ? stats.map(function(a) {
                const pct = Math.round((a.contacts / Config.rules.maxContactsPerDay) * 100);
                const lastAction = a.actions.length ? a.actions[0] : null;
                return `
                  <tr>
                    <td><span class="lead-name">${escHtml(a.agent)}</span></td>
                    <td>
                      <div style="display:flex;align-items:center;gap:10px">
                        <span class="td-mono">${a.contacts}</span>
                        <div class="load-bar-wrap" style="flex:1;max-width:80px">
                          <div class="load-bar ${pct >= 100 ? "load-full" : pct >= 80 ? "load-high" : ""}" style="width:${Math.min(100,pct)}%"></div>
                        </div>
                      </div>
                    </td>
                    <td><span class="status-badge status-sold">${a.sold}</span></td>
                    <td class="td-mono">${a.contacts}/${Config.rules.maxContactsPerDay}</td>
                    <td class="td-mono" style="color:var(--text-3)">${lastAction ? formatDateTime(lastAction.timestamp) : "—"}</td>
                  </tr>
                `;
              }).join("") : `<tr><td colspan="5" class="empty-state">No agent activity recorded today yet.</td></tr>`}
            </tbody>
          </table>
        </div>
      </div>

      <div class="card">
        <div class="card-header"><h2 class="card-title">Today's Sales</h2></div>
        ${renderLeadsTable(State.todaySales, true)}
      </div>
    `;
    window._reportStats = stats;
  } catch (err) {
    UI.showToast("Failed to load report: " + err.message, "error");
  }
}

function exportReportCSV() {
  const stats = window._reportStats || [];
  const headers = ["Agent","Contacts Today","Sales Today","Date"];
  const today = new Date().toISOString().split("T")[0];
  const rows = stats.map(function(a) {
    return [a.agent, a.contacts, a.sold, today].map(function(v) { return '"' + String(v||"").replace(/"/g,'""') + '"'; });
  });
  const csv = [headers.join(",")].concat(rows.map(function(r) { return r.join(","); })).join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "raimak-daily-report-" + today + ".csv";
  a.click();
  UI.showToast("Report exported!", "success");
}

// ============================================================
//  CONTRACTORS VIEW
// ============================================================

function renderContractors() {
  const { contractors, leads } = State;
  const { maxLeadsPerAgent } = Config.rules;
  const main = document.getElementById("main-content");
  main.innerHTML = `
    <div class="view-header">
      <h1 class="view-title">Contractors</h1>
      <span class="view-subtitle">${contractors.length} agents</span>
    </div>
    <div class="contractor-grid">
      ${contractors.length ? contractors.map(function(c) {
        const count = leads.filter(function(l) { return l.assignedTo === c.name && !Config.terminalStatuses.includes(l.status); }).length;
        const pct   = Math.min(100, Math.round((count / maxLeadsPerAgent) * 100));
        const contactsToday = Graph.agentContactsToday(c.name, State.activityLog);
        return `
          <div class="contractor-card">
            <div class="contractor-header">
              <div class="contractor-avatar">${c.name[0].toUpperCase()}</div>
              <div><div class="contractor-name">${escHtml(c.name)}</div><div class="contractor-role">${escHtml(c.role)}</div></div>
              <span class="status-dot ${c.active ? "dot-active" : "dot-inactive"}"></span>
            </div>
            <div class="contractor-email">${escHtml(c.email || "No email")}</div>
            <div class="contractor-stats">
              <div class="load-label"><span>Lead Load</span><span class="${count >= maxLeadsPerAgent ? "text-danger" : ""}">${count}/${maxLeadsPerAgent}</span></div>
              <div class="load-bar-wrap"><div class="load-bar ${pct >= 100 ? "load-full" : pct >= 80 ? "load-high" : ""}" style="width:${pct}%"></div></div>
              <div class="load-label" style="margin-top:8px"><span>Contacts Today</span><span>${contactsToday}/${Config.rules.maxContactsPerDay}</span></div>
              <div class="load-bar-wrap"><div class="load-bar ${contactsToday >= Config.rules.maxContactsPerDay ? "load-full" : ""}" style="width:${Math.min(100, Math.round((contactsToday/Config.rules.maxContactsPerDay)*100))}%"></div></div>
            </div>
          </div>
        `;
      }).join("") : `<p class="empty-state">No contractors found.</p>`}
    </div>
  `;
}

// ============================================================
//  ACTIVITY LOG
// ============================================================

function renderActivity() {
  const { activityLog } = State;
  const main = document.getElementById("main-content");
  main.innerHTML = `
    <div class="view-header">
      <h1 class="view-title">Activity Log</h1>
      <span class="view-subtitle">${activityLog.length} entries</span>
    </div>
    <div class="card">
      <div class="table-wrap">
        <table class="data-table">
          <thead><tr><th>Time</th><th>Lead</th><th>Action</th><th>Agent</th><th>Notes</th></tr></thead>
          <tbody>
            ${activityLog.length ? activityLog.map(function(entry) { return `
              <tr>
                <td class="td-mono">${formatDateTime(entry.timestamp)}</td>
                <td>${escHtml(entry.leadName || entry.leadId || "—")}</td>
                <td><span class="action-badge">${escHtml(entry.action || "—")}</span></td>
                <td>${escHtml(entry.agent || "—")}</td>
                <td class="td-notes">${escHtml(entry.notes || "")}</td>
              </tr>
            `; }).join("") : `<tr><td colspan="5" class="empty-state">No activity logged yet.</td></tr>`}
          </tbody>
        </table>
      </div>
    </div>
  `;
}

// ============================================================
//  LEAD MODAL (Admin only - Add/Edit)
// ============================================================

function openAddLeadModal() {
  if (!isAdmin()) return;
  State.editingLeadId = null;
  renderLeadModal(null);
}

function openEditLeadModal(id) {
  if (!isAdmin()) return;
  const lead = State.leads.find(function(l) { return l.id === id; });
  if (!lead) return;
  State.editingLeadId = id;
  renderLeadModal(lead);
}

function renderLeadModal(lead) {
  const isEdit = !!lead;
  const contractors = State.contractors.map(function(c) { return c.name; });

  document.getElementById("modal").innerHTML = `
    <div class="modal-header">
      <h2>${isEdit ? "Edit Lead" : "Add New Lead"}</h2>
      <button class="btn-icon" onclick="closeModal()">
        <svg width="18" height="18" fill="none" viewBox="0 0 24 24"><line x1="18" y1="6" x2="6" y2="18" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><line x1="6" y1="6" x2="18" y2="18" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
      </button>
    </div>
    <div class="modal-form">
      <div class="form-row">
        <div class="form-group"><label>Full Name *</label><input type="text" id="f-name" class="form-input" value="${escHtml((lead&&lead.name)||"")}" placeholder="Jane Smith"></div>
        <div class="form-group"><label>Company</label><input type="text" id="f-company" class="form-input" value="${escHtml((lead&&lead.company)||"")}" placeholder="Acme Corp"></div>
      </div>
      <div class="form-row">
        <div class="form-group"><label>Email</label><input type="email" id="f-email" class="form-input" value="${escHtml((lead&&lead.email)||"")}" placeholder="jane@acme.com"></div>
        <div class="form-group"><label>Phone</label><input type="tel" id="f-phone" class="form-input" value="${escHtml((lead&&lead.phone)||"")}" placeholder="+1 555 0100"></div>
      </div>
      <div class="form-row">
        <div class="form-group">
          <label>Status</label>
          <select id="f-status" class="form-input">
            ${Config.leadStatuses.map(function(s) { return `<option value="${s}" ${((lead&&lead.status)||"New")===s?"selected":""}>${s}</option>`; }).join("")}
          </select>
        </div>
        <div class="form-group">
          <label>Source</label>
          <select id="f-source" class="form-input">
            <option value="">Select source</option>
            ${Config.leadSources.map(function(s) { return `<option value="${s}" ${lead&&lead.source===s?"selected":""}>${s}</option>`; }).join("")}
          </select>
        </div>
      </div>
      <div class="form-row">
        <div class="form-group">
          <label>Assigned To</label>
          <select id="f-assigned" class="form-input">
            <option value="">Unassigned</option>
            ${contractors.map(function(c) {
              const canTake = Graph.canAgentTakeLead(c, State.leads.filter(function(l) { return l.id !== (lead&&lead.id); }));
              return `<option value="${c}" ${lead&&lead.assignedTo===c?"selected":""} ${!canTake&&lead&&lead.assignedTo!==c?"disabled":""}>${c}${!canTake&&lead&&lead.assignedTo!==c?" (full)":""}</option>`;
            }).join("")}
          </select>
        </div>
        <div class="form-group"><label>Deal Value</label><input type="number" id="f-value" class="form-input" value="${(lead&&lead.value)||""}" placeholder="0.00"></div>
      </div>
      <div class="form-row">
        <div class="form-group"><label>Last Contacted</label><input type="date" id="f-lastcontacted" class="form-input" value="${lead&&lead.lastContacted ? lead.lastContacted.split("T")[0] : ""}"></div>
      </div>
      <div class="form-group form-group-full">
        <label>Notes</label>
        <textarea id="f-notes" class="form-input form-textarea">${escHtml((lead&&lead.notes)||"")}</textarea>
      </div>
    </div>
    <div class="modal-footer">
      <button class="btn-ghost" onclick="closeModal()">Cancel</button>
      <button class="btn-primary" onclick="${isEdit ? "submitEditLead()" : "submitAddLead()"}">
        ${isEdit ? "Save Changes" : "Add Lead"}
      </button>
    </div>
  `;
  document.getElementById("modal-overlay").style.display = "flex";
}

async function submitAddLead() {
  const fields = collectLeadForm();
  if (!fields) return;
  setLoading(true);
  try {
    const newLead = await Graph.addLead(fields);
    await Graph.logActivity({ LeadId: newLead.id, LeadName: fields.Title, Action: "Lead Created", Agent: (State.currentUser&&State.currentUser.name)||"" });
    await refreshData();
    closeModal();
    UI.showToast("Lead added!", "success");
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally { setLoading(false); }
}

async function submitEditLead() {
  const fields = collectLeadForm();
  if (!fields) return;
  setLoading(true);
  try {
    await Graph.updateLead(State.editingLeadId, fields);
    await Graph.logActivity({ LeadId: State.editingLeadId, LeadName: fields.Title, Action: "Lead Updated", Agent: (State.currentUser&&State.currentUser.name)||"" });
    await refreshData();
    closeModal();
    UI.showToast("Lead updated!", "success");
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally { setLoading(false); }
}

function collectLeadForm() {
  const name = (document.getElementById("f-name")||{}).value;
  if (!name || !name.trim()) { UI.showToast("Name is required.", "error"); return null; }
  return {
    Title:         name.trim(),
    Company:       ((document.getElementById("f-company")||{}).value||"").trim(),
    Email:         ((document.getElementById("f-email")||{}).value||"").trim(),
    Phone:         ((document.getElementById("f-phone")||{}).value||"").trim(),
    Status:        (document.getElementById("f-status")||{}).value || "New",
    LeadSource:    (document.getElementById("f-source")||{}).value || "",
    AssignedTo:    (document.getElementById("f-assigned")||{}).value || "",
    DealValue:     (document.getElementById("f-value")||{}).value || "",
    LastContacted: (document.getElementById("f-lastcontacted")||{}).value || "",
    Notes:         ((document.getElementById("f-notes")||{}).value||"").trim(),
  };
}

async function deleteLead(id) {
  const lead = State.leads.find(function(l) { return l.id === id; });
  if (!confirm("Delete lead \"" + (lead&&lead.name) + "\"? This cannot be undone.")) return;
  setLoading(true);
  try {
    await Graph.deleteLead(id);
    await refreshData();
    UI.showToast("Lead deleted.", "success");
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally { setLoading(false); }
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
  const n = State.leads.filter(function(l) { return l.flags && (l.flags.includes("needs_recycle") || l.flags.includes("agent_overloaded")); }).length;
  const badge = document.getElementById("badge-leads");
  if (badge) { badge.textContent = n > 0 ? n : ""; badge.style.display = n > 0 ? "inline-flex" : "none"; }
}

function exportCSV() {
  const leads = getFilteredLeads();
  const headers = ["Name","Company","Email","Phone","Status","Source","Assigned To","Value","Last Contacted","Created","Notes"];
  const rows = leads.map(function(l) {
    return [l.name,l.company,l.email,l.phone,l.status,l.source,l.assignedTo,l.value,l.lastContacted,l.createdAt,l.notes]
      .map(function(v) { return '"' + String(v||"").replace(/"/g,'""') + '"'; });
  });
  const csv = [headers.join(",")].concat(rows.map(function(r) { return r.join(","); })).join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "raimak-leads-" + new Date().toISOString().slice(0,10) + ".csv";
  a.click();
  UI.showToast("Exported!", "success");
}

function flagLabel(flag) {
  return { cool_off: "Cool-off", needs_recycle: "Recycle", agent_overloaded: "Overloaded" }[flag] || flag;
}

function formatDate(d) {
  if (!d) return "";
  return new Date(d).toLocaleDateString("en-GB", { day:"2-digit", month:"short", year:"numeric" });
}

function formatTime(d) {
  if (!d) return "";
  return new Date(d).toLocaleTimeString("en-GB", { hour:"2-digit", minute:"2-digit" });
}

function formatDateTime(d) {
  if (!d) return "";
  return new Date(d).toLocaleString("en-GB", { day:"2-digit", month:"short", year:"numeric", hour:"2-digit", minute:"2-digit" });
}

function escHtml(str) {
  return String(str||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}

// ── UI Helpers ─────────────────────────────────────────────────
const UI = {
  showToast: function(msg, type) {
    type = type || "info";
    const container = document.getElementById("toast-container");
    if (!container) return;
    const toast = document.createElement("div");
    toast.className = "toast toast-" + type;
    toast.textContent = msg;
    container.appendChild(toast);
    setTimeout(function() { toast.classList.add("show"); }, 10);
    setTimeout(function() { toast.classList.remove("show"); setTimeout(function() { toast.remove(); }, 300); }, 4000);
  },
  showConfetti: function() {
    // Simple CSS confetti burst on sale
    const el = document.createElement("div");
    el.className = "confetti-burst";
    el.innerHTML = "&#127881; SOLD! &#127881;";
    document.body.appendChild(el);
    setTimeout(function() { el.remove(); }, 2500);
  },
};
