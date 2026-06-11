// Raimak LMS - App Logic v3.0
window._isWorkingCallback = false;
const cachedSkips = sessionStorage.getItem("_skippedSessionLeads");
window._skippedSessionLeads = cachedSkips ? JSON.parse(cachedSkips) : [];
const savedSyncDate = localStorage.getItem("RaimakActivityLastSyncDate");
const savedLeadsSyncDate = localStorage.getItem("RaimakLeadsLastSyncDate");

const State = {
  leads: [],
  contractors: [],
  activityLog: [],
  todaySales: [],
  drafts: {},
  agentScores: [],
  currentView: "dashboard",
  filters: { status: "all", search: "", assignedTo: "all" },
  editingLeadId: null,
  loading: false,
  role: "agent",
  currentUser: null,
  salesFeedTimer: null,
  dripLead: null,
  selectedLeads: new Set(),

  // 🚀 THE NEW TIME-BASED TRACKER
  lastSyncDate: savedSyncDate || null,
};

const stateTimezones = {
  AL: "America/Chicago",
  AK: "America/Anchorage",
  AZ: "America/Phoenix",
  AR: "America/Chicago",
  CA: "America/Los_Angeles",
  CO: "America/Denver",
  CT: "America/New_York",
  DE: "America/New_York",
  FL: "America/New_York",
  GA: "America/New_York",
  HI: "America/Honolulu",
  ID: "America/Boise",
  IL: "America/Chicago",
  IN: "America/Indianapolis",
  IA: "America/Chicago",
  KS: "America/Chicago",
  KY: "America/New_York",
  LA: "America/Chicago",
  ME: "America/New_York",
  MD: "America/New_York",
  MA: "America/New_York",
  MI: "America/Detroit",
  MN: "America/Chicago",
  MS: "America/Chicago",
  MO: "America/Chicago",
  MT: "America/Denver",
  NE: "America/Chicago",
  NV: "America/Los_Angeles",
  NH: "America/New_York",
  NJ: "America/New_York",
  NM: "America/Denver",
  NY: "America/New_York",
  NC: "America/New_York",
  ND: "America/Chicago",
  OH: "America/New_York",
  OK: "America/Chicago",
  OR: "America/Los_Angeles",
  PA: "America/New_York",
  RI: "America/New_York",
  SC: "America/New_York",
  SD: "America/Chicago",
  TN: "America/Chicago",
  TX: "America/Chicago",
  UT: "America/Denver",
  VT: "America/New_York",
  VA: "America/New_York",
  WA: "America/Los_Angeles",
  WV: "America/New_York",
  WI: "America/Chicago",
  WY: "America/Denver",
};

// ==========================================
// 📊 PIPELINE INSIGHTS MODULE
// ==========================================
window._activeInsightChart = null;

const PipelineInsights = {
  currentLeads: [],
  currentDataArr: [],
  currentMode: "",
  colorMap: {},
  currentTotal: 0,

  getColorForLabel: function (label) {
    const colors = [
      "#0ea5e9",
      "#10b981",
      "#8b5cf6",
      "#f59e0b",
      "#f43f5e",
      "#14b8a6",
      "#64748b",
      "#f97316",
      "#3b82f6",
      "#84cc16",
      "#d946ef",
      "#06b6d4",
    ];
    if (!this.colorMap[label]) {
      const assignedCount = Object.keys(this.colorMap).length;
      this.colorMap[label] = colors[assignedCount % colors.length];
    }
    return this.colorMap[label];
  },

  aggregate: function (leads, mode) {
    const counts = {};
    if (mode === "assigned") {
      let assigned = 0,
        unassigned = 0;
      leads.forEach((l) => (l.assignedTo ? assigned++ : unassigned++));
      return [
        ["Assigned", assigned],
        ["Unassigned", unassigned],
      ].filter((x) => x[1] > 0);
    }
    leads.forEach((lead) => {
      let key = "Unknown";
      if (mode === "status") key = lead.status || "New";
      if (mode === "state") key = (lead.state || "Unknown").toUpperCase();
      if (mode === "type") key = lead.leadType || "None";
      if (key.trim() !== "") counts[key] = (counts[key] || 0) + 1;
    });
    return Object.entries(counts).sort((a, b) => b[1] - a[1]);
  },

  render: function (canvasId, leads, mode) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;

    if (this.currentMode !== mode) {
      this.currentMode = mode;
      this.colorMap = {};
    }

    this.currentDataArr = this.aggregate(leads, mode);
    const labels = this.currentDataArr.map((item) => item[0]);
    const data = this.currentDataArr.map((item) => item[1]);
    const backgroundColors = labels.map((label) =>
      this.getColorForLabel(label),
    );

    const newTotal = data.reduce((a, b) => a + b, 0);

    if (
      window._activeInsightChart &&
      document.body.contains(window._activeInsightChart.canvas)
    ) {
      window._activeInsightChart.data.labels = labels;
      window._activeInsightChart.data.datasets[0].data = data;
      window._activeInsightChart.data.datasets[0].backgroundColor =
        backgroundColors;
      this.currentTotal = newTotal;
      window._activeInsightChart.update();
      return;
    }

    if (window._activeInsightChart) {
      window._activeInsightChart.destroy();
      window._activeInsightChart = null;
    }

    this.currentTotal = newTotal;

    window._activeInsightChart = new Chart(canvas, {
      type: "doughnut",
      data: {
        labels: labels,
        datasets: [
          {
            data: data,
            backgroundColor: backgroundColors,
            borderWidth: 2,
            borderColor: "#ffffff",
            hoverOffset: 6,
          },
        ],
      },
      plugins: [
        {
          id: "centerTotalText",
          beforeDraw: (chart) => {
            const { ctx, chartArea } = chart;
            const meta = chart.getDatasetMeta(0);
            if (!chartArea || !meta) return;

            ctx.restore();

            let centerX, centerY;
            if (meta.data && meta.data.length > 0) {
              centerX = meta.data[0].x;
              centerY = meta.data[0].y;
            } else {
              centerX = (chartArea.left + chartArea.right) / 2;
              centerY = (chartArea.top + chartArea.bottom) / 2;
            }

            ctx.font = "bold 28px var(--font-mono, sans-serif)";
            ctx.textBaseline = "middle";
            ctx.fillStyle = "#0D1B3E";

            const text = PipelineInsights.currentTotal.toLocaleString();
            const textX = Math.round(centerX - ctx.measureText(text).width / 2);
            const textY = centerY - 6;
            ctx.fillText(text, textX, textY);

            ctx.font = "600 11px sans-serif";
            ctx.fillStyle = "#64748B";
            const label = "TOTAL";
            const labelX = Math.round(
              centerX - ctx.measureText(label).width / 2,
            );
            ctx.fillText(label, labelX, centerY + 18);
            ctx.save();
          },
        },
      ],
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: "65%",
        animation: {
          onProgress: () => {
            if (window._activeInsightChart) window._activeInsightChart.draw();
          },
        },
        onClick: (event, elements) => {
          if (!elements.length) return;
          const clickedLabel =
            PipelineInsights.currentDataArr[elements[0].index][0];

          // 🚀 THE FIX: We now tell the chart to look for BOTH the standard IDs and the "opt-" quarantined IDs!
          let targetIds = [];
          if (PipelineInsights.currentMode === "state")
            targetIds = [
              "filter-state",
              "opt-filter-state",
              "bulk-state-select",
            ];
          if (PipelineInsights.currentMode === "type")
            targetIds = ["filter-type", "opt-filter-type", "bulk-type-select"];
          if (PipelineInsights.currentMode === "status")
            targetIds = ["filter-status", "opt-filter-status"];

          for (let id of targetIds) {
            const selectEl = document.getElementById(id);
            if (selectEl) {
              for (let i = 0; i < selectEl.options.length; i++) {
                const opt = selectEl.options[i];
                if (
                  opt.value.toUpperCase() === clickedLabel.toUpperCase() ||
                  opt.text.toUpperCase() === clickedLabel.toUpperCase()
                ) {
                  selectEl.value = opt.value;
                  selectEl.dispatchEvent(
                    new Event("change", { bubbles: true }),
                  );
                  return;
                }
              }
            }
          }
        },
        plugins: {
          legend: {
            position: "right",
            labels: {
              color: "#475569",
              font: { size: 12, family: "var(--font-mono)" },
              padding: 16,
            },
            onClick: null,
          },
        },
      },
    });
  },

  init: function (selectId, canvasId, leads, isMainPage) {
    const selector = document.getElementById(selectId);
    if (!selector) return;
    this.currentLeads = leads;
    let options = `<option value="type">Lead Type Distribution</option><option value="state">Geographic (State)</option>`;
    if (isMainPage)
      options =
        `<option value="status">Pipeline Status</option><option value="assigned">Assigned vs Unassigned</option>` +
        options;
    if (!selector.innerHTML.trim()) selector.innerHTML = options;
    this.render(canvasId, this.currentLeads, selector.value);
    selector.addEventListener("change", (e) =>
      this.render(canvasId, this.currentLeads, e.target.value),
    );
  },

  updateLive: function (selectId, canvasId, leads) {
    const selector = document.getElementById(selectId);
    this.currentLeads = leads;
    if (selector && window._activeInsightChart)
      this.render(canvasId, this.currentLeads, selector.value);
  },
};

function isAdmin() {
  return State.role === "admin";
}

function detectRole(user) {
  if (!user) return "agent";
  const email = (user.email || "").toLowerCase();
  const admins = (Config.roles.admins || []).map(function (a) {
    return a.toLowerCase();
  });
  return admins.includes(email) ? "admin" : "agent";
}

// ── Boot ──────────────────────────────────────────────────────
window.addEventListener("DOMContentLoaded", async function () {
  //separated login page html and js
  const loginBtn = document.getElementById("ms-login-btn");
  if (loginBtn) {
    loginBtn.addEventListener("click", function () {
      Auth.signIn();
    });
  }

  try {
    const redirectResult = await Auth.init();

    if (!Auth.isSignedIn()) {
      showLoginScreen();
      return;
    }

    if (redirectResult) {
      window.history.replaceState({}, document.title, window.location.pathname);
    }

    await LocalDB.init();
    State.currentUser = Auth.getUser();
    State.role = detectRole(State.currentUser);

    // If they aren't suspended, load the rest of the app!
    showAppShell();
    Points.initHUDAutoHider();
    // ==========================================
    // 🛑 THE SUSPENSION GATEKEEPER
    // ==========================================
    const userEmail = State.currentUser ? State.currentUser.email : null;

    if (userEmail) {
      // 🚀 Now returns the expiration date instead of true/false
      const suspensionExpiration = await Graph.checkSuspensionStatus(userEmail);

      if (suspensionExpiration) {
        const mainContent = document.getElementById("main-content");
        const template = document.getElementById("tmpl-suspended");
        const sidebar = document.getElementById("sidebar");

        if (sidebar) sidebar.style.display = "none";

        if (mainContent && template) {
          mainContent.style.marginLeft = "0";
          mainContent.style.width = "100%";
          mainContent.innerHTML = "";
          mainContent.appendChild(template.content.cloneNode(true));

          // 🚀 Start the clock!
          startSuspensionCountdown(suspensionExpiration);
        }

        return;
      }
    }
    await loadAllData();
    Points.updateHUD();
    renderDashboard();
    Ticker.update();
  } catch (err) {
    console.error("Boot error:", err);
    showLoginScreen();
  }
});

async function loadAllData() {
  setLoading(true);
  UI.showToast("Syncing floor data...", "info");

  try {
    // 1. Resolve IDs once
    await Graph.resolveSiteIds();

    // 🚀 THE FIX Part 1: We pull the Activity Log OUT of the concurrent race.
    // Leads are heavy, but Contractors and Points are tiny, so they can safely race together!
    const [rawLeads, contractors, pointsData] = await Promise.all([
      Graph.getLeads(savedLeadsSyncDate, State.leads).then((data) => {
        UI.showToast("✅ Leads synced!", "success");
        return data;
      }),
      Graph.getContractors().then((data) => {
        UI.showToast("✅ Contractors synced!", "success");
        return data;
      }),
      Points.fetchBalances().then((data) => {
        UI.showToast("✅ Points balances synced!", "success");
        return data;
      }),
    ]);

    State.contractors = contractors;
    State.leads = Graph.applyBusinessRules(rawLeads, contractors);

    // 🚀 THE FIX Part 2: The Smart Activity Fetch
    // Now that the massive Leads download is finished, we safely ask for the Logs
    let todayLogs = [];

    if (isAdmin()) {
      UI.showToast("Syncing historical admin logs...", "info");

      // Admins use the hyper-fast Delta Sync to get the entire database
      // 🚀 UPDATED: Swapped highestActivityId for lastSyncDate
      const logData = await Graph.getActivityLog(
        State.lastSyncDate,
        State.activityLog,
      );

      State.activityLog = logData.updatedLogs;
      // 🚀 UPDATED: Swapped newHighestId for newSyncDate
      State.lastSyncDate = logData.newSyncDate;

      // Extract today's logs purely from RAM so we don't have to fetch them again!
      const todayStr = new Date().toDateString();
      todayLogs = State.activityLog.filter(
        (log) =>
          log.timestamp && new Date(log.timestamp).toDateString() === todayStr,
      );

      UI.showToast("✅ Admin logs synced!", "success");
    } else {
      UI.showToast("Syncing today's activity...", "info");

      // Standard agents just get the fast daily log
      const logData = await Graph.getActivityLog(
        State.lastSyncDate,
        State.activityLog,
      );

      State.activityLog = logData.updatedLogs;
      // 🚀 UPDATED: Swapped newHighestId for newSyncDate
      State.lastSyncDate = logData.newSyncDate;

      UI.showToast("✅ Activity synced!", "success");
    }

    // 3. Instant, synchronous math!
    State.todaySales = Graph.getTodaySales(todayLogs);
  } catch (err) {
    UI.showToast("Failed to load data: " + err.message, "error");
    console.error("Data Load Error:", err);
  } finally {
    setLoading(false);
  }
}

// ============================================================
//  LOGIN
// ============================================================
function showLoginScreen() {
  // 1. Hide the main app wrapper
  document.getElementById("app-shell").style.display = "none";

  // 2. Show the login view
  document.getElementById("login-view").style.display = "flex";

  // 3. Safely inject the version number
  if (typeof Config !== "undefined" && Config.rules) {
    document.getElementById("app-version-text").textContent =
      Config.rules.appVersion;
  }
}

// ============================================================
//  APP SHELL
// ============================================================
function showAppShell() {
  const user = State.currentUser;

  // 1. Hide Login, Show App Shell
  document.getElementById("login-view").style.display = "none";
  document.getElementById("app-shell").style.display = "flex";

  // 2. Populate User Data dynamically into the HTML we just made
  if (user) {
    document.getElementById("ui-user-name").textContent = user.name || "User";
    document.getElementById("ui-user-email").textContent = user.email || "";
    document.getElementById("ui-user-initial").textContent = (user.name ||
      "U")[0].toUpperCase();
  }

  // 3. The "Cardboard" Security Guard (Frontend Role Check)
  // Hide all elements with the 'admin-only' class if they aren't an admin.
  if (!isAdmin()) {
    document.querySelectorAll(".admin-only").forEach((el) => {
      el.style.display = "none";
    });
  } else {
    document.querySelectorAll(".admin-only").forEach((el) => {
      el.style.display = "flex"; // Or 'block', depending on your CSS
    });
  }
}

// ============================================================
//  NAVIGATION
// ============================================================
function navigate(view) {
  if (window._clockTimer) clearInterval(window._clockTimer);
  const adminOnly = [
    "leads",
    "drip",
    "assign",
    "report",
    "contractors",
    "activity",
  ];
  if (!isAdmin() && adminOnly.includes(view)) {
    view = "myleads";
  }
  State.currentView = view;
  document.querySelectorAll(".nav-item").forEach(function (el) {
    el.classList.remove("active");
  });
  const navEl = document.querySelector("[data-view='" + view + "']");
  if (navEl) navEl.classList.add("active");
  switch (view) {
    case "dashboard":
      renderDashboard();
      break;
    case "leads":
      renderLeads();
      break;
    case "myleads":
      renderMyLeads();
      break;
    case "callbacks":
      renderCallBacks();
      break;

    case "stats":
      renderStats();
      break;

    case "drip":
      renderDripFeed();
      break;
    case "assign":
      renderAssignLeads();
      break;
    case "contractors":
      renderContractors();
      break;
    case "activity":
      renderActivity();
      break;
    case "report":
      renderDailyReport();
      break;
  }
}

// ============================================================
//  DASHBOARD
// ============================================================
function renderDashboard() {
  const leads = State.leads;
  const todaySales = State.todaySales;
  const total = leads.length;

  const active = leads.filter(
    (l) => !Config.terminalStatuses.includes(l.status),
  ).length;
  const sold = leads.filter((l) => l.status === "Sold").length;
  const needRecycle = leads.filter(
    (l) => l.flags && l.flags.includes("needs_recycle"),
  ).length;
  const coolOff = leads.filter(
    (l) => l.flags && l.flags.includes("cool_off"),
  ).length;

  const statusCounts = {};
  Config.leadStatuses.forEach((s) => {
    statusCounts[s] = leads.filter((l) => l.status === s).length;
  });

  const agentSales = {};
  todaySales.forEach((l) => {
    if (l.assignedTo)
      agentSales[l.assignedTo] = (agentSales[l.assignedTo] || 0) + 1;
  });

  const top5 = Object.entries(agentSales)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);
  // ==========================================
  //  THE NEW RENDER LOGIC
  // ==========================================
  const mainContent = document.getElementById("main-content");
  mainContent.innerHTML = "";

  // 1. Clone the HTML blueprint
  const template = document.getElementById("tmpl-dashboard");
  const clone = template.content.cloneNode(true);

  // 2. Handle Admin Security
  if (!isAdmin()) {
    clone.querySelectorAll(".admin-only").forEach((el) => el.remove());
  }

  // 3. Populate Header & KPIs
  clone.getElementById("dash-subtitle").textContent =
    `${isAdmin() ? "// ADMIN VIEW" : "// AGENT VIEW"} · v${Config.rules.appVersion}`;
  clone.getElementById("kpi-total").textContent = total;
  clone.getElementById("kpi-active-sub").textContent =
    `${active} active in pipeline`;
  clone.getElementById("kpi-sold-today").textContent = todaySales.length;
  clone.getElementById("kpi-close-rate").textContent =
    `${total ? Math.round((sold / total) * 100) : 0}% all-time close rate`;
  clone.getElementById("kpi-cooloff").textContent = coolOff;
  clone.getElementById("kpi-cooloff-sub").textContent =
    `${Config.rules.coolOffDays}-day rule active`;

  const coolOffCard = clone.getElementById("kpi-cooloff-card");
  coolOffCard.className = `kpi-card ${coolOff > 0 ? "kpi-info" : "kpi-neutral"}`;

  if (isAdmin()) {
    const recycleCard = clone.getElementById("kpi-recycle-card");
    clone.getElementById("kpi-recycle-count").textContent = needRecycle;
    clone.getElementById("kpi-recycle-sub").textContent =
      needRecycle > 0 ? "↓ See recycle queue below" : "All leads current";
    recycleCard.className = `kpi-card admin-only ${needRecycle > 0 ? "kpi-warn" : "kpi-neutral"}`;
    if (needRecycle > 0) {
      recycleCard.style.cursor = "pointer";
      recycleCard.onclick = () =>
        document
          .getElementById("dash-recycle-section")
          .scrollIntoView({ behavior: "smooth" });
    }
  }

  // 4. Inject Dynamic Lists
  if (isAdmin()) {
    clone.getElementById("dash-pipeline-status").innerHTML = Config.leadStatuses
      .map((s) => {
        const cls =
          "status-" +
          s
            .toLowerCase()
            .replace(/\s+/g, "-")
            .replace(/[^a-z0-9-]/g, "");
        return `<div class="status-row">
                <span class="status-badge ${cls}">${s}</span>
                <div class="status-bar-wrap"><div class="status-bar" style="width:${total ? (statusCounts[s] / total) * 100 : 0}%"></div></div>
                <span class="status-count">${statusCounts[s]}</span>
              </div>`;
      })
      .join("");
  }

  clone.getElementById("dash-top5").innerHTML = top5.length
    ? top5
        .map(
          (e, i) => `
      <div class="top5-row">
        <span class="top5-rank rank-${i + 1}">${i + 1}</span>
        <span class="top5-name">${escHtml(e[0])}</span>
        <span class="top5-count">${e[1]} sale${e[1] !== 1 ? "s" : ""}</span>
      </div>`,
        )
        .join("")
    : `<p class="empty-state">No sales yet today.</p>`;

  const todayStr = new Date().toLocaleDateString();
  const agentUniqueLeads = {};

  const aliasMap = {
    "j.torres@raimak.com": "JULIAN TORRES",
  };

  (State.activityLog || []).forEach((log) => {
    let isToday = false;
    if (log.timestamp) {
      isToday = new Date(log.timestamp).toLocaleDateString() === todayStr;
    }

    const actionStr = log.action || log.ActionType || "";
    const isContact =
      actionStr.startsWith("Status:") ||
      actionStr === "1st Contact" ||
      actionStr === "2nd Contact" ||
      actionStr === "3rd Contact";

    if (isToday && isContact) {
      let rawAgent = (log.agent || log.AgentEmail || "Unknown")
        .toLowerCase()
        .trim();

      let displayName = aliasMap[rawAgent];

      if (!displayName) {
        const contractor = State.contractors.find(
          (c) =>
            (c.email || "").toLowerCase().trim() === rawAgent ||
            (c.name || "").toLowerCase().trim() === rawAgent,
        );
        displayName = contractor
          ? contractor.name
          : log.agent || log.AgentEmail || "Unknown";
      }

      const leadId = log.leadId || log.LeadID;

      if (!agentUniqueLeads[displayName])
        agentUniqueLeads[displayName] = new Set();
      if (leadId) agentUniqueLeads[displayName].add(leadId);
    }
  });

  const top5Contacts = Object.entries(agentUniqueLeads)
    .map(([name, leadSet]) => [name, leadSet.size])
    .sort((a, b) => b[1] - a[1])
    .slice(0, 5);

  const dashContactsEl = clone.getElementById("dash-top5-contacts");
  if (dashContactsEl) {
    dashContactsEl.innerHTML = top5Contacts.length
      ? top5Contacts
          .map(
            (e, i) => `
        <div class="top5-row">
          <span class="top5-rank rank-${i + 1}">${i + 1}</span>
          <span class="top5-name">${escHtml(e[0])}</span>
          <span class="top5-count" style="color: var(--blue, #3b82f6);">${e[1]} contact${e[1] !== 1 ? "s" : ""}</span>
        </div>`,
          )
          .join("")
      : `<p class="empty-state">No contacts logged yet today.</p>`;
  }

  // ✂️ REMOVED: The dash-recent-table injection

  // 5. Inject the Recycle Queue Table
  if (isAdmin() && needRecycle > 0) {
    clone.getElementById("dash-recycle-title").textContent =
      `⚠️ Recycle Queue — ${needRecycle} lead${needRecycle !== 1 ? "s" : ""} ready`;

    const recycleQueue = leads.filter(
      (l) => l.flags && l.flags.includes("needs_recycle"),
    );

    clone.getElementById("dash-recycle-table").innerHTML = `
      <table class="data-table">
        <thead>
          <tr>
            <th>Name</th><th>Address</th><th>Previously Assigned To</th><th>Last Contacted</th><th>Action</th>
          </tr>
        </thead>
        <tbody>
          ${recycleQueue
            .map(
              (l) => `
            <tr>
              <td><span class="lead-name">${escHtml(l.name)}</span></td>
              <td class="td-mono" style="font-size:11px">${escHtml(l.address || "—")}${l.city ? ", " + escHtml(l.city) : ""}</td>
              <td>
                <div style="display:flex;flex-direction:column;gap:2px">
                  ${l.assignedTo ? `<span style="font-size:13px;font-weight:600;color:#1A2640">${escHtml(l.assignedTo)}</span>` : "—"}
                  ${l.previousAgents ? `<span style="font-family:var(--font-mono);font-size:10px;color:#8EA5C8">Previously: ${escHtml(l.previousAgents)}</span>` : ""}
                </div>
              </td>
              <td class="td-mono">${formatDate(l.lastContacted) || "—"}</td>
              <td>
                <button class="btn-primary" style="padding:6px 14px;font-size:12px" onclick="recycleLeadAction('${l.id}','${escHtml(l.assignedTo || "")}','${escHtml(l.name)}')">
                  Recycle
                </button>
              </td>
            </tr>
          `,
            )
            .join("")}
        </tbody>
      </table>
    `;
  } else if (isAdmin()) {
    clone.getElementById("dash-recycle-section").style.display = "none";
  }

  // 6. Mount it to the screen!
  mainContent.appendChild(clone);

  updateBadges();
  startSalesFeedPolling();
}

async function recycleLeadAction(leadId, currentAgent, leadName) {
  if (
    !confirm(
      'Recycle "' +
        leadName +
        '"?\n\nThis will:\n• Reset status to New\n• Unassign from ' +
        (currentAgent || "current agent") +
        "\n• Record previous assignment history\n\nThe lead can then be reassigned to a different agent.",
    )
  )
    return;
  setLoading(true);
  try {
    await Graph.recycleLead(leadId, currentAgent);
    await Graph.logActivity({
      LeadID: leadId,
      Title: leadName,
      ActionType: "Recycled",
      AgentEmail: (State.currentUser && State.currentUser.email) || "",
      Notes:
        "Recycled by admin — previous agent: " + (currentAgent || "unknown"),
    });
    UI.showToast(leadName + " recycled and ready to reassign!", "success");
    await loadAllData();
    renderDashboard();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

async function recycleAllLeads() {
  const recycleLeads = State.leads.filter(function (l) {
    return l.flags && l.flags.includes("needs_recycle");
  });
  if (!recycleLeads.length) {
    UI.showToast("No leads to recycle.", "info");
    return;
  }
  if (
    !confirm(
      "Recycle ALL " +
        recycleLeads.length +
        " lead" +
        (recycleLeads.length !== 1 ? "s" : "") +
        " in the queue?\n\nThis will:\n• Reset all their statuses to New\n• Unassign them from their current agents\n• Record previous assignment history\n\nThis cannot be undone.",
    )
  )
    return;
  setLoading(true);
  try {
    for (var i = 0; i < recycleLeads.length; i++) {
      const lead = recycleLeads[i];
      await Graph.recycleLead(lead.id, lead.assignedTo || "");
      await Graph.logActivity({
        LeadID: lead.id,
        Title: lead.name,
        ActionType: "Recycled",
        AgentEmail: (State.currentUser && State.currentUser.email) || "",
        Notes:
          "Bulk recycled by admin — previous agent: " +
          (lead.assignedTo || "unknown"),
      });
    }
    UI.showToast(
      "Recycled " +
        recycleLeads.length +
        " lead" +
        (recycleLeads.length !== 1 ? "s" : "") +
        " successfully!",
      "success",
    );
    await loadAllData();
    renderDashboard();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

function startSalesFeedPolling() {
  if (State.salesFeedTimer) clearInterval(State.salesFeedTimer);

  let knownSaleIds = new Set(
    (State.todaySales || []).map(function (l) {
      return l.id;
    }),
  );

  async function pollSalesData() {
    if (document.visibilityState === "hidden") return;

    try {
      // 🚀 1. Background Sync: Keep the data fresh regardless of the view
      const leadsSyncDate = localStorage.getItem("RaimakLeadsLastSyncDate");

      const [updatedLeads, logData] = await Promise.all([
        Graph.getLeads(leadsSyncDate, State.leads),
        Graph.getActivityLog(State.lastSyncDate, State.activityLog),
      ]);

      // State is updated so the next time they click a tab, the data is current
      State.leads = updatedLeads;
      State.activityLog = logData.updatedLogs;
      State.lastSyncDate = logData.newSyncDate;

      // 🚀 2. Process Sales
      const newSales = Graph.getTodaySales(State.activityLog);
      State.todaySales = newSales;

      // 3. Confetti/Ticker logic (Global triggers)
      const newOnes = newSales.filter((l) => !knownSaleIds.has(l.id));
      if (newOnes.length > 0) {
        if (Ticker && Ticker.update) Ticker.update();
        if (UI && UI.showConfetti) UI.showConfetti();
        newOnes.forEach((l) => knownSaleIds.add(l.id));
      }

      // 🚀 4. EXCLUSIVE UI UPDATE: Only the Dashboard Feed
      // This avoids the "UI jumping" on the Lead tables or Admin reports
      if (State.currentView === "dashboard") {
        updateDashboardUI(newSales);
      }
    } catch (e) {
      console.error("Sync polling error:", e);
    }
  }

  // UI update helper stays the same
  function updateDashboardUI(newSales) {
    const feed = document.getElementById("dash-sales-feed");
    const time = document.getElementById("sales-feed-time");
    if (!feed) return;

    if (time) {
      time.textContent = "Updated " + formatTime(new Date().toISOString());
    }

    if (!newSales || !newSales.length) {
      feed.innerHTML = `<p class="empty-state" style="padding:24px; text-align:center;">No sales yet today.</p>`;
      return;
    }

    feed.innerHTML = [...newSales]
      .sort((a, b) => new Date(b.saleTime) - new Date(a.saleTime))
      .slice(0, 6)
      .map(function (l) {
        const displayAgent = l.soldBy || "Unassigned";
        const displayName = l.name || "Unknown Lead";

        return `
          <div class="sale-entry">
            <div class="sale-icon">🎉</div>
            <div class="sale-info">
              <span class="sale-name">${escHtml(displayName)}</span>
              <span class="sale-agent">${escHtml(displayAgent)}</span>
            </div>
            <span class="sale-time">${formatTime(l.saleTime)}</span>
          </div>`;
      })
      .join("");
  }

  pollSalesData();
  State.salesFeedTimer = setInterval(
    pollSalesData,
    Config.salesFeedInterval || 60000,
  );
}

// ============================================================
//  ADMIN — DRIP FEED
// ============================================================
function renderDripFeed() {
  // 1. Security & Data Prep (Kept exactly the same)
  if (!isAdmin()) {
    navigate("myleads");
    return;
  }

  const unassigned = State.leads.filter(
    (l) => !l.assignedTo && !Config.terminalStatuses.includes(l.status),
  );

  if (!State.dripLead && unassigned.length) {
    State.dripLead = unassigned[0];
  }

  const lead = State.dripLead;
  const remaining = unassigned.length;

  // 2. Setup Template
  const mainContent = document.getElementById("main-content");
  mainContent.innerHTML = "";

  const template = document.getElementById("tmpl-drip-feed");
  const clone = template.content.cloneNode(true);

  clone.getElementById("drip-subtitle").textContent =
    `// ASSIGN ONE LEAD AT A TIME · ${remaining} unassigned`;
  if (!lead) {
    clone.getElementById("drip-empty-state").style.display = "block";
    clone.getElementById("drip-header-skip").style.display = "none";
  } else {
    clone.getElementById("drip-active-state").style.display = "block";

    const typeBadge = clone.getElementById("drip-lead-type");
    if (lead.leadType) {
      typeBadge.textContent = lead.leadType;
      typeBadge.className = `lead-type-badge lead-type-${lead.leadType.toLowerCase()}`;
    } else {
      typeBadge.style.display = "none";
    }

    const statusBadge = clone.getElementById("drip-lead-status");
    statusBadge.textContent = lead.status;
    statusBadge.className = `status-badge status-${lead.status
      .toLowerCase()
      .replace(/\s+/g, "-")
      .replace(/[^a-z0-9-]/g, "")}`;

    // Build Text Fields
    clone.getElementById("drip-lead-name").textContent = lead.name;

    const notesEl = clone.getElementById("drip-notes");
    if (lead.notes) notesEl.textContent = lead.notes;
    else notesEl.style.display = "none";

    // Build Meta Icons (Phone, Email, etc.)
    let metaHtml = "";
    if (lead.phone)
      metaHtml += `<span class="feed-meta"><svg width="13" height="13" fill="none" viewBox="0 0 24 24"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07A19.5 19.5 0 0 1 4.69 12 19.79 19.79 0 0 1 1.61 3.38 2 2 0 0 1 3.6 1.22h3a2 2 0 0 1 2 1.72c.127.96.361 1.903.7 2.81a2 2 0 0 1-.45 2.11L7.91 8.96a16 16 0 0 0 6 6l.92-.92a2 2 0 0 1 2.11-.45c.907.339 1.85.573 2.81.7A2 2 0 0 1 21.73 16.92z" stroke="currentColor" stroke-width="2"/></svg>${escHtml(lead.phone)}</span>`;
    if (lead.email)
      metaHtml += `<span class="feed-meta"><svg width="13" height="13" fill="none" viewBox="0 0 24 24"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z" stroke="currentColor" stroke-width="2"/><polyline points="22,6 12,13 2,6" stroke="currentColor" stroke-width="2"/></svg>${escHtml(lead.email)}</span>`;
    if (lead.currentMRC)
      metaHtml += `<span class="feed-meta">MRC: $${escHtml(lead.currentMRC)}/mo</span>`;
    if (lead.currentProducts)
      metaHtml += `<span class="feed-meta">Has: ${escHtml(lead.currentProducts)}</span>`;
    clone.getElementById("drip-meta-container").innerHTML = metaHtml;

    // Build Agent Dropdown
    const selectEl = clone.getElementById("drip-agent-select");
    let optionsHtml = `<option value="">Select an agent...</option>`;
    State.contractors.forEach((c) => {
      const count = State.leads.filter(
        (l) =>
          l.assignedTo === c.name &&
          !Config.terminalStatuses.includes(l.status),
      ).length;
      const full = count >= Config.rules.maxLeadsPerAgent;
      optionsHtml += `<option value="${escHtml(c.name)}" ${full ? "disabled" : ""}>${escHtml(c.name)} — ${count}/${Config.rules.maxLeadsPerAgent}${full ? " (FULL)" : ""}</option>`;
    });
    selectEl.innerHTML = optionsHtml;

    // Wire up the dynamic ID to the assign button
    clone.getElementById("drip-assign-btn").onclick = () =>
      confirmDripAssign(lead.id);

    // Build Remaining Table
    clone.getElementById("drip-remaining-title").textContent =
      `Remaining Unassigned (${remaining})`;
    clone
      .getElementById("drip-remaining-table")
      .replaceChildren(renderLeadsTable(unassigned.slice(0, 10), true));
  }

  // 4. Mount
  mainContent.appendChild(clone);
}

async function confirmDripAssign(leadId) {
  const select = document.getElementById("drip-agent-select");
  const agent = select && select.value;
  if (!agent) {
    UI.showToast("Please select an agent first.", "error");
    return;
  }
  if (!Graph.canAgentTakeLead(agent, State.leads)) {
    UI.showToast(
      agent + " is at the " + Config.rules.maxLeadsPerAgent + "-lead limit.",
      "error",
    );
    return;
  }
  const lead = State.leads.find(function (l) {
    return l.id === leadId;
  });
  setLoading(true);
  try {
    await Graph.assignAgent(leadId, agent);
    await Graph.logActivity({
      LeadID: leadId,
      Title: lead ? lead.name : "",
      ActionType: "Drip Assigned",
      AgentEmail: agent,
      Notes:
        "Drip-assigned by " +
        ((State.currentUser && State.currentUser.name) || "Admin"),
    });
    UI.showToast(lead.name + " assigned to " + agent, "success");
    await loadAllData();
    const remaining = State.leads.filter(function (l) {
      return !l.assignedTo && !Config.terminalStatuses.includes(l.status);
    });
    State.dripLead = remaining.length ? remaining[0] : null;
    renderDripFeed();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

function skipDripLead() {
  const unassigned = State.leads.filter(function (l) {
    return !l.assignedTo && !Config.terminalStatuses.includes(l.status);
  });
  const currentIdx = unassigned.findIndex(function (l) {
    return State.dripLead && l.id === State.dripLead.id;
  });
  const nextIdx = (currentIdx + 1) % unassigned.length;
  State.dripLead = unassigned[nextIdx] || null;
  renderDripFeed();
}

// ============================================================
//  AGENT — MY LEADS
// ============================================================
function getStatusColor(status) {
  const colors = Config.statusColors || {};
  if ((colors.red || []).includes(status)) return "#FF4444";
  if ((colors.yellow || []).includes(status)) return "#FFD700";
  if ((colors.green || []).includes(status)) return "#00FF88";
  if ((colors.blue || []).includes(status)) return "#4D79FF";
  if ((colors.cyan || []).includes(status)) return "#00E5FF";
  if ((colors.white || []).includes(status)) return "#FFFFFF";
  return "#7A98C8";
}

function getStatusDot(status) {
  const color = getStatusColor(status);
  return `<span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:${color};box-shadow:0 0 6px ${color};flex-shrink:0;margin-right:6px"></span>`;
}

let _leadSaved = false;
let _currentFeedIndex = 0;

function renderMyLeads() {
  if (typeof _leadSaved === "undefined") window._leadSaved = false;

  // ==========================================
  // 🛑 THE SHRINKING ARRAY FIX
  // ==========================================
  if (!window._forceShowLead) {
    window._currentFeedIndex = 0;
  }

  // 🕒 KEEPING THE CLOCK LOGIC AT THE TOP
  const updateClock = () => {
    const clockEl = document.getElementById("myleads-clock");
    if (!clockEl) return;

    const activeLead = (window._myLeads || [])[_currentFeedIndex];
    const leadState =
      activeLead && activeLead.state
        ? activeLead.state.toUpperCase().trim()
        : null;

    let tz = Intl.DateTimeFormat().resolvedOptions().timeZone;
    if (leadState && stateTimezones[leadState]) {
      tz = stateTimezones[leadState];
    }
    try {
      clockEl.textContent = new Date().toLocaleTimeString([], {
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
        timeZone: tz,
        timeZoneName: "short",
      });
    } catch (e) {
      clockEl.textContent = new Date().toLocaleTimeString([], {
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
      });
    }
  };

  const user = State.currentUser;
  const userName = ((user && user.name) || "").toLowerCase().trim();
  const userEmail = ((user && user.email) || "").toLowerCase().trim();

  const contractor = State.contractors.find((c) => {
    return (
      (c.email || "").toLowerCase().trim() === userEmail ||
      (c.name || "").toLowerCase().trim() === userName
    );
  });
  const agentName = contractor
    ? contractor.name.toLowerCase().trim()
    : userName;

  // ==========================================
  //  THE STRICT BOUNCER
  // ==========================================
  let myLeads;
  let hiddenByTimezone = 0; // 🕵️ NEW: Our tracker for sleeping leads!

  const filterNow = new Date();
  const filterTodayMidnight = new Date(filterNow);
  filterTodayMidnight.setHours(0, 0, 0, 0);

  const tzHourCache = {};

  if (window._forceShowLead && window._myLeads && window._myLeads.length > 0) {
    myLeads = window._myLeads;
  } else {
    myLeads = State.leads.filter((l) => {
      // 1. Agent Match
      const assigned = (l.assignedTo || "")
        .toLowerCase()
        .replace(/\s+/g, " ")
        .trim();
      const matchesAgent =
        assigned &&
        (assigned === agentName.replace(/\s+/g, " ") ||
          assigned === userName.replace(/\s+/g, " ") ||
          assigned === userEmail.replace(/\s+/g, " "));

      if (!matchesAgent) return false;

      // 2. Terminal Status
      if (Config.terminalStatuses.includes(l.status)) return false;

      // 2.5 🛑 3rd Contact Eviction
      if (l.status === "3rd Contact") return false;

      // 3. Dismissed Leads
      if (
        window._skippedSessionLeads &&
        window._skippedSessionLeads.includes(l.id)
      ) {
        return false;
      }

      // 3.5 Patch for the lead polling
      if (window._sessionWorkedLeads && window._sessionWorkedLeads.has(l.id)) {
        const savedTime = window._sessionWorkedLeads.get(l.id);
        const minutesSinceSave = (Date.now() - savedTime) / 60000;

        if (minutesSinceSave < 5) {
          return false;
        } else {
          window._sessionWorkedLeads.delete(l.id);
        }
      }

      // 4. Callback Math
      let waitingForDate = false;
      let isDueCallback = false;

      if (l.callbackAt) {
        const scheduledDate = new Date(l.callbackAt);
        scheduledDate.setHours(0, 0, 0, 0);

        if (l.status === "Pending Order") {
          if (filterTodayMidnight <= scheduledDate) waitingForDate = true;
        } else {
          if (filterTodayMidnight < scheduledDate) waitingForDate = true;
        }
        if (!waitingForDate) isDueCallback = true;
      }

      if (waitingForDate) return false;

      // 5. Cool-Off Shield
      const inCoolOff = Graph.isInCoolOff(l);
      const passedCoolOff = isDueCallback ? true : !inCoolOff;

      if (!passedCoolOff) return false;

      // ==========================================
      // 🚀 THE TIMEZONE SHIELD
      // ==========================================
      if (l.state) {
        let tz = "America/New_York";
        if (typeof stateTimezones !== "undefined" && stateTimezones[l.state]) {
          tz = stateTimezones[l.state];
        }

        if (tzHourCache[tz] === undefined) {
          try {
            tzHourCache[tz] = parseInt(
              new Intl.DateTimeFormat("en-US", {
                hour: "numeric",
                hour12: false,
                timeZone: tz,
              }).format(filterNow),
              10,
            );
          } catch (e) {
            console.warn("Timezone calculation failed for tz:", tz);
            tzHourCache[tz] = 12;
          }
        }

        const localHour = tzHourCache[tz];
        const isAwake = localHour >= 8 && localHour < 20;

        if (!isAwake) {
          hiddenByTimezone++; // 🎯 THE TRACKER: Catch the sleeping lead before it drops
          return false;
        }
      }

      return true;
    });

    const statusPriority = {
      "3rd Contact": 1,
      "2nd Contact": 2,
      "1st Contact": 3,
    };

    myLeads.sort((a, b) => {
      const aIsCallback = !!a.callbackAt;
      const bIsCallback = !!b.callbackAt;
      if (aIsCallback && !bIsCallback) return -1;
      if (!aIsCallback && bIsCallback) return 1;
      if (aIsCallback && bIsCallback)
        return new Date(a.callbackAt) - new Date(b.callbackAt);

      const aWeight = statusPriority[a.status] || 4;
      const bWeight = statusPriority[b.status] || 4;

      if (aWeight !== bWeight) {
        return aWeight - bWeight;
      }

      return a.id.localeCompare(b.id);
    });
  }

  const rawMyLeads = State.leads.filter((l) => {
    const assigned = (l.assignedTo || "")
      .toLowerCase()
      .replace(/\s+/g, " ")
      .trim();
    return (
      assigned &&
      (assigned === agentName ||
        assigned === userName ||
        assigned === userEmail)
    );
  });

  console.log("--- QUEUE DIAGNOSTIC ---");
  console.log(
    `Total leads technically assigned to this agent in RAM: ${rawMyLeads.length}`,
  );
  if (hiddenByTimezone > 0)
    console.log(`🌙 Sleeping Leads Caught: ${hiddenByTimezone}`);
  console.log("------------------------");

  window._myLeads = myLeads;
  window._agentName = agentName;
  _leadSaved = false;
  window._forceShowLead = false;

  // ==========================================
  //  ☕ THE FINISH LINE (Preserving the UI)
  // ==========================================
  const mainContent = document.getElementById("main-content");
  mainContent.innerHTML = "";

  const template = document.getElementById("tmpl-my-leads");
  const clone = template.content.cloneNode(true);

  const contactsToday =
    typeof getMyContactsToday === "function" ? getMyContactsToday() : 0;
  const textEl = clone.getElementById("myleads-contact-text");
  if (textEl) textEl.textContent = contactsToday;

  const subtitleEl = clone.getElementById("myleads-subtitle");
  const feedWrap = clone.getElementById("lead-feed-wrap");

  // If the queue is empty, inject the empty state INTO the feed wrapper
  if (myLeads.length === 0 || _currentFeedIndex >= myLeads.length) {
    _currentFeedIndex = 0;
    if (window._clockTimer) clearInterval(window._clockTimer);

    if (subtitleEl) subtitleEl.textContent = `// 0 remaining`;

    if (feedWrap) {
      // 🎨 Dynamic HTML: Inject the banner ONLY if there are sleeping leads
      const sleepingBannerHTML =
        hiddenByTimezone > 0
          ? `<div style="background: var(--blue-light, #e0e7ff); color: var(--blue-dark, #3730a3); padding: 10px 16px; border-radius: 8px; display: inline-block; margin-bottom: 24px; font-size: 14px; font-weight: 600;">
             🌙 ${hiddenByTimezone} lead${hiddenByTimezone !== 1 ? "s" : ""} resting outside 8 AM - 8 PM dialing hours.
           </div><br>`
          : ``;

      feedWrap.innerHTML = `
        <div class="card" style="text-align:center; padding:60px 20px;">
          <div style="font-size:4rem; margin-bottom:20px;">☕</div>
          <h2 class="view-title">Queue Clear!</h2>
          <p style="color:var(--text-3); margin-bottom:${hiddenByTimezone > 0 ? "16px" : "24px"};">Worked everything available for now.</p>
          
          ${sleepingBannerHTML}

          <button class="btn-primary" onclick="this.innerHTML='Syncing...'; this.disabled=true; typeof loadAllData === 'function' ? loadAllData() : window.location.reload();">Check for Updates</button>
        </div>`;
    }

    mainContent.appendChild(clone);
    return;
  }

  // ==========================================
  //  THE RENDER LOGIC (If there are leads)
  // ==========================================
  if (subtitleEl) {
    subtitleEl.textContent = `// ${myLeads.length} remaining · lead ${_currentFeedIndex + 1} of ${myLeads.length}`;
  }

  if (feedWrap) {
    feedWrap.innerHTML = "";
    feedWrap.appendChild(renderLeadFeedCard(myLeads));
  }

  mainContent.appendChild(clone);

  updateClock();
  if (window._clockTimer) clearInterval(window._clockTimer);
  window._clockTimer = setInterval(updateClock, 1000);
}

function searchMyLeads(q) {
  const wrap = document.getElementById("my-leads-table");
  if (!wrap) return;

  if (!q.trim()) {
    wrap.innerHTML = "";
    return;
  }

  const queryLower = q.trim().toLowerCase();
  const toggleEl = document.getElementById("toggle-search-all");
  const searchAll = toggleEl ? toggleEl.checked : true;

  // Grab the current agent's identity (matching the logic from your render function)
  const agentName = (window._agentName || "")
    .toLowerCase()
    .replace(/\s+/g, " ");
  const userName = ((State.currentUser && State.currentUser.name) || "")
    .toLowerCase()
    .replace(/\s+/g, " ");
  const userEmail = ((State.currentUser && State.currentUser.email) || "")
    .toLowerCase()
    .replace(/\s+/g, " ");

  // THE FIX: Always search the master database so we catch cool-off and terminal leads!
  const filtered = (State.leads || []).filter((l) => {
    const assigned = (l.assignedTo || "")
      .toLowerCase()
      .replace(/\s+/g, " ")
      .trim();

    let passesAssignmentFilter = false;

    if (searchAll) {
      // Searching everything: just verify it's assigned to someone
      passesAssignmentFilter = assigned !== "";
    } else {
      // Searching "My Leads": verify it belongs to THIS agent specifically
      passesAssignmentFilter =
        assigned &&
        (assigned === agentName ||
          assigned === userName ||
          assigned === userEmail);
    }

    const matchesSearch =
      (l.name && l.name.toLowerCase().includes(queryLower)) ||
      (l.phone && l.phone.includes(queryLower)) ||
      (l.btn && l.btn.includes(queryLower)) ||
      (l.cbr && l.cbr.includes(queryLower)) ||
      (l.address && l.address.toLowerCase().includes(queryLower));

    return passesAssignmentFilter && matchesSearch;
  });

  if (filtered.length) {
    // THE QOL UPGRADE: Slice down to 25 max for instant rendering
    const displayLeads = filtered.slice(0, 25);

    // Draw the table
    wrap.replaceChildren(renderLeadsTable(displayLeads, false, true));

    // Add a helpful hint if we truncated the list
    if (filtered.length > 25) {
      const hint = document.createElement("div");
      hint.style.cssText =
        "text-align:center; padding:12px; font-size:12px; color:#64748b; font-style:italic; border-top:1px solid #e2e8f0;";
      hint.textContent = `+ ${filtered.length - 25} more matches. Keep typing to narrow it down.`;
      wrap.appendChild(hint);
    }
  } else {
    wrap.innerHTML = `<div class="empty-state">No leads found for "${escHtml(q)}"</div>`;
  }
}

function getMyContactsToday() {
  const logs = State.activityLog || [];
  const user = State.currentUser || {};

  const myEmail = (user.email || "").toLowerCase().trim();
  let myName = (user.name || "").toLowerCase().trim();

  const contractor = State.contractors.find(
    (c) =>
      (c.email || "").toLowerCase().trim() === myEmail ||
      (c.name || "").toLowerCase().trim() === myName,
  );
  if (contractor) myName = contractor.name.toLowerCase().trim();

  // 1. Get today's local date (e.g., "4/18/2026")
  const todayString = new Date().toLocaleDateString();

  const uniqueLeads = new Set();

  logs.forEach((log) => {
    const entryAgent = (log.agent || log.AgentEmail || "").toLowerCase().trim();
    const actionStr = log.action || log.ActionType || "";
    const leadId = log.leadId || log.LeadID;

    // 2. Convert the log's timestamp to a local date string and compare
    let isToday = false;
    if (log.timestamp) {
      isToday = new Date(log.timestamp).toLocaleDateString() === todayString;
    }

    const isMyLog = entryAgent === myEmail || entryAgent === myName;
    const isContact =
      actionStr.startsWith("Status:") ||
      actionStr === "1st Contact" ||
      actionStr === "2nd Contact" ||
      actionStr === "3rd Contact";

    // 3. The new requirement: It MUST happen today
    if (isMyLog && isContact && isToday && leadId) {
      uniqueLeads.add(leadId);
    }
  });

  return uniqueLeads.size;
}
let _stagedStatus = null;

function renderLeadFeedCard(myLeads) {
  // 1. FIXED LOGIC: Grab the exact lead we are supposed to be looking at!
  let lead = myLeads[_currentFeedIndex];

  const isCoolOff = lead ? Graph.isInCoolOff(lead) : false;

  _stagedStatus = null;
  window._forceShowLead = false;

  // 2. Setup Template
  const template = document.getElementById("tmpl-lead-feed-card");
  const clone = template.content.cloneNode(true);

  const emptyState = clone.getElementById("feed-card-empty");
  const activeState = clone.getElementById("feed-card-active");

  // 3. Handle Empty State
  if (!lead) {
    emptyState.style.display = "flex";
    clone.getElementById("feed-empty-text").textContent =
      myLeads.length > 0
        ? "Remaining leads are in the cool-off period."
        : "No leads assigned yet — ask your manager.";

    const wrapper = document.createElement("div");
    wrapper.appendChild(clone);
    return wrapper;
  }

  // 4. Handle Active Lead State
  activeState.style.display = "block";

  // Badges & Name
  const typeBadge = clone.getElementById("feed-lead-type");
  if (lead.leadType) {
    typeBadge.textContent = lead.leadType;
    typeBadge.className = `lead-type-badge lead-type-${lead.leadType.toLowerCase()}`;
  } else {
    typeBadge.style.display = "none";
  }

  const statusBadge = clone.getElementById("feed-current-status");
  statusBadge.textContent = lead.status;
  statusBadge.className = `status-badge status-${lead.status
    .toLowerCase()
    .replace(/\s+/g, "-")
    .replace(/[^a-z0-9-]/g, "")}`;

  clone.getElementById("feed-lead-name").textContent = lead.name;

  // Cooloff Alert
  if (isCoolOff) {
    const alert = clone.getElementById("feed-cooloff-alert");
    alert.style.display = "block";
    alert.textContent = `⏱ This lead is in the ${Config.rules.coolOffDays}-day cool-off period — you can still update it if the customer reached out.`;
  }

  // ==========================================
  // 1. THE CALLBACK / INSTALL ALERT BADGE
  // ==========================================
  const callbackAlert = clone.getElementById("feed-callback-alert");
  if (callbackAlert && lead.callbackAt) {
    const targetDate = new Date(lead.callbackAt);
    const today = new Date();

    const todayMidnight = new Date(today);
    todayMidnight.setHours(0, 0, 0, 0);
    const targetMidnight = new Date(targetDate);
    targetMidnight.setHours(0, 0, 0, 0);

    if (todayMidnight >= targetMidnight) {
      const timeString = targetDate.toLocaleTimeString([], {
        hour: "2-digit",
        minute: "2-digit",
      });

      if (lead.status === "Pending Order") {
        callbackAlert.innerHTML =
          "📅 INSTALL DUE: Check if fiber is active today!";
      } else {
        callbackAlert.innerHTML = `📅 ACTION REQUIRED: Scheduled follow-up today at ${timeString}`;
      }
      callbackAlert.style.display = "block";
    }
  }

  // Meta Row (Icons)
  let metaHtml = "";
  if (lead.phone)
    metaHtml += `<span class="feed-meta">📞 ${escHtml(lead.phone)}</span>`;
  if (lead.email)
    metaHtml += `<span class="feed-meta">✉️ ${escHtml(lead.email)}</span>`;
  if (lead.address)
    metaHtml += `<span class="feed-meta">📍 ${escHtml(lead.address)}${lead.city ? ", " + escHtml(lead.city) : ""}${lead.state ? " " + escHtml(lead.state) : ""}${lead.zip ? " " + escHtml(lead.zip) : ""}</span>`;
  clone.getElementById("feed-meta-container").innerHTML = metaHtml;

  // Form Inputs
  clone.getElementById("feed-btn").value = lead.btn || "";
  clone.getElementById("feed-mrc").value = lead.currentMRC || "";
  clone.getElementById("feed-cbr").value = lead.cbr || "";

  // ==========================================
  // 2. PULLING THE SAVED CALLBACK DATE UI
  // ==========================================
  const callbackInput = clone.getElementById("f-callback-date");
  const callbackWrap = clone.getElementById("callback-wrapper");
  const callbackBtn = clone.getElementById("btn-toggle-callback");
  const callbackLabel = clone.getElementById("callback-label");

  if (callbackInput && lead.callbackAt) {
    const localDate = new Date(lead.callbackAt);
    const tzOffset = localDate.getTimezoneOffset() * 60000;
    const localISOTime = new Date(localDate - tzOffset)
      .toISOString()
      .slice(0, 16);

    callbackInput.value = localISOTime;

    if (callbackWrap) {
      callbackWrap.style.width = "200px";
      callbackWrap.style.opacity = "1";
      callbackWrap.style.overflow = "visible";
      callbackWrap.dataset.manuallyOpened = "false";
    }

    if (lead.status === "Pending Order" && callbackBtn) {
      if (callbackLabel)
        callbackLabel.innerHTML =
          'Scheduled Install <span style="color: var(--red)">*</span>';
      callbackBtn.disabled = true;
      callbackBtn.style.opacity = "0.4";
      callbackBtn.style.cursor = "not-allowed";
      callbackInput.required = true;
    }
  }

  if (callbackInput) {
    callbackInput.addEventListener("change", (e) => {
      const selectedVal = e.target.value;

      if (callbackWrap) {
        callbackWrap.dataset.manuallyOpened = "true";
      }

      if (selectedVal) {
        callbackInput.style.transition =
          "background-color 0.3s, border-color 0.3s";
        callbackInput.style.backgroundColor = "var(--green-dim, #e6f8f3)";
        callbackInput.style.borderColor = "var(--green, #10b981)";

        setTimeout(() => {
          callbackInput.style.backgroundColor = "";
        }, 600);

        if (lead.status !== "Pending Order" && callbackBtn) {
          if (callbackLabel) {
            const d = new Date(selectedVal);
            const formattedStr =
              d.toLocaleDateString([], { month: "short", day: "numeric" }) +
              " @ " +
              d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            callbackLabel.innerHTML = `✅ Set: ${formattedStr}`;
            callbackLabel.style.color = "var(--green, #10b981)";
          }

          callbackBtn.disabled = true;
          callbackBtn.style.opacity = "0.4";
          callbackBtn.style.cursor = "not-allowed";

          setTimeout(() => {
            if (lead.status !== "Pending Order") {
              callbackBtn.disabled = false;
              callbackBtn.style.opacity = "1";
              callbackBtn.style.cursor = "pointer";
            }
          }, 2000);
        }
      } else {
        callbackInput.style.borderColor = "";
        if (callbackLabel) {
          callbackLabel.innerHTML = "Callback date and time";
          callbackLabel.style.color = "#6b85b0";
        }
      }
    });
  }

  // Products Dropdown
  const productsSelect = clone.getElementById("feed-products");
  Config.currentProducts.forEach((p) => {
    const option = document.createElement("option");
    option.value = p;
    option.textContent = p;
    if (lead.currentProducts === p) option.selected = true;
    productsSelect.appendChild(option);
  });

  // Sold By Dropdown
  const soldBySelect = clone.getElementById("feed-sold-by");
  State.contractors.forEach((c) => {
    const option = document.createElement("option");
    option.value = c.name;
    option.textContent = c.name;
    soldBySelect.appendChild(option);
  });

  // THE DRAFT PEEK
  const draft = State.drafts[lead.id] || {};
  const activeAutoPay =
    draft.autoPay !== undefined ? draft.autoPay : lead.autoPay;

  // AutoPay Radios
  const autoPayContainer = clone.getElementById("feed-autopay-container");
  ["ACH - Debit Card", "ACH - Credit Card", "No Auto Pay"].forEach((opt) => {
    const isChecked = activeAutoPay === opt ? "checked" : "";
    autoPayContainer.innerHTML += `
      <label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-size:13px;color:#1A2640;background:#F4F7FD;border:1px solid #D0DCF0;padding:8px 14px;border-radius:6px;">
        <input type="radio" name="feed-autopay" value="${opt}" ${isChecked} style="accent-color:#2563B0"> ${opt}
      </label>`;
  });

  // Status Buttons
  const statusContainer = clone.getElementById("feed-status-buttons");
  const hiddenStatuses = ["New", "TD Non-Reg", "D2D Lead"];

  Config.leadStatuses
    .filter((s) => !hiddenStatuses.includes(s))
    .forEach((s) => {
      const isTDM = s === "TDM" ? " ↩" : "";
      const cls =
        "status-btn-" +
        s
          .toLowerCase()
          .replace(/\s+/g, "-")
          .replace(/[^a-z0-9-]/g, "");
      statusContainer.innerHTML += `<button class="status-btn ${cls}" id="sbtn-${s.replace(/\s+/g, "-")}" onclick="stageStatus('${lead.id}','${s}')">${s}${isTDM}</button>`;
    });

  clone.getElementById("feed-today-date").textContent =
    new Date().toLocaleDateString("en-US", {
      month: "2-digit",
      day: "2-digit",
      year: "2-digit",
    });

  // THE RESTORED LEGACY NOTES PARSER
  const pastNotesContainer = clone.getElementById("feed-past-notes-container");
  if (lead.notes && lead.notes.trim()) {
    pastNotesContainer.style.display = "block";
    const notesHtml = lead.notes
      .split("\n")
      .filter((l) => l.trim())
      .map((line) => {
        const match = line.match(/^\[(\d{2}\/\d{2}(?:\/\d{2})?)(.*?)\]\s*(.*)/);
        if (match) {
          const date = match[1];
          const agent = match[2] ? match[2].replace(/^\s*-\s*/, "") : "";
          const text = match[3];
          return `
            <div style="margin-bottom:8px;padding-bottom:8px;border-bottom:1px solid #E8EFF8">
              <div style="display:flex;gap:8px;align-items:center;margin-bottom:3px">
                <span style="font-family:var(--font-mono);font-size:10px;color:#2563B0;font-weight:700;background:#E8F0FF;padding:1px 6px;border-radius:3px">${date}</span>
                ${agent ? `<span style="font-family:var(--font-mono);font-size:10px;color:#6B85B0">${escHtml(agent)}</span>` : ""}
              </div>
              <span style="font-size:13px;color:#1A2640">${escHtml(text)}</span>
            </div>`;
        }
        return `
          <div style="margin-bottom:8px;padding-bottom:8px;border-bottom:1px solid #E8EFF8">
            <div style="margin-bottom:3px">
              <span style="font-family:var(--font-mono);font-size:10px;color:#8EA5C8;background:#F4F7FD;padding:1px 6px;border-radius:3px">Legacy note — author unknown</span>
            </div>
            <span style="font-size:13px;color:#4A6080">${escHtml(line)}</span>
          </div>`;
      })
      .join("");
    pastNotesContainer.innerHTML = notesHtml;
  }

  // Save Button Action
  clone.getElementById("feed-save-btn").onclick = () => agentSaveAll(lead.id);

  // THE DRAFT MEMORY ENABLER
  const inputsToDraft = [
    { id: "feed-btn", key: "btn" },
    { id: "feed-mrc", key: "mrc" },
    { id: "feed-cbr", key: "cbr" },
    { id: "feed-notes", key: "notes" },
    { id: "feed-products", key: "products" },
    { id: "feed-sold-by", key: "soldBy" },
  ];

  inputsToDraft.forEach((item) => {
    const el = clone.getElementById(item.id);
    if (el) {
      if (draft[item.key] !== undefined) {
        el.value = draft[item.key];
      }
      const eventType = el.tagName === "SELECT" ? "change" : "input";
      el.addEventListener(eventType, (e) =>
        updateLeadDraft(lead.id, item.key, e.target.value),
      );
    }
  });

  const allAutoPayRadios = clone.querySelectorAll('input[name="feed-autopay"]');
  if (allAutoPayRadios.length > 0) {
    allAutoPayRadios.forEach((r) => {
      r.addEventListener("change", (e) =>
        updateLeadDraft(lead.id, "autoPay", e.target.value),
      );
    });
  }

  // Wrap and Return
  const wrapper = document.createElement("div");
  wrapper.appendChild(clone);
  return wrapper;
}

function toggleCallbackDate() {
  const wrap = document.getElementById("callback-wrapper");
  const btn = document.getElementById("btn-toggle-callback");
  const input = document.getElementById("f-callback-date");
  const label = document.getElementById("callback-label");

  // Prevent toggling if it's locked (like for Pending Orders or during our 2-second cooldown!)
  if (btn.disabled) return;

  if (wrap.style.width === "0px" || wrap.style.width === "") {
    // Slide & Fade IN
    wrap.style.width = "200px";
    wrap.style.opacity = "1";

    // THE MEMORY: Remember that the agent specifically asked for this to be open
    wrap.dataset.manuallyOpened = "true";

    setTimeout(() => {
      wrap.style.overflow = "visible";
    }, 300);
  } else {
    // Slide & Fade OUT
    wrap.style.width = "0px";
    wrap.style.opacity = "0";
    wrap.style.overflow = "hidden";

    // THE MEMORY: Remember that the agent closed it
    wrap.dataset.manuallyOpened = "false";

    // Scrub the data and reset the UI *after* the menu finishes sliding shut
    setTimeout(() => {
      if (input) {
        input.value = "";
        input.style.borderColor = "";
        input.style.backgroundColor = "";
      }

      // Reset the label ONLY if it's our dynamic "✅ Set:" text
      if (label && label.innerHTML.includes("✅ Set:")) {
        label.innerHTML = "Callback date and time"; // <-- Updated to match template!
        label.style.color = "#6b85b0";
      }
    }, 300);
  }
}

function updateCallbackUIForStatus(status) {
  const wrap = document.getElementById("callback-wrapper");
  const label = document.getElementById("callback-label");
  const btn = document.getElementById("btn-toggle-callback");
  const dateInput = document.getElementById("f-callback-date");

  if (!wrap || !label || !btn || !dateInput) return;

  if (status === "Pending Order") {
    // Force the menu open instantly and lock the button
    wrap.style.width = "200px";
    wrap.style.opacity = "1";
    wrap.style.overflow = "visible";

    label.innerHTML =
      'Scheduled Install <span style="color: var(--red)">*</span>';
    btn.disabled = true;
    btn.style.opacity = "0.4";
    btn.style.cursor = "not-allowed";
    dateInput.required = true;
  } else {
    // Return everything to normal callback mode
    label.textContent = "Callback date and time";
    btn.disabled = false;
    btn.style.opacity = "1";
    btn.style.cursor = "pointer";
    dateInput.required = false;

    // THE FIX: Check the memory! If they didn't manually open it before, slide it shut.
    if (wrap.dataset.manuallyOpened !== "true") {
      wrap.style.width = "0px";
      wrap.style.opacity = "0";
      wrap.style.overflow = "hidden";

      // Wait for the slide animation to finish before clearing the data
      setTimeout(() => {
        if (wrap.style.width === "0px") dateInput.value = "";
      }, 300);
    }
  }
}

function stageStatus(leadId, newStatus) {
  const lead = State.leads.find(function (l) {
    return l.id === leadId;
  });
  if (!lead) return;

  if (Graph.isInCoolOff(lead) && !Config.terminalStatuses.includes(newStatus)) {
    UI.showToast(
      "Note: this lead is in the " +
        Config.rules.coolOffDays +
        "-day cool-off period.",
      "info",
    );
  }

  _stagedStatus = newStatus;
  updateCallbackUIForStatus(newStatus);
  document.querySelectorAll(".status-btn").forEach(function (btn) {
    btn.style.borderColor = "";
    btn.style.color = "";
    btn.style.background = "";
    btn.style.boxShadow = "";
  });
  const selectedBtn = document.getElementById(
    "sbtn-" + newStatus.replace(/\s+/g, "-"),
  );
  if (selectedBtn) {
    selectedBtn.style.borderColor = "var(--cyan)";
    selectedBtn.style.color = "var(--cyan)";
    selectedBtn.style.background = "var(--cyan-dim)";
    selectedBtn.style.boxShadow = "0 0 12px var(--cyan-glow)";
  }

  const notice = document.getElementById("feed-staged-notice");
  if (notice) {
    notice.style.display = "block";
    notice.textContent =
      '⚡ "' + newStatus + '" staged — click Save to confirm';
  }

  const badge = document.getElementById("feed-current-status");
  if (badge) {
    badge.textContent = newStatus + " (staged)";
    badge.style.opacity = "0.7";
  }
}

async function agentSaveAll(leadId) {
  const user = State.currentUser;
  window._sessionWorkedLeads = window._sessionWorkedLeads || new Map();
  window._sessionWorkedLeads.set(leadId, Date.now());
  const lead = State.leads.find((l) => l.id === leadId);
  if (!lead) return;

  const newStatus = _stagedStatus || lead.status;

  // 1. Grab UI Values
  const mrc = (document.getElementById("feed-mrc") || {}).value || "";
  const productsSelectEl = document.getElementById("feed-products");
  let products = "";
  if (
    productsSelectEl &&
    productsSelectEl.options.length > 0 &&
    productsSelectEl.selectedIndex !== -1
  ) {
    products = productsSelectEl.options[productsSelectEl.selectedIndex].value;
  }
  const newNote = (document.getElementById("feed-notes") || {}).value || "";
  const cbr = (document.getElementById("feed-cbr") || {}).value || "";
  const btn = (document.getElementById("feed-btn") || {}).value || "";
  const autoPayEl = document.querySelector(
    'input[name="feed-autopay"]:checked',
  );
  const autoPay = autoPayEl ? autoPayEl.value : "";
  const soldByEl = document.getElementById("feed-sold-by");
  const soldByName = soldByEl ? soldByEl.value : "";
  let rawCallbackDate =
    (document.getElementById("f-callback-date") || {}).value || "";

  // ==========================================
  // 🛑 THE BOUNCER: Prevent saving as "New"
  // ==========================================
  if (newStatus === "New") {
    return UI.showToast("Please update the lead status from 'New'.", "warning");
  }

  // Validation (Terminal Bypass)
  const isTerminal = Config.terminalStatuses.includes(newStatus);
  if (!isTerminal) {
    if (!autoPay) return UI.showToast("Select AutoPay", "error");
    if (newStatus === Config.soldStatus && !soldByName)
      return UI.showToast("Select Sold By", "error");
    if (!mrc) return UI.showToast("Enter MRC", "error");
    if (btn.replace(/\D/g, "").length !== 10)
      return UI.showToast("Valid BTN required", "error");
  }

  // Note Stamping
  let notes = lead.notes || "";
  if (newNote.trim()) {
    const today = new Date();
    const dateStamp =
      (today.getMonth() + 1).toString().padStart(2, "0") +
      "/" +
      today.getDate().toString().padStart(2, "0") +
      "/" +
      String(today.getFullYear()).slice(-2);
    const agentTag = user && user.name ? " - " + user.name : "";
    const stamped = `[${dateStamp}${agentTag}] ${newStatus} - ${newNote.trim()}`;
    notes = notes ? stamped + "\n" + notes : stamped;
  }

  // Activity Log Email Resolution
  const soldByContractor = soldByName
    ? State.contractors.find((c) => c.name === soldByName)
    : null;
  const soldByEmail = soldByContractor
    ? soldByContractor.email || soldByName
    : (user && user.email) || "";
  const activityEmail =
    newStatus === Config.soldStatus ? soldByEmail : (user && user.email) || "";

  // 2. Setup Payload for SharePoint
  const todayDate = new Date().toISOString().split("T")[0];
  const saveFields = {
    Status: newStatus,
    LastTouchedOn: todayDate,
    Notes: notes,
  };

  if (mrc) saveFields["MonthlyRecurringCharge_x0028_MRC"] = mrc;
  if (products) saveFields["CurrentProducts"] = products;
  if (cbr) saveFields["CBR"] = cbr;
  if (btn) saveFields["BTN"] = btn;
  if (autoPay) saveFields["AutoPay"] = autoPay;

  saveFields["CallbackDateTime"] = rawCallbackDate
    ? new Date(rawCallbackDate).toISOString()
    : newStatus === "Pending Order"
      ? lead.callbackAt
      : null;

  if (newStatus === "TDM") {
    saveFields["Agent_x0020_Assigned"] = null;
  }
  setLoading(true);
  try {
    const logEntry = {
      LeadID: leadId,
      Title: lead.name || "Unknown Lead",
      ActionType: "Status: " + newStatus,
      AgentEmail: activityEmail,
      Notes:
        notes +
        (newStatus === Config.soldStatus && soldByName
          ? ` [Sold by ${soldByName}]`
          : ""),
    };

    await Promise.all([
      Graph.updateLead(leadId, saveFields),
      Graph.logActivity(logEntry),
    ]);

    // LOCAL STATE: Fixed leadName mapping
    State.activityLog.push({
      id: "local-" + Date.now(),
      leadId: leadId,
      leadName: lead.name || "Unknown Lead",
      agent: activityEmail,
      agentEmail: activityEmail,
      action: "Status: " + newStatus,
      notes: notes,
      timestamp: new Date().toISOString(),
    });

    // 3. Update RAM (Optimistic UI)
    lead.status = newStatus;
    lead.notes = notes;
    if (mrc) lead.currentMRC = mrc;
    if (products) lead.currentProducts = products;
    if (cbr) lead.cbr = cbr;
    if (btn) lead.btn = btn;
    if (autoPay) lead.autoPay = autoPay;
    lead.callbackAt = rawCallbackDate || null;

    Points.awardPoints(newStatus, leadId);

    if (newStatus === "TDM") {
      UI.showToast("TDM — lead returned to admin queue.", "info");
    } else {
      UI.showToast("Saved!", "success");
    }

    // UI State
    _stagedStatus = null;
    _leadSaved = true;

    // Update Save Button
    const saveBtn = document.getElementById("feed-save-btn");
    if (saveBtn) {
      // 1. Trigger the Success State
      saveBtn.textContent = "Saved ✓";
      saveBtn.disabled = true;
      saveBtn.style.background = "var(--green, #10b981)";
      saveBtn.style.borderColor = "var(--green, #10b981)";
      saveBtn.style.cursor = "default";

      // 2. The 2-Second Cooldown & Reset
      setTimeout(() => {
        saveBtn.textContent = "Save";
        saveBtn.disabled = false;
        saveBtn.style.background = "";
        saveBtn.style.borderColor = "";
        saveBtn.style.cursor = "pointer";
      }, 2000);
    }

    // Show Next Row
    const nextRow = document.getElementById("feed-next-row");
    if (nextRow) {
      nextRow.style.display = "block";
      const nextBtn = nextRow.querySelector("button");
      if (nextBtn) {
        if (window._isWorkingCallback) {
          nextBtn.innerHTML = "Complete Callback ✓";
          nextBtn.classList.replace("btn-cyan", "btn-green");
          nextBtn.onclick = () => {
            window._isWorkingCallback = false;
            completeCallbackLead(leadId);
          };
        } else {
          nextBtn.onclick = () => advanceToNextLead();
        }
      }
    }

    if (Ticker && Ticker.update) Ticker.update();
  } catch (err) {
    console.error("Save Error:", err);
    UI.showToast("Failed to save: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

async function skipOutOfHoursLead(leadId) {
  const lead = State.leads.find((l) => l.id === leadId);
  if (!lead) return;

  // 1. Grab any data they managed to scrub/find
  const mrc = (document.getElementById("feed-mrc") || {}).value || "";
  const products = (document.getElementById("feed-products") || {}).value || "";
  const cbr = (document.getElementById("feed-cbr") || {}).value || "";
  const btn = (document.getElementById("feed-btn") || {}).value || "";
  const newNote = (document.getElementById("feed-notes") || {}).value || "";

  // Note Stamping
  let notes = lead.notes || "";
  if (newNote.trim()) {
    const today = new Date();
    const dateStamp =
      (today.getMonth() + 1).toString().padStart(2, "0") +
      "/" +
      today.getDate().toString().padStart(2, "0") +
      "/" +
      String(today.getFullYear()).slice(-2);
    const stamped = `[${dateStamp} - Skipped] Out of Hours Scrub - ${newNote.trim()}`;
    notes = notes ? stamped + "\n" + notes : stamped;
  }

  // 2. Setup the "Stealth" Payload
  // Notice we DO NOT change the Status, and we DO NOT update LastContacted.
  //const todayDate = new Date().toISOString().split("T")[0];
  const saveFields = {
    //LastTouchedOn: todayDate, // Resets the recycle clock so they keep it!
    Notes: notes,
  };

  if (mrc) saveFields["MonthlyRecurringCharge_x0028_MRC"] = mrc;
  if (products) saveFields["CurrentProducts"] = products;
  if (cbr) saveFields["CBR"] = cbr;
  if (btn) saveFields["BTN"] = btn;

  setLoading(true);
  try {
    // 3. Save silently to SharePoint
    await Graph.updateLead(leadId, saveFields);

    // 4. Update Local RAM
    lead.notes = notes;
    if (mrc) lead.currentMRC = mrc;
    if (products) lead.currentProducts = products;
    if (cbr) lead.cbr = cbr;
    if (btn) lead.btn = btn;

    UI.showToast("Data scrubbed and lead skipped for later!", "info");

    // 5. Move to the next lead without burning this one
    advanceToNextLead();
  } catch (err) {
    console.error("Skip Error:", err);
    UI.showToast("Failed to skip: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

function advanceToNextLead() {
  const currentLead = window._myLeads[_currentFeedIndex];

  if (currentLead) {
    let isDismissedCallback = false;

    if (currentLead.callbackAt) {
      // 1. Initialize the memory bank just in case
      if (!window._skippedSessionLeads) window._skippedSessionLeads = [];

      // 2. Prevent duplicates and memorize the dismiss
      if (!window._skippedSessionLeads.includes(currentLead.id)) {
        window._skippedSessionLeads.push(currentLead.id);
        sessionStorage.setItem(
          "_skippedSessionLeads",
          JSON.stringify(window._skippedSessionLeads),
        );
      }
      isDismissedCallback = true;
    }

    const isTerminal = Config.terminalStatuses.includes(currentLead.status);
    const isExhausted = currentLead.status === "3rd Contact";
    const inCoolOff = Graph.isInCoolOff(currentLead);

    // 🛑 THE FIX: Did the agent literally JUST click the save button on this lead?
    const justWorked =
      window._sessionWorkedLeads &&
      window._sessionWorkedLeads.has(currentLead.id);

    // 3. Only increment the index if the lead is DEFINITELY staying in the active queue!
    if (
      !isTerminal &&
      !isExhausted &&
      !inCoolOff &&
      !isDismissedCallback &&
      !justWorked
    ) {
      _currentFeedIndex++;
    }
  }

  _leadSaved = false;
  renderMyLeads();
}

// ============================================================
//  ASSIGN LEADS (Admin only)
// ============================================================
function renderAssignLeads() {
  if (!isAdmin()) {
    navigate("myleads");
    return;
  }

  const { leads, contractors } = State;
  const unassigned = leads.filter(function (l) {
    const isValidLead = l && l.id && (l.name || l.phone || l.BTN || l.btn);
    const isAvailable =
      !l.assignedTo && !Config.terminalStatuses.includes(l.status);
    return isValidLead && isAvailable;
  });

  const agentCounts = {};
  contractors.forEach((c) => (agentCounts[c.name] = 0));
  leads.forEach((l) => {
    if (
      l.assignedTo &&
      !Config.terminalStatuses.includes(l.status) &&
      agentCounts[l.assignedTo] !== undefined
    ) {
      agentCounts[l.assignedTo]++;
    }
  });

  // ==========================================
  // 🧠 BATCH CLUSTERER
  // ==========================================
  const batches = [];
  let currentBatch = null;
  let legacyCount = 0;

  const sortedForBatches = [...unassigned].sort((a, b) => {
    return (
      (b.createdAt ? new Date(b.createdAt).getTime() : 0) -
      (a.createdAt ? new Date(a.createdAt).getTime() : 0)
    );
  });

  sortedForBatches.forEach((l) => {
    if (!l.createdAt) {
      l._batchId = "legacy";
      legacyCount++;
      return;
    }
    const lTime = new Date(l.createdAt).getTime();
    if (!currentBatch) {
      currentBatch = { time: lTime, count: 1 };
      batches.push(currentBatch);
    } else {
      const diffMinutes = Math.abs(currentBatch.time - lTime) / 60000;
      if (diffMinutes <= 30) {
        currentBatch.count++;
      } else {
        currentBatch = { time: lTime, count: 1 };
        batches.push(currentBatch);
      }
    }
    l._batchId = currentBatch.time.toString();
  });

  const batchOptionsHTML = batches
    .map((b) => {
      const d = new Date(b.time);
      return `<option value="${b.time}">${d.getMonth() + 1}/${d.getDate()} @ ${d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })} (${b.count} leads)</option>`;
    })
    .join("");

  const legacyHTML =
    legacyCount > 0
      ? `<option value="legacy">Legacy / Unknown Date (${legacyCount} leads)</option>`
      : "";

  // ==========================================
  // 🎨 THE NEW LAYOUT RESTRUCTURE
  // ==========================================
  document.getElementById("main-content").innerHTML = `
    <div class="view-header" style="margin-bottom: 16px;">
      <div>
        <h1 class="view-title">Assign Leads</h1>
        <span class="view-subtitle">// ${unassigned.length} total unassigned</span>
      </div>
      <div style="display:flex;gap:8px">
        <button class="btn-cyan" onclick="navigate('drip')">Drip Feed Mode</button>
        <button class="btn-primary" onclick="autoAssignLeads()">Auto-Assign Evenly</button>
      </div>
    </div>

    <div style="display:flex; gap:10px; align-items:center; flex-wrap:wrap; margin-bottom: 24px;">
      <select id="bulk-batch-select" class="form-input" style="width: 160px;">
        <option value="all">Any Batch</option>
        ${batchOptionsHTML}
        ${legacyHTML}
      </select>

      <select id="bulk-type-select" class="form-input" style="width: 120px;"><option value="all">Any Type</option></select>
      <select id="bulk-state-select" class="form-input" style="width: 120px;"><option value="all">Any State</option></select>
      
      <select id="bulk-sort-select" class="form-input" style="width: 150px;">
        <option value="newest">Sort: Newest First</option>
        <option value="least_worked">Sort: Least Worked</option>
        <option value="most_worked">Sort: Most Worked</option>
      </select>

      <label style="display:flex; align-items:center; gap:6px; font-size:13px; color:#0D1B3E; cursor:pointer; margin-left:8px;">
        <input type="checkbox" id="bulk-unworked-check" style="cursor:pointer; width:15px; height:15px;"> Unworked Only
      </label>

      <button id="assign-reset-filters" class="btn-ghost" style="padding: 6px; margin-left: auto;" title="Reset Filters">
        <svg width="18" height="18" fill="none" viewBox="0 0 24 24" stroke="#64748B" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" /></svg>
      </button>
    </div>

    <div style="margin-bottom: 28px; padding: 0 4px;">
      <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
        <h2 style="font-size: 13px; font-weight: 700; color: #64748B; text-transform: uppercase; letter-spacing: 0.5px; margin: 0;">Pipeline Insights</h2>
        <select id="insights-selector" class="form-input" style="width: 220px; padding: 6px 12px; height: auto; font-size: 13px;"></select>
      </div>
      <div style="height: 220px; width: 100%;">
        <canvas id="insights-canvas"></canvas>
      </div>
    </div>

    <div class="card" style="margin-bottom:24px;border-color:#2563B0">
      <div class="card-header" style="background:#EEF4FB; padding: 12px 20px;">
        <h2 class="card-title" style="color:#0D1B3E; font-size: 14px;">Bulk Assign Filtered Leads</h2>
      </div>
      <div style="padding:16px 20px; display:flex; gap:16px; align-items:center; flex-wrap:wrap;">
        
        <select id="bulk-agent-select" class="form-input" style="min-width: 220px;">
          <option value="">Select Agent...</option>
          ${contractors.map((c) => `<option value="${escHtml(c.name)}">${escHtml(c.name)} (${agentCounts[c.name]} assigned)</option>`).join("")}
        </select>

        <span id="agent-readiness-badge" style="display:none; align-items:center; padding: 4px 10px; border-radius: 6px; font-size: 13px; font-weight: 700; transition: all 0.2s ease; white-space: nowrap;"></span>

        <div style="display:flex; align-items:center; gap:12px; margin-left: auto;">
          <input type="number" id="bulk-agent-qty" class="form-input" min="1" max="${unassigned.length}" placeholder="Qty" style="width: 90px;">
          <span id="bulk-type-count" style="font-size: 13px; color: #6B85B0; white-space: nowrap;">of ${unassigned.length} available</span>
          <button class="btn-primary" onclick="bulkAssignToSelectedAgent()">Assign Leads</button>
        </div>
      </div>
    </div>

    <div class="card">
      <div class="card-header" style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:12px;">
        <div>
          <h2 class="card-title">Unassigned Leads Preview</h2>
          <span class="card-meta" id="table-meta-count">Loading...</span>
        </div>
        <div style="display:flex; gap:8px; align-items:center;">
          <button id="btn-prev-page" class="btn-secondary" style="padding: 4px 10px;">&larr; Prev</button>
          <span id="page-indicator" style="font-family:var(--font-mono); font-size: 13px; font-weight: 600; color: #0D1B3E;">Page 1</span>
          <button id="btn-next-page" class="btn-secondary" style="padding: 4px 10px;">Next &rarr;</button>
        </div>
      </div>
      
      <div class="table-wrap">
        <table class="data-table">
          <thead><tr><th>Name</th><th>Type</th><th>BTN</th><th>Status</th><th style="text-align: right;">Assign To</th></tr></thead>
          <tbody id="assign-tbody"></tbody>
        </table>
      </div>
    </div>
  `;

  // 3. Internal Pointers
  let currentPage = 1;
  const itemsPerPage = 25;
  const unworkedCheck = document.getElementById("bulk-unworked-check");
  const typeSelect = document.getElementById("bulk-type-select");
  const stateSelect = document.getElementById("bulk-state-select");
  const batchSelect = document.getElementById("bulk-batch-select");
  const sortSelect = document.getElementById("bulk-sort-select");
  const qtyInput = document.getElementById("bulk-agent-qty");
  const countDisplay = document.getElementById("bulk-type-count");
  const resetFiltersBtn = document.getElementById("assign-reset-filters");

  // ==========================================
  // 🧠 CASCADING DROPDOWNS
  // ==========================================
  function updateDynamicDropdowns() {
    const selectedBatch = batchSelect ? batchSelect.value : "all";
    const currentType = typeSelect ? typeSelect.value : "all";
    const currentState = stateSelect ? stateSelect.value : "all";

    const batchLeads = unassigned.filter(
      (l) => selectedBatch === "all" || l._batchId === selectedBatch,
    );

    const availableTypes = [
      ...new Set(
        batchLeads.map((l) => (l.leadType || "").trim()).filter(Boolean),
      ),
    ].sort();
    const availableStates = [
      ...new Set(
        batchLeads
          .map((l) => (l.state || "").trim().toUpperCase())
          .filter(Boolean),
      ),
    ].sort();

    if (typeSelect) {
      typeSelect.innerHTML =
        `<option value="all">Any Type</option>` +
        availableTypes
          .map((t) => `<option value="${escHtml(t)}">${escHtml(t)}</option>`)
          .join("");
      typeSelect.value = availableTypes.includes(currentType)
        ? currentType
        : "all";
    }

    if (stateSelect) {
      stateSelect.innerHTML =
        `<option value="all">Any State</option>` +
        availableStates
          .map((s) => `<option value="${escHtml(s)}">${escHtml(s)}</option>`)
          .join("");
      stateSelect.value = availableStates.includes(currentState)
        ? currentState
        : "all";
    }
  }

  // ==========================================
  // 🧠 TABLE RENDERING & MATH
  // ==========================================
  function updateTableAndMath() {
    const selectedType = typeSelect ? typeSelect.value : "all";
    const selectedState = stateSelect ? stateSelect.value : "all";
    const selectedBatch = batchSelect ? batchSelect.value : "all";
    const selectedSort = sortSelect ? sortSelect.value : "newest";
    const requireUnworked = unworkedCheck ? unworkedCheck.checked : false;

    const filteredLeads = unassigned.filter(function (l) {
      return (
        (selectedType === "all" ||
          (l.leadType &&
            l.leadType.toLowerCase() === selectedType.toLowerCase())) &&
        (selectedState === "all" ||
          (l.state && l.state.toUpperCase() === selectedState.toUpperCase())) &&
        (selectedBatch === "all" || l._batchId === selectedBatch) &&
        (!requireUnworked || (!l.previousAgents && !l.currentMRC))
      );
    });

    filteredLeads.sort((a, b) => {
      const countA = a.previousAgents
        ? a.previousAgents.split(",").filter((x) => x.trim()).length
        : 0;
      const countB = b.previousAgents
        ? b.previousAgents.split(",").filter((x) => x.trim()).length
        : 0;

      if (selectedSort === "least_worked")
        return (
          countA - countB ||
          new Date(b.createdAt || 0) - new Date(a.createdAt || 0)
        );
      if (selectedSort === "most_worked")
        return (
          countB - countA ||
          new Date(b.createdAt || 0) - new Date(a.createdAt || 0)
        );
      return new Date(b.createdAt || 0) - new Date(a.createdAt || 0);
    });

    const total = filteredLeads.length;
    if (countDisplay) countDisplay.textContent = `of ${total} available`;
    if (qtyInput) {
      qtyInput.max = total;
      if (parseInt(qtyInput.value, 10) > total) qtyInput.value = total;
    }

    const totalPages = Math.max(1, Math.ceil(total / itemsPerPage));
    if (currentPage > totalPages) currentPage = totalPages;
    if (currentPage < 1) currentPage = 1;

    document.getElementById("page-indicator").textContent =
      `Page ${currentPage} of ${totalPages}`;
    document.getElementById("table-meta-count").textContent =
      `Showing ${total} leads matching filters`;

    document.getElementById("btn-prev-page").disabled = currentPage === 1;
    document.getElementById("btn-prev-page").style.opacity =
      currentPage === 1 ? "0.4" : "1";
    document.getElementById("btn-next-page").disabled =
      currentPage === totalPages;
    document.getElementById("btn-next-page").style.opacity =
      currentPage === totalPages ? "0.4" : "1";

    const startIndex = (currentPage - 1) * itemsPerPage;
    const displayLeads = filteredLeads.slice(
      startIndex,
      startIndex + itemsPerPage,
    );

    const tbody = document.getElementById("assign-tbody");
    if (displayLeads.length === 0) {
      tbody.innerHTML = `<tr><td colspan="5" class="empty-state">No unassigned leads match these filters!</td></tr>`;
    } else {
      tbody.innerHTML = displayLeads
        .map(function (lead) {
          const prevArray = lead.previousAgents
            ? lead.previousAgents.split(",").filter((a) => a.trim() !== "")
            : [];
          const prevBadge =
            prevArray.length > 0
              ? `<span title="${escHtml(lead.previousAgents)}" style="font-size:10px; background:#f1f5f9; color:#64748b; padding:2px 6px; border-radius:4px; margin-left:8px; font-weight:600; cursor:help;">↺ ${prevArray.length} prev agents</span>`
              : "";

          return `
        <tr>
          <td>
            <div style="display:flex; align-items:center;">
              <span class="lead-name">${escHtml(lead.name)}</span>${prevBadge}
            </div>
          </td>
          <td>${lead.leadType ? `<span class="lead-type-badge lead-type-${(lead.leadType || "").toLowerCase()}">${escHtml(lead.leadType)}</span>` : "—"}</td>
          <td class="td-mono">${escHtml(lead.BTN || lead.btn || lead.phone || "—")}</td>
          <td><span class="status-badge status-${lead.status
            .toLowerCase()
            .replace(/\s+/g, "-")
            .replace(/[^a-z0-9-]/g, "")}">${lead.status}</span></td>
          <td>
            <div class="assign-select-row" style="display:flex; gap:6px; align-items:center; justify-content: flex-end;">
              <select class="filter-select assign-select" id="assign-${lead.id}">
                <option value="">Select agent</option>
                ${contractors.map((c) => `<option value="${escHtml(c.name)}">${escHtml(c.name)} (${agentCounts[c.name]} assigned)</option>`).join("")}
              </select>
              <button class="btn-primary" style="padding:6px 14px;font-size:12px" onclick="assignLead('${lead.id}')">Assign</button>
            </div>
          </td>
        </tr>`;
        })
        .join("");
    }

    PipelineInsights.updateLive(
      "insights-selector",
      "insights-canvas",
      filteredLeads,
    );
  }

  PipelineInsights.init(
    "insights-selector",
    "insights-canvas",
    unassigned,
    false,
  );

  // 6. Attach Event Listeners
  const triggers = [unworkedCheck, typeSelect, stateSelect, sortSelect];
  triggers.forEach(
    (el) =>
      el &&
      el.addEventListener("change", () => {
        currentPage = 1;
        updateTableAndMath();
      }),
  );

  // 🚀 AGENT READINESS LISTENER
  const bulkAgentSelect = document.getElementById("bulk-agent-select");
  const readinessBadge = document.getElementById("agent-readiness-badge");

  if (bulkAgentSelect && readinessBadge) {
    bulkAgentSelect.addEventListener("change", (e) => {
      const agentName = e.target.value;

      // If they unselect the agent, hide the badge
      if (!agentName) {
        readinessBadge.style.display = "none";
        return;
      }

      // 1. Grab all leads assigned to this specific agent
      const agentLeads = State.leads.filter((l) => l.assignedTo === agentName);

      // 2. Filter out the terminal statuses
      const nonTerminalLeads = agentLeads.filter(
        (l) => !Config.terminalStatuses.includes(l.status || "New"),
      );

      // 3. Filter out cool-off leads using the graph.js logic
      const actionableLeads = nonTerminalLeads.filter(
        (l) => !Graph.isInCoolOff(l),
      );

      const actionableCount = actionableLeads.length;

      // 4. Update the UI
      readinessBadge.style.display = "inline-flex";

      if (actionableCount === 0) {
        readinessBadge.textContent = "🟢 Ready for Leads";
        readinessBadge.style.backgroundColor = "rgba(16, 185, 129, 0.15)"; // Soft Green
        readinessBadge.style.color = "#059669"; // Emerald Text
      } else {
        readinessBadge.textContent = `🔴 Not Ready (${actionableCount} Active)`;
        readinessBadge.style.backgroundColor = "rgba(244, 63, 94, 0.15)"; // Soft Red
        readinessBadge.style.color = "#E11D48"; // Rose Text
      }
    });
  }

  if (batchSelect) {
    batchSelect.addEventListener("change", () => {
      updateDynamicDropdowns();
      currentPage = 1;
      updateTableAndMath();
    });
  }

  if (resetFiltersBtn) {
    resetFiltersBtn.addEventListener("click", () => {
      if (batchSelect) batchSelect.value = "all";
      if (typeSelect) typeSelect.value = "all";
      if (stateSelect) stateSelect.value = "all";
      if (sortSelect) sortSelect.value = "newest";
      if (unworkedCheck) unworkedCheck.checked = false;

      currentPage = 1;
      updateDynamicDropdowns();
      updateTableAndMath();
    });
  }

  document.getElementById("btn-prev-page").addEventListener("click", () => {
    if (currentPage > 1) {
      currentPage--;
      updateTableAndMath();
    }
  });
  document.getElementById("btn-next-page").addEventListener("click", () => {
    currentPage++;
    updateTableAndMath();
  });

  // 7. Initial Draw
  updateDynamicDropdowns();
  updateTableAndMath();
}

async function assignLead(leadId) {
  const select = document.getElementById("assign-" + leadId);
  const agent = select && select.value;
  if (!agent) {
    UI.showToast("Please select an agent.", "error");
    return;
  }
  if (!Graph.canAgentTakeLead(agent, State.leads)) {
    UI.showToast(agent + " is at the lead limit.", "error");
    return;
  }
  const lead = State.leads.find(function (l) {
    return l.id === leadId;
  });
  setLoading(true);
  try {
    await Graph.assignAgent(leadId, agent);
    await Graph.logActivity({
      LeadID: leadId,
      Title: lead ? lead.name : "",
      ActionType: "Assigned",
      AgentEmail: (State.currentUser && State.currentUser.email) || "",
      Notes:
        "Assigned by " +
        ((State.currentUser && State.currentUser.name) || "Admin"),
    });
    UI.showToast("Assigned to " + agent, "success");
    await loadAllData();
    renderAssignLeads();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

async function bulkAssignToSelectedAgent() {
  // Grab all the dropdown values (including the new Batch and Sort!)
  const agentSelect = document.getElementById("bulk-agent-select");
  const agentName = agentSelect ? agentSelect.value : "";

  const qtyInput = document.getElementById("bulk-agent-qty");
  const qty = qtyInput ? parseInt(qtyInput.value, 10) : 0;

  const typeSelect = document.getElementById("bulk-type-select");
  const selectedType = typeSelect ? typeSelect.value : "all";

  const stateSelect = document.getElementById("bulk-state-select");
  const selectedState = stateSelect ? stateSelect.value : "all";

  const batchSelect = document.getElementById("bulk-batch-select");
  const selectedBatch = batchSelect ? batchSelect.value : "all";

  const sortSelect = document.getElementById("bulk-sort-select");
  const selectedSort = sortSelect ? sortSelect.value : "newest";

  const unworkedCheck = document.getElementById("bulk-unworked-check");
  const requireUnworked = unworkedCheck ? unworkedCheck.checked : false;

  if (!agentName) {
    UI.showToast("Please select an agent first.", "warning");
    return;
  }
  if (!qty || qty <= 0) {
    UI.showToast("Please enter a valid number of leads.", "warning");
    return;
  }

  // 1. Get the base unassigned pool
  const unassigned = State.leads.filter(function (l) {
    const isValidLead = l && l.id && (l.name || l.phone || l.BTN || l.btn);
    const isAvailable =
      !l.assignedTo && !Config.terminalStatuses.includes(l.status);
    return isValidLead && isAvailable;
  });

  // 2. Filter using the EXACT same rules as the table preview
  const validLeads = unassigned.filter(function (l) {
    const typeMatch =
      selectedType === "all" ||
      (l.leadType && l.leadType.toLowerCase() === selectedType.toLowerCase());
    const stateMatch =
      selectedState === "all" ||
      (l.state && l.state.toUpperCase() === selectedState.toUpperCase());

    // Check our new dynamic batch tags
    const batchMatch = selectedBatch === "all" || l._batchId === selectedBatch;

    // Check if it's completely untouched
    const unworkedMatch =
      !requireUnworked || (!l.previousAgents && !l.currentMRC);

    // 🛑 CRITICAL BOUNCER: Make sure the selected agent hasn't worked this lead before!
    const prevAgents = (l.previousAgents || "").toLowerCase();
    const agentMatch = !prevAgents.includes(agentName.toLowerCase());

    return typeMatch && stateMatch && batchMatch && unworkedMatch && agentMatch;
  });

  // 3. SORT USING THE EXACT SAME RULES AS THE TABLE PREVIEW 🚀
  validLeads.sort((a, b) => {
    const countA = a.previousAgents
      ? a.previousAgents.split(",").filter((x) => x.trim()).length
      : 0;
    const countB = b.previousAgents
      ? b.previousAgents.split(",").filter((x) => x.trim()).length
      : 0;

    if (selectedSort === "least_worked") {
      return (
        countA - countB ||
        new Date(b.createdAt || 0) - new Date(a.createdAt || 0)
      );
    } else if (selectedSort === "most_worked") {
      return (
        countB - countA ||
        new Date(b.createdAt || 0) - new Date(a.createdAt || 0)
      );
    } else {
      return new Date(b.createdAt || 0) - new Date(a.createdAt || 0);
    }
  });

  // Dynamic labels for the Toast notification
  const stateLabel = selectedState === "all" ? "" : `${selectedState} `;
  const typeLabel = selectedType === "all" ? "leads" : `${selectedType} leads`;
  const combinedLabel = `${stateLabel}${typeLabel}`.trim();

  // 4. Validation Checks
  if (validLeads.length === 0) {
    UI.showToast(
      `${agentName} has no available ${combinedLabel} left to work with these filters!`,
      "warning",
    );
    return;
  }

  if (qty > validLeads.length) {
    UI.showToast(
      `Only ${validLeads.length} ${combinedLabel} available for ${agentName} (already worked the rest).`,
      "warning",
    );
    return;
  }

  // 5. Slice from the newly sorted top!
  const leadsToAssign = validLeads.slice(0, qty);

  setLoading(true);
  try {
    await Promise.all(
      leadsToAssign.map(async (lead) => {
        await Graph.updateLead(lead.id, { Agent_x0020_Assigned: agentName });
        lead.assignedTo = agentName;
      }),
    );

    UI.showToast(
      `Successfully assigned ${qty} ${combinedLabel} to ${agentName}!`,
      "success",
    );
    renderAssignLeads();
  } catch (err) {
    console.error("Bulk Assign Error:", err);
    UI.showToast("Failed to assign leads: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

async function bulkAssignByQuantity() {
  const { leads, contractors } = State;
  const unassigned = leads.filter(function (l) {
    return !l.assignedTo && !Config.terminalStatuses.includes(l.status);
  });

  const plan = [];
  contractors.forEach(function (c) {
    const qty =
      parseInt((document.getElementById("qty-" + c.name) || {}).value || "0") ||
      0;
    if (qty > 0) plan.push({ agent: c.name, qty: qty });
  });

  if (!plan.length) {
    UI.showToast("Please enter a quantity for at least one agent.", "error");
    return;
  }

  const totalRequested = plan.reduce(function (s, p) {
    return s + p.qty;
  }, 0);
  if (totalRequested > unassigned.length) {
    UI.showToast(
      "Total (" +
        totalRequested +
        ") exceeds unassigned leads (" +
        unassigned.length +
        "). Reduce quantities.",
      "error",
    );
    return;
  }

  const summary = plan
    .map(function (p) {
      return p.qty + " → " + p.agent;
    })
    .join("\n");
  if (
    !confirm(
      "Assign leads by quantity?\n\n" +
        summary +
        "\n\nTotal: " +
        totalRequested +
        " leads",
    )
  )
    return;

  setLoading(true);
  try {
    let idx = 0;
    for (var p = 0; p < plan.length; p++) {
      for (var q = 0; q < plan[p].qty; q++) {
        if (idx >= unassigned.length) break;
        await Graph.assignAgent(unassigned[idx].id, plan[p].agent);
        idx++;
      }
    }
    UI.showToast(
      "Assigned " + totalRequested + " leads successfully!",
      "success",
    );
    await loadAllData();
    renderAssignLeads();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

// ============================================================
// CALLBACKS
// ============================================================
function renderCallBacks() {
  const mainContent = document.getElementById("main-content");
  mainContent.innerHTML = ""; // Clear existing screen

  // 1. Setup Template
  const template = document.getElementById("tmpl-callbacks-page");
  const clone = template.content.cloneNode(true);
  const wrap = clone.getElementById("callbacks-list-wrap");

  // 2. Identify the current agent
  const userName = ((State.currentUser && State.currentUser.name) || "")
    .toLowerCase()
    .trim();
  const userEmail = ((State.currentUser && State.currentUser.email) || "")
    .toLowerCase()
    .trim();

  const contractor = (State.contractors || []).find((c) => {
    return (
      (c.email || "").toLowerCase().trim() === userEmail ||
      (c.name || "").toLowerCase().trim() === userName
    );
  });

  const agentName = contractor
    ? contractor.name.toLowerCase().trim()
    : userName;

  // 3. Filter the master database
  const callbacks = (State.leads || []).filter((l) => {
    const assigned = (l.assignedTo || "")
      .toLowerCase()
      .replace(/\s+/g, " ")
      .trim();

    const isAssignedToMe =
      assigned &&
      (assigned === agentName.replace(/\s+/g, " ") ||
        assigned === userName.replace(/\s+/g, " ") ||
        assigned === userEmail.replace(/\s+/g, " "));

    const isTerminal =
      l.status === "TDM" || l.status === (Config.soldStatus || "Sold");

    return isAssignedToMe && !isTerminal && l.callbackAt;
  });

  // 4. Sort chronologically
  callbacks.sort((a, b) => {
    const dateA = new Date(a.callbackAt);
    if (a.status === "Pending Order") dateA.setDate(dateA.getDate() + 1);

    const dateB = new Date(b.callbackAt);
    if (b.status === "Pending Order") dateB.setDate(dateB.getDate() + 1);

    return dateA - dateB;
  });

  // 5. Handle the Empty State
  if (callbacks.length === 0) {
    // NEW: Added the animation class to the empty state so it fades in too!
    wrap.innerHTML = `
      <div class="animate-fade-up" style="padding: 60px 20px; text-align: center; color: #64748b;">
        <div style="font-size: 40px; margin-bottom: 16px;">📅</div>
        <h3 style="margin: 0 0 8px 0; color: #0a1a3f; font-size: 18px;">Pipeline Clear</h3>
        <p style="margin: 0; font-size: 14px;">You have no upcoming callbacks or installations scheduled.</p>
      </div>
    `;
  } else {
    // 6. Build the Table
    let html = `
      <table class="animate-fade-up" style="width: 100%; border-collapse: collapse; text-align: left;">
        <thead>
          <tr style="background: #f8fafc; border-bottom: 1px solid #e2e8f0; font-size: 11px; color: #64748b; text-transform: uppercase; letter-spacing: 1px;">
            <th style="padding: 14px 16px; font-weight: 600;">Scheduled For</th>
            <th style="padding: 14px 16px; font-weight: 600;">Customer Info</th>
            <th style="padding: 14px 16px; font-weight: 600;">Status</th>
            <th style="padding: 14px 16px; font-weight: 600; text-align: right;">Action</th>
          </tr>
        </thead>
        <tbody>
    `;

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // NEW: Added the 'index' parameter to the loop so we can multiply it for the staggered delay
    callbacks.forEach((l, index) => {
      // 1. Calculate the actual Action Date
      const actionDate = new Date(l.callbackAt);
      const isInstall = l.status === "Pending Order";

      if (isInstall) {
        // Push the required action to the day AFTER the install
        actionDate.setDate(actionDate.getDate() + 1);
      }

      const actionMidnight = new Date(actionDate);
      actionMidnight.setHours(0, 0, 0, 0);

      let dateColor = "#1e293b";
      let dateWeight = "normal";
      let badge = "";

      // 2. Check the new action date against Today
      if (actionMidnight < today) {
        dateColor = "var(--red, #ef4444)";
        dateWeight = "bold";
        badge = `<span style="background: #fee2e2; color: #991b1b; padding: 2px 6px; border-radius: 4px; font-size: 10px; margin-left: 8px; font-weight: bold;">OVERDUE</span>`;
      } else if (actionMidnight.getTime() === today.getTime()) {
        dateColor = "var(--blue, #2563B0)";
        dateWeight = "bold";
        badge = `<span style="background: #e0f2fe; color: #0369a1; padding: 2px 6px; border-radius: 4px; font-size: 10px; margin-left: 8px; font-weight: bold;">TODAY</span>`;
      }

      // 3. Format the strings for the UI
      const dateStr = actionDate.toLocaleDateString([], {
        weekday: "short",
        month: "short",
        day: "numeric",
      });
      const timeStr = actionDate.toLocaleTimeString([], {
        hour: "2-digit",
        minute: "2-digit",
      });

      // ---> THE UI TOUCH SNIPPET <---
      // Give the agent clear context on the original install date
      const typeStr = isInstall
        ? `Install Check (Installed ${new Date(l.callbackAt).toLocaleDateString([], { month: "short", day: "numeric" })})`
        : "Scheduled Call";

      const statusCls = `status-${(l.status || "")
        .toLowerCase()
        .replace(/\s+/g, "-")
        .replace(/[^a-z0-9-]/g, "")}`;

      // NEW: Added the animate-fade-up class and the dynamic staggered delay using the index
      html += `
        <tr class="animate-row-fade" style="border-bottom: 1px solid #f1f5f9; transition: background 0.2s; animation-delay: ${index * 0.05 + 0.1}s;" onmouseover="this.style.background='#f8fafc'" onmouseout="this.style.background='white'">
          <td style="padding: 16px; font-size: 13px; color: ${dateColor}; font-weight: ${dateWeight};">
            <div style="display: flex; align-items: center; font-size: 14px;">${dateStr} @ ${timeStr} ${badge}</div>
            <div style="font-size: 11px; color: #64748b; margin-top: 4px; font-weight: normal; text-transform: uppercase;">${typeStr}</div>
          </td>
          <td style="padding: 16px;">
            <div style="font-size: 14px; font-weight: 600; color: #0a1a3f;">${escHtml(l.name || "Unknown")}</div>
            <div style="font-size: 12px; color: #64748b; margin-top: 2px;">📞 ${escHtml(l.cbr || "No Phone Number")}</div>
          </td>
          <td style="padding: 16px;">
            <span class="status-badge ${statusCls}">${l.status}</span>
          </td>
          <td style="padding: 16px; text-align: right; white-space: nowrap;">
            <button class="btn-secondary" style="font-size: 12px; padding: 6px 12px; margin-right: 8px;" onclick="viewCallbackLead('${l.id}')">
              View Callback
            </button>
            <button class="btn-primary" style="font-size: 12px; padding: 6px 12px;" onclick="workCallbackLead('${l.id}')">
              Work Callback
            </button>
          </td>
          </td>
        </tr>
      `;
    });

    html += `</tbody></table>`;
    wrap.innerHTML = html;
  }

  // 7. Return the fully built DOM element for your router to mount!
  mainContent.appendChild(clone);
}

function viewCallbackLead(leadId) {
  const targetLead = State.leads.find((l) => l.id === leadId);
  if (!targetLead) return;

  const currentQueue = window._myLeads || [];
  const filteredQueue = currentQueue.filter((l) => l.id !== leadId);
  window._myLeads = [targetLead, ...filteredQueue];

  window._forceShowLead = true;
  window._currentFeedIndex = 0;

  // NEW: Make sure the app knows we are just viewing
  window._isWorkingCallback = false;

  navigate("myleads");
}

function workCallbackLead(leadId) {
  const targetLead = State.leads.find((l) => l.id === leadId);
  if (!targetLead) return;
  const currentQueue = window._myLeads || [];
  const filteredQueue = currentQueue.filter((l) => l.id !== leadId);
  window._myLeads = [targetLead, ...filteredQueue];

  window._forceShowLead = true;
  window._currentFeedIndex = 0;

  // NEW: Tell the app to intercept the "Next Lead" button
  window._isWorkingCallback = true;

  navigate("myleads");
}

function completeCallbackLead(leadId) {
  const targetLead = State.leads.find((l) => l.id === leadId);

  if (targetLead) {
    // 1. NOW we wipe the date since the interaction is over
    targetLead.callbackAt = null;

    // 2. Secretly update SharePoint in the background
    if (window.Graph && Graph.updateLead) {
      Graph.updateLead(leadId, { CallbackDateTime: null }).catch((err) =>
        console.error("Failed to clear callback", err),
      );
    }
  }

  // 3. Reset all our bypass flags so the Bouncer wakes back up
  window._isWorkingCallback = false;
  window._forceShowLead = false;

  // 4. Send them back to their pipeline
  navigate("callbacks");
}
// ============================================================
//  LEADS VIEW (Admin only)
// ============================================================
function renderLeads() {
  if (!isAdmin()) {
    navigate("myleads");
    return;
  }

  // 1. Security & Data Prep
  State.selectedLeads.clear();
  const contractors = State.contractors.map((c) => c.name);

  // ==========================================
  // 🧠 THE BATCH CLUSTERER
  // ==========================================
  const batches = [];
  let currentBatch = null;
  let legacyCount = 0;

  const sortedForBatches = [...State.leads].sort((a, b) => {
    const timeA = a.createdAt ? new Date(a.createdAt).getTime() : 0;
    const timeB = b.createdAt ? new Date(b.createdAt).getTime() : 0;
    return timeB - timeA;
  });

  sortedForBatches.forEach((l) => {
    // 🚀 CACHE THE HEAVY MATH ONCE ON LOAD
    l._time = l.createdAt ? new Date(l.createdAt).getTime() : 0;
    l._prevCount = l.previousAgents
      ? l.previousAgents.split(",").filter((x) => x.trim()).length
      : 0;

    if (!l.createdAt) {
      l._batchId = "legacy";
      legacyCount++;
      return;
    }

    const lTime = new Date(l.createdAt).getTime();

    if (!currentBatch) {
      currentBatch = { time: lTime, count: 1 };
      batches.push(currentBatch);
    } else {
      const diffMinutes = Math.abs(currentBatch.time - lTime) / 60000;
      if (diffMinutes <= 30) {
        currentBatch.count++;
      } else {
        currentBatch = { time: lTime, count: 1 };
        batches.push(currentBatch);
      }
    }
    l._batchId = currentBatch.time.toString();
  });

  const uniqueStates = [
    ...new Set(
      State.leads
        .map((l) => (l.state || "").trim().toUpperCase())
        .filter(Boolean),
    ),
  ].sort();

  // 2. Setup Template
  const mainContent = document.getElementById("main-content");
  mainContent.innerHTML = "";

  const template = document.getElementById("tmpl-all-leads");
  const clone = template.content.cloneNode(true);

  // 3. Header
  clone.getElementById("leads-subtitle").textContent =
    `// ${State.leads.length} total`;

  // ==========================================
  // 🚀 4. CSV IMPORTER WIRING
  // ==========================================
  const importBtn = clone.getElementById("importLeadsBtn");
  const fileInput = clone.getElementById("leadFileInput");

  if (importBtn && fileInput) {
    importBtn.onclick = () => {
      const overlay = document.createElement("div");
      overlay.style.cssText =
        "position:fixed; top:0; left:0; width:100vw; height:100vh; background:rgba(13, 27, 62, 0.6); z-index:9999; display:flex; align-items:center; justify-content:center; backdrop-filter: blur(3px);";

      const modal = document.createElement("div");
      modal.style.cssText =
        "background:#fff; padding:24px; border-radius:12px; width:320px; box-shadow:0 10px 25px rgba(0,0,0,0.2); display:flex; flex-direction:column; gap:16px;";

      modal.innerHTML = `
        <div>
          <h3 style="margin:0 0 4px 0; font-size:18px; color:#0D1B3E;">Upload Leads</h3>
          <p style="margin:0; font-size:13px; color:#666;">What type of leads are in this file?</p>
        </div>
        <select id="tempLeadType" class="filter-select" style="width:100%; padding:10px; border-radius:6px;">
          <option value="OFS">OFS Leads</option>
          <option value="MLR">MLR Leads</option>
          <option value="Forced">Forced Leads</option>
        </select>
        <div style="display:flex; justify-content:flex-end; gap:10px; margin-top:8px;">
          <button id="cancelTypeBtn" class="btn-ghost" style="padding:8px 16px;">Cancel</button>
          <button id="confirmTypeBtn" class="btn-primary" style="padding:8px 16px;">Choose File</button>
        </div>
      `;

      overlay.appendChild(modal);
      document.body.appendChild(overlay);

      document.getElementById("cancelTypeBtn").onclick = () => overlay.remove();
      document.getElementById("confirmTypeBtn").onclick = () => {
        const selectedType = document.getElementById("tempLeadType").value;
        fileInput.dataset.leadType = selectedType;
        overlay.remove();
        fileInput.click();
      };
    };

    fileInput.addEventListener("change", handleFileSelect, false);
  }

  // ==========================================
  // 5. Populate Bulk Bar Dropdowns
  // ==========================================
  const bulkAssignSelect = clone.getElementById("bulk-assign-select");

  const unassignOption = document.createElement("option");
  unassignOption.value = "";
  unassignOption.textContent = "-- Unassign Leads --";
  unassignOption.style.fontStyle = "italic";
  unassignOption.style.color = "#64748b";
  bulkAssignSelect.appendChild(unassignOption);

  contractors.forEach((c) => {
    const option = document.createElement("option");
    option.value = c;
    option.textContent = c;
    bulkAssignSelect.appendChild(option);
  });

  const bulkTypeSelect = clone.getElementById("bulk-type-select");
  Config.leadTypes.forEach((t) => {
    const option = document.createElement("option");
    option.value = t;
    option.textContent = t;
    bulkTypeSelect.appendChild(option);
  });

  // ==========================================
  // 6. Populate Filters & Dynamic Injection
  // ==========================================
  const searchInput = clone.getElementById("opt-search-input");
  if (searchInput) searchInput.value = State.filters.search || "";

  const statusFilter = clone.getElementById("opt-filter-status");
  Config.leadStatuses.forEach((s) => {
    const option = document.createElement("option");
    option.value = s;
    option.textContent = s;
    if (State.filters.status === s) option.selected = true;
    statusFilter.appendChild(option);
  });

  const agentFilter = clone.getElementById("opt-filter-agent");
  contractors.forEach((c) => {
    const option = document.createElement("option");
    option.value = c;
    option.textContent = c;
    if (State.filters.assignedTo === c) option.selected = true;
    agentFilter.appendChild(option);
  });

  if (agentFilter && agentFilter.parentNode) {
    const batchOptionsHTML = batches
      .map((b) => {
        const d = new Date(b.time);
        const dateStr = `${d.getMonth() + 1}/${d.getDate()}`;
        const timeStr = d.toLocaleTimeString([], {
          hour: "2-digit",
          minute: "2-digit",
        });
        return `<option value="${b.time}">${dateStr} @ ${timeStr} (${b.count})</option>`;
      })
      .join("");

    const legacyHTML =
      legacyCount > 0
        ? `<option value="legacy">Legacy (${legacyCount})</option>`
        : "";
    const stateOptionsHTML = uniqueStates
      .map((s) => `<option value="${s}">${s}</option>`)
      .join("");
    const typeOptionsHTML = (Config.leadTypes || [])
      .map((t) => `<option value="${t}">${t}</option>`)
      .join("");

    agentFilter.parentNode.insertAdjacentHTML(
      "beforeend",
      `
      <select id="filter-batch" class="filter-select">
        <option value="all">Any Batch</option>
        ${batchOptionsHTML}
        ${legacyHTML}
      </select>
      <select id="filter-type" class="filter-select">
        <option value="all">Any Type</option>
        ${typeOptionsHTML}
      </select>
      <select id="filter-state" class="filter-select">
        <option value="all">Any State</option>
        ${stateOptionsHTML}
      </select>
      <select id="filter-sort" class="filter-select">
        <option value="newest">Sort: Newest First</option>
        <option value="least_worked">Sort: Least Worked</option>
        <option value="most_worked">Sort: Most Worked</option>
      </select>
      
      <button id="leads-reset-filters" class="btn-ghost" style="padding: 6px; margin-left: 4px;" title="Reset Filters">
        <svg width="18" height="18" fill="none" viewBox="0 0 24 24" stroke="#64748B" stroke-width="2"><path stroke-linecap="round" stroke-linejoin="round" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" /></svg>
      </button>
    `,
    );
  }

  const batchFilter = clone.querySelector("#filter-batch");
  const stateFilter = clone.querySelector("#filter-state");
  const typeFilter = clone.querySelector("#filter-type");
  const sortFilter = clone.querySelector("#filter-sort");
  const resetFiltersBtn = clone.querySelector("#leads-reset-filters"); // 🚀 Reset Pointer

  if (batchFilter && State.filters.batch)
    batchFilter.value = State.filters.batch;
  if (stateFilter && State.filters.state)
    stateFilter.value = State.filters.state;
  if (typeFilter && State.filters.type) typeFilter.value = State.filters.type;
  if (sortFilter && State.filters.sort) sortFilter.value = State.filters.sort;

  // ==========================================
  // 📊 PIPELINE INSIGHTS INJECTION
  // ==========================================
  const tableWrap = clone.getElementById("opt-leads-table-wrap");

  if (tableWrap) {
    const insightsContainer = document.createElement("div");
    insightsContainer.style.cssText = "margin-bottom: 28px; padding: 0 4px;";
    insightsContainer.innerHTML = `
      <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
        <h2 style="font-size: 13px; font-weight: 700; color: #64748B; text-transform: uppercase; letter-spacing: 0.5px; margin: 0;">Pipeline Insights</h2>
        <select id="insights-selector-main" class="form-input" style="width: 220px; padding: 6px 12px; height: auto; font-size: 13px;"></select>
      </div>
      <div style="height: 220px; width: 100%;">
        <canvas id="insights-canvas-main"></canvas>
      </div>
    `;
    tableWrap.parentNode.insertBefore(insightsContainer, tableWrap);
  }

  // ==========================================
  // 🧠 CASCADING DROPDOWNS
  // ==========================================
  function updateDynamicDropdowns() {
    const selectedBatch = batchFilter ? batchFilter.value : "all";
    const currentState = stateFilter ? stateFilter.value : "all";

    const batchLeads = State.leads.filter(
      (l) => selectedBatch === "all" || l._batchId === selectedBatch,
    );
    const availableStates = [
      ...new Set(
        batchLeads
          .map((l) => (l.state || "").trim().toUpperCase())
          .filter(Boolean),
      ),
    ].sort();

    if (stateFilter) {
      stateFilter.innerHTML =
        `<option value="all">Any State</option>` +
        availableStates
          .map((s) => `<option value="${escHtml(s)}">${escHtml(s)}</option>`)
          .join("");
      stateFilter.value = availableStates.includes(currentState)
        ? currentState
        : "all";
    }
  }

  // ==========================================
  // 7. THE SMART PAGINATION RENDERER
  // ==========================================
  let currentPage = 1;
  const itemsPerPage = 50;

  const paginationWrap = document.createElement("div");
  paginationWrap.style.cssText =
    "display:flex; gap:6px; align-items:center; justify-content:flex-end; padding: 12px 0; margin-top: 8px;";
  paginationWrap.innerHTML = `
    <button id="leads-prev-page" class="btn-secondary" style="padding: 6px 12px; font-size:13px;">&larr; Prev</button>
    <span id="leads-page-indicator" style="font-family:var(--font-mono); font-size:13px; font-weight:600; color:#0D1B3E; min-width: 80px; text-align: center;">Pg 1</span>
    <button id="leads-next-page" class="btn-secondary" style="padding: 6px 12px; font-size:13px;">Next &rarr;</button>
  `;

  if (tableWrap) {
    tableWrap.parentNode.insertBefore(paginationWrap, tableWrap.nextSibling);
  }

  const prevBtn = paginationWrap.querySelector("#leads-prev-page");
  const nextBtn = paginationWrap.querySelector("#leads-next-page");
  const pageIndicator = paginationWrap.querySelector("#leads-page-indicator");

  function updateTable() {
    let filtered =
      typeof getFilteredLeads === "function" ? getFilteredLeads() : State.leads;

    const selectedBatch = batchFilter ? batchFilter.value : "all";
    const selectedState = stateFilter ? stateFilter.value : "all";
    const selectedType = typeFilter ? typeFilter.value : "all";
    const selectedSort = sortFilter ? sortFilter.value : "newest";

    filtered = filtered.filter((l) => {
      const batchMatch =
        selectedBatch === "all" || l._batchId === selectedBatch;
      const stateMatch =
        selectedState === "all" ||
        (l.state && l.state.toUpperCase() === selectedState.toUpperCase());
      const typeMatch =
        selectedType === "all" ||
        (l.leadType && l.leadType.toLowerCase() === selectedType.toLowerCase());

      return batchMatch && stateMatch && typeMatch;
    });
    /* OLD UNPERFORMANT SORT
    filtered.sort((a, b) => {
      const countA = a.previousAgents
        ? a.previousAgents.split(",").filter((x) => x.trim()).length
        : 0;
      const countB = b.previousAgents
        ? b.previousAgents.split(",").filter((x) => x.trim()).length
        : 0;

      if (selectedSort === "least_worked") {
        return (
          countA - countB ||
          new Date(b.createdAt || 0) - new Date(a.createdAt || 0)
        );
      } else if (selectedSort === "most_worked") {
        return (
          countB - countA ||
          new Date(b.createdAt || 0) - new Date(a.createdAt || 0)
        );
      } else {
        return new Date(b.createdAt || 0) - new Date(a.createdAt || 0);
      }
    });
  */
    // 🚀 BLAZING FAST SORT (Uses the cached variables)
    filtered.sort((a, b) => {
      if (selectedSort === "least_worked") {
        return a._prevCount - b._prevCount || b._time - a._time;
      } else if (selectedSort === "most_worked") {
        return b._prevCount - a._prevCount || b._time - a._time;
      } else {
        return b._time - a._time;
      }
    });

    const total = filtered.length;
    const totalPages = Math.max(1, Math.ceil(total / itemsPerPage));

    if (currentPage > totalPages) currentPage = totalPages;
    if (currentPage < 1) currentPage = 1;

    if (pageIndicator)
      pageIndicator.textContent = `Pg ${currentPage} / ${totalPages}`;

    if (prevBtn) {
      prevBtn.disabled = currentPage === 1;
      prevBtn.style.opacity = currentPage === 1 ? "0.4" : "1";
    }
    if (nextBtn) {
      nextBtn.disabled = currentPage === totalPages;
      nextBtn.style.opacity = currentPage === totalPages ? "0.4" : "1";
    }

    const startIndex = (currentPage - 1) * itemsPerPage;
    const displayLeads = filtered.slice(startIndex, startIndex + itemsPerPage);

    if (tableWrap) {
      tableWrap.replaceChildren(renderLeadsTable(displayLeads));
    }
    // 🚀 NON-BLOCKING CHART RENDER: Paints the table instantly, calculates chart math next frame
    setTimeout(() => {
      PipelineInsights.updateLive(
        "insights-selector-main",
        "insights-canvas-main",
        filtered,
      );
    }, 0);
  }

  if (prevBtn)
    prevBtn.addEventListener("click", () => {
      if (currentPage > 1) {
        currentPage--;
        updateTable();
      }
    });
  if (nextBtn)
    nextBtn.addEventListener("click", () => {
      currentPage++;
      updateTable();
    });

  const handleFilterChange = (e) => {
    // 1. Instantly save the filter selections to state
    State.filters.search = searchInput ? searchInput.value : "";
    State.filters.status = statusFilter ? statusFilter.value : "all";
    State.filters.assignedTo = agentFilter ? agentFilter.value : "all";
    State.filters.batch = batchFilter ? batchFilter.value : "all";
    State.filters.state = stateFilter ? stateFilter.value : "all";
    State.filters.type = typeFilter ? typeFilter.value : "all";
    State.filters.sort = sortFilter ? sortFilter.value : "newest";

    currentPage = 1;

    // If the Batch filter changed, update the cascading options instantly
    if (e && e.target === batchFilter) {
      updateDynamicDropdowns();
    }

    // 🚀 THE PERFORMANCE FIX: Defer the table calculation to the next CPU cycle.
    // This allows the dropdown menu to visually close smoothly before processing data.
    requestAnimationFrame(() => {
      setTimeout(() => {
        updateTable();
      }, 0);
    });
  };

  if (searchInput) {
    let searchTimeout;
    searchInput.addEventListener("input", (e) => {
      clearTimeout(searchTimeout);
      searchTimeout = setTimeout(() => {
        handleFilterChange(e);
      }, 300); // Waits 300ms after the last keystroke before updating
    });
  }
  if (statusFilter) statusFilter.addEventListener("change", handleFilterChange);
  if (agentFilter) agentFilter.addEventListener("change", handleFilterChange);
  if (batchFilter) batchFilter.addEventListener("change", handleFilterChange);
  if (stateFilter) stateFilter.addEventListener("change", handleFilterChange);
  if (typeFilter) typeFilter.addEventListener("change", handleFilterChange);
  if (sortFilter) sortFilter.addEventListener("change", handleFilterChange);

  // 🚀 THE RESET LISTENER (Resets all inputs AND the global State object)
  if (resetFiltersBtn) {
    resetFiltersBtn.addEventListener("click", () => {
      if (searchInput) searchInput.value = "";
      if (statusFilter) statusFilter.value = "all";
      if (agentFilter) agentFilter.value = "all";
      if (batchFilter) batchFilter.value = "all";
      if (typeFilter) typeFilter.value = "all";
      if (stateFilter) stateFilter.value = "all";
      if (sortFilter) sortFilter.value = "newest";

      State.filters.search = "";
      State.filters.status = "all";
      State.filters.assignedTo = "all";
      State.filters.batch = "all";
      State.filters.type = "all";
      State.filters.state = "all";
      State.filters.sort = "newest";

      currentPage = 1;
      updateDynamicDropdowns();
      updateTable();
    });
  }

  updateDynamicDropdowns();

  mainContent.appendChild(clone);

  PipelineInsights.init(
    "insights-selector-main",
    "insights-canvas-main",
    State.leads,
    true,
  );

  updateTable();
}

function toggleLeadSelect(id, checked) {
  if (checked) {
    State.selectedLeads.add(id);
  } else {
    State.selectedLeads.delete(id);
  }
  updateBulkBar();
}

function toggleSelectAll(checked) {
  const checkboxes = document.querySelectorAll(".lead-checkbox");
  checkboxes.forEach(function (cb) {
    cb.checked = checked;
    if (checked) {
      State.selectedLeads.add(cb.dataset.id);
    } else {
      State.selectedLeads.delete(cb.dataset.id);
    }
  });
  updateBulkBar();
}

function updateBulkBar() {
  const bar = document.getElementById("bulk-bar");
  const count = document.getElementById("bulk-count");
  const n = State.selectedLeads.size;
  if (!bar) return;
  bar.style.display = n > 0 ? "flex" : "none";
  if (count)
    count.textContent = n + " lead" + (n !== 1 ? "s" : "") + " selected";
  const allCbs = document.querySelectorAll(".lead-checkbox");
  const selAll = document.getElementById("select-all-cb");
  if (selAll && allCbs.length) {
    selAll.indeterminate = n > 0 && n < allCbs.length;
    selAll.checked = n === allCbs.length;
  }
}

function clearSelection() {
  State.selectedLeads.clear();
  document.querySelectorAll(".lead-checkbox").forEach(function (cb) {
    cb.checked = false;
  });
  const selAll = document.getElementById("select-all-cb");
  if (selAll) {
    selAll.checked = false;
    selAll.indeterminate = false;
  }
  updateBulkBar();
}

async function bulkDelete() {
  const ids = Array.from(State.selectedLeads);
  if (!ids.length) return;
  if (
    !confirm(
      "Permanently delete " +
        ids.length +
        " lead" +
        (ids.length !== 1 ? "s" : "") +
        "? This cannot be undone.",
    )
  )
    return;
  setLoading(true);
  try {
    for (var i = 0; i < ids.length; i++) {
      await Graph.deleteLead(ids[i]);
    }
    UI.showToast(
      "Deleted " + ids.length + " lead" + (ids.length !== 1 ? "s" : ""),
      "success",
    );
    State.selectedLeads.clear();
    await loadAllData();
    renderLeads();
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

async function bulkAssign() {
  const ids = Array.from(State.selectedLeads);
  const agentSelect = document.getElementById("bulk-assign-select");

  // Safely grab the value. If the dropdown is missing, it falls back to undefined.
  const agent = agentSelect ? agentSelect.value : undefined;

  if (!ids.length) {
    UI.showToast("Please select at least one lead first.", "warning");
    return;
  }

  // 🛡️ THE BOUNCER: Allow the empty string ("") to pass through!
  if (agent === undefined || agent === null) {
    UI.showToast("Please select an assignment option.", "error");
    return;
  }

  const isUnassigning = agent === "";
  const actionWord = isUnassigning ? "Unassign" : "Assign";
  const targetWord = isUnassigning ? "" : ` to ${agent}`;
  const leadWord = ids.length === 1 ? "lead" : "leads";

  // Dynamic popup: "Unassign 5 leads?" vs "Assign 5 leads to Michael?"
  if (!confirm(`${actionWord} ${ids.length} ${leadWord}${targetWord}?`)) {
    return;
  }

  setLoading(true);
  try {
    // 🚀 THE UPGRADE: Promise.all processes the entire batch concurrently instead of one-by-one
    await Promise.all(
      ids.map(async (id) => {
        // Pass null to SharePoint if unassigning, otherwise pass the agent's name
        const targetAgent = isUnassigning ? null : agent;
        await Graph.assignAgent(id, targetAgent);
      }),
    );

    const successAction = isUnassigning ? "Unassigned" : `Assigned to ${agent}`;
    UI.showToast(
      `Successfully ${successAction}: ${ids.length} ${leadWord}!`,
      "success",
    );

    State.selectedLeads.clear();

    // Refresh the local UI state
    await loadAllData();
    renderLeads();
  } catch (err) {
    console.error("Bulk Assign Error:", err);
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

async function bulkAssignType() {
  const ids = Array.from(State.selectedLeads);
  const typeSelect = document.getElementById("bulk-type-select");

  // Safely grab the value in case the DOM element is missing
  const type = typeSelect ? typeSelect.value : undefined;

  if (!ids.length) {
    UI.showToast("Please select at least one lead first.", "warning");
    return;
  }

  if (!type) {
    UI.showToast("Please select a lead type first.", "error");
    return;
  }

  const leadWord = ids.length === 1 ? "lead" : "leads";

  if (!confirm(`Set type to "${type}" for ${ids.length} ${leadWord}?`)) {
    return;
  }

  setLoading(true);
  try {
    // 🚀 THE UPGRADE: Promise.all blasts these requests out simultaneously
    await Promise.all(
      ids.map(async (id) => {
        await Graph.updateLead(id, { Lead_x0020_Type: type });
      }),
    );

    UI.showToast(
      `Successfully set ${ids.length} ${leadWord} to type: ${type}!`,
      "success",
    );

    State.selectedLeads.clear();
    await loadAllData();
    renderLeads();
  } catch (err) {
    console.error("Bulk Type Assign Error:", err);
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

function bulkExportSelected() {
  const ids = Array.from(State.selectedLeads);
  const leads = State.leads.filter((l) => ids.includes(l.id));

  if (!leads.length) return;

  const today = new Date().toISOString().slice(0, 10);

  // 📝 THE UPGRADE: Stripped Email, added BTN & CBR
  const headers = [
    "Name",
    "Type",
    "BTN",
    "CBR",
    "Status",
    "Source",
    "Assigned To",
    "MRC",
    "Current Products",
    "Last Contacted",
    "Notes",
  ];

  const csv = [headers.join(",")]
    .concat(
      leads.map((l) => {
        // 🛡️ Safe fallbacks to catch any capitalization quirks
        const btn = l.BTN || l.btn || l.phone || "";
        const cbr = l.CBR || l.cbr || l.altPhone || "";
        const mrc = l.currentMRC || l.mrc || "";
        const products = l.currentProducts || l.products || "";

        return [
          l.name,
          l.leadType,
          btn,
          cbr,
          l.status,
          l.source,
          l.assignedTo,
          mrc,
          products,
          l.lastContacted,
          l.notes,
        ]
          .map((v) => '"' + String(v || "").replace(/"/g, '""') + '"')
          .join(",");
      }),
    )
    .join("\n");

  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
  a.download = `raimak-leads-selected-${today}.csv`;
  a.click();

  // 🧹 Clean up the local memory blob
  URL.revokeObjectURL(a.href);

  UI.showToast(`Exported ${leads.length} leads!`, "success");
}

function getFilteredLeads() {
  const { status, search, assignedTo } = State.filters;
  const q = (search || "").trim().toLowerCase();

  // 🚀 FAST PATH: If nothing is filtered, bypass the entire array loop instantly
  if (!q && status === "all" && assignedTo === "all") {
    return State.leads;
  }

  return State.leads.filter(function (l) {
    // 1. Status Match (Instant boolean escape)
    if (status !== "all" && l.status !== status) return false;

    // 2. Strict Agent Match (Instant boolean escape)
    if (assignedTo !== "all" && l.assignedTo !== assignedTo) return false;

    // 3. Targeted Search Match (Only hits this if the lead passed steps 1 & 2)
    if (q) {
      const nameMatch = l.name && l.name.toLowerCase().includes(q);
      const phoneMatch = l.phone && l.phone.includes(q);

      // Smart matching for your updated billing telephone and contact number metrics
      const btnMatch = (l.BTN || l.btn) && String(l.BTN || l.btn).includes(q);
      const cbrMatch = l.cbr && String(l.cbr).includes(q);

      // If it doesn't match any of our target fields, throw it away
      if (!nameMatch && !phoneMatch && !btnMatch && !cbrMatch) return false;
    }

    // Lead passed all tests!
    return true;
  });
}

function applyFilters() {
  State.filters.search =
    (document.getElementById("search-input") || {}).value || "";
  State.filters.status =
    (document.getElementById("filter-status") || {}).value || "all";
  State.filters.assignedTo =
    (document.getElementById("filter-agent") || {}).value || "all";

  const wrap = document.getElementById("leads-table-wrap");

  if (wrap) {
    wrap.replaceChildren(renderLeadsTable(getFilteredLeads()));
  }
}

function renderLeadsTable(leads, compact = false, agentView = false) {
  // 1. Handle Empty State
  if (!leads.length) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.innerHTML = "<p>No leads found.</p>";
    return empty;
  }

  // 2. Setup Template
  const template = document.getElementById("tmpl-leads-table");
  const clone = template.content.cloneNode(true);

  // We wrap it so we can return the outer element properly
  const wrapper = document.createElement("div");
  wrapper.appendChild(clone);

  const thead = wrapper.querySelector("#table-header-row");
  const tbody = wrapper.querySelector("#table-body");

  // 3. Build Headers dynamically based on 'compact' mode
  let headers = "<tr>";
  if (!compact)
    headers += `<th style="width:36px"><input type="checkbox" id="select-all-cb" class="lead-cb" onchange="toggleSelectAll(this.checked)" title="Select all"></th>`;
  headers += `<th>Name</th><th>Type</th><th>Status</th><th>Phone</th><th>Assigned To</th><th>Address</th><th>Last Contacted</th>`;
  if (!compact) headers += `<th>CBR</th><th>BTN</th><th>Flags</th><th></th>`;
  headers += "</tr>";
  thead.innerHTML = headers;

  // 4. Build Rows instantly using our helper
  const admin = isAdmin();
  tbody.innerHTML = leads
    .map((lead) => buildLeadRowHtml(lead, compact, agentView, admin))
    .join("");

  return wrapper.firstElementChild; // Return the living DOM element!
}

function buildLeadRowHtml(lead, compact, agentView, admin) {
  const statusCls =
    "status-" +
    lead.status
      .toLowerCase()
      .replace(/\s+/g, "-")
      .replace(/[^a-z0-9-]/g, "");
  const typeCls = lead.leadType
    ? "lead-type-" + lead.leadType.toLowerCase()
    : "";
  const isChecked = State.selectedLeads.has(lead.id);

  const rowClick = agentView
    ? `loadLeadInFeed('${lead.id}')`
    : admin
      ? `openEditLeadModal('${lead.id}')`
      : "";
  const rowStyle = agentView ? "cursor:pointer" : "";
  const warnCls =
    lead.flags && lead.flags.includes("needs_recycle") ? "row-warn" : "";
  const selCls = isChecked ? "row-selected" : "";

  // 🚀 THE BADGE CALCULATOR
  const prevArray = lead.previousAgents
    ? lead.previousAgents.split(",").filter((a) => a.trim() !== "")
    : [];
  const prevCount = prevArray.length;
  const prevBadge =
    prevCount > 0
      ? `<span title="${escHtml(lead.previousAgents)}" style="font-size:10px; background:#f1f5f9; color:#64748b; padding:2px 6px; border-radius:4px; margin-left:8px; font-weight:600; white-space:nowrap; cursor:help;">↺ ${prevCount} prev agent${prevCount > 1 ? "s" : ""}</span>`
      : "";

  let html = `<tr class="lead-row ${warnCls} ${selCls}" onclick="${rowClick}" style="${rowStyle}">`;

  if (!compact) {
    html += `<td onclick="event.stopPropagation()" style="width:36px">
              <input type="checkbox" class="lead-checkbox lead-cb" data-id="${lead.id}" ${isChecked ? "checked" : ""} onchange="toggleLeadSelect('${lead.id}',this.checked)">
             </td>`;
  }

  // 🚀 THE UPGRADED NAME COLUMN
  html += `<td>
            <div style="display:flex; align-items:center;">
              <span class="lead-name" style="font-weight: 600; color: #0D1B3E;">${escHtml(lead.name)}</span>
              ${prevBadge}
            </div>
           </td>`;

  html += `<td>${lead.leadType ? `<span class="lead-type-badge ${typeCls}">${escHtml(lead.leadType)}</span>` : "—"}</td>`;
  html += `<td><span class="status-badge ${statusCls}">${lead.status}</span></td>`;
  html += `<td class="td-mono">${escHtml(lead.phone || "—")}</td>`;
  html += `<td>${escHtml(lead.assignedTo || "—")}</td>`;
  html += `<td class="td-mono" style="font-size:11px">${lead.address ? escHtml(lead.address) : "—"}${lead.city ? ", " + escHtml(lead.city) : ""}${lead.state ? " " + escHtml(lead.state) : ""}</td>`;
  html += `<td class="td-mono">${formatDate(lead.lastContacted) || "—"}</td>`;

  if (!compact) {
    html += `<td class="td-mono">${escHtml(lead.cbr || "—")}</td>`;
    html += `<td class="td-mono">${escHtml(lead.btn || "—")}</td>`;

    const flagsHtml = (lead.flags || [])
      .map((f) => `<span class="flag flag-${f}">${flagLabel(f)}</span>`)
      .join("");
    html += `<td class="td-flags">${flagsHtml}</td>`;

    html += `<td class="td-actions">`;
    if (admin) {
      html += `
        <button class="btn-icon" onclick="event.stopPropagation();openEditLeadModal('${lead.id}')" title="Edit">
          <svg width="13" height="13" fill="none" viewBox="0 0 24 24"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
        </button>
        <button class="btn-icon btn-danger" onclick="event.stopPropagation();deleteLead('${lead.id}')" title="Delete">
          <svg width="13" height="13" fill="none" viewBox="0 0 24 24"><polyline points="3,6 5,6 21,6" stroke="currentColor" stroke-width="2" stroke-linecap="round"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></svg>
        </button>`;
    }
    html += `</td>`;
  }

  html += `</tr>`;
  return html;
}

function loadLeadInFeed(leadId) {
  let realIndex = (window._myLeads || []).findIndex((l) => l.id === leadId);
  if (realIndex === -1) {
    const masterList = State.leads || State.allLeads || [];
    const globalLead = masterList.find((l) => l.id === leadId);

    if (globalLead) {
      // Temporarily inject it at the very beginning of their personal feed
      window._myLeads = window._myLeads || [];
      window._myLeads.unshift(globalLead);
      realIndex = 0;
    } else {
      if (window.UI && UI.showToast)
        UI.showToast("Lead not found in database.", "error");
      return;
    }
  }

  // 3. Render the card (now guaranteed to work)
  _leadSaved = false;
  _currentFeedIndex = realIndex;
  window._forceShowLead = true;
  renderMyLeads();
  window.scrollTo({ top: 0, behavior: "smooth" });
}

async function handleFileSelect(event) {
  const files = event.target.files;
  if (!files || files.length === 0) return;

  // 🎯 Grab the Lead Type we saved from the modal
  const selectedLeadType = event.target.dataset.leadType || "OFS";
  let combinedCSVData = [];

  UI.showToast(`📄 Reading ${files.length} file(s)...`, "info");

  // Helper function to turn the FileReader into an awaitable Promise
  const readSingleFile = (file) => {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        // Use your existing parseCSV function to turn the text into objects
        const parsed = parseCSV(e.target.result);
        resolve(parsed);
      };
      reader.readAsText(file);
    });
  };

  // Loop through every file they selected and parse it
  for (let i = 0; i < files.length; i++) {
    const parsedData = await readSingleFile(files[i]);
    // Merge the new data into our master list
    combinedCSVData = combinedCSVData.concat(parsedData);
  }

  console.log(
    `📂 Merged ${files.length} files. Total raw leads: ${combinedCSVData.length}`,
  );

  // Send the massive combined list to your uploader
  // Your "Bouncer" inside this function will automatically filter duplicates across all the files!
  await uploadLeadsToSharePoint(combinedCSVData, selectedLeadType);

  // Reset input so they can upload again later
  event.target.value = "";
}

// The CSV Parser (Handles quotes and commas flawlessly)
function parseCSV(text) {
  const lines = text.split("\n").filter((line) => line.trim() !== "");
  const headers = lines[0].split(",").map((h) => h.trim().replace(/"/g, ""));
  const data = [];

  for (let i = 1; i < lines.length; i++) {
    const values = lines[i]
      .split(/,(?=(?:(?:[^"]*"){2})*[^"]*$)/)
      .map((v) => v.trim().replace(/"/g, ""));
    let row = {};

    headers.forEach((h, index) => {
      row[h] = values[index] || "";
    });
    data.push(row);
  }
  return data;
}

async function uploadLeadsToSharePoint(csvData, leadType) {
  // ==========================================
  // 🛡️ 1. THE IN-MEMORY BOUNCER (Deduplication)
  // ==========================================
  const generateKey = (first, last, address) => {
    const clean = (str) =>
      (str || "")
        .replace(/[^\w\s]/gi, "")
        .toLowerCase()
        .trim();
    return `${clean(first)}|${clean(last)}|${clean(address)}`;
  };

  const existingKeys = new Set();
  (State.leads || []).forEach((lead) => {
    const key = generateKey(lead.firstName, lead.lastName, lead.address);
    existingKeys.add(key);
  });

  const validLeads = [];
  let duplicateCount = 0;

  csvData.forEach((row) => {
    const key = generateKey(
      row["FirstName"],
      row["LastName"],
      row["StreetAddress"],
    );
    if (existingKeys.has(key)) {
      duplicateCount++;
    } else {
      validLeads.push(row);
      existingKeys.add(key);
    }
  });

  const totalLeads = validLeads.length;

  if (totalLeads === 0) {
    UI.showToast(
      `❌ Upload aborted: All ${csvData.length} leads are already in the system.`,
      "error",
    );
    return;
  }

  UI.showToast(
    `🚀 Starting batch upload of ${totalLeads} leads... (Skipped ${duplicateCount} duplicates)`,
    "info",
  );

  // ==========================================
  // 🛑 2. HIJACK THE UI
  // ==========================================
  const importBtn = document.getElementById("importLeadsBtn");
  const originalBtnHTML = importBtn ? importBtn.innerHTML : "Import CSV";
  if (importBtn) {
    importBtn.disabled = true;
    importBtn.style.cursor = "not-allowed";
  }

  // ==========================================
  // 🔀 3. URL SETUP FOR BATCHING
  // ==========================================
  const host = Config.sharePoint.hostname;
  const sitePath = Config.sharePoint.sites.team;
  const listId = Config.sharePoint.lists.leadsList;
  const batchUrl = `${Config.sharePoint.graphBase}/$batch`;
  const relativeUploadUrl = `/sites/${host}:/${sitePath}:/lists/${listId}/items`;

  let successCount = 0;
  let failCount = 0;
  const batchSize = 20;

  // ==========================================
  // 📦 4. THE BATCHING ENGINE
  // ==========================================
  for (let i = 0; i < totalLeads; i += batchSize) {
    const chunk = validLeads.slice(i, i + batchSize);
    const currentBatchNum = Math.ceil(i / batchSize) + 1;
    const totalBatches = Math.ceil(totalLeads / batchSize);

    if (importBtn) {
      importBtn.innerHTML = `⏳ Batch ${currentBatchNum}/${totalBatches}... (${i} saved)`;
    }

    const batchRequests = chunk.map((row, index) => {
      return {
        id: String(index + 1),
        method: "POST",
        url: relativeUploadUrl,
        headers: { "Content-Type": "application/json" },
        body: {
          fields: {
            FirstName: row["FirstName"],
            LastName: row["LastName"],
            WorkAddress: row["StreetAddress"],
            WorkCity: row["City"],
            State: row["State"],
            Zip: row["Zip"],
            Lead_x0020_Type: leadType,
            Agent_x0020_Assigned: "",
            Status: "New",
          },
        },
      };
    });

    try {
      // 🚀 THE UPGRADE: Aligning with the new apiFetch options object
      const batchResponse = await Graph.apiFetch(batchUrl, {
        method: "POST",
        body: { requests: batchRequests },
      });

      if (batchResponse && batchResponse.responses) {
        batchResponse.responses.forEach((res) => {
          if (res.status >= 200 && res.status < 300) {
            successCount++;
          } else {
            console.error(`❌ Batch item ${res.id} failed:`, res.body);
            failCount++;
          }
        });
      }
    } catch (error) {
      console.error("❌ Entire batch failed:", error);
      failCount += chunk.length;
      UI.showToast(`⚠️ Network error on batch ${currentBatchNum}.`, "error");
    }
  }

  // ✅ 5. RESTORE THE UI & REPORT
  if (importBtn) {
    importBtn.disabled = false;
    importBtn.style.cursor = "pointer";
    importBtn.innerHTML = originalBtnHTML;
  }

  if (failCount === 0) {
    UI.showToast(`✅ Upload complete! ${successCount} leads added.`, "success");
  } else {
    UI.showToast(
      `⚠️ Finished: ${successCount} added, ${failCount} failed.`,
      "warning",
    );
  }

  await loadAllData();
}
// ============================================================
//  DAILY REPORT (Admin only)
// ============================================================
async function renderDailyReport() {
  if (!isAdmin()) {
    navigate("myleads");
    return;
  }

  // 🎨 HELPER: Generates the "Stoplight" color shift
  function getDynamicColor(contacts) {
    let hue = 0;
    if (contacts <= 25) {
      hue = (contacts / 25) * 60; // Red to Yellow
    } else if (contacts <= 50) {
      hue = 60 + ((contacts - 25) / 25) * 60; // Yellow to Green
    } else {
      hue = 120; // Goal Green
    }
    return `hsl(${hue}, 80%, 45%)`;
  }

  document.getElementById("main-content").innerHTML = `
    <div class="view-header"><h1 class="view-title">Daily Report</h1></div>
    <div class="card"><div class="empty-state" style="padding:40px">Loading report...</div></div>`;

  try {
    const stats = await Graph.getDailyStats();
    const today = new Date().toLocaleDateString("en-GB", {
      weekday: "long",
      year: "numeric",
      month: "long",
      day: "numeric",
    });

    // Generate Agent Options for the dropdown
    const agentOptions = (State.contractors || [])
      .map((c) => `<option value="${c.email}">${c.name}</option>`)
      .join("");

    document.getElementById("main-content").innerHTML = `
      <div class="view-header" style="display: flex; justify-content: space-between; align-items: flex-end; flex-wrap: wrap; gap: 15px;">
        <div>
          <h1 class="view-title">Daily Report</h1>
          <span class="view-subtitle">// ${today}</span>
        </div>
        
        <div style="display: flex; flex-direction: column; gap: 10px; align-items: flex-end;">
          <div style="display: flex; gap: 10px; align-items: center; background: #f8fafc; padding: 10px 15px; border-radius: 8px; border: 1px solid #e2e8f0;">
            
            <div style="display: flex; flex-direction: column; gap: 2px;">
               <span style="font-size: 11px; font-weight: 700; color: #64748b; text-transform: uppercase;">Agent Filter</span>
               <select id="export-agent" style="padding: 4px 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 13px; width: 150px;">
                 <option value="ALL">All Agents</option>
                 ${agentOptions}
               </select>
            </div>

            <div style="display: flex; flex-direction: column; gap: 2px;">
               <span style="font-size: 11px; font-weight: 700; color: #64748b; text-transform: uppercase;">Start Date</span>
               <input type="date" id="export-start" style="padding: 4px 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 13px;">
            </div>

            <div style="display: flex; flex-direction: column; gap: 2px;">
               <span style="font-size: 11px; font-weight: 700; color: #64748b; text-transform: uppercase;">End Date</span>
               <input type="date" id="export-end" style="padding: 4px 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 13px;">
            </div>

            <button class="btn-ghost" style="font-size: 13px; margin-top: 15px;" onclick="exportDateRangeCSV()" id="btn-export-range">Download CSV</button>
          </div>
          <button class="btn-ghost" style="font-size: 13px;" onclick="exportReportCSV()">Export Today Only (.csv)</button>
        </div>
      </div>

      <div class="kpi-grid">
        <div class="kpi-card kpi-primary">
          <span class="kpi-label">Total Contacts Today</span>
          <span class="kpi-value">${stats.reduce((s, a) => s + a.contacts, 0)}</span>
        </div>
        <div class="kpi-card kpi-success">
          <span class="kpi-label">Total Sales Today</span>
          <span class="kpi-value">${State.todaySales ? State.todaySales.length : 0}</span>
        </div>
        <div class="kpi-card kpi-info">
          <span class="kpi-label">Active Agents</span>
          <span class="kpi-value">${stats.length}</span>
        </div>
        <div class="kpi-card kpi-neutral">
          <span class="kpi-label">Avg Contacts/Agent</span>
          <span class="kpi-value">${
            stats.length
              ? Math.round(
                  stats.reduce((s, a) => s + a.contacts, 0) / stats.length,
                )
              : 0
          }</span>
        </div>
      </div>

      <div class="card">
        <div class="card-header"><h2 class="card-title">Agent Breakdown</h2></div>
        <div class="table-wrap">
          <table class="data-table">
            <thead>
              <tr>
                <th>Agent</th>
                <th style="text-align:center;">Contacts</th>
                <th style="text-align:center;">Sales</th>
                <th style="text-align:center;">Avg. Contact Time</th>
                <th style="text-align:right;">Last Action</th>
              </tr>
            </thead>
            <tbody>
              ${
                stats.length
                  ? stats
                      .map(function (a) {
                        // Sort actions to find the most recent one
                        const last = a.actions.length
                          ? a.actions.sort(
                              (x, y) =>
                                new Date(y.timestamp) - new Date(x.timestamp),
                            )[0]
                          : null;

                        const statusColor = getDynamicColor(a.contacts);

                        return `<tr>
                        <td><span class="lead-name" style="font-weight:600;">${escHtml(a.agent)}</span></td>
                        <td style="text-align:center;">
                          <span class="td-mono" style="color:${statusColor}; font-size:1.1rem; font-weight:800;">${a.contacts}</span>
                        </td>
                        <td style="text-align:center;">
                          <span class="status-badge status-sold" style="font-weight:700; padding: 4px 12px;">${a.sold}</span>
                        </td>
                        <td style="text-align:center;" class="td-mono">${a.avgTime || "—"}</td>
                        <td style="text-align:right;" class="td-mono">
                          <span style="color:var(--text-3); font-size: 0.85rem;">${last ? formatDateTime(last.timestamp) : "—"}</span>
                        </td>
                      </tr>`;
                      })
                      .join("")
                  : `<tr><td colspan="5" class="empty-state">No activity today yet.</td></tr>`
              }
            </tbody>
          </table>
        </div>
      </div>`;

    window._reportStats = stats;
  } catch (err) {
    UI.showToast("Failed to load report: " + err.message, "error");
    console.error(err);
  }
}

async function exportDateRangeCSV() {
  const startVal = document.getElementById("export-start").value;
  const endVal = document.getElementById("export-end").value;
  const agentFilter = document.getElementById("export-agent").value;
  const btn = document.getElementById("btn-export-range");

  if (!startVal || !endVal)
    return UI.showToast("Select a date range.", "warning");

  // 🛑 THE FIX: Force Local Time by splitting the YYYY-MM-DD string
  const [sYear, sMonth, sDay] = startVal.split("-");
  const startObj = new Date(sYear, sMonth - 1, sDay, 0, 0, 0, 0);

  const [eYear, eMonth, eDay] = endVal.split("-");
  const endObj = new Date(eYear, eMonth - 1, eDay, 23, 59, 59, 999);

  if (startObj > endObj) return UI.showToast("Invalid date range.", "warning");

  if (btn) {
    btn.disabled = true;
    btn.textContent = "Processing...";
  }

  try {
    // 1. Filter Logs by Date AND Agent
    let logs = (State.activityLog || []).filter((log) => {
      if (!log.timestamp) return false;
      const logDate = new Date(log.timestamp);
      const isWithinDate = logDate >= startObj && logDate <= endObj;

      if (!isWithinDate) return false;
      if (agentFilter !== "ALL") {
        return log.agentEmail === agentFilter || log.agent === agentFilter;
      }
      return true;
    });

    // 2. Generate a reference list of every date in the range
    const rangeDates = [];
    let curr = new Date(startObj);
    while (curr <= endObj) {
      rangeDates.push(curr.toLocaleDateString());
      curr.setDate(curr.getDate() + 1);
    }

    const agentData = {};

    // 3. Tally Data
    logs.forEach((log) => {
      const email = log.agentEmail || log.agent || "Unknown";
      const localDateStr = new Date(log.timestamp).toLocaleDateString();

      if (!agentData[email]) {
        agentData[email] = {
          agent: email,
          totalContacts: 0,
          totalSales: 0,
          dailyCounts: {},
          dailySales: {},
          daysWorked: new Set(),
          timestampsByDay: {},
        };
      }

      agentData[email].totalContacts++;
      agentData[email].dailyCounts[localDateStr] =
        (agentData[email].dailyCounts[localDateStr] || 0) + 1;

      if (
        log.action &&
        (log.action.includes("Sold") || log.action.includes("Sale"))
      ) {
        agentData[email].totalSales++;
        agentData[email].dailySales[localDateStr] =
          (agentData[email].dailySales[localDateStr] || 0) + 1;
      }

      agentData[email].daysWorked.add(localDateStr);
      if (!agentData[email].timestampsByDay[localDateStr])
        agentData[email].timestampsByDay[localDateStr] = [];
      agentData[email].timestampsByDay[localDateStr].push(
        new Date(log.timestamp).getTime(),
      );
    });

    // 4. Build CSV Rows (Block Format per Agent)
    const csvRows = [];

    // Global Header for the Base Stats
    csvRows.push(
      [
        "Agent",
        "Total Contacts",
        "Total Sales",
        "Overall Close Rate",
        "Days Logged In",
        "Avg Contacts/Active Day",
        "Avg Time Between Contacts",
      ].join(","),
    );

    // 5. Build Each Agent's Scorecard Block
    Object.values(agentData).forEach((data) => {
      const activeDays = data.daysWorked.size || 1;

      // Calculate Time Gaps (Ignoring lunch/breaks over 2hrs)
      let totalDiffMs = 0,
        diffCount = 0;
      Object.values(data.timestampsByDay).forEach((dayTimes) => {
        if (dayTimes.length > 1) {
          dayTimes.sort((a, b) => a - b);
          for (let i = 1; i < dayTimes.length; i++) {
            const diff = dayTimes[i] - dayTimes[i - 1];
            if (diff < 7200000) {
              totalDiffMs += diff;
              diffCount++;
            }
          }
        }
      });

      let avgTimeStr = "—";
      if (diffCount > 0) {
        const avgMs = totalDiffMs / diffCount;
        avgTimeStr = `${Math.floor(avgMs / 60000)}m ${Math.floor((avgMs % 60000) / 1000)}s`;
      }

      const overallCloseRate =
        data.totalContacts > 0
          ? Math.round((data.totalSales / data.totalContacts) * 100) + "%"
          : "0%";

      // ROW 1: The Agent's Base Stats
      csvRows.push(
        [
          `"${data.agent}"`,
          data.totalContacts,
          data.totalSales,
          `"${overallCloseRate}"`,
          activeDays,
          (data.totalContacts / activeDays).toFixed(1),
          `"${avgTimeStr}"`,
        ].join(","),
      );

      // Filter out dates where this specific agent made 0 contacts
      const activeDatesOnly = rangeDates.filter(
        (dateStr) => (data.dailyCounts[dateStr] || 0) > 0,
      );

      if (activeDatesOnly.length > 0) {
        // Separates Base Stats from Date Headers
        csvRows.push("");

        // ROW 2: Date Headers
        csvRows.push(
          ["Date", ...activeDatesOnly.map((d) => `"${d}"`)].join(","),
        );

        // ROW 3: Contacts
        csvRows.push(
          ["Contacts", ...activeDatesOnly.map((d) => data.dailyCounts[d])].join(
            ",",
          ),
        );

        // ROW 4: Sales
        csvRows.push(
          [
            "Sales",
            ...activeDatesOnly.map((d) => data.dailySales[d] || 0),
          ].join(","),
        );

        // ROW 5: Close Rate
        csvRows.push(
          [
            "Close Rate",
            ...activeDatesOnly.map((d) => {
              const c = data.dailyCounts[d];
              const s = data.dailySales[d] || 0;
              return `"${Math.round((s / c) * 100)}%"`;
            }),
          ].join(","),
        );
      }

      // Pushing two empty strings creates a double line break between agents
      csvRows.push("", "");
    });

    if (csvRows.length === 1) {
      // Only headers generated
      if (btn) {
        btn.disabled = false;
        btn.textContent = "Download CSV";
      }
      return UI.showToast("No activity found.", "warning");
    }

    // 6. Download Trigger
    const csvContent = csvRows.join("\n");
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Report_${agentFilter}_${startVal}_to_${endVal}.csv`;
    a.click();
    URL.revokeObjectURL(url);

    UI.showToast("Report downloaded!", "success");
  } catch (err) {
    console.error("Export failed:", err);
    UI.showToast("Export failed: " + err.message, "error");
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.textContent = "Download CSV";
    }
  }
}

function exportReportCSV() {
  const now = new Date();

  // Force local time boundaries for exactly TODAY
  const startObj = new Date(
    now.getFullYear(),
    now.getMonth(),
    now.getDate(),
    0,
    0,
    0,
    0,
  );
  const endObj = new Date(
    now.getFullYear(),
    now.getMonth(),
    now.getDate(),
    23,
    59,
    59,
    999,
  );
  const todayStr = now.toLocaleDateString();

  // 1. Filter logs for TODAY only
  let todayLogs = (State.activityLog || []).filter((log) => {
    if (!log.timestamp) return false;
    const logDate = new Date(log.timestamp);
    return logDate >= startObj && logDate <= endObj;
  });

  const agentData = {};

  // 2. Tally Data (Using the same advanced math as the Date Range tool)
  todayLogs.forEach((log) => {
    const email = log.agentEmail || log.agent || "Unknown";

    if (!agentData[email]) {
      agentData[email] = {
        agent: email,
        totalContacts: 0,
        totalSales: 0,
        timestamps: [],
      };
    }

    agentData[email].totalContacts++;

    if (
      log.action &&
      (log.action.includes("Sold") || log.action.includes("Sale"))
    ) {
      agentData[email].totalSales++;
    }

    agentData[email].timestamps.push(new Date(log.timestamp).getTime());
  });

  // 3. Build CSV Rows
  const csvRows = [];

  // Global Header
  csvRows.push(
    [
      "Agent",
      "Contacts Today",
      "Sales Today",
      "Close Rate",
      "Avg Time Between Contacts",
      "Date",
    ].join(","),
  );

  // 4. Calculate Advanced Metrics per Agent
  Object.values(agentData).forEach((data) => {
    // Calculate Time Gaps (Ignoring lunch/breaks over 2hrs)
    let totalDiffMs = 0,
      diffCount = 0;
    if (data.timestamps.length > 1) {
      data.timestamps.sort((a, b) => a - b);
      for (let i = 1; i < data.timestamps.length; i++) {
        const diff = data.timestamps[i] - data.timestamps[i - 1];
        if (diff < 7200000) {
          totalDiffMs += diff;
          diffCount++;
        }
      }
    }

    let avgTimeStr = "—";
    if (diffCount > 0) {
      const avgMs = totalDiffMs / diffCount;
      avgTimeStr = `${Math.floor(avgMs / 60000)}m ${Math.floor((avgMs % 60000) / 1000)}s`;
    }

    const closeRate =
      data.totalContacts > 0
        ? Math.round((data.totalSales / data.totalContacts) * 100) + "%"
        : "0%";

    csvRows.push(
      [
        `"${data.agent}"`,
        data.totalContacts,
        data.totalSales,
        `"${closeRate}"`,
        `"${avgTimeStr}"`,
        `"${todayStr}"`,
      ].join(","),
    );
  });

  if (csvRows.length === 1) {
    return UI.showToast("No activity found for today.", "warning");
  }

  // 5. Download Trigger
  const csvContent = csvRows.join("\n");
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;

  // Format the filename safely (e.g. 5-29-2026)
  const safeDate = todayStr.replace(/\//g, "-");
  a.download = `raimak-daily-report-${safeDate}.csv`;

  a.click();
  URL.revokeObjectURL(url);

  UI.showToast("Daily report exported!", "success");
}

// ============================================================
//  RAIMAK TEAM (Admin only)
// ============================================================
function renderContractors() {
  if (!isAdmin()) {
    navigate("myleads");
    return;
  }
  const { contractors, leads } = State;
  const max = Config.rules.maxLeadsPerAgent;
  document.getElementById("main-content").innerHTML = `
    <div class="view-header">
      <h1 class="view-title">Raimak Team</h1>
      <span class="view-subtitle">// ${contractors.length} agents</span>
    </div>
    <div class="contractor-grid">
      ${contractors
        .map(function (c) {
          const count = leads.filter(function (l) {
            return (
              l.assignedTo === c.name &&
              !Config.terminalStatuses.includes(l.status)
            );
          }).length;
          const pct = Math.min(100, Math.round((count / max) * 100));
          const contacts = Graph.agentContactsToday(
            c.email || c.name,
            State.activityLog,
          );
          return `
          <div class="contractor-card">
            <div class="contractor-header">
              <div class="contractor-avatar">${c.name[0].toUpperCase()}</div>
              <div><div class="contractor-name">${escHtml(c.name)}</div><div class="contractor-role">${escHtml(c.role)}</div></div>
              <span class="status-dot ${c.active ? "dot-active" : "dot-inactive"}"></span>
            </div>
            <div class="contractor-email">${escHtml(c.email || "No email")}</div>
            <div class="load-label"><span>Lead Load</span><span class="${count >= max ? "text-danger" : ""}">${count}/${max}</span></div>
            <div class="load-bar-wrap"><div class="load-bar ${pct >= 100 ? "load-full" : pct >= 80 ? "load-high" : ""}" style="width:${pct}%"></div></div>
            <div class="load-label" style="margin-top:10px"><span>Contacts Today</span><span>${contacts}/${Config.rules.maxContactsPerDay}</span></div>
            <div class="load-bar-wrap"><div class="load-bar ${contacts >= Config.rules.maxContactsPerDay ? "load-full" : ""}" style="width:${Math.min(100, Math.round((contacts / Config.rules.maxContactsPerDay) * 100))}%"></div></div>
          </div>`;
        })
        .join("")}
    </div>`;
}

// ============================================================
//  ACTIVITY LOG (Admin only)
// ============================================================
function renderActivity() {
  if (!isAdmin()) {
    navigate("myleads");
    return;
  }

  const { activityLog, contractors } = State;

  // 1. THE PREDICTABLE IDENTITY MAP
  const identityMap = {};

  function getStandardEmail(name) {
    if (!name) return "";
    const parts = name.trim().toLowerCase().split(/\s+/);
    if (parts.length >= 2) {
      return `${parts[0].charAt(0)}.${parts[parts.length - 1]}@raimak.com`;
    }
    return "";
  }

  function registerPerson(name, knownEmail) {
    if (!name) return;
    const officialName = name.trim();
    const lowerName = officialName.toLowerCase();

    identityMap[lowerName] = officialName;

    const generatedEmail = getStandardEmail(officialName);
    if (generatedEmail) identityMap[generatedEmail] = officialName;

    if (knownEmail) identityMap[knownEmail.trim().toLowerCase()] = officialName;
  }

  (contractors || []).forEach((c) => registerPerson(c.name, c.email));

  const user = State.currentUser;
  if (user) registerPerson(user.name, user.email);

  // 2. THE BULLETPROOF NORMALIZER
  function getCleanAgentName(rawString) {
    if (!rawString) return "";
    const cleanString = rawString.trim().toLowerCase();

    if (identityMap[cleanString]) {
      return identityMap[cleanString];
    }

    return rawString.trim().replace(/\w\S*/g, function (txt) {
      return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
    });
  }

  // 3. EXTRACT UNIQUE DATA FOR DROPDOWNS
  const uniqueAgents = [
    ...new Set(
      activityLog.map((e) => getCleanAgentName(e.agent)).filter(Boolean),
    ),
  ].sort();

  const uniqueActions = [
    ...new Set(activityLog.map((e) => (e.action || "").trim()).filter(Boolean)),
  ].sort();

  // Draw the layout skeleton
  document.getElementById("main-content").innerHTML = `
    <div class="view-header">
      <div>
        <h1 class="view-title">Activity Log</h1>
        <span class="view-subtitle">// ${activityLog.length} total entries</span>
      </div>
    </div>

    <div class="card">
      <div class="card-header" style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:16px;">
        
        <div>
          <h2 class="card-title">Recent Activity</h2>
          <span class="card-meta" id="activity-meta-count">Loading...</span>
        </div>
        
        <div style="display:flex; gap:16px; align-items:center; flex-wrap:wrap; justify-content:flex-end; flex:1;">
          
          <div style="display:flex; gap:8px; align-items:center;">
            
            <div id="date-inputs-container" style="display:flex; align-items:center; gap:6px; max-width:0px; opacity:0; overflow:hidden; transition:all 0.3s ease-out; white-space:nowrap; pointer-events:none;">
              <input type="date" id="filter-start-date" class="form-input" style="padding:6px 10px; font-size:13px;">
              <span style="font-size:13px; color:#666; font-weight:600;">to</span>
              <input type="date" id="filter-end-date" class="form-input" style="padding:6px 10px; font-size:13px;">
            </div>

            <label style="display:flex; align-items:center; gap:4px; font-size:13px; font-weight:600; cursor:pointer; color:#0D1B3E; margin:0; white-space:nowrap;">
              <input type="checkbox" id="toggle-date-filter" style="cursor:pointer; margin:0; width:14px; height:14px;">
              Date
            </label>
          </div>

          <select id="filter-agent" class="form-input" style="padding:6px 24px 6px 10px; font-size:13px; min-width:130px; max-width:160px;">
            <option value="all">All Agents</option>
            ${uniqueAgents.map((a) => `<option value="${escHtml(a)}">${escHtml(a)}</option>`).join("")}
          </select>

          <select id="filter-action" class="form-input" style="padding:6px 24px 6px 10px; font-size:13px; min-width:130px; max-width:160px;">
            <option value="all">All Actions</option>
            ${uniqueActions.map((a) => `<option value="${escHtml(a)}">${escHtml(a)}</option>`).join("")}
          </select>

          <select id="sort-date" class="form-input" style="padding:6px 24px 6px 10px; font-size:13px; min-width:130px; max-width:160px;">
            <option value="desc">Newest First</option>
            <option value="asc">Oldest First</option>
          </select>

          <div style="display:flex; gap:6px; align-items:center; white-space:nowrap; border-left: 1px solid #e2e8f0; padding-left: 12px; margin-left: 4px;">
            <button id="btn-prev-page" class="btn-secondary" style="padding: 6px 12px; font-size:13px;">&larr; Prev</button>
            <span id="page-indicator" style="font-family:var(--font-mono); font-size:13px; font-weight:600; color:#0D1B3E; min-width: 50px; text-align: center;">Pg 1</span>
            <button id="btn-next-page" class="btn-secondary" style="padding: 6px 12px; font-size:13px;">Next &rarr;</button>
          </div>

        </div>
      </div>
      
      <div class="table-wrap">
        <table class="data-table">
          <thead><tr><th>Time</th><th>Lead</th><th>Action</th><th>Agent</th><th>Notes</th></tr></thead>
          <tbody id="activity-tbody">
            </tbody>
        </table>
      </div>
    </div>
  `;

  // Internal State & DOM Pointers
  let currentPage = 1;
  const itemsPerPage = 50;

  const tbody = document.getElementById("activity-tbody");
  const prevBtn = document.getElementById("btn-prev-page");
  const nextBtn = document.getElementById("btn-next-page");
  const pageIndicator = document.getElementById("page-indicator");

  const agentFilter = document.getElementById("filter-agent");
  const actionFilter = document.getElementById("filter-action");
  const dateSort = document.getElementById("sort-date");
  const metaCount = document.getElementById("activity-meta-count");

  const dateToggle = document.getElementById("toggle-date-filter");
  const dateContainer = document.getElementById("date-inputs-container");
  const startDateFilter = document.getElementById("filter-start-date");
  const endDateFilter = document.getElementById("filter-end-date");

  // The Smart Table Renderer
  function updateTable() {
    const selectedAgent = agentFilter ? agentFilter.value : "all";
    const selectedAction = actionFilter ? actionFilter.value : "all";
    const sortOrder = dateSort ? dateSort.value : "desc";

    const isDateActive = dateToggle ? dateToggle.checked : false;
    const startDateStr = startDateFilter ? startDateFilter.value : "";
    const endDateStr = endDateFilter ? endDateFilter.value : "";

    const startTimestamp =
      isDateActive && startDateStr
        ? new Date(startDateStr + "T00:00:00").getTime()
        : 0;
    const endTimestamp =
      isDateActive && endDateStr
        ? new Date(endDateStr + "T23:59:59").getTime()
        : Infinity;

    // Step A: Filter by Agent, Action, AND Date
    let processedLog = activityLog.filter(function (e) {
      const matchAgent =
        selectedAgent === "all" || getCleanAgentName(e.agent) === selectedAgent;
      const matchAction =
        selectedAction === "all" || (e.action || "").trim() === selectedAction;

      const logTime = new Date(e.timestamp || 0).getTime();
      const matchDate = logTime >= startTimestamp && logTime <= endTimestamp;

      return matchAgent && matchAction && matchDate;
    });

    // Step B: Sort by Date
    processedLog.sort(function (a, b) {
      const timeA = new Date(a.timestamp || 0).getTime();
      const timeB = new Date(b.timestamp || 0).getTime();
      return sortOrder === "desc" ? timeB - timeA : timeA - timeB;
    });

    // Step C: Pagination Math
    const total = processedLog.length;
    const totalPages = Math.max(1, Math.ceil(total / itemsPerPage));

    if (currentPage > totalPages) currentPage = totalPages;
    if (currentPage < 1) currentPage = 1;

    // Update UI Text
    pageIndicator.textContent = `Pg ${currentPage} / ${totalPages}`;
    if (metaCount) metaCount.textContent = `Showing ${total} entries`;

    prevBtn.disabled = currentPage === 1;
    prevBtn.style.opacity = currentPage === 1 ? "0.4" : "1";

    nextBtn.disabled = currentPage === totalPages;
    nextBtn.style.opacity = currentPage === totalPages ? "0.4" : "1";

    // Step D: Slice and Draw HTML
    const startIndex = (currentPage - 1) * itemsPerPage;
    const displayLog = processedLog.slice(
      startIndex,
      startIndex + itemsPerPage,
    );

    if (displayLog.length === 0) {
      tbody.innerHTML = `<tr><td colspan="5" class="empty-state">No activity matches these filters.</td></tr>`;
      return;
    }

    // 🚀 THE UPGRADE: Inject the raw array index and add a pointer cursor
    tbody.innerHTML = displayLog
      .map(function (e) {
        const rawIndex = activityLog.indexOf(e);
        return `
        <tr data-raw-index="${rawIndex}" style="cursor: pointer; transition: background 0.1s;" onmouseover="this.style.background='#f8fafc'" onmouseout="this.style.background='transparent'">
          <td class="td-mono">${formatDateTime(e.timestamp)}</td>
          <td>${escHtml(e.leadName || e.leadId || "—")}</td>
          <td><span class="action-badge">${escHtml(e.action || "—")}</span></td>
          <td>${escHtml(getCleanAgentName(e.agent) || "—")}</td>
          <td class="td-notes" style="max-width: 200px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${escHtml(e.notes || "")}</td>
        </tr>`;
      })
      .join("");
  }

  // ==========================================
  // 🔍 THE MODAL LOGIC & TIME CALCULATOR
  // ==========================================
  function showActivityModal(rawIndex) {
    const entry = activityLog[rawIndex];
    if (!entry) return;

    const currentAgent = getCleanAgentName(entry.agent);

    // Grab all logs for this specific agent and sort newest -> oldest
    const agentLogs = activityLog
      .filter((e) => getCleanAgentName(e.agent) === currentAgent)
      .sort(
        (a, b) =>
          new Date(b.timestamp || 0).getTime() -
          new Date(a.timestamp || 0).getTime(),
      );

    const entryIndex = agentLogs.indexOf(entry);
    let timeSinceLast = "— (First logged action)";

    // If it's not their oldest action, look at the NEXT older item in their array
    if (entryIndex < agentLogs.length - 1) {
      const previousEntry = agentLogs[entryIndex + 1];
      const diffMs =
        new Date(entry.timestamp || 0).getTime() -
        new Date(previousEntry.timestamp || 0).getTime();

      const diffMins = Math.floor(diffMs / 60000);
      const diffSecs = Math.floor((diffMs % 60000) / 1000);

      if (diffMins > 60) {
        const diffHours = Math.floor(diffMins / 60);
        const remainMins = diffMins % 60;
        timeSinceLast = `${diffHours}h ${remainMins}m since last action`;
      } else {
        timeSinceLast = `${diffMins}m ${diffSecs}s since last action`;
      }
    }

    // Build and inject the modal
    const overlay = document.createElement("div");
    overlay.style.cssText =
      "position:fixed; top:0; left:0; width:100vw; height:100vh; background:rgba(13, 27, 62, 0.6); z-index:9999; display:flex; align-items:center; justify-content:center; backdrop-filter: blur(3px);";

    const modal = document.createElement("div");
    // 🛡️ THE FIX: Re-applying dark text color to ensure readability
    modal.style.cssText =
      "background:#fff; padding:24px; border-radius:12px; width:450px; max-width:90vw; box-shadow:0 10px 25px rgba(0,0,0,0.2); display:flex; flex-direction:column; gap:16px; color:#0D1B3E;";

    modal.innerHTML = `
      <div style="display:flex; justify-content:space-between; align-items:flex-start; border-bottom:1px solid #e2e8f0; padding-bottom:12px;">
        <div>
          <h3 style="margin:0 0 4px 0; font-size:18px; color:#0D1B3E;">Action: ${escHtml(entry.action || "Unknown")}</h3>
          <p style="margin:0; font-size:13px; color:#64748b; font-family:var(--font-mono);">${formatDateTime(entry.timestamp)}</p>
        </div>
        <button id="closeModalBtn" style="background:none; border:none; font-size:20px; cursor:pointer; color:#94a3b8;">&times;</button>
      </div>

      <div style="display:grid; grid-template-columns: 1fr 1fr; gap:12px; font-size:14px; color:#0D1B3E;">
        <div>
          <span style="color:#64748b; font-size:12px; text-transform:uppercase; font-weight:600;">Agent</span>
          <div style="font-weight:500;">${escHtml(currentAgent)}</div>
        </div>
        <div>
          <span style="color:#64748b; font-size:12px; text-transform:uppercase; font-weight:600;">Lead Info</span>
          <div style="font-weight:500;">${escHtml(entry.leadName || entry.leadId || "—")}</div>
        </div>
      </div>

      <div style="background:#f1f5f9; border-radius:6px; padding:12px;">
        <span style="color:#64748b; font-size:12px; text-transform:uppercase; font-weight:600;">Pacing Metric</span>
        <div style="font-family:var(--font-mono); font-size:13px; margin-top:4px; color:#3b82f6; font-weight:600;">
          ⏱️ ${timeSinceLast}
        </div>
      </div>

      <div>
        <span style="color:#64748b; font-size:12px; text-transform:uppercase; font-weight:600;">Full Note</span>
        <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:6px; padding:12px; margin-top:4px; font-size:14px; min-height:80px; white-space:pre-wrap; line-height:1.5; color:#1a1a1a;">${escHtml(entry.notes || "No notes provided.")}</div>
      </div>
    `;

    overlay.appendChild(modal);
    document.body.appendChild(overlay);

    // Close logic
    const closeIt = () => overlay.remove();
    document.getElementById("closeModalBtn").onclick = closeIt;
    overlay.onclick = (e) => {
      if (e.target === overlay) closeIt();
    };
  }

  // Bind the delegated listener to the table body
  if (tbody) {
    tbody.addEventListener("click", (e) => {
      const tr = e.target.closest("tr");
      if (!tr || !tr.dataset.rawIndex) return;
      showActivityModal(parseInt(tr.dataset.rawIndex, 10));
    });
  }

  // Attach Event Listeners
  if (prevBtn)
    prevBtn.addEventListener("click", () => {
      if (currentPage > 1) {
        currentPage--;
        updateTable();
      }
    });
  if (nextBtn)
    nextBtn.addEventListener("click", () => {
      currentPage++;
      updateTable();
    });

  if (agentFilter)
    agentFilter.addEventListener("change", () => {
      currentPage = 1;
      updateTable();
    });
  if (actionFilter)
    actionFilter.addEventListener("change", () => {
      currentPage = 1;
      updateTable();
    });
  if (dateSort)
    dateSort.addEventListener("change", () => {
      currentPage = 1;
      updateTable();
    });

  if (startDateFilter)
    startDateFilter.addEventListener("change", () => {
      currentPage = 1;
      updateTable();
    });
  if (endDateFilter)
    endDateFilter.addEventListener("change", () => {
      currentPage = 1;
      updateTable();
    });

  // THE SMOOTH ANIMATION TOGGLE
  if (dateToggle) {
    dateToggle.addEventListener("change", (e) => {
      if (e.target.checked) {
        dateContainer.style.maxWidth = "350px";
        dateContainer.style.opacity = "1";
        dateContainer.style.pointerEvents = "auto";
      } else {
        dateContainer.style.maxWidth = "0px";
        dateContainer.style.opacity = "0";
        dateContainer.style.pointerEvents = "none";

        if (startDateFilter) startDateFilter.value = "";
        if (endDateFilter) endDateFilter.value = "";
        currentPage = 1;
        updateTable();
      }
    });
  }

  // Initialize on first load
  updateTable();
}

// ============================================================
//  STATS PAGE
// ============================================================
function renderStats() {
  const currentUser = State.currentUser;
  if (!currentUser || !currentUser.email) return;

  const appContainer = document.getElementById("main-content");
  appContainer.innerHTML = "";

  const template = document.getElementById("tmpl-stats-page");
  const clone = template.content.cloneNode(true);
  appContainer.appendChild(clone);

  const adminControls = document.getElementById("admin-stats-controls");
  const agentSelect = document.getElementById("stats-agent-select");
  const timeframeSelect = document.getElementById("stats-timeframe-select");

  const refreshData = () => {
    // Now grabs 'all' if selected, otherwise target email
    const targetEmail =
      isAdmin() && agentSelect.value ? agentSelect.value : currentUser.email;
    const timeframe = timeframeSelect.value;
    paintStats(targetEmail, timeframe);
  };

  timeframeSelect.addEventListener("change", refreshData);

  if (isAdmin()) {
    adminControls.style.display = "flex";

    // 🌟 THE FIX: Inject the "All Agents" option first
    const allOption = document.createElement("option");
    allOption.value = "all";
    allOption.textContent = "-- All Agents (Floor Total) --";
    agentSelect.appendChild(allOption);

    const agents = (State.agentScores || [])
      .map((s) => ({
        name: s.AgentName || "Unknown",
        email: s.AgentEmail || "",
      }))
      .filter((a) => a.email !== "")
      .sort((a, b) => a.name.localeCompare(b.name));

    const seenEmails = new Set();

    agents.forEach((agent) => {
      const safeEmail = agent.email.toLowerCase().trim();
      if (seenEmails.has(safeEmail)) return;
      seenEmails.add(safeEmail);

      const option = document.createElement("option");
      option.value = agent.email;
      option.textContent = agent.name;

      // Select the current user by default so they see their own stats first
      if (safeEmail === currentUser.email.toLowerCase().trim()) {
        option.selected = true;
      }
      agentSelect.appendChild(option);
    });

    agentSelect.addEventListener("change", refreshData);
  }

  refreshData();
}

function paintStats(email, timeframe) {
  const stats = getAgentStats(email, timeframe);

  let timeLabel = "(Today)";
  if (timeframe === "week") timeLabel = "(Past 7 Days)";
  if (timeframe === "month") timeLabel = "(Past 30 Days)";
  if (timeframe === "all") timeLabel = "(All Time)";

  let salesSubtext = "";
  let touchesSubtext = "";
  let rateSubtext = "";

  if (email === "all" && stats.leaderboard && stats.leaderboard.length > 0) {
    const topSales = [...stats.leaderboard].sort(
      (a, b) => b.sales - a.sales,
    )[0];
    const topTouches = [...stats.leaderboard].sort(
      (a, b) => b.touches - a.touches,
    )[0];

    // 🛡️ THE TROPHY THRESHOLD: Require at least 5 touches to win the Close Rate award
    const eligibleForRate = stats.leaderboard.filter((a) => a.touches >= 5);
    const topRate =
      eligibleForRate.length > 0
        ? eligibleForRate.sort((a, b) => b.rate - a.rate)[0]
        : null;

    if (topSales && topSales.sales > 0) {
      salesSubtext = `<br><span style="font-size:12px; font-weight:600; color:#10b981;">🏆 Top: ${escHtml(topSales.name)} (${topSales.sales})</span>`;
    }
    if (topTouches && topTouches.touches > 0) {
      touchesSubtext = `<br><span style="font-size:12px; font-weight:600; color:#f59e0b;">🔥 Top: ${escHtml(topTouches.name)} (${topTouches.touches})</span>`;
    }
    if (topRate && topRate.rate > 0) {
      rateSubtext = `<br><span style="font-size:12px; font-weight:600; color:#8b5cf6;">🎯 Top: ${escHtml(topRate.name)} (${topRate.rate}%)</span>`;
    }
  } else if (
    email !== "all" &&
    timeframe === "day" &&
    stats.personalBestSales > 0
  ) {
    salesSubtext = `<br><span style="font-size:12px; font-weight:600; color:#8b5cf6;">⭐ Record: ${stats.personalBestSales} (${stats.personalBestDate})</span>`;
  }

  document.getElementById("label-sales").innerHTML =
    `Sales ${timeLabel} ${salesSubtext}`;
  document.getElementById("label-touches").innerHTML =
    `Leads Touched ${timeLabel} ${touchesSubtext}`;
  document.getElementById("label-conversion").innerHTML =
    `Close Rate ${timeLabel} ${rateSubtext}`;

  const animateValue = (id, endVal, isPercent = false) => {
    const obj = document.getElementById(id);
    if (!obj) return;

    if (
      endVal === undefined ||
      endVal === null ||
      (typeof endVal === "number" && isNaN(endVal))
    ) {
      endVal = 0;
    }

    let startTimestamp = null;
    const duration = 600;

    const step = (timestamp) => {
      if (!startTimestamp) startTimestamp = timestamp;
      const progress = Math.min((timestamp - startTimestamp) / duration, 1);

      if (isPercent) {
        const rawPercent = parseFloat(endVal) || 0;
        const currentPercent = (progress * rawPercent).toFixed(1);
        obj.textContent = currentPercent + "%";
      } else {
        const currentVal = Math.floor(progress * endVal);
        obj.textContent = currentVal;
      }

      if (progress < 1) window.requestAnimationFrame(step);
      else obj.textContent = endVal;
    };
    window.requestAnimationFrame(step);
  };

  animateValue("stat-sales-timeframe", stats.salesTimeframe);
  animateValue("stat-touches-timeframe", stats.uniqueLeadsTimeframe);
  animateValue("stat-sales-total", stats.salesTotal);
  animateValue("stat-conversion", stats.conversionRate, true);

  document.getElementById("stat-actions-timeframe").textContent =
    stats.touchesTimeframe;
  animateValue("stat-current-points", stats.currentPoints);

  if (timeframe === "day") {
    document.getElementById("stat-avg-daily").textContent = "--";
  } else {
    animateValue("stat-avg-daily", stats.avgDailyLeads);
  }

  // ==========================================
  // CHART GENERATION LOGIC
  // ==========================================
  const chartsContainer = document.getElementById("stats-charts-container");

  if (timeframe === "day" || stats.trends.labels.length <= 1) {
    chartsContainer.style.display = "none";
  } else {
    chartsContainer.style.display = "block";
    window._chartInstances = window._chartInstances || {};

    const drawChart = (
      canvasId,
      title,
      dataArray,
      colorStr,
      isPercent = false,
    ) => {
      if (window._chartInstances[canvasId]) {
        window._chartInstances[canvasId].destroy();
      }

      const canvasEl = document.getElementById(canvasId);
      if (!canvasEl) return;

      const ctx = canvasEl.getContext("2d");
      const gradient = ctx.createLinearGradient(0, 0, 0, 300);
      gradient.addColorStop(0, colorStr + "66");
      gradient.addColorStop(1, colorStr + "00");

      window._chartInstances[canvasId] = new Chart(ctx, {
        type: "line",
        data: {
          labels: stats.trends.labels,
          datasets: [
            {
              label: title,
              data: dataArray,
              borderColor: colorStr,
              backgroundColor: gradient,
              borderWidth: 3,
              fill: true,
              pointBackgroundColor: "#ffffff",
              pointBorderColor: colorStr,
              pointBorderWidth: 2,
              pointRadius: 4,
              pointHoverRadius: 6,
              tension: 0.4,
            },
          ],
        },
        options: {
          responsive: true,
          plugins: {
            legend: { display: false },
            title: { display: true, text: title },
          },
          scales: {
            y: {
              beginAtZero: true,
              ticks: {
                callback: function (value) {
                  return isPercent ? value + "%" : value;
                },
              },
            },
          },
        },
      });
    };

    drawChart("chart-sales", "Daily Sales", stats.trends.sales, "#10b981");
    drawChart(
      "chart-touches",
      "Daily Leads Touched",
      stats.trends.touches,
      "#f59e0b",
    );
    drawChart(
      "chart-conversion",
      "Daily Close Rate (%)",
      stats.trends.rates,
      "#8b5cf6",
      true,
    );
  }

  // ==========================================
  // LEADERBOARD INJECTION LOGIC
  // ==========================================
  let lbContainer = document.getElementById("stats-leaderboard-card");

  if (!lbContainer) {
    lbContainer = document.createElement("div");
    lbContainer.id = "stats-leaderboard-card";
    lbContainer.className = "card";
    lbContainer.style.marginTop = "24px";

    if (chartsContainer) {
      chartsContainer.parentNode.insertBefore(
        lbContainer,
        chartsContainer.nextSibling,
      );
    } else {
      document.getElementById("main-content").appendChild(lbContainer);
    }
  }

  if (email !== "all" || !stats.leaderboard || stats.leaderboard.length === 0) {
    lbContainer.style.display = "none";
  } else {
    lbContainer.style.display = "block";
    const topPerformers = stats.leaderboard.slice(0, 5);

    lbContainer.innerHTML = `
      <div class="card-header">
        <h2 class="card-title">Top Performers ${timeLabel}</h2>
      </div>
      <div class="table-wrap">
        <table class="data-table">
          <thead>
            <tr>
              <th style="width: 60px;">Rank</th>
              <th>Agent</th>
              <th>Sales</th>
              <th>Leads Touched</th>
              <th>Close Rate</th>
            </tr>
          </thead>
          <tbody>
            ${topPerformers
              .map((agent, index) => {
                const rankDisplay = index === 0 ? "🏆 1" : index + 1;
                return `
              <tr>
                <td><span class="top5-rank rank-${index + 1}">${rankDisplay}</span></td>
                <td style="font-weight: 600; color: #0D1B3E;">${escHtml(agent.name)}</td>
                <td style="color: #10b981; font-weight: bold;">${agent.sales}</td>
                <td class="td-mono">${agent.touches}</td>
                <td style="color: #8b5cf6; font-weight: 600;">${agent.rate}%</td>
              </tr>
            `;
              })
              .join("")}
          </tbody>
        </table>
      </div>
    `;
  }
}

function getAgentStats(targetEmail, timeframe = "day") {
  const target = targetEmail.toLowerCase().trim();
  const isAllAgents = target === "all";

  const now = Date.now();
  let startMs = 0;

  if (timeframe === "day") {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    startMs = today.getTime();
  } else if (timeframe === "week") {
    startMs = now - 7 * 24 * 60 * 60 * 1000;
  } else if (timeframe === "month") {
    startMs = now - 30 * 24 * 60 * 60 * 1000;
  }

  let salesTimeframe = 0;
  let salesTotal = 0;
  let touchesTimeframe = 0;
  let uniqueLeadsTouchedTimeframe = new Set();

  let uniqueSalesTotal = new Set();
  let uniqueSalesTimeframe = new Set();
  let seenDailyActions = new Set();

  let dailyMap = {};
  let agentMap = {};

  (State.activityLog || []).forEach((log) => {
    let cleanRaw = (log.agentEmail || log.agent || "").toLowerCase().trim();
    if (!cleanRaw) return;

    // 🛑 THE IDENTITY RESOLVER: Match raw logs to official contractors
    const aliasMap = {
      "everett henry": "h.gatlin@raimak.com",
      "julian torres": "j.torres@raimak.com",
      "tory mathis": "t.mathis@raimak.com",
      "stephanie balleste": "s.balleste@raimak.com",
      "brianna woodall": "b.woodall@raimak.com",
      // You can easily add any future rogue names right here!
    };

    if (aliasMap[cleanRaw]) {
      cleanRaw = aliasMap[cleanRaw];
    }

    let finalEmail = cleanRaw;
    let knownName = "";

    // ... (the rest of the matching logic stays exactly the same)

    const matchedContractor = (State.contractors || []).find(
      (c) =>
        (c.email || "").toLowerCase().trim() === cleanRaw ||
        (c.name || "").toLowerCase().trim() === cleanRaw,
    );

    if (matchedContractor) {
      finalEmail = (matchedContractor.email || finalEmail).toLowerCase().trim();
      knownName = matchedContractor.name;
    } else {
      // Fallback formatters if they somehow aren't in the contractor list
      if (!cleanRaw.includes("@")) {
        const parts = cleanRaw.split(/\s+/);
        if (parts.length >= 2)
          finalEmail = `${parts[0][0]}.${parts[parts.length - 1]}@raimak.com`;
        knownName = cleanRaw.replace(
          /\w\S*/g,
          (txt) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase(),
        );
      } else {
        const prefix = cleanRaw.split("@")[0].split(".");
        knownName = prefix
          .map((p) => p.charAt(0).toUpperCase() + p.slice(1))
          .join(" ");
      }
    }

    const logEmail = finalEmail;
    if (!isAllAgents && logEmail !== target) return;
    if (!log.action || !log.action.includes("Status:")) return;

    const logMs =
      typeof log.timestamp === "number"
        ? log.timestamp
        : Date.parse(log.timestamp);
    if (isNaN(logMs)) return;

    const isInTimeframe = logMs >= startMs;

    let dateKey = "";
    if (isInTimeframe) {
      const d = new Date(logMs);
      dateKey = `${d.getMonth() + 1}/${d.getDate()}`;
      if (!dailyMap[dateKey]) {
        dailyMap[dateKey] = {
          sales: 0,
          touches: 0,
          uniqueLeads: new Set(),
          timestamp: d.setHours(0, 0, 0, 0),
        };
      }

      if (isAllAgents && !agentMap[logEmail]) {
        agentMap[logEmail] = {
          name: knownName || log.agentName || logEmail,
          sales: 0,
          uniqueLeads: new Set(),
        };
      }
    }

    const trueLeadId = log.leadId || log.LeadId || log.LeadID;
    const fallbackName = log.leadName || log.Title || "Unknown";
    const rawKey = trueLeadId || fallbackName;
    const uniqueKey = String(rawKey).toLowerCase().trim();

    if (!uniqueKey || uniqueKey === "unknown") return;

    const dedupeKey = `${uniqueKey}-${log.action}-${dateKey}`;
    if (seenDailyActions.has(dedupeKey)) return;
    seenDailyActions.add(dedupeKey);

    if (isInTimeframe) {
      touchesTimeframe++;
      uniqueLeadsTouchedTimeframe.add(uniqueKey);
      if (dateKey) dailyMap[dateKey].uniqueLeads.add(uniqueKey);

      if (isAllAgents && agentMap[logEmail]) {
        agentMap[logEmail].uniqueLeads.add(uniqueKey);
      }
    }

    if (log.action.includes("Sold")) {
      if (!uniqueSalesTotal.has(uniqueKey)) {
        uniqueSalesTotal.add(uniqueKey);
        salesTotal++;
      }
      if (isInTimeframe && !uniqueSalesTimeframe.has(uniqueKey)) {
        uniqueSalesTimeframe.add(uniqueKey);
        salesTimeframe++;
        if (dateKey) dailyMap[dateKey].sales++;

        if (isAllAgents && agentMap[logEmail]) {
          agentMap[logEmail].sales++;
        }
      }
    }
  });

  const sortedDays = Object.values(dailyMap).sort(
    (a, b) => a.timestamp - b.timestamp,
  );
  const trends = { labels: [], sales: [], touches: [], rates: [] };

  sortedDays.forEach((day) => {
    const d = new Date(day.timestamp);
    trends.labels.push(`${d.getMonth() + 1}/${d.getDate()}`);
    trends.sales.push(day.sales);
    trends.touches.push(day.uniqueLeads.size);
    trends.rates.push(
      day.uniqueLeads.size > 0
        ? parseFloat(((day.sales / day.uniqueLeads.size) * 100).toFixed(1))
        : 0,
    );
  });

  const activeDays =
    Object.keys(dailyMap).length > 0 ? Object.keys(dailyMap).length : 1;
  let totalDailyUniqueTouches = 0;

  Object.values(dailyMap).forEach((day) => {
    totalDailyUniqueTouches += day.uniqueLeads.size;
  });

  const avgDailyLeads = Math.round(totalDailyUniqueTouches / activeDays);

  let currentPoints = 0;
  if (isAllAgents) {
    currentPoints = (State.agentScores || []).reduce(
      (sum, s) => sum + (s.CurrentPoints || 0),
      0,
    );
  } else {
    const scoreRow = (State.agentScores || []).find(
      (s) => (s.AgentEmail || "").toLowerCase().trim() === target,
    );
    currentPoints = scoreRow ? scoreRow.CurrentPoints : 0;
  }

  let leaderboard = [];
  if (isAllAgents) {
    leaderboard = Object.values(agentMap)
      .map((a) => {
        const touches = a.uniqueLeads.size;
        const rate = touches > 0 ? ((a.sales / touches) * 100).toFixed(1) : 0;
        return {
          name: a.name,
          sales: a.sales,
          touches: touches,
          rate: parseFloat(rate),
        };
      })
      .sort((a, b) => b.sales - a.sales || b.rate - a.rate);
  }

  let personalBestSales = 0;
  let personalBestDate = "";

  if (!isAllAgents && timeframe === "day") {
    const allTimeSalesMap = {};
    const recordTracker = new Set();

    (State.activityLog || []).forEach((log) => {
      let rawAgent = (log.agentEmail || log.agent || "").toLowerCase().trim();
      if (rawAgent && !rawAgent.includes("@")) {
        const nameParts = rawAgent.split(/\s+/);
        if (nameParts.length >= 2)
          rawAgent = `${nameParts[0][0]}.${nameParts[nameParts.length - 1]}@raimak.com`;
      }

      if (rawAgent !== target || !log.action || !log.action.includes("Sold"))
        return;

      const logMs =
        typeof log.timestamp === "number"
          ? log.timestamp
          : Date.parse(log.timestamp);
      if (isNaN(logMs)) return;

      const uniqueKey = String(
        log.leadId || log.LeadId || log.leadName || log.Title || "unknown",
      )
        .toLowerCase()
        .trim();
      if (!uniqueKey || uniqueKey === "unknown") return;

      const d = new Date(logMs);
      const dateKey = `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
      const dedupeKey = `${uniqueKey}-${dateKey}`;

      if (!recordTracker.has(dedupeKey)) {
        recordTracker.add(dedupeKey);
        allTimeSalesMap[dateKey] = (allTimeSalesMap[dateKey] || 0) + 1;

        if (allTimeSalesMap[dateKey] > personalBestSales) {
          personalBestSales = allTimeSalesMap[dateKey];
          personalBestDate = `${d.getMonth() + 1}/${d.getDate()}`;
        }
      }
    });
  }

  return {
    salesTimeframe,
    salesTotal,
    touchesTimeframe,
    uniqueLeadsTimeframe: uniqueLeadsTouchedTimeframe.size,
    conversionRate:
      uniqueLeadsTouchedTimeframe.size > 0
        ? ((salesTimeframe / uniqueLeadsTouchedTimeframe.size) * 100).toFixed(
            1,
          ) + "%"
        : "0%",
    avgDailyLeads,
    currentPoints,
    trends,
    leaderboard,
    personalBestSales,
    personalBestDate,
  };
}

// ============================================================
//  LEAD MODAL (Admin — Add/Edit)
// ============================================================
function openAddLeadModal() {
  if (!isAdmin()) return;
  State.editingLeadId = null;
  renderLeadModal(null);
}

function openEditLeadModal(id) {
  if (!isAdmin()) return;
  const lead = State.leads.find(function (l) {
    return l.id === id;
  });
  if (!lead) return;
  State.editingLeadId = id;
  renderLeadModal(lead);
}

function renderLeadModal(lead) {
  const isEdit = !!lead;
  const contractors = State.contractors.map((c) => c.name);
  const modalContainer = document.getElementById("modal");
  modalContainer.innerHTML = "";
  const template = document.getElementById("tmpl-lead-modal");
  const clone = template.content.cloneNode(true);

  // 2. Header & Button Logic
  clone.getElementById("modal-title").textContent = isEdit
    ? "Edit Lead"
    : "New Lead";
  const submitBtn = clone.getElementById("modal-submit-btn");
  submitBtn.textContent = isEdit ? "Save Changes" : "Add Lead";
  submitBtn.onclick = () => (isEdit ? submitEditLead() : submitAddLead());

  // 3. Populate Standard Text Inputs
  const safeVal = (val) => val || "";

  clone.getElementById("f-firstname").value = safeVal(lead?.firstName);
  clone.getElementById("f-lastname").value = safeVal(lead?.lastName);

  // Maps BTN securely, falling back to legacy phone data if needed
  clone.getElementById("f-btn").value = safeVal(
    lead?.BTN || lead?.btn || lead?.phone,
  );
  clone.getElementById("f-cbr").value = safeVal(lead?.cbr);

  clone.getElementById("f-address").value = safeVal(lead?.address);
  clone.getElementById("f-city").value = safeVal(lead?.city);
  clone.getElementById("f-state").value = safeVal(lead?.state);
  clone.getElementById("f-zip").value = safeVal(lead?.zip);
  clone.getElementById("f-mrc").value = safeVal(lead?.currentMRC);

  if (lead && lead.lastContacted) {
    clone.getElementById("f-lastcontacted").value =
      lead.lastContacted.split("T")[0];
  }

  // 4. Populate Dropdowns dynamically
  const leadTypeSelect = clone.getElementById("f-leadtype");
  Config.leadTypes.forEach((t) => {
    const opt = document.createElement("option");
    opt.value = t;
    opt.textContent = t;
    if (lead && lead.leadType === t) opt.selected = true;
    leadTypeSelect.appendChild(opt);
  });

  const statusSelect = clone.getElementById("f-status");
  Config.leadStatuses.forEach((s) => {
    const opt = document.createElement("option");
    opt.value = s;
    opt.textContent = s;
    if ((lead?.status || "New") === s) opt.selected = true;
    statusSelect.appendChild(opt);
  });

  const assignedSelect = clone.getElementById("f-assigned");
  contractors.forEach((c) => {
    const opt = document.createElement("option");
    opt.value = c;
    opt.textContent = c;
    if (lead && lead.assignedTo === c) opt.selected = true;
    assignedSelect.appendChild(opt);
  });

  const productsSelect = clone.getElementById("f-products");
  Config.currentProducts.forEach((p) => {
    const opt = document.createElement("option");
    opt.value = p;
    opt.textContent = p;
    if (lead && lead.currentProducts === p) opt.selected = true;
    productsSelect.appendChild(opt);
  });

  // 5. Build AutoPay Radios
  // 🛡️ THE FIX: Set text color to #cbd5e1 (slate-300) so it's readable on dark
  const autopayContainer = clone.getElementById("f-autopay-container");
  ["ACH - Debit Card", "ACH - Credit Card", "No Auto Pay"].forEach((opt) => {
    const isChecked = lead && lead.autoPay === opt ? "checked" : "";
    autopayContainer.innerHTML += `
      <label style="display:flex;align-items:center;gap:6px;font-size:12px;cursor:pointer;color:#cbd5e1;">
        <input type="radio" name="f-autopay" value="${opt}" ${isChecked} style="accent-color:#0ea5e9;width:13px;height:13px">
        ${opt}
      </label>`;
  });

  // 6. Notes History Parser
  // 🛡️ THE FIX: Replaced light mode backgrounds with translucent dark slate
  const notesHistory = clone.getElementById("modal-notes-history");
  if (lead && lead.notes && lead.notes.trim()) {
    const notesHtml = lead.notes
      .split("\n")
      .filter((l) => l.trim())
      .map((line) => {
        const match = line.match(/^\[(\d{2}\/\d{2}(?:\/\d{2})?)(.*?)\]\s*(.*)/);
        if (match) {
          const date = match[1];
          const agent = match[2] ? match[2].replace(/^\s*-\s*/, "") : "";
          const text = match[3];
          return `
            <div style="margin-bottom:6px;padding-bottom:6px;border-bottom:1px solid rgba(255,255,255,0.05)">
              <div style="display:flex;gap:6px;align-items:center;margin-bottom:2px">
                <span style="font-family:var(--font-mono);font-size:10px;color:#38bdf8;font-weight:700;background:rgba(56, 189, 248, 0.1);padding:2px 4px;border-radius:4px">${date}</span>
                ${agent ? `<span style="font-family:var(--font-mono);font-size:10px;color:#94a3b8;font-weight:600;">${escHtml(agent)}</span>` : ""}
              </div>
              <span style="font-size:12px;color:#e2e8f0;line-height:1.3;">${escHtml(text)}</span>
            </div>`;
        }
        return `
          <div style="margin-bottom:6px;padding-bottom:6px;border-bottom:1px solid rgba(255,255,255,0.05)">
            <div style="margin-bottom:2px">
              <span style="font-family:var(--font-mono);font-size:10px;color:#94a3b8;background:rgba(255,255,255,0.05);padding:2px 4px;border-radius:4px">Legacy note</span>
            </div>
            <span style="font-size:12px;color:#cbd5e1;line-height:1.3;">${escHtml(line)}</span>
          </div>`;
      })
      .join("");

    notesHistory.innerHTML = `<div style="background:rgba(0,0,0,0.2);border:1px solid rgba(255,255,255,0.1);border-radius:6px;padding:8px 10px;margin-bottom:8px;max-height:140px;overflow-y:auto">${notesHtml}</div>`;
  } else {
    notesHistory.innerHTML = `<div style="font-size:12px;color:#64748b;margin-bottom:8px;font-style:italic;">No notes yet.</div>`;
  }

  // 7. Mount & Display
  modalContainer.appendChild(clone);
  document.getElementById("modal-overlay").style.display = "flex";
}

async function submitAddLead() {
  const fields = collectLeadForm();
  if (!fields) return;
  const agentName = fields._agentName;
  delete fields._agentName;
  setLoading(true);
  try {
    const newLead = await Graph.addLead(fields);
    if (agentName) await Graph.assignAgent(newLead.id, agentName);
    await Graph.logActivity({
      LeadID: newLead.id,
      Title: fields.Title,
      ActionType: "Lead Created",
      AgentEmail: (State.currentUser && State.currentUser.email) || "",
    });
    await refreshData();
    closeModal();
    UI.showToast("Lead added!", "success");
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

async function submitEditLead() {
  const fields = collectLeadForm();
  if (!fields) return;
  const agentName = fields._agentName;
  delete fields._agentName;
  setLoading(true);
  try {
    await Graph.updateLead(State.editingLeadId, fields);
    if (agentName) await Graph.assignAgent(State.editingLeadId, agentName);
    await Graph.logActivity({
      LeadID: State.editingLeadId,
      Title: fields.Title,
      ActionType: "Lead Updated",
      AgentEmail: (State.currentUser && State.currentUser.email) || "",
    });
    await refreshData();
    closeModal();
    UI.showToast("Lead updated!", "success");
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}

function collectLeadForm() {
  const firstName = (
    (document.getElementById("f-firstname") || {}).value || ""
  ).trim();
  const lastName = (
    (document.getElementById("f-lastname") || {}).value || ""
  ).trim();
  const agentName = (document.getElementById("f-assigned") || {}).value || "";
  const nameEl = document.getElementById("f-name");
  const fullName = nameEl
    ? (nameEl.value || "").trim()
    : (firstName + " " + lastName).trim();

  if (!firstName && !lastName && !fullName) {
    UI.showToast("Name is required.", "error");
    return null;
  }

  const fields = { _agentName: agentName };
  if (firstName) fields["FirstName"] = firstName;
  if (lastName) fields["LastName"] = lastName;
  if (fullName && !firstName) fields["Title"] = fullName;

  const add = function (key, elId, trim) {
    const el = document.getElementById(elId);
    const val = el ? (trim ? (el.value || "").trim() : el.value || "") : "";
    if (val) fields[key] = val;
  };

  add("Lead_x0020_Type", "f-leadtype");
  add("Email", "f-email", true);
  add("Phone", "f-phone", true);
  add("Status", "f-status");
  add("LastTouchedOn", "f-lastcontacted");
  add("MonthlyRecurringCharge_x0028_MRC", "f-mrc", true);
  add("CurrentProducts", "f-products");
  add("CBR", "f-cbr", true);
  add("BTN", "f-btn", true);
  add("WorkAddress", "f-address", true);
  add("WorkCity", "f-city", true);
  add("State", "f-state", true);
  add("Zip", "f-zip", true);
  add("AutoPay", "f-autopay");

  const notesEl = document.getElementById("f-notes");
  if (notesEl && notesEl.value.trim()) {
    const today = new Date();
    const dateStamp =
      (today.getMonth() + 1).toString().padStart(2, "0") +
      "/" +
      today.getDate().toString().padStart(2, "0") +
      "/" +
      String(today.getFullYear()).slice(-2);
    const adminName =
      State.currentUser && State.currentUser.name
        ? " - " + State.currentUser.name
        : "";
    const lead = State.editingLeadId
      ? State.leads.find(function (l) {
          return l.id === State.editingLeadId;
        })
      : null;
    const existing = (lead && lead.notes) || "";
    const stamped = "[" + dateStamp + adminName + "] " + notesEl.value.trim();
    fields["Notes"] = existing ? stamped + "\n" + existing : stamped;
  }

  if (!fields.Status) fields.Status = "New";
  return fields;
}

/*async function deleteLead(id) {
  const lead = State.leads.find(function (l) {
    return l.id === id;
  });
  if (!confirm('Delete "' + (lead && lead.name) + '"? This cannot be undone.'))
    return;
  setLoading(true);
  try {
    await Graph.deleteLead(id);
    await refreshData();
    UI.showToast("Lead deleted.", "success");
  } catch (err) {
    UI.showToast("Failed: " + err.message, "error");
  } finally {
    setLoading(false);
  }
}*/

function closeModal(event) {
  if (event && event.target !== document.getElementById("modal-overlay"))
    return;
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
  const o = document.getElementById("loading-overlay");
  if (o) o.style.display = on ? "flex" : "none";
}

function updateBadges() {
  const n = State.leads.filter(function (l) {
    return (
      l.flags &&
      (l.flags.includes("needs_recycle") ||
        l.flags.includes("agent_overloaded"))
    );
  }).length;
  const b = document.getElementById("badge-leads");
  if (b) {
    b.textContent = n > 0 ? n : "";
    b.style.display = n > 0 ? "inline-flex" : "none";
  }
}

async function exportD2DLeads() {
  const groupA_Unassigned = [];
  const groupB_Assigned = [];

  (State.leads || []).forEach((l) => {
    // 🛑 Skip leads that are already terminal (using your global Config) or ALREADY exported & flagged
    if (l.Status && (Config.terminalStatuses || []).includes(l.Status)) return;
    if (l.flaggedForExport === true || l.FlaggedForExport === true) return;

    let count = 0;
    if (Array.isArray(l.previousAgents)) {
      count = l.previousAgents.length;
    } else if (typeof l.previousAgents === "string") {
      count = l.previousAgents
        .split(",")
        .map((a) => a.trim())
        .filter((a) => a !== "").length;
    } else {
      count = parseInt(l.previousAgents) || 0;
    }

    const isUnassigned = !l.assignedTo || l.assignedTo.trim() === "";

    // 🎯 Group 1: Unassigned & 3+ touches
    if (isUnassigned && count >= 3) {
      groupA_Unassigned.push(l);
    }
    // 🎯 Group 2: Assigned & 5+ touches
    else if (!isUnassigned && count >= 5) {
      groupB_Assigned.push(l);
    }
  });

  const allExportLeads = [...groupA_Unassigned, ...groupB_Assigned];

  if (allExportLeads.length === 0) {
    UI.showToast("No leads meet the criteria for D2D export.", "warning");
    return;
  }

  // --- 📝 CSV GENERATION ---
  const headers = [
    "First Name",
    "Last Name",
    "Address",
    "City",
    "State",
    "BTN",
    "CBR",
    "currentMRC",
    "currentProducts",
  ];

  const rows = allExportLeads.map((l) => {
    const firstName = l.firstName || "";
    const lastName = l.lastName || "";
    const address = l.address || "";
    const city = l.city || "";
    const state = l.state || "";
    const btn = l.BTN || l.btn || l.phone || "";
    const cbr = l.CBR || l.cbr || l.altPhone || "";
    const mrc = l.currentMRC || l.mrc || "";
    const products = l.currentProducts || l.products || "";

    return `"${firstName}","${lastName}","${address}","${city}","${state}","${btn}","${cbr}","${mrc}","${products}"`;
  });

  const csvContent = headers.join(",") + "\n" + rows.join("\n");
  const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.setAttribute("href", url);
  link.setAttribute(
    "download",
    `D2D_Export_${new Date().toLocaleDateString().replace(/\//g, "-")}.csv`,
  );
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);

  UI.showToast(
    `📁 Exported ${allExportLeads.length} leads! (Unassigned: ${groupA_Unassigned.length}, Flagged: ${groupB_Assigned.length})`,
    "info",
  );

  // --- 🔄 SHAREPOINT BATCH UPDATES ---
  const host = Config.sharePoint.hostname;
  const sitePath = Config.sharePoint.sites.team;
  const listId = Config.sharePoint.lists.leadsList;
  const batchUrl = `${Config.sharePoint.graphBase}/$batch`;

  const batchSize = 20;
  let updateCount = 0;

  for (let i = 0; i < allExportLeads.length; i += batchSize) {
    const chunk = allExportLeads.slice(i, i + batchSize);

    const batchRequests = chunk.map((lead, index) => {
      // Intelligently route the update payload based on which group the lead belongs to
      const isGroupA = groupA_Unassigned.includes(lead);
      const updateFields = isGroupA
        ? { Status: "D2D Lead" }
        : { FlaggedForExport: true }; // Flags the assigned leads silently

      return {
        id: String(index + 1),
        method: "PATCH",
        url: `/sites/${host}:/${sitePath}:/lists/${listId}/items/${lead.id}`,
        headers: { "Content-Type": "application/json", "If-Match": "*" },
        body: {
          fields: updateFields,
        },
      };
    });

    try {
      await Graph.apiFetch(batchUrl, {
        method: "POST",
        body: { requests: batchRequests },
      });
      updateCount += chunk.length;
    } catch (error) {
      console.error("❌ Batch update failed:", error);
    }
  }

  UI.showToast(
    `✅ Successfully updated ${updateCount} leads in SharePoint!`,
    "success",
  );
  await loadAllData();
}

function updateLeadDraft(leadId, fieldName, value) {
  // If this lead doesn't have a draft object yet, create one
  if (!State.drafts[leadId]) {
    State.drafts[leadId] = {};
  }
  // Save the keystroke
  State.drafts[leadId][fieldName] = value;
}
function flagLabel(f) {
  return (
    {
      cool_off: "Cool-off",
      needs_recycle: "Recycle",
      agent_overloaded: "Overloaded",
    }[f] || f
  );
}
function formatDate(d) {
  if (!d) return "";
  return new Date(d).toLocaleDateString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
  });
}
function formatTime(d) {
  if (!d) return "";
  return new Date(d).toLocaleTimeString("en-GB", {
    hour: "2-digit",
    minute: "2-digit",
  });
}
function formatDateTime(d) {
  if (!d) return "";
  return new Date(d).toLocaleString("en-GB", {
    day: "2-digit",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}
function escHtml(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function startSuspensionCountdown(expirationDate) {
  const end = expirationDate.getTime();

  // Run immediately so there's no 1-second lag, then loop
  const timer = setInterval(() => {
    const now = new Date().getTime();
    const distance = end - now;

    // If the countdown hits zero, refresh the page to let them in!
    if (distance < 0) {
      clearInterval(timer);
      window.location.reload();
      return;
    }

    // Math magic
    const days = Math.floor(distance / (1000 * 60 * 60 * 24));
    const hours = Math.floor(
      (distance % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60),
    );
    const minutes = Math.floor((distance % (1000 * 60 * 60)) / (1000 * 60));
    const seconds = Math.floor((distance % (1000 * 60)) / 1000);

    // Update the DOM safely (adding leading zeros for style)
    const elDays = document.getElementById("cd-days");
    const elHours = document.getElementById("cd-hours");
    const elMins = document.getElementById("cd-mins");
    const elSecs = document.getElementById("cd-secs");

    if (elDays) elDays.innerText = String(days).padStart(2, "0");
    if (elHours) elHours.innerText = String(hours).padStart(2, "0");
    if (elMins) elMins.innerText = String(minutes).padStart(2, "0");
    if (elSecs) elSecs.innerText = String(seconds).padStart(2, "0");
  }, 1000);
}

const UI = {
  showToast: function (msg, type) {
    type = type || "info";

    let c = document.getElementById("toast-container");

    if (!c) {
      c = document.createElement("div");
      c.id = "toast-container";

      // THE FIX: Nuke the z-index to a billion, and ensure absolute highest priority
      c.style.cssText =
        "position: fixed !important; bottom: 20px !important; right: 20px !important; z-index: 2147483647 !important; display: flex !important; flex-direction: column !important; gap: 10px !important; pointer-events: none !important;";

      // THE FIX: Attach it directly to the HTML document element, bypassing the body entirely
      document.documentElement.appendChild(c);
    }

    const t = document.createElement("div");
    t.className = "toast toast-" + type;
    t.textContent = msg;
    t.style.pointerEvents = "auto";

    c.appendChild(t);

    setTimeout(function () {
      t.classList.add("show");
    }, 10);

    setTimeout(function () {
      t.classList.remove("show");
      setTimeout(function () {
        t.remove();
      }, 300);
    }, 4000);
  },
  showConfetti: function () {
    const el = document.createElement("div");
    el.className = "confetti-burst";
    el.innerHTML = "&#127881; SOLD! &#127881;";
    document.body.appendChild(el);
    setTimeout(function () {
      el.remove();
    }, 2600);
  },
};

const Ticker = {
  update: function () {
    const tickerEl = document.getElementById("sales-ticker-content");
    if (!tickerEl) return;

    // Grab today's date string (e.g., "Fri Apr 10 2026")
    const todayStr = new Date().toDateString();

    // 1. Find all leads sold TODAY
    const soldToday = State.leads.filter((l) => {
      const isSold = l.status && l.status.toLowerCase().includes("sold");
      // Check if the lead was last updated today
      const isFromToday =
        l.lastContacted &&
        new Date(l.lastContacted).toDateString() === todayStr;

      return isSold && isFromToday;
    });

    // 2. Calculate Top 5 Agents (Today Only)
    const agentSales = {};
    soldToday.forEach((l) => {
      const seller = l.soldBy || l.assignedTo;
      if (seller) {
        agentSales[seller] = (agentSales[seller] || 0) + 1;
      }
    });

    const topAgents = Object.entries(agentSales)
      .sort((a, b) => b[1] - a[1]) // Sort highest to lowest
      .slice(0, 5) // Grab top 5
      .map(
        (entry, i) =>
          `<strong>#${i + 1} ${escHtml(entry[0])}</strong> (${entry[1]})`,
      );

    // 3. Grab the 5 most recent sales (Today Only)
    const recentSales = soldToday
      .slice(-5)
      .reverse()
      .map((l) => {
        const soldBy = l.soldBy || l.assignedTo || "Someone";
        const forAgent = l.assignedTo;

        if (forAgent && soldBy !== forAgent) {
          return `🎉 <strong>${escHtml(soldBy)}</strong> just made a sale for <strong>${escHtml(forAgent)}</strong> — ${escHtml(l.name)}`;
        } else {
          return `🎉 <strong>${escHtml(soldBy)}</strong> just closed a sale! — ${escHtml(l.name)}`;
        }
      });

    // 4. Build the string with new "TODAY" labels
    let textParts = [];
    if (recentSales.length > 0) {
      textParts.push(`🔥 TODAY'S RECENT: ${recentSales.join("  •  ")}`);
    }
    if (topAgents.length > 0) {
      textParts.push(`🏆 TODAY'S LEADERS: ${topAgents.join("  •  ")}`);
    }

    // 5. Inject it
    tickerEl.innerHTML =
      textParts.length > 0
        ? textParts.join("  |  ")
        : "🚀 Let's make some sales today!";
  },
};
