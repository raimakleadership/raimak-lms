// ============================================================
//  Raimak LMS — Graph API / SharePoint Data Layer
// ============================================================

const Graph = (() => {

  const base = Config.sharePoint.graphBase;
  const host  = Config.sharePoint.hostname;
  const lists  = Config.sharePoint.lists;

  // ── Site IDs (resolved at runtime) ───────────────────────
  let siteIds = { leadship: null, team: null };

  // ── Generic Fetch Helper ──────────────────────────────────
  async function apiFetch(url, method = "GET", body = null) {
    const token = await Auth.getToken();
    if (!token) throw new Error("Not authenticated");

    const opts = {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        Prefer: "HidePersonalData=false",
      },
    };
    if (body) opts.body = JSON.stringify(body);

    const res = await fetch(url, opts);
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(err?.error?.message || `HTTP ${res.status}`);
    }
    if (res.status === 204) return null;
    return res.json();
  }

  // ── Resolve Site IDs ──────────────────────────────────────
  async function resolveSiteIds() {
    if (siteIds.leadship && siteIds.team) return;
    const [s1, s2] = await Promise.all([
      apiFetch(`${base}/sites/${host}:/${Config.sharePoint.sites.leadship}`),
      apiFetch(`${base}/sites/${host}:/${Config.sharePoint.sites.team}`),
    ]);
    siteIds.leadship = s1.id;
    siteIds.team     = s2.id;
  }

  // ── Paginate Through All Items ─────────────────────────────
  async function getAllItems(url) {
    let items = [];
    let next  = url;
    while (next) {
      const data = await apiFetch(next);
      items = items.concat(data.value || []);
      next  = data["@odata.nextLink"] || null;
    }
    return items;
  }

  // ============================================================
  //  LEADS
  // ============================================================

  async function getLeads() {
    await resolveSiteIds();
    const url = `${base}/sites/${siteIds.team}/lists/${lists.leadsList}/items?expand=fields&$top=500`;
    const raw = await getAllItems(url);
    return raw.map(normalizeLeadItem);
  }

  async function addLead(fields) {
    await resolveSiteIds();
    const url = `${base}/sites/${siteIds.team}/lists/${lists.leadsList}/items`;
    const res = await apiFetch(url, "POST", { fields });
    return normalizeLeadItem(res);
  }

  async function updateLead(itemId, fields) {
    await resolveSiteIds();
    const url = `${base}/sites/${siteIds.team}/lists/${lists.leadsList}/items/${itemId}/fields`;
    await apiFetch(url, "PATCH", fields);
  }

  async function deleteLead(itemId) {
    await resolveSiteIds();
    const url = `${base}/sites/${siteIds.team}/lists/${lists.leadsList}/items/${itemId}`;
    await apiFetch(url, "DELETE");
  }

  // Normalise raw SharePoint item → clean lead object
  function normalizeLeadItem(item) {
    const f = item.fields || {};
    return {
      id:           item.id,
      name:         f.Title        || f.LeadName    || "",
      company:      f.Company      || "",
      email:        f.Email        || f.EmailAddress || "",
      phone:        f.Phone        || f.PhoneNumber  || "",
      status:       f.Status       || "New",
      source:       f.LeadSource   || f.Source       || "",
      assignedTo:   f.AssignedTo   || f.Agent        || "",
      notes:        f.Notes        || "",
      value:        f.DealValue    || f.Value        || "",
      lastContacted:f.LastContacted || null,
      createdAt:    item.createdDateTime || f.Created || null,
      modified:     item.lastModifiedDateTime || null,
    };
  }

  // ============================================================
  //  CONTRACTORS / AGENTS
  // ============================================================

  async function getContractors() {
    await resolveSiteIds();
    const url = `${base}/sites/${siteIds.team}/lists/${lists.contractorList}/items?expand=fields&$top=500`;
    const raw = await getAllItems(url);
    return raw.map(item => {
      const f = item.fields || {};
      return {
        id:       item.id,
        name:     f.Title        || f.ContractorName || "",
        email:    f.Email        || "",
        phone:    f.Phone        || "",
        role:     f.Role         || "Agent",
        active:   f.Active !== undefined ? f.Active : true,
      };
    });
  }

  // ============================================================
  //  ACTIVITY LOG
  // ============================================================

  async function getActivityLog(limit = 200) {
    await resolveSiteIds();
    const url = `${base}/sites/${siteIds.leadship}/lists/${lists.activityLog}/items?expand=fields&$orderby=createdDateTime desc&$top=${limit}`;
    const raw = await getAllItems(url);
    return raw.map(item => {
      const f = item.fields || {};
      return {
        id:        item.id,
        leadId:    f.LeadId     || f.LeadID    || "",
        leadName:  f.LeadName   || f.Title     || "",
        action:    f.Action     || f.Activity  || "",
        agent:     f.Agent      || f.AssignedTo || "",
        notes:     f.Notes      || "",
        timestamp: item.createdDateTime || f.Created || null,
      };
    });
  }

  async function logActivity(entry) {
    await resolveSiteIds();
    const url = `${base}/sites/${siteIds.leadship}/lists/${lists.activityLog}/items`;
    await apiFetch(url, "POST", { fields: entry });
  }

  // ============================================================
  //  BUSINESS RULES
  // ============================================================

  function applyBusinessRules(leads, contractors) {
    const now = new Date();
    const { coolOffDays, maxLeadsPerAgent, recycleAfterDays } = Config.rules;

    // Count leads per agent
    const agentCounts = {};
    for (const lead of leads) {
      if (lead.assignedTo && !["Won","Lost","Recycled"].includes(lead.status)) {
        agentCounts[lead.assignedTo] = (agentCounts[lead.assignedTo] || 0) + 1;
      }
    }

    return leads.map(lead => {
      const flags = [];

      // Cool-off check
      if (lead.lastContacted) {
        const last   = new Date(lead.lastContacted);
        const daysSince = (now - last) / 86400000;
        if (daysSince < coolOffDays && !["Won","Lost"].includes(lead.status)) {
          flags.push(`cool_off`);
        }
      }

      // Recycle check
      const ref = lead.lastContacted || lead.createdAt;
      if (ref && !["Won","Lost","Recycled"].includes(lead.status)) {
        const daysSince = (now - new Date(ref)) / 86400000;
        if (daysSince > recycleAfterDays) flags.push("needs_recycle");
      }

      // Over-assigned agent
      if (lead.assignedTo && agentCounts[lead.assignedTo] > maxLeadsPerAgent) {
        flags.push("agent_overloaded");
      }

      return { ...lead, flags, agentLeadCount: agentCounts[lead.assignedTo] || 0 };
    });
  }

  function canAgentTakeLead(agentName, leads) {
    const count = leads.filter(
      l => l.assignedTo === agentName && !["Won","Lost","Recycled"].includes(l.status)
    ).length;
    return count < Config.rules.maxLeadsPerAgent;
  }

  function isInCoolOff(lead) {
    if (!lead.lastContacted) return false;
    const daysSince = (new Date() - new Date(lead.lastContacted)) / 86400000;
    return daysSince < Config.rules.coolOffDays;
  }

  return {
    getLeads, addLead, updateLead, deleteLead,
    getContractors,
    getActivityLog, logActivity,
    applyBusinessRules, canAgentTakeLead, isInCoolOff,
  };
})();
