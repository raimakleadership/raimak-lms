// ============================================================
//  Raimak LMS — App Configuration
//  Values sourced from AppSettings.csv + .iqy SharePoint files
// ============================================================

const Config = {

  // --- Azure AD / MSAL ---
  // Replace these after registering your app in portal.azure.com
  azure: {
    clientId:   ad1b153f-8b6a-4f1c-8ab2-58fcf03cf5c2,   // App Registration → Application (client) ID
    tenantId:   39e14190-0b23-4ecd-99f9-606ad1215881,              // Azure AD → Overview → Tenant ID
    redirectUri: window.location.origin + window.location.pathname,
  },

  // --- SharePoint ---
  sharePoint: {
    hostname: "raimak.sharepoint.com",

    // Sites
    sites: {
      leadship: "sites/RaimakLeadship",
      team:     "TeamSite",
    },

    // List GUIDs (from .iqy files — do not change)
    lists: {
      activityLog:    "2adb1260-e635-45cd-bb3b-87dd57a2d022",  // Activity_Log.iqy
      contractorList: "bd5df38a-9cb6-411d-87e8-3e79934213d3",  // Contractor_List.iqy
      leadsList:      "5a01419d-e2c9-4aad-8484-6ed97233f305",  // Leads_List.iqy
    },

    // Graph API base
    graphBase: "https://graph.microsoft.com/v1.0",
  },

  // --- Business Rules (from AppSettings.csv) ---
  rules: {
    coolOffDays:       2,   // Minimum days before re-contacting a lead
    maxLeadsPerAgent:  15,  // Max active leads assigned per agent at once
    recycleAfterDays:  30,  // Days of inactivity before lead is recycled
    appVersion:        "1.0",
  },

  // --- Lead Statuses ---
  leadStatuses: ["New", "Contacted", "Qualified", "Proposal Sent", "Negotiating", "Won", "Lost", "Recycled"],

  // --- Lead Sources ---
  leadSources: ["Web Form", "Referral", "Cold Call", "Email Campaign", "Social Media", "Trade Show", "Other"],

  // --- Microsoft Graph Scopes ---
  scopes: [
    "Sites.ReadWrite.All",
    "User.Read",
  ],
};
