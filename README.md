# Raimak Lead Management System

A static web app that manages leads stored in SharePoint via Microsoft Graph API.

## Setup

### 1. Register in Azure AD
1. Go to [portal.azure.com](https://portal.azure.com)
2. Azure Active Directory → App Registrations → **New Registration**
3. Name: Raimak LMS
4. Redirect URI: https://richardzacker.github.io/raimak-lms/
5. Copy the **Client ID** and **Tenant ID**

### 2. Configure the App
Edit js/config.js and replace:
- YOUR_AZURE_APP_CLIENT_ID → your Client ID
- YOUR_TENANT_ID → your Tenant ID

### 3. Grant Permissions (Azure Portal)
In your App Registration → API Permissions:
- Add Sites.ReadWrite.All (Microsoft Graph, Delegated)
- Grant admin consent

### 4. Deploy
Push to the main branch — GitHub Pages deploys automatically.

## Business Rules (from AppSettings.csv)
| Rule | Value |
|------|-------|
| Cool-off between contacts | 2 days |
| Max leads per agent | 15 |
| Recycle after inactivity | 30 days |

## SharePoint Lists
| List | Purpose |
|------|---------|
| 5a01419d-... | Leads |
| d5df38a-... | Contractors |
| 2adb1260-... | Activity Log |
