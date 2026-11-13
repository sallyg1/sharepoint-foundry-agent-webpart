# SharePoint Online (SPFx) Web Part for Azure AI Foundry Agent

A SharePoint Framework (SPFx v1.24) React client-side web part that surfaces an **Azure AI Foundry Prompt Agent** directly on SharePoint Online pages across **cross-tenant and cross-cloud sovereign boundaries**.

---

## 🎯 Target Scenario: Cross-Tenant / Cross-Cloud (GCC & Azure Government)

This solution addresses the common public sector / enterprise pattern where:
- **SharePoint Online** lives in **Tenant A** (e.g. Microsoft 365 GCC Moderate, backed by commercial Microsoft Entra ID).
- **Azure AI Foundry Agent** lives in **Tenant B** (e.g. Azure Government, backed by Microsoft Entra US Gov).

### Why this pattern is needed
Direct out-of-the-box publishing from Azure AI Foundry or Copilot Studio to SharePoint Online **cannot cross sovereign cloud or tenant identity boundaries** (commercial Entra vs Entra Gov). 

To solve this without directory peering or B2B guest account management, this client-side web part pairs with a middle-tier **zero-trust authentication broker**:

> 🔗 **Backend Broker Repository**:  
> **[sallyg1/crosstenant-auth-bridge](https://github.com/sallyg1/crosstenant-auth-bridge)**  
> *(Reference pattern for cross-tenant, cross-cloud access to Azure AI Foundry agents: validates an Entra token in code against public JWKS, and calls Foundry using a Managed Identity).*

---

## 🏛️ End-to-End Architecture

```text
[Tenant A: M365 GCC Moderate / Commercial]
       │
       ▼ (SharePoint Online Page)
[SPFx React Web Part]
       │
       ▼ (AadHttpClientFactory)
[Tenant A Entra ID] ──> Issues Bearer Token for exposed API scope (access_as_user)
       │
       │   HTTPS POST with Bearer Token (CORS allowed for SharePoint origin)
       ▼
─────────────────────────────────────────────────────────────────────────────
[Cloud / Tenant Boundary]
─────────────────────────────────────────────────────────────────────────────
       ▼
[Tenant B: Azure Government / Commercial]
[Azure Function Authentication Broker (sallyg1/crosstenant-auth-bridge)]
       │
       ├── 1. Validates Tenant A JWT in C# code using Tenant A's public JWKS
       │      (No Easy Auth / No directory peering required)
       ├── 2. Extracts caller identity claims (oid, upn, name)
       │
       ▼ 3. Calls Azure AI Foundry using Tenant B Managed Identity
[Azure AI Foundry Project]
       │
       ▼
[Foundry Prompt Agent with File Search & Knowledge]
```

---

## 🔗 Related Repositories
- **Front-end Web Part (This repo)**: [sharepoint-foundry-agent-webpart](https://github.com/sallyg1/sharepoint-foundry-agent-webpart)
- **Back-end Auth Broker (Azure Function)**: [crosstenant-auth-bridge](https://github.com/sallyg1/crosstenant-auth-bridge)

---

## 📋 Prerequisites

- **Node.js**: v18.x or v20.x
- **SharePoint Online Tenant** (Commercial or GCC) with App Catalog access and SharePoint Administrator rights.
- **Entra ID App Registration** exposing an API scope (`access_as_user`).
- **Azure Function App** running the backend authentication broker and configured with CORS.
- **Azure AI Foundry Prompt Agent** deployed in Azure.

---

## ⚙️ Configuration Guide

Before building and deploying, configure your environment identifiers:

### 1. Entra ID App Registration
1. In the Microsoft Entra admin center for your M365 tenant, create an App Registration (e.g. `Dallas-Foundry-bridge`).
2. Navigate to **Expose an API**:
   - Set the Application ID URI: `api://<YOUR_API_CLIENT_ID>`
   - Add a scope: `access_as_user` (Who can consent: Admins and users).

### 2. Configure Solution Manifest (`config/package-solution.json`)
Open `config/package-solution.json` and set `resource` to your exact Entra App Registration **display name**:

```json
"webApiPermissionRequests": [
  {
    "resource": "<YOUR_ENTRA_APP_REGISTRATION_NAME>",
    "scope": "access_as_user"
  }
]
```

### 3. Configure Web Part Defaults (`DallasPermitChatWebPart.ts`)
Open `src/webparts/dallasPermitChat/DallasPermitChatWebPart.ts` and update the fallback defaults (or configure them dynamically via the web part property pane on the page):

```typescript
apiAppIdUri: this.properties.apiAppIdUri || 'api://<YOUR_ENTRA_API_APP_CLIENT_ID>',
functionEndpoint: this.properties.functionEndpoint || 'https://<YOUR_FUNCTION_APP_NAME>.azurewebsites.net/api/chat',
```

---

## 🚀 Build and Deployment

### Install Dependencies
```bash
npm install
```

### Build & Package (.sppkg)
```bash
npm run build
```
*(Runs `heft test --clean --production && heft package-solution --production`).*

The compiled production package will be created at:
`sharepoint/solution/dallas-permit-chat.sppkg`

### Deploy to SharePoint App Catalog
1. Go to your **SharePoint Admin Center** (`https://<tenant>-admin.sharepoint.com`).
2. Navigate to **More features** → **Apps** → **Open App Catalog** (or `https://<tenant>.sharepoint.com/sites/appcatalog/AppCatalog`).
3. Upload `dallas-permit-chat.sppkg`.
4. Check **Enable this app and add it to all sites** (or specific site collections) and click **Enable app**.

### Approve API Access
1. In the **SharePoint Admin Center**, go to **Advanced** → **API access**.
2. Locate the pending request for your app registration (e.g. `Dallas-Foundry-bridge` / `access_as_user`).
3. Click **Approve**.

### Configure Backend Function CORS
In the Azure Portal for your Azure Function App:
1. Navigate to **API** → **CORS**.
2. Add your SharePoint origin (e.g. `https://<tenant>.sharepoint.com` without a trailing slash).
3. Ensure **Enable Access-Control-Allow-Credentials** is **unchecked**.
4. Click **Save**.

---

## 🧪 Testing on SharePoint Page

1. Open your target SharePoint site (e.g. `https://<tenant>.sharepoint.com/sites/HR`).
2. Create a new page or edit an existing page.
3. Add the **DallasPermitChat** web part (or use the full-page app template).
4. Click the **Pencil icon** on the web part to verify properties:
   - **Entra API App ID URI**: `api://<YOUR_ENTRA_API_APP_CLIENT_ID>` (or raw GUID).
   - **Azure Function Chat Endpoint**: `https://<YOUR_FUNCTION_APP_NAME>.azurewebsites.net/api/chat`.
5. Save/Publish the page and submit a prompt.

---

## 🛠️ Tech Stack
- **Framework**: SharePoint Framework (SPFx) v1.24.1
- **UI Library**: React 18 & TypeScript
- **HTTP Client**: `@microsoft/sp-http` (`AadHttpClient`)
- **Build Tooling**: `@rushstack/heft` and Webpack 5