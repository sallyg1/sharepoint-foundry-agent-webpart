# SharePoint Online (SPFx) Web Part for Azure AI Foundry Agent

A SharePoint Framework (SPFx v1.24) React client-side web part that integrates an **Azure AI Foundry Prompt Agent** into SharePoint Online pages.

This architecture enables cross-cloud and cross-tenant deployments (such as **Microsoft 365 GCC Moderate to Azure Government** or multi-tenant commercial) using a zero-trust middle-tier authentication broker.

---

## 🏛️ Architecture Overview

```text
[SharePoint Online Page]
       │
       ▼ (SPFx React Web Part)
[AadHttpClientFactory] ──> Acquires Entra Bearer Token for Tenant A API
       │
       ▼ (HTTPS POST with Bearer Token & CORS)
[Azure Function Authentication Broker]
       │   ├── Validates Tenant A JWT in C# code against public JWKS
       │   └── Extracts caller identity (oid / name / email)
       ▼ (Invokes via Tenant B Managed Identity)
[Azure AI Foundry Project]
       │
       ▼
[Foundry Prompt Agent with File Search & Tools]
```

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