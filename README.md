# AzureReader — Avontus Cost & Resource Report

A production-grade React dashboard for comprehensive Azure cost accounting and resource usage across all subscriptions in the `avontus.com` tenant.

## Features

- **Real-time Cost Data** — Azure Cost Management API (MTD spend, 12-month trends)
- **Resource Inventory** — Azure Resource Graph queries (by type, location, resource group)
- **Multi-Subscription** — Aggregates across all subscriptions in the tenant
- **4 Views**: Overview, Costs, Resources, Subscriptions
- **Dark enterprise UI** with interactive charts

## App Registration Details

| Field | Value |
|---|---|
| Display Name | AzureReader |
| Application (Client) ID | `3977e66a-cdf1-419d-9d0d-70e8cf3a76ed` |
| Directory (Tenant) ID | `bd98204b-b981-4d03-8796-356d537927eb` |
| Cert Expiry | 3/5/2028 |

## Required API Permissions

Ensure these permissions are granted in Azure Portal > App Registrations > AzureReader > API Permissions:

| Permission | Type | Reason |
|---|---|---|
| `Microsoft.Management/managementGroups/read` | Azure RBAC | List subscriptions |
| `Microsoft.CostManagement/query/read` | Azure RBAC | Read cost data |
| `Microsoft.ResourceGraph/resources/read` | Azure RBAC | Query resources |

**Assign these roles at the tenant/subscription level:**
- **Reader** — on each subscription (for Cost Management + Resource Graph)
- **Cost Management Reader** — on each subscription (for cost queries)

## Local Development

```bash
# 1. Install dependencies
npm install

# 2. Create local env file
cp .env.example .env.local
# Edit .env.local and add your client secret

# 3. Start dev server
npm run dev
# Open http://localhost:3000
```

## Deploy to Azure Static Web Apps

### Option A: Azure Portal (Manual)

1. **Create Static Web App**
   ```
   Azure Portal → Create a resource → Static Web App
   - Name: azure-reader
   - Region: East US 2
   - Source: GitHub (connect your repo)
   - Build preset: Vite
   - App location: /
   - Output location: dist
   ```

2. **Set environment variable**
   ```
   Static Web App → Configuration → Application settings
   + Add: VITE_AZURE_CLIENT_SECRET = <your secret>
   ```

3. **Push to GitHub** — CI/CD deploys automatically

### Option B: GitHub Actions (Automated)

1. Get your Static Web Apps deployment token:
   ```
   Azure Portal → Static Web App → Manage deployment token
   ```

2. Add GitHub secrets:
   ```
   GitHub repo → Settings → Secrets and variables → Actions
   + AZURE_STATIC_WEB_APPS_API_TOKEN = <deployment token>
   + AZURE_CLIENT_SECRET = <client secret>
   ```

3. Push to `main` branch — workflow in `.github/workflows/deploy.yml` handles the rest.

### Option C: Azure CLI

```bash
# Build the app
npm run build

# Deploy to existing Static Web App
az staticwebapp upload \
  --name azure-reader \
  --resource-group your-rg \
  --source dist/
```

## Security Notes

⚠ **IMPORTANT**: This app uses client credentials flow. The client secret must NEVER be:
- Committed to source control (`.gitignore` protects `.env.local`)
- Hardcoded in source files
- Exposed in browser DevTools in production

**Production recommendation**: Proxy Azure API calls through an Azure Function with managed identity. The Function holds the credentials; the frontend calls the Function.

### Recommended Production Architecture

```
Browser → Azure Static Web App (React)
              ↓ calls
        Azure Function (Node.js, Managed Identity)
              ↓ calls
        Azure Management APIs
```

## CORS Note

Azure Cost Management and Resource Graph APIs do not support browser-side CORS by default. If you encounter CORS errors:
1. Deploy the API calls to an Azure Function (recommended)
2. Or configure API Management with CORS headers as a proxy

## Rotating the Client Secret

1. Azure Portal → App registrations → AzureReader → Certificates & secrets
2. Create new client secret
3. Update `VITE_AZURE_CLIENT_SECRET` in Static Web App configuration
4. Delete old secret

Current cert ID: `cc9ee9c5-50fc-48a7-90c3-222228439bde` (expires 3/5/2028)
