// Azure API Service Layer
// Uses client credentials flow (app secret) to authenticate
// IMPORTANT: In production, proxy these calls through a secure backend (Azure Function / API Management)
// to avoid exposing client secrets in the browser.

const AZURE_CONFIG = {
  tenantId: "bd98204b-b981-4d03-8796-356d537927eb",
  clientId: "3977e66a-cdf1-419d-9d0d-70e8cf3a76ed",
  // NOTE: Store this in Azure Key Vault / environment variable in production, never hardcode
  clientSecret: import.meta.env.VITE_AZURE_CLIENT_SECRET || "",
  avontusTenantId: "bd98204b-b981-4d03-8796-356d537927eb",
};

const ENDPOINTS = {
  token: `https://login.microsoftonline.com/${AZURE_CONFIG.tenantId}/oauth2/v2.0/token`,
  management: "https://management.azure.com",
  costManagement: "https://management.azure.com",
  resourceGraph: "https://management.azure.com/providers/Microsoft.ResourceGraph/resources",
};

// Cache token to avoid re-fetching
let cachedToken = null;
let tokenExpiry = null;

export async function getAccessToken(scope = "https://management.azure.com/.default") {
  if (cachedToken && tokenExpiry && Date.now() < tokenExpiry) {
    return cachedToken;
  }

  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: AZURE_CONFIG.clientId,
    client_secret: AZURE_CONFIG.clientSecret,
    scope,
  });

  const res = await fetch(ENDPOINTS.token, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  if (!res.ok) {
    const err = await res.json();
    throw new Error(`Auth failed: ${err.error_description || err.error}`);
  }

  const data = await res.json();
  cachedToken = data.access_token;
  tokenExpiry = Date.now() + (data.expires_in - 60) * 1000;
  return cachedToken;
}

async function azureGet(url, params = {}) {
  const token = await getAccessToken();
  const qs = new URLSearchParams({ "api-version": "2022-12-01", ...params });
  const res = await fetch(`${url}?${qs}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!res.ok) throw new Error(`Azure API error ${res.status}: ${url}`);
  return res.json();
}

async function azurePost(url, body, apiVersion = "2023-11-01") {
  const token = await getAccessToken();
  const res = await fetch(`${url}?api-version=${apiVersion}`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Azure POST error ${res.status}: ${err}`);
  }
  return res.json();
}

// ─── Subscriptions ───────────────────────────────────────────────────────────

export async function listSubscriptions() {
  const data = await azureGet(`${ENDPOINTS.management}/subscriptions`);
  return data.value || [];
}

// ─── Cost Management ─────────────────────────────────────────────────────────

export async function getSubscriptionCosts(subscriptionId, startDate, endDate) {
  const url = `${ENDPOINTS.costManagement}/subscriptions/${subscriptionId}/providers/Microsoft.CostManagement/query`;
  return azurePost(
    url,
    {
      type: "ActualCost",
      timeframe: "Custom",
      timePeriod: { from: startDate, to: endDate },
      dataset: {
        granularity: "Monthly",
        aggregation: { totalCost: { name: "Cost", function: "Sum" } },
        grouping: [
          { type: "Dimension", name: "ServiceName" },
          { type: "Dimension", name: "ResourceGroup" },
        ],
      },
    },
    "2023-11-01"
  );
}

export async function getMonthlyTrend(subscriptionId) {
  const url = `${ENDPOINTS.costManagement}/subscriptions/${subscriptionId}/providers/Microsoft.CostManagement/query`;
  const now = new Date();
  const startDate = new Date(now.getFullYear(), now.getMonth() - 11, 1)
    .toISOString()
    .split("T")[0];
  const endDate = now.toISOString().split("T")[0];

  return azurePost(
    url,
    {
      type: "ActualCost",
      timeframe: "Custom",
      timePeriod: { from: startDate, to: endDate },
      dataset: {
        granularity: "Monthly",
        aggregation: { totalCost: { name: "Cost", function: "Sum" } },
      },
    },
    "2023-11-01"
  );
}

export async function getBillingAccount() {
  const url = `${ENDPOINTS.management}/providers/Microsoft.Billing/billingAccounts`;
  try {
    return azureGet(url, { "api-version": "2020-05-01" });
  } catch {
    return null;
  }
}

// ─── Resource Graph ───────────────────────────────────────────────────────────

export async function queryResourceGraph(query, subscriptionIds = []) {
  const token = await getAccessToken();
  const res = await fetch(`${ENDPOINTS.resourceGraph}?api-version=2021-03-01`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      subscriptions: subscriptionIds,
      query,
    }),
  });
  if (!res.ok) throw new Error(`Resource Graph error ${res.status}`);
  return res.json();
}

export async function getResourcesByType(subscriptionIds) {
  return queryResourceGraph(
    `Resources
     | summarize Count=count() by type
     | order by Count desc
     | take 20`,
    subscriptionIds
  );
}

export async function getResourcesByLocation(subscriptionIds) {
  return queryResourceGraph(
    `Resources
     | summarize Count=count() by location
     | order by Count desc`,
    subscriptionIds
  );
}

export async function getResourcesByResourceGroup(subscriptionIds) {
  return queryResourceGraph(
    `Resources
     | summarize Count=count(), Types=dcount(type) by resourceGroup, subscriptionId
     | order by Count desc
     | take 30`,
    subscriptionIds
  );
}

export async function getAllResources(subscriptionIds) {
  return queryResourceGraph(
    `Resources
     | project name, type, location, resourceGroup, subscriptionId, tags
     | order by type asc`,
    subscriptionIds
  );
}

// ─── Aggregate Helper ─────────────────────────────────────────────────────────

export async function fetchAllData() {
  const subscriptions = await listSubscriptions();
  const subIds = subscriptions.map((s) => s.subscriptionId);

  const now = new Date();
  const firstOfMonth = new Date(now.getFullYear(), now.getMonth(), 1)
    .toISOString()
    .split("T")[0];
  const today = now.toISOString().split("T")[0];

  const [resourcesByType, resourcesByLocation, resourcesByRG, costs, trends] =
    await Promise.allSettled([
      getResourcesByType(subIds),
      getResourcesByLocation(subIds),
      getResourcesByResourceGroup(subIds),
      Promise.allSettled(
        subscriptions.map((s) =>
          getSubscriptionCosts(s.subscriptionId, firstOfMonth, today).then((d) => ({
            subscription: s,
            data: d,
          }))
        )
      ),
      Promise.allSettled(
        subscriptions.map((s) =>
          getMonthlyTrend(s.subscriptionId).then((d) => ({
            subscription: s,
            data: d,
          }))
        )
      ),
    ]);

  return {
    subscriptions,
    resourcesByType: resourcesByType.status === "fulfilled" ? resourcesByType.value : null,
    resourcesByLocation: resourcesByLocation.status === "fulfilled" ? resourcesByLocation.value : null,
    resourcesByRG: resourcesByRG.status === "fulfilled" ? resourcesByRG.value : null,
    costs: costs.status === "fulfilled" ? costs.value.filter((r) => r.status === "fulfilled").map((r) => r.value) : [],
    trends: trends.status === "fulfilled" ? trends.value.filter((r) => r.status === "fulfilled").map((r) => r.value) : [],
  };
}
