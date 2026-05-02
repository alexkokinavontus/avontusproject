const PROXY_BASE = "https://azurereader-api.azurewebsites.net/api/proxy";

async function proxyGet(path, params = {}) {
  const qs = new URLSearchParams({ "api-version": "2022-12-01", ...params });
  const res = await fetch(`${PROXY_BASE}/${path}?${qs}`);
  if (!res.ok) throw new Error(`API error ${res.status}: ${path}`);
  return res.json();
}

async function proxyPost(path, body, apiVersion = "2023-11-01") {
  const res = await fetch(`${PROXY_BASE}/${path}?api-version=${apiVersion}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  if (!res.ok) throw new Error(`API error ${res.status}: ${path}`);
  return res.json();
}

export async function listSubscriptions() {
  const data = await proxyGet("subscriptions");
  return data.value || [];
}

export async function getSubscriptionCosts(subscriptionId, startDate, endDate) {
  return proxyPost(
    `subscriptions/${subscriptionId}/providers/Microsoft.CostManagement/query`,
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
  const now = new Date();
  const startDate = new Date(now.getFullYear(), now.getMonth() - 11, 1).toISOString().split("T")[0];
  const endDate = now.toISOString().split("T")[0];
  return proxyPost(
    `subscriptions/${subscriptionId}/providers/Microsoft.CostManagement/query`,
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

async function resourceGraphQuery(query, subscriptionIds = []) {
  return proxyPost(
    "providers/Microsoft.ResourceGraph/resources",
    { subscriptions: subscriptionIds, query },
    "2021-03-01"
  );
}

export async function getResourcesByType(subscriptionIds) {
  return resourceGraphQuery(
    `Resources | summarize Count=count() by type | order by Count desc | take 20`,
    subscriptionIds
  );
}

export async function getResourcesByLocation(subscriptionIds) {
  return resourceGraphQuery(
    `Resources | summarize Count=count() by location | order by Count desc`,
    subscriptionIds
  );
}

export async function getResourcesByResourceGroup(subscriptionIds) {
  return resourceGraphQuery(
    `Resources | summarize Count=count(), Types=dcount(type) by resourceGroup, subscriptionId | order by Count desc | take 30`,
    subscriptionIds
  );
}

export async function fetchAllData() {
  const subscriptions = await listSubscriptions();
  const subIds = subscriptions.map((s) => s.subscriptionId);

  const now = new Date();
  const firstOfMonth = new Date(now.getFullYear(), now.getMonth(), 1).toISOString().split("T")[0];
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
