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

export async function getCostsByServiceAndResource(subscriptionId, startDate, endDate) {
  return proxyPost(
    `subscriptions/${subscriptionId}/providers/Microsoft.CostManagement/query`,
    {
      type: "ActualCost",
      timeframe: "Custom",
      timePeriod: { from: startDate, to: endDate },
      dataset: {
        granularity: "None",
        aggregation: { totalCost: { name: "Cost", function: "Sum" } },
        grouping: [
          { type: "Dimension", name: "ServiceName" },
          { type: "Dimension", name: "ResourceGroupName" },
          { type: "Dimension", name: "ResourceId" },
          { type: "Dimension", name: "ResourceType" },
        ],
      },
    },
    "2023-11-01"
  );
}

export async function getCostsByDay(subscriptionId, startDate, endDate) {
  return proxyPost(
    `subscriptions/${subscriptionId}/providers/Microsoft.CostManagement/query`,
    {
      type: "ActualCost",
      timeframe: "Custom",
      timePeriod: { from: startDate, to: endDate },
      dataset: {
        granularity: "Daily",
        aggregation: { totalCost: { name: "Cost", function: "Sum" } },
        grouping: [{ type: "Dimension", name: "ServiceName" }],
      },
    },
    "2023-11-01"
  );
}

export async function getMonthlyTrend(subscriptionId, startDate, endDate) {
  return proxyPost(
    `subscriptions/${subscriptionId}/providers/Microsoft.CostManagement/query`,
    {
      type: "ActualCost",
      timeframe: "Custom",
      timePeriod: { from: startDate, to: endDate },
      dataset: {
        granularity: "Monthly",
        aggregation: { totalCost: { name: "Cost", function: "Sum" } },
        grouping: [{ type: "Dimension", name: "ServiceName" }],
      },
    },
    "2023-11-01"
  );
}

export async function fetchAllData(startDate, endDate) {
  const subscriptions = await listSubscriptions();

  const [costs, trends, daily] = await Promise.allSettled([
    Promise.allSettled(
      subscriptions.map(s =>
        getCostsByServiceAndResource(s.subscriptionId, startDate, endDate)
          .then(d => ({ subscription: s, data: d }))
      )
    ),
    Promise.allSettled(
      subscriptions.map(s =>
        getMonthlyTrend(s.subscriptionId, startDate, endDate)
          .then(d => ({ subscription: s, data: d }))
      )
    ),
    Promise.allSettled(
      subscriptions.map(s =>
        getCostsByDay(s.subscriptionId, startDate, endDate)
          .then(d => ({ subscription: s, data: d }))
      )
    ),
  ]);

  return {
    subscriptions,
    costs: costs.status === "fulfilled"
      ? costs.value.filter(r => r.status === "fulfilled").map(r => r.value)
      : [],
    trends: trends.status === "fulfilled"
      ? trends.value.filter(r => r.status === "fulfilled").map(r => r.value)
      : [],
    daily: daily.status === "fulfilled"
      ? daily.value.filter(r => r.status === "fulfilled").map(r => r.value)
      : [],
  };
}
