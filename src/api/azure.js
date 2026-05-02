// Azure Cost Management API — via Azure Function proxy
// Proxy URL: https://azurereader-api.azurewebsites.net/api/proxy/

const BASE = 'https://azurereader-api.azurewebsites.net/api/proxy';

async function get(path, apiVersion = '2022-12-01') {
  const url = `${BASE}/${path}?api-version=${apiVersion}`;
  const res = await fetch(url);
  const text = await res.text();
  try { return JSON.parse(text); } catch { throw new Error(`Non-JSON from ${path}: ${text.slice(0, 200)}`); }
}

async function post(path, body, apiVersion = '2023-11-01') {
  const url = `${BASE}/${path}?api-version=${apiVersion}`;
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });
  const text = await res.text();
  try { return JSON.parse(text); } catch { throw new Error(`Non-JSON from ${path}: ${text.slice(0, 200)}`); }
}

// Parse Cost Management response into clean row objects
// Handles any column ordering the API returns
function parseRows(apiResp, subInfo) {
  const props = apiResp?.properties;
  if (!props) return [];
  const cols = (props.columns || []).map(c => c.name.toLowerCase());
  const rows = props.rows || [];

  const idx = name => cols.findIndex(c => c === name.toLowerCase());
  const costIdx = idx('cost') >= 0 ? idx('cost') : idx('pretaxcost') >= 0 ? idx('pretaxcost') : 0;
  const svcIdx = idx('servicename');
  const rgIdx = idx('resourcegroupname');
  const ridIdx = idx('resourceid');
  const rtIdx = idx('resourcetype');
  const dateIdx = idx('usagedate') >= 0 ? idx('usagedate') : idx('billingmonth') >= 0 ? idx('billingmonth') : idx('date');
  const meterIdx = idx('metercategory') >= 0 ? idx('metercategory') : idx('meter');
  const resourceNameIdx = idx('resourcename') >= 0 ? idx('resourcename') : -1;

  return rows.map(r => ({
    cost: costIdx >= 0 ? (parseFloat(r[costIdx]) || 0) : 0,
    service: svcIdx >= 0 ? (r[svcIdx] || 'Unknown') : 'Unknown',
    rg: rgIdx >= 0 ? (r[rgIdx] || '') : '',
    resourceId: ridIdx >= 0 ? (r[ridIdx] || '') : '',
    resourceType: rtIdx >= 0 ? (r[rtIdx] || '') : '',
    resourceName: resourceNameIdx >= 0 ? (r[resourceNameIdx] || '') : '',
    date: dateIdx >= 0 ? String(r[dateIdx] || '') : '',
    meter: meterIdx >= 0 ? (r[meterIdx] || '') : '',
    sub: subInfo?.displayName || '',
    subId: subInfo?.subscriptionId || '',
    currency: 'USD',
  })).filter(r => r.cost > 0.0001);
}

export async function listSubscriptions() {
  const d = await get('subscriptions');
  return (d.value || []).filter(s => s.state === 'Enabled' || !s.state);
}

// Get detailed costs: service + RG + resource + type breakdown
export async function getDetailedCosts(subId, start, end) {
  return post(
    `subscriptions/${subId}/providers/Microsoft.CostManagement/query`,
    {
      type: 'ActualCost',
      timeframe: 'Custom',
      timePeriod: { from: start, to: end },
      dataset: {
        granularity: 'None',
        aggregation: { totalCost: { name: 'Cost', function: 'Sum' } },
        grouping: [
          { type: 'Dimension', name: 'ServiceName' },
          { type: 'Dimension', name: 'ResourceGroupName' },
          { type: 'Dimension', name: 'ResourceType' },
          { type: 'Dimension', name: 'ResourceId' },
        ]
      }
    }
  );
}

// Monthly trend by service
export async function getMonthlyByService(subId, start, end) {
  return post(
    `subscriptions/${subId}/providers/Microsoft.CostManagement/query`,
    {
      type: 'ActualCost',
      timeframe: 'Custom',
      timePeriod: { from: start, to: end },
      dataset: {
        granularity: 'Monthly',
        aggregation: { totalCost: { name: 'Cost', function: 'Sum' } },
        grouping: [{ type: 'Dimension', name: 'ServiceName' }]
      }
    }
  );
}

// Daily total spend
export async function getDailySpend(subId, start, end) {
  return post(
    `subscriptions/${subId}/providers/Microsoft.CostManagement/query`,
    {
      type: 'ActualCost',
      timeframe: 'Custom',
      timePeriod: { from: start, to: end },
      dataset: {
        granularity: 'Daily',
        aggregation: { totalCost: { name: 'Cost', function: 'Sum' } },
      }
    }
  );
}

// Main data fetch — runs all queries in parallel across all subscriptions
export async function fetchAllData(start, end) {
  const subscriptions = await listSubscriptions();

  const results = await Promise.allSettled(
    subscriptions.map(async sub => {
      const [detailed, monthly, daily] = await Promise.allSettled([
        getDetailedCosts(sub.subscriptionId, start, end),
        getMonthlyByService(sub.subscriptionId, start, end),
        getDailySpend(sub.subscriptionId, start, end),
      ]);

      return {
        sub,
        detailedRows: detailed.status === 'fulfilled' ? parseRows(detailed.value, sub) : [],
        monthlyRows: monthly.status === 'fulfilled' ? parseRows(monthly.value, sub) : [],
        dailyRows: daily.status === 'fulfilled' ? parseRows(daily.value, sub) : [],
        errors: [detailed, monthly, daily]
          .filter(r => r.status === 'rejected')
          .map(r => r.reason?.message),
      };
    })
  );

  const subData = results
    .filter(r => r.status === 'fulfilled')
    .map(r => r.value);

  // Merge all rows
  const allDetailed = subData.flatMap(s => s.detailedRows);
  const allMonthly = subData.flatMap(s => s.monthlyRows);
  const allDaily = subData.flatMap(s => s.dailyRows);
  const errors = subData.flatMap(s => s.errors);

  return { subscriptions, allDetailed, allMonthly, allDaily, subData, errors };
}
