// Azure API — via Function proxy
// Confirmed column order from live API:
// Detailed (None granularity): [Cost, ServiceName, ResourceGroupName, ResourceType, Currency]
// Monthly: [Cost, ServiceName, BillingMonth, Currency]  
// Daily: [Cost, UsageDate, Currency]

const BASE = 'https://azurereader-api.azurewebsites.net/api/proxy';

async function get(path, apiVersion = '2022-12-01') {
  const res = await fetch(`${BASE}/${path}?api-version=${apiVersion}`);
  const text = await res.text();
  if (!text) throw new Error(`Empty response from ${path}`);
  try { return JSON.parse(text); }
  catch { throw new Error(`Invalid JSON from ${path}: ${text.slice(0, 200)}`); }
}

async function post(path, body, apiVersion = '2023-11-01') {
  const res = await fetch(`${BASE}/${path}?api-version=${apiVersion}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });
  const text = await res.text();
  if (!text) throw new Error(`Empty response from ${path}`);
  try { return JSON.parse(text); }
  catch { throw new Error(`Invalid JSON from ${path}: ${text.slice(0, 200)}`); }
}

export async function listSubscriptions() {
  const d = await get('subscriptions');
  return (d.value || []).filter(s => s.state === 'Enabled');
}

// Parse response using confirmed column names (case-insensitive lookup)
function parseRows(resp, sub) {
  const props = resp?.properties;
  if (!props) return [];
  
  const cols = (props.columns || []).map(c => c.name.toLowerCase());
  const rows = props.rows || [];
  
  // Find column indices by name
  const ci = name => cols.findIndex(c => c === name.toLowerCase());
  
  const costI = ci('cost') >= 0 ? ci('cost') : ci('pretaxcost');
  const svcI  = ci('servicename');
  const rgI   = ci('resourcegroupname');
  const rtI   = ci('resourcetype');
  const ridI  = ci('resourceid');
  const dateI = ci('billingmonth') >= 0 ? ci('billingmonth') : ci('usagedate') >= 0 ? ci('usagedate') : ci('date');
  const curI  = ci('currency');

  return rows
    .map(r => ({
      cost:         costI >= 0 ? (parseFloat(r[costI]) || 0) : 0,
      service:      svcI  >= 0 ? (r[svcI]  || 'Unknown') : 'Unknown',
      rg:           rgI   >= 0 ? (r[rgI]   || '') : '',
      resourceType: rtI   >= 0 ? (r[rtI]   || '') : '',
      resourceId:   ridI  >= 0 ? (r[ridI]  || '') : '',
      date:         dateI >= 0 ? String(r[dateI] || '') : '',
      currency:     curI  >= 0 ? (r[curI]  || 'USD') : 'USD',
      sub:          sub?.displayName || '',
      subId:        sub?.subscriptionId || '',
    }))
    .filter(r => r.cost > 0.001);
}

// Detailed cost: service + RG + type per subscription
async function fetchDetailed(subId, start, end) {
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
        ]
      }
    }
  );
}

// Monthly trend by service
async function fetchMonthly(subId, start, end) {
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

// Daily total
async function fetchDaily(subId, start, end) {
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

export async function fetchAllData(start, end) {
  const subscriptions = await listSubscriptions();

  const subResults = await Promise.allSettled(
    subscriptions.map(async sub => {
      const [det, mon, day] = await Promise.allSettled([
        fetchDetailed(sub.subscriptionId, start, end),
        fetchMonthly(sub.subscriptionId, start, end),
        fetchDaily(sub.subscriptionId, start, end),
      ]);
      return {
        sub,
        detailedRows: det.status === 'fulfilled' ? parseRows(det.value, sub) : [],
        monthlyRows:  mon.status === 'fulfilled' ? parseRows(mon.value, sub) : [],
        dailyRows:    day.status === 'fulfilled' ? parseRows(day.value, sub) : [],
        errors: [det, mon, day].filter(r => r.status === 'rejected').map(r => r.reason?.message),
      };
    })
  );

  const subData = subResults
    .filter(r => r.status === 'fulfilled')
    .map(r => r.value);

  return {
    subscriptions,
    allDetailed: subData.flatMap(s => s.detailedRows),
    allMonthly:  subData.flatMap(s => s.monthlyRows),
    allDaily:    subData.flatMap(s => s.dailyRows),
    subData,
    errors: subData.flatMap(s => s.errors).filter(Boolean),
  };
}
