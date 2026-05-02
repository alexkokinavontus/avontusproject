const BASE = 'https://azurereader-api.azurewebsites.net/api/proxy';

const sleep = ms => new Promise(r => setTimeout(r, ms));

async function get(path, apiVersion = '2022-12-01') {
  const res = await fetch(`${BASE}/${path}?api-version=${apiVersion}`);
  const text = await res.text();
  if (!text) throw new Error(`Empty response from ${path}`);
  try { return JSON.parse(text); }
  catch { throw new Error(`Bad JSON from ${path}: ${text.slice(0,200)}`); }
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
  catch { throw new Error(`Bad JSON from ${path}: ${text.slice(0,200)}`); }
}

// Parse date from Azure — handles 20260401 (number) or "2026-04-01" (string)
export function parseAzureDate(raw) {
  const s = String(raw || '');
  if (!s || s === 'null') return null;
  if (/^\d{8}$/.test(s)) return `${s.slice(0,4)}-${s.slice(4,6)}-${s.slice(6,8)}`;
  if (s.includes('-')) return s.slice(0,10);
  return null;
}

// Parse Cost Management response into clean row objects
// Confirmed column order: [Cost, ServiceName, ResourceGroupName, ResourceType, Currency]
export function parseRows(resp, sub) {
  const props = resp?.properties;
  if (!props) return [];
  const cols = (props.columns || []).map(c => c.name.toLowerCase());
  const rows = props.rows || [];

  const ci = name => cols.findIndex(c => c === name.toLowerCase());
  const costI = ci('cost') >= 0 ? ci('cost') : ci('pretaxcost');
  const svcI  = ci('servicename');
  const rgI   = ci('resourcegroupname');
  const rtI   = ci('resourcetype');
  const ridI  = ci('resourceid');
  const dateI = ci('billingmonth') >= 0 ? ci('billingmonth') : ci('usagedate') >= 0 ? ci('usagedate') : ci('date');

  return rows.map(r => ({
    cost:         costI >= 0 ? (parseFloat(r[costI]) || 0) : 0,
    service:      svcI  >= 0 ? (r[svcI]  || 'Unknown') : 'Unknown',
    rg:           rgI   >= 0 ? (r[rgI]   || '') : '',
    resourceType: rtI   >= 0 ? (r[rtI]   || '') : '',
    resourceId:   ridI  >= 0 ? (r[ridI]  || '') : '',
    date:         dateI >= 0 ? String(r[dateI] || '') : '',
    sub:          sub?.displayName || '',
    subId:        sub?.subscriptionId || '',
  })).filter(r => r.cost > 0.001);
}

export async function listSubscriptions() {
  const d = await get('subscriptions');
  return (d.value || []).filter(s => s.state === 'Enabled');
}

// Fetch one subscription sequentially to avoid 429s
async function fetchOneSub(sub, start, end) {
  const sid = sub.subscriptionId;

  // Fetch detailed first
  let detRows = [];
  try {
    const det = await post(
      `subscriptions/${sid}/providers/Microsoft.CostManagement/query`,
      { type:'ActualCost', timeframe:'Custom', timePeriod:{from:start,to:end},
        dataset:{ granularity:'None',
          aggregation:{totalCost:{name:'Cost',function:'Sum'}},
          grouping:[
            {type:'Dimension',name:'ServiceName'},
            {type:'Dimension',name:'ResourceGroupName'},
            {type:'Dimension',name:'ResourceType'},
          ]
        }
      }
    );
    detRows = parseRows(det, sub);
  } catch(e) { console.warn('detailed failed:', sub.displayName, e.message); }

  await sleep(500); // pace requests

  // Monthly trend
  let monRows = [];
  try {
    const mon = await post(
      `subscriptions/${sid}/providers/Microsoft.CostManagement/query`,
      { type:'ActualCost', timeframe:'Custom', timePeriod:{from:start,to:end},
        dataset:{ granularity:'Monthly',
          aggregation:{totalCost:{name:'Cost',function:'Sum'}},
          grouping:[{type:'Dimension',name:'ServiceName'}]
        }
      }
    );
    monRows = parseRows(mon, sub);
  } catch(e) { console.warn('monthly failed:', sub.displayName, e.message); }

  await sleep(500);

  // Daily spend
  let dayRows = [];
  try {
    const day = await post(
      `subscriptions/${sid}/providers/Microsoft.CostManagement/query`,
      { type:'ActualCost', timeframe:'Custom', timePeriod:{from:start,to:end},
        dataset:{ granularity:'Daily',
          aggregation:{totalCost:{name:'Cost',function:'Sum'}},
        }
      }
    );
    dayRows = parseRows(day, sub);
  } catch(e) { console.warn('daily failed:', sub.displayName, e.message); }

  return { sub, detRows, monRows, dayRows };
}

export async function fetchAllData(start, end, onProgress) {
  const subscriptions = await listSubscriptions();

  const allDetailed = [];
  const allMonthly  = [];
  const allDaily    = [];
  const errors      = [];

  // Sequential with 1s gap between subscriptions to avoid 429
  for (let i = 0; i < subscriptions.length; i++) {
    const sub = subscriptions[i];
    if (onProgress) onProgress(i, subscriptions.length, sub.displayName);
    try {
      const result = await fetchOneSub(sub, start, end);
      allDetailed.push(...result.detRows);
      allMonthly.push(...result.monRows);
      allDaily.push(...result.dayRows);
    } catch(e) {
      errors.push(`${sub.displayName}: ${e.message}`);
    }
    // Pause between subscriptions
    if (i < subscriptions.length - 1) await sleep(1000);
  }

  return { subscriptions, allDetailed, allMonthly, allDaily, errors };
}
