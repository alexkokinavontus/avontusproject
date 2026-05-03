// Multi-tenant Azure Cost Management API
// 4 tenants, 14 subscriptions

const BASE = 'https://azurereader-api.azurewebsites.net/api/proxy';
const sleep = ms => new Promise(r => setTimeout(r, ms));

const TENANTS = [
  { name: 'Avontus Software',  id: 'bd98204b-b981-4d03-8796-356d537927eb', color: '#60a5fa' },
  { name: 'Places2Swim',       id: 'afafd9ca-9af6-4f95-8032-f71fc87ef9e5', color: '#34d399' },
  { name: 'Azure-Internal',    id: 'd971099d-75b2-4a01-8d0d-507161733ea5', color: '#a78bfa' },
  { name: 'SmallBiz',          id: '5e9927b8-90dd-40c9-bdb8-3283e73304c6', color: '#f59e0b' },
];

export { TENANTS };

export function parseAzureDate(raw) {
  const s = String(raw || '');
  if (!s || s === 'null') return null;
  if (/^\d{8}$/.test(s)) return `${s.slice(0,4)}-${s.slice(4,6)}-${s.slice(6,8)}`;
  if (s.includes('-')) return s.slice(0,10);
  return null;
}

export function parseRows(resp, sub, tenant) {
  const props = resp?.properties;
  if (!props) return [];
  const cols = (props.columns || []).map(c => c.name.toLowerCase());
  const rows = props.rows || [];
  const ci = name => cols.findIndex(c => c === name.toLowerCase());
  const costI = ci('cost') >= 0 ? ci('cost') : ci('pretaxcost');
  const svcI  = ci('servicename');
  const rgI   = ci('resourcegroupname');
  const rtI   = ci('resourcetype');
  const dateI = ci('billingmonth') >= 0 ? ci('billingmonth') : ci('usagedate') >= 0 ? ci('usagedate') : -1;

  return rows.map(r => ({
    cost:         costI >= 0 ? (parseFloat(r[costI]) || 0) : 0,
    service:      svcI  >= 0 ? (r[svcI]  || 'Unknown') : 'Unknown',
    rg:           rgI   >= 0 ? (r[rgI]   || '') : '',
    resourceType: rtI   >= 0 ? (r[rtI]   || '') : '',
    date:         dateI >= 0 ? String(r[dateI] || '') : '',
    sub:          sub?.displayName || '',
    subId:        sub?.subscriptionId || '',
    tenant:       tenant?.name || '',
    tenantId:     tenant?.id || '',
    tenantColor:  tenant?.color || '#60a5fa',
  })).filter(r => r.cost > 0.001);
}

async function proxyGet(tenantId, path, apiVersion = '2022-12-01') {
  const res = await fetch(`${BASE}/${path}?api-version=${apiVersion}&tenantId=${tenantId}`);
  const text = await res.text();
  if (!text) throw new Error(`Empty response from ${path}`);
  try { return JSON.parse(text); }
  catch { throw new Error(`Bad JSON: ${text.slice(0, 200)}`); }
}

async function proxyPost(tenantId, path, body, apiVersion = '2023-11-01') {
  const res = await fetch(`${BASE}/${path}?api-version=${apiVersion}&tenantId=${tenantId}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });
  const text = await res.text();
  if (!text) throw new Error(`Empty response from ${path}`);
  try { return JSON.parse(text); }
  catch { throw new Error(`Bad JSON: ${text.slice(0, 200)}`); }
}

async function getSubscriptionsForTenant(tenant) {
  try {
    const d = await proxyGet(tenant.id, 'subscriptions');
    return (d.value || [])
      .filter(s => s.state === 'Enabled')
      .map(s => ({ ...s, tenant, tenantName: tenant.name }));
  } catch(e) {
    console.warn(`Failed to get subs for ${tenant.name}:`, e.message);
    return [];
  }
}

async function fetchOneSub(sub, start, end) {
  const tid = sub.tenant.id;
  const sid = sub.subscriptionId;
  const costPath = `subscriptions/${sid}/providers/Microsoft.CostManagement/query`;

  let detRows = [], monRows = [], dayRows = [];

  try {
    const det = await proxyPost(tid, costPath, {
      type: 'ActualCost', timeframe: 'Custom',
      timePeriod: { from: start, to: end },
      dataset: { granularity: 'None',
        aggregation: { totalCost: { name: 'Cost', function: 'Sum' } },
        grouping: [
          { type: 'Dimension', name: 'ServiceName' },
          { type: 'Dimension', name: 'ResourceGroupName' },
          { type: 'Dimension', name: 'ResourceType' },
        ]
      }
    });
    detRows = parseRows(det, sub, sub.tenant);
  } catch(e) { console.warn('det failed:', sub.displayName, e.message); }

  await sleep(600);

  try {
    const mon = await proxyPost(tid, costPath, {
      type: 'ActualCost', timeframe: 'Custom',
      timePeriod: { from: start, to: end },
      dataset: { granularity: 'Monthly',
        aggregation: { totalCost: { name: 'Cost', function: 'Sum' } },
        grouping: [{ type: 'Dimension', name: 'ServiceName' }]
      }
    });
    monRows = parseRows(mon, sub, sub.tenant);
  } catch(e) { console.warn('mon failed:', sub.displayName, e.message); }

  await sleep(600);

  try {
    const day = await proxyPost(tid, costPath, {
      type: 'ActualCost', timeframe: 'Custom',
      timePeriod: { from: start, to: end },
      dataset: { granularity: 'Daily',
        aggregation: { totalCost: { name: 'Cost', function: 'Sum' } },
      }
    });
    dayRows = parseRows(day, sub, sub.tenant);
  } catch(e) { console.warn('day failed:', sub.displayName, e.message); }

  return { sub, detRows, monRows, dayRows };
}

export async function fetchAllData(start, end, onProgress) {
  // Get all subscriptions across all tenants first
  const allSubsNested = await Promise.all(TENANTS.map(t => getSubscriptionsForTenant(t)));
  const allSubs = allSubsNested.flat();

  const allDetailed = [], allMonthly = [], allDaily = [], errors = [];

  for (let i = 0; i < allSubs.length; i++) {
    const sub = allSubs[i];
    if (onProgress) onProgress(i, allSubs.length, `${sub.tenant.name} · ${sub.displayName}`);
    try {
      const result = await fetchOneSub(sub, start, end);
      allDetailed.push(...result.detRows);
      allMonthly.push(...result.monRows);
      allDaily.push(...result.dayRows);
    } catch(e) {
      errors.push(`${sub.displayName}: ${e.message}`);
    }
    if (i < allSubs.length - 1) await sleep(800);
  }

  return {
    subscriptions: allSubs,
    tenants: TENANTS,
    allDetailed, allMonthly, allDaily, errors
  };
}


// ── Billing / Invoices ────────────────────────────────────────────────────────

export async function getBillingAccounts(tenantId) {
  try {
    const res = await proxyGet(tenantId, 'providers/Microsoft.Billing/billingAccounts', '2020-05-01');
    return res.value || [];
  } catch(e) {
    console.warn('billingAccounts failed:', e.message);
    return [];
  }
}

export async function getInvoicesForBillingAccount(tenantId, billingAccountName) {
  try {
    const res = await proxyGet(
      tenantId,
      `providers/Microsoft.Billing/billingAccounts/${billingAccountName}/invoices`,
      '2020-05-01'
    );
    return (res.value || []).map(i => normalizeInvoice(i, tenantId));
  } catch(e) {
    console.warn('billing account invoices failed:', e.message);
    return [];
  }
}

export async function getInvoicesForSubscription(tenantId, subscriptionId) {
  const now = new Date();
  const start = new Date(now.getFullYear() - 1, 0, 1).toISOString().split('T')[0];
  const end = now.toISOString().split('T')[0];
  try {
    const res = await proxyGet(
      tenantId,
      `subscriptions/${subscriptionId}/providers/Microsoft.Billing/invoices`,
      '2020-09-01&periodStartDate=' + start + '&periodEndDate=' + end
    );
    return (res.value || []).map(i => normalizeInvoice(i, tenantId, subscriptionId));
  } catch(e) {
    console.warn('sub invoices failed:', subscriptionId, e.message);
    return [];
  }
}

function normalizeInvoice(inv, tenantId, subscriptionId) {
  const p = inv.properties || {};
  return {
    id: inv.name || inv.id,
    name: inv.name,
    status: p.status || p.invoiceStatus || 'Unknown',
    dueDate: p.dueDate || p.invoicePeriodEndDate || '',
    invoiceDate: p.invoiceDate || p.invoicePeriodStartDate || '',
    periodStart: p.billingProfileDisplayName ? p.invoicePeriodStartDate : p.invoicePeriodStartDate,
    periodEnd: p.invoicePeriodEndDate || '',
    amount: p.amountDue?.value ?? p.subTotal?.value ?? p.totalAmount?.value ?? 0,
    currency: p.amountDue?.currency ?? p.currency ?? 'USD',
    downloadUrl: p.invoiceDocuments?.[0]?.url || p.downloadUrl || null,
    billingProfileName: p.billingProfileDisplayName || '',
    subscriptionId: subscriptionId || '',
    tenantId,
  };
}

export async function fetchAllInvoices(subscriptions, tenants) {
  const all = [];

  for (const tenant of tenants) {
    // Try billing account level first
    const billingAccounts = await getBillingAccounts(tenant.id);
    for (const ba of billingAccounts) {
      const invs = await getInvoicesForBillingAccount(tenant.id, ba.name);
      invs.forEach(i => { i.tenant = tenant.name; i.tenantColor = tenant.color; i.source = 'billing-account'; });
      all.push(...invs);
      await sleep(500);
    }

    // Also try subscription level for each sub in this tenant
    const tenantSubs = subscriptions.filter(s => s.tenantId === tenant.id || s.tenant?.id === tenant.id);
    for (const sub of tenantSubs) {
      const invs = await getInvoicesForSubscription(tenant.id, sub.subscriptionId);
      invs.forEach(i => {
        i.tenant = tenant.name;
        i.tenantColor = tenant.color;
        i.subName = sub.displayName;
        i.source = 'subscription';
      });
      all.push(...invs);
      await sleep(400);
    }
  }

  // Deduplicate by invoice name
  const seen = new Set();
  return all.filter(i => {
    if (seen.has(i.id)) return false;
    seen.add(i.id);
    return true;
  }).sort((a, b) => (b.dueDate || b.invoiceDate || '').localeCompare(a.dueDate || a.invoiceDate || ''));
}
