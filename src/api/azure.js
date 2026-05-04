// Multi-tenant Azure Cost Management API
// 4 tenants, 14 subscriptions

import { getUserMgmtToken } from '../auth/msal';

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

async function proxyGet(tenantId, path, apiVersion = '2022-12-01', userToken = null) {
  const headers = { 'Content-Type': 'application/json' };
  if (userToken) headers['X-User-Token'] = userToken;
  const res = await fetch(`${BASE}/${path}?api-version=${apiVersion}&tenantId=${tenantId}`, { headers });
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



// ── Invoices ─────────────────────────────────────────────────────────────────
// Fetches real Microsoft invoices per subscription
// Matches exactly what Azure Portal shows under Billing > Invoices

async function tryBillingAccountInvoices(tenantId, billingAccountName, userToken) {
  try {
    const res = await proxyGet(tenantId,
      `providers/Microsoft.Billing/billingAccounts/${billingAccountName}/invoices`,
      '2020-05-01', userToken
    );
    return res.value || [];
  } catch(e) { return []; }
}

function normalizeSubInvoice(inv, sub, tenant) {
  const p = inv.properties || {};
  return {
    id: inv.name,
    name: inv.name,
    type: 'azure-invoice',
    source: 'subscription',
    status: p.status || 'Unknown',
    invoiceDate: p.invoiceDate || '',
    periodStart: p.billingPeriodStartDate || p.invoicePeriodStartDate || '',
    periodEnd: p.billingPeriodEndDate || p.invoicePeriodEndDate || '',
    dueDate: p.dueDate || '',
    amountDue: p.amountDue?.value ?? 0,
    totalAmount: p.totalAmount?.value ?? p.amountDue?.value ?? 0,
    amount: p.totalAmount?.value ?? p.amountDue?.value ?? 0,
    currency: p.amountDue?.currency ?? p.totalAmount?.currency ?? 'USD',
    downloadUrl: p.invoiceDocuments?.find(d => d.kind === 'Invoice' || d.kind === 'invoice')?.url
              || p.invoiceDocuments?.[0]?.url
              || null,
    invoiceType: p.invoiceType || 'AzureServices',
    subName: sub?.displayName || '',
    subId: sub?.subscriptionId || '',
    tenant: tenant?.name || '',
    tenantColor: tenant?.color || '#60a5fa',
    tenantId: tenant?.id || '',
    documents: p.invoiceDocuments || [],
    // keep bySub/byTenant empty for real invoices
    bySub: null,
    byTenant: null,
  };
}

export async function fetchAllInvoices(subscriptions, tenants) {
  const allInvoices = [];
  const userToken = await getUserMgmtToken();

  // Fetch real invoices for every subscription
  for (const sub of subscriptions) {
    const tenantId = sub.tenantId || sub.tenant?.id || tenants[0].id;
    const tenant = tenants.find(t => t.id === tenantId);
    try {
      const res = await proxyGet(
        tenantId,
        `subscriptions/${sub.subscriptionId}/providers/Microsoft.Billing/invoices`,
        '2020-05-01',
        userToken
      );
      const invoices = res.value || [];
      invoices.forEach(inv => {
        allInvoices.push(normalizeSubInvoice(inv, sub, tenant));
      });
    } catch(e) {
      console.warn(`Invoices failed for ${sub.displayName}:`, e.message);
    }
    await sleep(300);
  }

  // Also try billing account level
  for (const tenant of tenants) {
    try {
      const baRes = await proxyGet(tenant.id, 'providers/Microsoft.Billing/billingAccounts', '2020-05-01', userToken);
      for (const ba of baRes.value || []) {
        const invs = await tryBillingAccountInvoices(tenant.id, ba.name, userToken);
        invs.forEach(inv => {
          const p = inv.properties || {};
          // Only add if not already present
          if (!allInvoices.find(i => i.name === inv.name)) {
            allInvoices.push({
              id: inv.name,
              name: inv.name,
              type: 'azure-invoice',
              source: 'billing-account',
              status: p.status || 'Unknown',
              invoiceDate: p.invoiceDate || '',
              periodStart: p.invoicePeriodStartDate || '',
              periodEnd: p.invoicePeriodEndDate || '',
              dueDate: p.dueDate || '',
              amountDue: p.amountDue?.value ?? 0,
              totalAmount: p.totalAmount?.value ?? p.amountDue?.value ?? 0,
              amount: p.totalAmount?.value ?? p.amountDue?.value ?? 0,
              currency: p.amountDue?.currency ?? 'USD',
              downloadUrl: p.invoiceDocuments?.find(d => d.kind === 'Invoice')?.url || p.invoiceDocuments?.[0]?.url || null,
              subName: ba.properties?.displayName || ba.name,
              tenant: tenant.name,
              tenantColor: tenant.color,
              tenantId: tenant.id,
              documents: p.invoiceDocuments || [],
              bySub: null,
              byTenant: null,
            });
          }
        });
        await sleep(300);
      }
    } catch(e) { /* no billing account access */ }
  }

  // Sort by invoice date descending (newest first)
  return allInvoices.sort((a, b) =>
    (b.invoiceDate || b.periodEnd || '').localeCompare(a.invoiceDate || a.periodEnd || '')
  );
}
