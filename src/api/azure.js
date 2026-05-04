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
// Confirmed field structure from live API:
// amountDue, billedAmount, billingProfileDisplayName, documents[{kind,url}],
// dueDate, invoiceDate, invoicePeriodStartDate, invoicePeriodEndDate,
// isMonthlyInvoice, payments[], status, subTotal, taxAmount, totalAmount

const BILLING_ACCOUNTS = [
  "6598a801-67b3-5760-fd9c-b6559890b00e:180c8926-382d-41a9-a1da-8eae6ee475ae_2019-05-31",
  "6598a801-67b3-5760-fd9c-b6559890b00e:729679fa-ac97-4460-a0f5-090bd55ea2b6_2019-05-31",
  "6598a801-67b3-5760-fd9c-b6559890b00e:7a95c33a-18e4-4bb3-91b2-83121e02bb41_2019-05-31",
  "6598a801-67b3-5760-fd9c-b6559890b00e:f41edd95-cb04-4d47-a6f3-97941354cd28_2019-05-31",
];

function normalizeBillingAccountInvoice(inv, tenant) {
  const p = inv.properties || {};
  const downloadDoc = (p.documents || []).find(d => d.kind === 'Invoice') || (p.documents || [])[0];
  const payment = (p.payments || [])[0];
  return {
    id: inv.name,
    name: inv.name,
    type: 'azure-invoice',
    source: 'billing-account',
    status: p.status || 'Unknown',
    invoiceDate: p.invoiceDate ? p.invoiceDate.slice(0, 10) : '',
    periodStart: p.invoicePeriodStartDate ? p.invoicePeriodStartDate.slice(0, 10) : '',
    periodEnd: p.invoicePeriodEndDate ? p.invoicePeriodEndDate.slice(0, 10) : '',
    dueDate: p.dueDate ? p.dueDate.slice(0, 10) : '',
    // Real amounts
    amountDue: p.amountDue?.value ?? 0,
    billedAmount: p.billedAmount?.value ?? 0,
    subTotal: p.subTotal?.value ?? 0,
    taxAmount: p.taxAmount?.value ?? 0,
    totalAmount: p.totalAmount?.value ?? p.billedAmount?.value ?? 0,
    amount: p.totalAmount?.value ?? p.billedAmount?.value ?? 0,
    currency: p.amountDue?.currency ?? 'USD',
    // Download URL - direct from documents array
    downloadUrl: downloadDoc?.url || null,
    documentType: p.documentType || 'Invoice',
    isMonthly: p.isMonthlyInvoice || false,
    billingProfile: p.billingProfileDisplayName || '',
    // Payment info
    paymentMethod: payment ? `${payment.paymentMethodType || payment.paymentMethodFamily || ''}`.trim() : '',
    paymentDate: payment?.date ? payment.date.slice(0, 10) : '',
    paymentAmount: payment?.amount?.value ?? 0,
    // Tenant info
    tenant: tenant?.name || 'Avontus Software',
    tenantColor: tenant?.color || '#60a5fa',
    tenantId: tenant?.id || '',
    subName: p.billingProfileDisplayName || '',
    documents: p.documents || [],
    bySub: null,
    byTenant: null,
  };
}

export async function fetchAllInvoices(subscriptions, tenants) {
  const allInvoices = [];
  const userToken = await getUserMgmtToken();

  const now = new Date();
  const startDate = new Date(now.getFullYear() - 2, 0, 1).toISOString().slice(0, 10);
  const endDate = now.toISOString().slice(0, 10);

  // Fetch invoices for every tenant
  for (const tenant of tenants) {
    try {
      // Get all billing accounts for this tenant
      let billingAccounts = [];
      if (tenant.name === 'Avontus Software') {
        // Use hardcoded known billing accounts for Avontus
        billingAccounts = BILLING_ACCOUNTS.map(name => ({ name, tenantId: tenant.id }));
      } else {
        try {
          const baRes = await proxyGet(tenant.id, 'providers/Microsoft.Billing/billingAccounts', '2020-05-01', userToken);
          billingAccounts = (baRes.value || []).map(ba => ({ name: ba.name, tenantId: tenant.id }));
        } catch(e) { /* no access */ }
      }

      for (const ba of billingAccounts) {
        // 1. Get invoices at billing account level
        try {
          const res = await proxyGet(
            tenant.id,
            `providers/Microsoft.Billing/billingAccounts/${ba.name}/invoices`,
            `2020-05-01&periodStartDate=${startDate}&periodEndDate=${endDate}`,
            userToken
          );
          (res.value || []).forEach(inv => {
            allInvoices.push({ ...normalizeBillingAccountInvoice(inv, tenant), billingAccountName: ba.name });
          });
        } catch(e) { console.warn('BA invoices failed:', e.message); }
        await sleep(200);

        // 2. Get billing profiles and their invoices (catches invoices like G151448138)
        try {
          const profilesRes = await proxyGet(
            tenant.id,
            `providers/Microsoft.Billing/billingAccounts/${ba.name}/billingProfiles`,
            '2020-05-01',
            userToken
          );
          for (const profile of profilesRes.value || []) {
            try {
              const profInvRes = await proxyGet(
                tenant.id,
                `providers/Microsoft.Billing/billingAccounts/${ba.name}/billingProfiles/${profile.name}/invoices`,
                `2020-05-01&periodStartDate=${startDate}&periodEndDate=${endDate}`,
                userToken
              );
              (profInvRes.value || []).forEach(inv => {
                const p = inv.properties || {};
                // Override billingProfile name with the actual profile
                allInvoices.push({
                  ...normalizeBillingAccountInvoice(inv, tenant),
                  billingAccountName: ba.name,
                  billingProfileName: profile.name,
                  billingProfile: p.billingProfileDisplayName || profile.properties?.displayName || profile.name,
                  subName: profile.properties?.displayName || profile.name,
                });
              });
            } catch(e) { /* skip profile */ }
            await sleep(150);
          }
        } catch(e) { /* no profile access */ }
      }
    } catch(e) {
      console.warn(`Tenant ${tenant.name} invoices failed:`, e.message);
    }
  }

  // Deduplicate by invoice ID and sort newest first
  const seen = new Set();
  return allInvoices
    .filter(i => { if (seen.has(i.id)) return false; seen.add(i.id); return true; })
    .sort((a, b) => (b.invoiceDate || '').localeCompare(a.invoiceDate || ''));
}

// ── Invoice Transactions ──────────────────────────────────────────────────────
export async function fetchInvoiceTransactions(inv, userToken) {
  const tenantId = inv.tenantId || 'bd98204b-b981-4d03-8796-356d537927eb';
  const invoiceId = inv.name || inv.id;
  const results = [];

  // Try billing profile path first (MCA accounts)
  if (inv.billingAccountName && inv.billingProfileName) {
    try {
      const res = await proxyGet(
        tenantId,
        `providers/Microsoft.Billing/billingAccounts/${inv.billingAccountName}/billingProfiles/${inv.billingProfileName}/invoices/${invoiceId}/transactions`,
        '2020-05-01',
        userToken
      );
      if ((res.value || []).length > 0) return res.value;
    } catch(e) { console.warn('profile txn failed:', e.message); }
  }

  // Try direct billing account path
  for (const BA of inv.billingAccountName ? [inv.billingAccountName] : BILLING_ACCOUNTS) {
    try {
      const res = await proxyGet(
        tenantId,
        `providers/Microsoft.Billing/billingAccounts/${BA}/invoices/${invoiceId}/transactions`,
        '2020-05-01',
        userToken
      );
      if ((res.value || []).length > 0) return res.value;
    } catch(e) { /* try next */ }
    await sleep(200);
  }

  return [];
}
