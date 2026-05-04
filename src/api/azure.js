// Avontus Accounting Portal - Azure API
// Lighthouse enabled: single Avontus tenant token covers all 4 tenants
// All 14 subscriptions accessible via bd98204b tenant

import { getUserMgmtToken, getAllTenantTokens } from '../auth/msal';

const BASE = 'https://azurereader-api.azurewebsites.net/api/proxy';
const AVONTUS_TENANT = 'bd98204b-b981-4d03-8796-356d537927eb';
const sleep = ms => new Promise(r => setTimeout(r, ms));

// Tenant config for display
export const TENANTS = [
  { name: 'Avontus Software',  id: 'bd98204b-b981-4d03-8796-356d537927eb', color: '#60a5fa' },
  { name: 'Places2Swim',       id: 'afafd9ca-9af6-4f95-8032-f71fc87ef9e5', color: '#34d399' },
  { name: 'Azure-Internal',    id: 'd971099d-75b2-4a01-8d0d-507161733ea5', color: '#a78bfa' },
  { name: 'SmallBiz',          id: '5e9927b8-90dd-40c9-bdb8-3283e73304c6', color: '#f59e0b' },
];

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

async function proxyGet(path, apiVersion = '2022-12-01', userToken = null) {
  const headers = { 'Content-Type': 'application/json' };
  if (userToken) headers['X-User-Token'] = userToken;
  // Always use Avontus tenant - Lighthouse handles cross-tenant
  const res = await fetch(`${BASE}/${path}?api-version=${apiVersion}&tenantId=${AVONTUS_TENANT}`, { headers });
  const text = await res.text();
  if (!text) throw new Error(`Empty response from ${path}`);
  try { return JSON.parse(text); }
  catch { throw new Error(`Bad JSON from ${path}: ${text.slice(0,200)}`); }
}

async function proxyPost(path, body, apiVersion = '2023-11-01', userToken = null) {
  const headers = { 'Content-Type': 'application/json' };
  if (userToken) headers['X-User-Token'] = userToken;
  const res = await fetch(`${BASE}/${path}?api-version=${apiVersion}&tenantId=${AVONTUS_TENANT}`, {
    method: 'POST',
    headers,
    body: JSON.stringify(body)
  });
  const text = await res.text();
  if (!text) throw new Error(`Empty response from ${path}`);
  try { return JSON.parse(text); }
  catch { throw new Error(`Bad JSON from ${path}: ${text.slice(0,200)}`); }
}

// List ALL subscriptions - Lighthouse makes them all visible from Avontus tenant
export async function listSubscriptions() {
  const d = await proxyGet('subscriptions');
  const allSubs = d.value || [];
  // Map tenantId to our tenant config for display
  return allSubs.filter(s => s.state === 'Enabled').map(s => {
    const tenant = TENANTS.find(t => t.id === s.tenantId) || TENANTS[0];
    return { ...s, tenant, tenantName: tenant.name };
  });
}

async function fetchOneSub(sub, start, end) {
  const sid = sub.subscriptionId;
  const tenant = sub.tenant || TENANTS[0];
  let detRows = [], monRows = [], dayRows = [];

  try {
    const det = await proxyPost(
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
    detRows = parseRows(det, sub, tenant);
  } catch(e) { console.warn('det failed:', sub.displayName, e.message); }

  await sleep(500);

  try {
    const mon = await proxyPost(
      `subscriptions/${sid}/providers/Microsoft.CostManagement/query`,
      { type:'ActualCost', timeframe:'Custom', timePeriod:{from:start,to:end},
        dataset:{ granularity:'Monthly',
          aggregation:{totalCost:{name:'Cost',function:'Sum'}},
          grouping:[{type:'Dimension',name:'ServiceName'}]
        }
      }
    );
    monRows = parseRows(mon, sub, tenant);
  } catch(e) { console.warn('mon failed:', sub.displayName, e.message); }

  await sleep(500);

  try {
    const day = await proxyPost(
      `subscriptions/${sid}/providers/Microsoft.CostManagement/query`,
      { type:'ActualCost', timeframe:'Custom', timePeriod:{from:start,to:end},
        dataset:{ granularity:'Daily',
          aggregation:{totalCost:{name:'Cost',function:'Sum'}},
        }
      }
    );
    dayRows = parseRows(day, sub, tenant);
  } catch(e) { console.warn('day failed:', sub.displayName, e.message); }

  return { sub, detRows, monRows, dayRows };
}

export async function fetchAllData(start, end, onProgress) {
  const subscriptions = await listSubscriptions();
  const allDetailed = [], allMonthly = [], allDaily = [], errors = [];

  for (let i = 0; i < subscriptions.length; i++) {
    const sub = subscriptions[i];
    if (onProgress) onProgress(i, subscriptions.length, `${sub.tenant?.name} · ${sub.displayName}`);
    try {
      const result = await fetchOneSub(sub, start, end);
      allDetailed.push(...result.detRows);
      allMonthly.push(...result.monRows);
      allDaily.push(...result.dayRows);
    } catch(e) { errors.push(`${sub.displayName}: ${e.message}`); }
    if (i < subscriptions.length - 1) await sleep(800);
  }

  return { subscriptions, tenants: TENANTS, allDetailed, allMonthly, allDaily, errors };
}

// ── Invoices ──────────────────────────────────────────────────────────────────

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
    invoiceDate: p.invoiceDate ? p.invoiceDate.slice(0,10) : '',
    periodStart: p.invoicePeriodStartDate ? p.invoicePeriodStartDate.slice(0,10) : '',
    periodEnd: p.invoicePeriodEndDate ? p.invoicePeriodEndDate.slice(0,10) : '',
    dueDate: p.dueDate ? p.dueDate.slice(0,10) : '',
    amountDue: p.amountDue?.value ?? 0,
    billedAmount: p.billedAmount?.value ?? 0,
    subTotal: p.subTotal?.value ?? 0,
    taxAmount: p.taxAmount?.value ?? 0,
    totalAmount: p.totalAmount?.value ?? p.billedAmount?.value ?? 0,
    amount: p.totalAmount?.value ?? p.billedAmount?.value ?? 0,
    currency: p.amountDue?.currency ?? 'USD',
    downloadUrl: downloadDoc?.url || null,
    documentType: p.documentType || 'Invoice',
    isMonthly: p.isMonthlyInvoice || false,
    billingProfile: p.billingProfileDisplayName || '',
    paymentMethod: payment ? `${payment.paymentMethodType || payment.paymentMethodFamily || ''}`.trim() : '',
    paymentDate: payment?.date ? payment.date.slice(0,10) : '',
    paymentAmount: payment?.amount?.value ?? 0,
    tenant: tenant?.name || 'Avontus Software',
    tenantColor: tenant?.color || '#60a5fa',
    tenantId: tenant?.id || AVONTUS_TENANT,
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
  const startDate = new Date(now.getFullYear() - 2, 0, 1).toISOString().slice(0,10);
  const endDate = now.toISOString().slice(0,10);
  const avontusTenant = tenants.find(t => t.name === 'Avontus Software') || tenants[0];

  // Fetch from all billing accounts + profiles
  for (const ba of BILLING_ACCOUNTS) {
    try {
      const profilesRes = await proxyGet(
        `providers/Microsoft.Billing/billingAccounts/${ba}/billingProfiles`,
        '2020-05-01', userToken
      );
      for (const profile of profilesRes.value || []) {
        try {
          const res = await proxyGet(
            `providers/Microsoft.Billing/billingAccounts/${ba}/billingProfiles/${profile.name}/invoices`,
            `2020-05-01&periodStartDate=${startDate}&periodEndDate=${endDate}`,
            userToken
          );
          (res.value || []).forEach(inv => {
            allInvoices.push({
              ...normalizeBillingAccountInvoice(inv, avontusTenant),
              billingAccountName: ba,
              billingProfileName: profile.name,
              billingProfile: profile.properties?.displayName || profile.name,
              subName: profile.properties?.displayName || profile.name,
            });
          });
        } catch(e) { /* skip */ }
        await sleep(100);
      }
    } catch(e) {
      // Fallback to billing account level
      try {
        const res = await proxyGet(
          `providers/Microsoft.Billing/billingAccounts/${ba}/invoices`,
          `2020-05-01&periodStartDate=${startDate}&periodEndDate=${endDate}`,
          userToken
        );
        (res.value || []).forEach(inv => {
          allInvoices.push({ ...normalizeBillingAccountInvoice(inv, avontusTenant), billingAccountName: ba });
        });
      } catch(e2) { /* skip */ }
    }
    await sleep(150);
  }

  // Also discover billing accounts in other tenants via Lighthouse
  // Since Lighthouse delegates subscription access but not billing accounts,
  // we still need to enumerate BAs per tenant using the user token
  const tenantTokens = await getAllTenantTokens();
  for (const tenant of tenants.filter(t => t.id !== AVONTUS_TENANT)) {
    const tToken = tenantTokens[tenant.id] || userToken;
    try {
      const baRes = await fetch(
        `${BASE}/providers/Microsoft.Billing/billingAccounts?api-version=2020-05-01&tenantId=${tenant.id}`,
        { headers: { 'Content-Type': 'application/json', 'X-User-Token': tToken } }
      ).then(r => r.json());

      for (const ba of baRes.value || []) {
        try {
          const profilesRes = await fetch(
            `${BASE}/providers/Microsoft.Billing/billingAccounts/${ba.name}/billingProfiles?api-version=2020-05-01&tenantId=${tenant.id}`,
            { headers: { 'Content-Type': 'application/json', 'X-User-Token': tToken } }
          ).then(r => r.json());

          for (const profile of profilesRes.value || []) {
            try {
              const res = await fetch(
                `${BASE}/providers/Microsoft.Billing/billingAccounts/${ba.name}/billingProfiles/${profile.name}/invoices?api-version=2020-05-01&periodStartDate=${startDate}&periodEndDate=${endDate}&tenantId=${tenant.id}`,
                { headers: { 'Content-Type': 'application/json', 'X-User-Token': tToken } }
              ).then(r => r.json());
              (res.value || []).forEach(inv => {
                allInvoices.push({
                  ...normalizeBillingAccountInvoice(inv, tenant),
                  billingAccountName: ba.name,
                  billingProfileName: profile.name,
                  billingProfile: profile.properties?.displayName || profile.name,
                  subName: profile.properties?.displayName || profile.name,
                });
              });
            } catch(e) { /* skip */ }
            await sleep(100);
          }
        } catch(e) { /* skip */ }
        await sleep(150);
      }
    } catch(e) { console.warn(`${tenant.name} billing accounts failed:`, e.message); }
  }

  // Deduplicate and sort newest first
  const seen = new Set();
  return allInvoices
    .filter(i => { if (seen.has(i.id)) return false; seen.add(i.id); return true; })
    .sort((a, b) => (b.invoiceDate || '').localeCompare(a.invoiceDate || ''));
}

export async function fetchInvoiceTransactions(inv, userToken) {
  const invoiceId = inv.name || inv.id;

  // Try billing profile path first
  if (inv.billingAccountName && inv.billingProfileName) {
    try {
      const res = await proxyGet(
        `providers/Microsoft.Billing/billingAccounts/${inv.billingAccountName}/billingProfiles/${inv.billingProfileName}/invoices/${invoiceId}/transactions`,
        '2020-05-01', userToken
      );
      if ((res.value || []).length > 0) return res.value;
    } catch(e) { /* try next */ }
  }

  // Try billing account path
  for (const BA of inv.billingAccountName ? [inv.billingAccountName] : BILLING_ACCOUNTS) {
    try {
      const res = await proxyGet(
        `providers/Microsoft.Billing/billingAccounts/${BA}/invoices/${invoiceId}/transactions`,
        '2020-05-01', userToken
      );
      if ((res.value || []).length > 0) return res.value;
    } catch(e) { /* try next */ }
    await sleep(200);
  }

  return [];
}
