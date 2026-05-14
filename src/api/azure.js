// Avontus Accounting Portal - Azure API
// Lighthouse enabled: single Avontus tenant token covers all 4 tenants
// Performance: parallel batched fetching with caching

import { getUserMgmtToken, getAllTenantTokens } from '../auth/msal';

const BASE = 'https://azurereader-api.azurewebsites.net/api/proxy';
const AVONTUS_TENANT = 'bd98204b-b981-4d03-8796-356d537927eb';
const sleep = ms => new Promise(r => setTimeout(r, ms));

// ── Cache ─────────────────────────────────────────────────────────────────────
const CACHE_TTL = 5 * 60 * 1000; // 5 minutes
const cache = new Map();
function cacheGet(key) {
  const hit = cache.get(key);
  if (hit && Date.now() - hit.ts < CACHE_TTL) return hit.val;
  return null;
}
function cacheSet(key, val) { cache.set(key, { val, ts: Date.now() }); return val; }

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
  const costI = cols.findIndex(c => c === 'cost' || c === 'pretaxcost');
  const svcI  = cols.findIndex(c => c === 'servicename');
  const rgI   = cols.findIndex(c => c === 'resourcegroupname');
  const rtI   = cols.findIndex(c => c === 'resourcetype');
  const dateI = cols.findIndex(c => c === 'billingmonth') >= 0
    ? cols.findIndex(c => c === 'billingmonth')
    : cols.findIndex(c => c === 'usagedate');
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

async function proxyFetch(path, options = {}) {
  const { apiVersion = '2022-12-01', method = 'GET', body, userToken } = options;
  const headers = { 'Content-Type': 'application/json' };
  if (userToken) headers['X-User-Token'] = userToken;
  const url = `${BASE}/${path}?api-version=${apiVersion}&tenantId=${AVONTUS_TENANT}`;
  const res = await fetch(url, {
    method,
    headers,
    ...(body ? { body: JSON.stringify(body) } : {})
  });
  const text = await res.text();
  if (!text) throw new Error(`Empty response: ${path}`);
  try { return JSON.parse(text); }
  catch { throw new Error(`Bad JSON: ${text.slice(0, 200)}`); }
}

const proxyGet  = (path, apiVersion, userToken) => proxyFetch(path, { apiVersion, userToken });
const proxyPost = (path, body, apiVersion = '2023-11-01') => proxyFetch(path, { method: 'POST', body, apiVersion });

export async function listSubscriptions() {
  const cached = cacheGet('subscriptions');
  if (cached) return cached;
  const d = await proxyGet('subscriptions');
  const subs = (d.value || [])
    .filter(s => s.state === 'Enabled')
    .map(s => {
      const tenant = TENANTS.find(t => t.id === s.tenantId) || TENANTS[0];
      return { ...s, tenant, tenantName: tenant.name };
    });
  return cacheSet('subscriptions', subs);
}

// ── Parallel cost fetcher ────────────────────────────────────────────────────
// Fetches all 3 query types for one subscription in parallel
async function fetchOneSubParallel(sub, start, end) {
  const sid = sub.subscriptionId;
  const tenant = sub.tenant || TENANTS[0];
  const costPath = `subscriptions/${sid}/providers/Microsoft.CostManagement/query`;

  const makeBody = (granularity, grouping) => ({
    type: 'ActualCost',
    timeframe: 'Custom',
    timePeriod: { from: start, to: end },
    dataset: {
      granularity,
      aggregation: { totalCost: { name: 'Cost', function: 'Sum' } },
      ...(grouping ? { grouping } : {})
    }
  });

  // Check cache
  const cacheKey = `cost:${sid}:${start}:${end}`;
  const cached = cacheGet(cacheKey);
  if (cached) return { sub, ...cached };

  // Fire all 3 queries in parallel
  const [detRes, monRes, dayRes] = await Promise.allSettled([
    proxyPost(costPath, makeBody('None', [
      { type: 'Dimension', name: 'ServiceName' },
      { type: 'Dimension', name: 'ResourceGroupName' },
      { type: 'Dimension', name: 'ResourceType' },
    ])),
    proxyPost(costPath, makeBody('Monthly', [{ type: 'Dimension', name: 'ServiceName' }])),
    proxyPost(costPath, makeBody('Daily')),
  ]);

  const detRows = detRes.status === 'fulfilled' ? parseRows(detRes.value, sub, tenant) : [];
  const monRows = monRes.status === 'fulfilled' ? parseRows(monRes.value, sub, tenant) : [];
  const dayRows = dayRes.status === 'fulfilled' ? parseRows(dayRes.value, sub, tenant) : [];

  const result = { detRows, monRows, dayRows };
  cacheSet(cacheKey, result);
  return { sub, ...result };
}

// Rate limit: Azure allows ~10 concurrent CostManagement queries
// Process in batches of 4 with a small delay between batches
async function fetchInBatches(subscriptions, start, end, onProgress) {
  const BATCH_SIZE = 4;
  const BATCH_DELAY = 1200; // ms between batches
  const allDetailed = [], allMonthly = [], allDaily = [], errors = [];
  let completed = 0;

  for (let i = 0; i < subscriptions.length; i += BATCH_SIZE) {
    const batch = subscriptions.slice(i, i + BATCH_SIZE);

    // Report progress for first sub in batch
    if (onProgress) onProgress(completed, subscriptions.length,
      batch.map(s => s.displayName.split(':')[0].trim()).join(', '));

    // Fetch batch in parallel
    const results = await Promise.allSettled(
      batch.map(sub => fetchOneSubParallel(sub, start, end))
    );

    results.forEach((res, j) => {
      completed++;
      if (res.status === 'fulfilled') {
        allDetailed.push(...res.value.detRows);
        allMonthly.push(...res.value.monRows);
        allDaily.push(...res.value.dayRows);
      } else {
        errors.push(`${batch[j].displayName}: ${res.reason?.message || 'failed'}`);
      }
    });

    // Delay between batches (except last)
    if (i + BATCH_SIZE < subscriptions.length) {
      await sleep(BATCH_DELAY);
    }
  }

  return { allDetailed, allMonthly, allDaily, errors };
}

export async function fetchAllData(start, end, onProgress) {
  const subscriptions = await listSubscriptions();
  const { allDetailed, allMonthly, allDaily, errors } =
    await fetchInBatches(subscriptions, start, end, onProgress);
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
    id: inv.name, name: inv.name, type: 'azure-invoice', source: 'billing-account',
    status: p.status || 'Unknown',
    invoiceDate: p.invoiceDate?.slice(0,10) || '',
    periodStart: p.invoicePeriodStartDate?.slice(0,10) || '',
    periodEnd: p.invoicePeriodEndDate?.slice(0,10) || '',
    dueDate: p.dueDate?.slice(0,10) || '',
    amountDue: p.amountDue?.value ?? 0,
    billedAmount: p.billedAmount?.value ?? 0,
    subTotal: p.subTotal?.value ?? 0,
    taxAmount: p.taxAmount?.value ?? 0,
    totalAmount: p.totalAmount?.value ?? p.billedAmount?.value ?? 0,
    amount: p.totalAmount?.value ?? p.billedAmount?.value ?? 0,
    currency: p.amountDue?.currency ?? 'USD',
    downloadUrl: downloadDoc?.url || null,
    billingProfile: p.billingProfileDisplayName || '',
    paymentMethod: payment ? `${payment.paymentMethodType || payment.paymentMethodFamily || ''}`.trim() : '',
    paymentDate: payment?.date?.slice(0,10) || '',
    paymentAmount: payment?.amount?.value ?? 0,
    tenant: tenant?.name || 'Avontus Software',
    tenantColor: tenant?.color || '#60a5fa',
    tenantId: tenant?.id || AVONTUS_TENANT,
    subName: p.billingProfileDisplayName || '',
    documents: p.documents || [],
    bySub: null, byTenant: null,
  };
}

export async function fetchAllInvoices(subscriptions, tenants) {
  const cached = cacheGet('invoices');
  if (cached) return cached;

  const allInvoices = [];
  const userToken = await getUserMgmtToken();
  const avontusTenant = tenants.find(t => t.name === 'Avontus Software') || tenants[0];
  const now = new Date();
  const startDate = new Date(now.getFullYear() - 2, 0, 1).toISOString().slice(0,10);
  const endDate = now.toISOString().slice(0,10);

  // Fetch all billing profiles in parallel per billing account
  await Promise.all(BILLING_ACCOUNTS.map(async ba => {
    try {
      const profilesRes = await proxyGet(
        `providers/Microsoft.Billing/billingAccounts/${ba}/billingProfiles`,
        '2020-05-01', userToken
      );
      await Promise.all((profilesRes.value || []).map(async profile => {
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
        } catch(e) { /* skip profile */ }
      }));
    } catch(e) { /* skip BA */ }
  }));

  // Also check other tenant billing accounts using per-tenant tokens
  const tenantTokens = await getAllTenantTokens();
  await Promise.all(tenants.filter(t => t.id !== AVONTUS_TENANT).map(async tenant => {
    const tToken = tenantTokens[tenant.id] || userToken;
    try {
      const res = await fetch(
        `${BASE}/providers/Microsoft.Billing/billingAccounts?api-version=2020-05-01&tenantId=${tenant.id}`,
        { headers: { 'Content-Type': 'application/json', 'X-User-Token': tToken } }
      ).then(r => r.json());
      await Promise.all((res.value || []).map(async ba => {
        try {
          const profilesRes = await fetch(
            `${BASE}/providers/Microsoft.Billing/billingAccounts/${ba.name}/billingProfiles?api-version=2020-05-01&tenantId=${tenant.id}`,
            { headers: { 'Content-Type': 'application/json', 'X-User-Token': tToken } }
          ).then(r => r.json());
          await Promise.all((profilesRes.value || []).map(async profile => {
            try {
              const invRes = await fetch(
                `${BASE}/providers/Microsoft.Billing/billingAccounts/${ba.name}/billingProfiles/${profile.name}/invoices?api-version=2020-05-01&periodStartDate=${startDate}&periodEndDate=${endDate}&tenantId=${tenant.id}`,
                { headers: { 'Content-Type': 'application/json', 'X-User-Token': tToken } }
              ).then(r => r.json());
              (invRes.value || []).forEach(inv => {
                allInvoices.push({
                  ...normalizeBillingAccountInvoice(inv, tenant),
                  billingAccountName: ba.name,
                  billingProfileName: profile.name,
                  billingProfile: profile.properties?.displayName || profile.name,
                  subName: profile.properties?.displayName || profile.name,
                });
              });
            } catch(e) { /* skip */ }
          }));
        } catch(e) { /* skip */ }
      }));
    } catch(e) { /* skip tenant */ }
  }));

  const seen = new Set();
  const result = allInvoices
    .filter(i => { if (seen.has(i.id)) return false; seen.add(i.id); return true; })
    .sort((a, b) => (b.invoiceDate || '').localeCompare(a.invoiceDate || ''));

  return cacheSet('invoices', result);
}

export async function fetchInvoiceTransactions(inv, userToken) {
  const invoiceId = inv.name || inv.id;
  const cacheKey = `txn:${invoiceId}`;
  const cached = cacheGet(cacheKey);
  if (cached) return cached;

  // Try billing profile path first
  if (inv.billingAccountName && inv.billingProfileName) {
    try {
      const res = await proxyGet(
        `providers/Microsoft.Billing/billingAccounts/${inv.billingAccountName}/billingProfiles/${inv.billingProfileName}/invoices/${invoiceId}/transactions`,
        '2020-05-01', userToken
      );
      if ((res.value || []).length > 0) return cacheSet(cacheKey, res.value);
    } catch(e) { /* try next */ }
  }

  // Try billing account paths
  for (const BA of inv.billingAccountName ? [inv.billingAccountName] : BILLING_ACCOUNTS) {
    try {
      const res = await proxyGet(
        `providers/Microsoft.Billing/billingAccounts/${BA}/invoices/${invoiceId}/transactions`,
        '2020-05-01', userToken
      );
      if ((res.value || []).length > 0) return cacheSet(cacheKey, res.value);
    } catch(e) { /* try next */ }
  }
  return [];
}
