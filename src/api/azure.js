// Avontus Accounting Portal - Azure API
// Optimized: 2-at-a-time batching with retry on 429, 5-min cache

import { getUserMgmtToken, getAllTenantTokens } from '../auth/msal';

const BASE = 'https://azurereader-api.azurewebsites.net/api/proxy';
const AVONTUS_TENANT = 'bd98204b-b981-4d03-8796-356d537927eb';
const sleep = ms => new Promise(r => setTimeout(r, ms));

// ── Cache ─────────────────────────────────────────────────────────────────────
const CACHE_TTL = 5 * 60 * 1000;
const cache = new Map();
const cacheGet = k => { const h = cache.get(k); return h && Date.now()-h.ts < CACHE_TTL ? h.val : null; };
const cacheSet = (k, v) => { cache.set(k, { val: v, ts: Date.now() }); return v; };

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
  const dateI = ['billingmonth','usagedate'].map(n => cols.findIndex(c=>c===n)).find(i=>i>=0) ?? -1;
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

// ── Core fetch with retry on 429 ──────────────────────────────────────────────
async function proxyFetch(path, options = {}, attempt = 0) {
  const { apiVersion = '2022-12-01', method = 'GET', body, userToken } = options;
  const headers = { 'Content-Type': 'application/json' };
  if (userToken) headers['X-User-Token'] = userToken;
  const url = `${BASE}/${path}?api-version=${apiVersion}&tenantId=${AVONTUS_TENANT}`;
  const res = await fetch(url, { method, headers, ...(body ? { body: JSON.stringify(body) } : {}) });

  // Handle 429 with exponential backoff
  if (res.status === 429 && attempt < 4) {
    const retryAfter = parseInt(res.headers.get('Retry-After') || '10');
    const wait = Math.max(retryAfter * 1000, (attempt + 1) * 5000);
    console.warn(`429 on ${path}, waiting ${wait}ms (attempt ${attempt+1})`);
    await sleep(wait);
    return proxyFetch(path, options, attempt + 1);
  }

  const text = await res.text();
  if (!text) throw new Error(`Empty response: ${path}`);
  try { return JSON.parse(text); }
  catch { throw new Error(`Bad JSON [${res.status}]: ${text.slice(0, 200)}`); }
}

const proxyGet  = (path, apiVersion, userToken) => proxyFetch(path, { apiVersion, userToken });
const proxyPost = (path, body, apiVersion = '2023-11-01') => proxyFetch(path, { method: 'POST', body, apiVersion });

// ── Run promises with limited concurrency ─────────────────────────────────────
async function withConcurrency(items, limit, fn) {
  const results = [];
  let i = 0;
  async function worker() {
    while (i < items.length) {
      const idx = i++;
      results[idx] = await fn(items[idx], idx).catch(e => ({ error: e.message }));
    }
  }
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, worker));
  return results;
}

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

// Fetch a single subscription — detailed + monthly + daily sequentially
// (sequential within sub to avoid hammering the same endpoint)
// Sequential fetching — proxy handles 429 retries server-side
// Client just needs to pace requests to avoid flooding
async function fetchOneSub(sub, start, end) {
  const sid = sub.subscriptionId;
  const tenant = sub.tenant || TENANTS[0];
  const costPath = `subscriptions/${sid}/providers/Microsoft.CostManagement/query`;
  const cacheKey = `cost:${sid}:${start}:${end}`;
  const cached = cacheGet(cacheKey);
  if (cached) return { sub, ...cached };

  const makeBody = (granularity, grouping) => ({
    type: 'ActualCost', timeframe: 'Custom', timePeriod: { from: start, to: end },
    dataset: {
      granularity,
      aggregation: { totalCost: { name: 'Cost', function: 'Sum' } },
      ...(grouping ? { grouping } : {})
    }
  });

  let detRows = [], monRows = [], dayRows = [];

  try {
    const det = await proxyPost(costPath, makeBody('None', [
      { type: 'Dimension', name: 'ServiceName' },
      { type: 'Dimension', name: 'ResourceGroupName' },
      { type: 'Dimension', name: 'ResourceType' },
    ]));
    detRows = parseRows(det, sub, tenant);
  } catch(e) { console.warn('det failed:', sub.displayName, e.message); }
  await sleep(400);

  try {
    const mon = await proxyPost(costPath,
      makeBody('Monthly', [{ type: 'Dimension', name: 'ServiceName' }]));
    monRows = parseRows(mon, sub, tenant);
  } catch(e) { console.warn('mon failed:', sub.displayName, e.message); }
  await sleep(400);

  try {
    const day = await proxyPost(costPath, makeBody('Daily'));
    dayRows = parseRows(day, sub, tenant);
  } catch(e) { console.warn('day failed:', sub.displayName, e.message); }

  const result = { detRows, monRows, dayRows };
  cacheSet(cacheKey, result);
  return { sub, ...result };
}

export async function fetchAllData(start, end, onProgress) {
  const subscriptions = await listSubscriptions();
  const allDetailed = [], allMonthly = [], allDaily = [], errors = [];

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
    // Small gap between subs — proxy handles per-request retries
    if (i < subscriptions.length - 1) await sleep(600);
  }

  return { subscriptions, tenants: TENANTS, allDetailed, allMonthly, allDaily, errors };
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

  // Sequential per billing account to avoid 429
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
        await sleep(200);
      }
    } catch(e) { /* skip BA */ }
    await sleep(300);
  }

  // Other tenant billing accounts
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
            await sleep(200);
          }
        } catch(e) { /* skip */ }
      }
    } catch(e) { /* skip */ }
  }

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

  if (inv.billingAccountName && inv.billingProfileName) {
    try {
      const res = await proxyGet(
        `providers/Microsoft.Billing/billingAccounts/${inv.billingAccountName}/billingProfiles/${inv.billingProfileName}/invoices/${invoiceId}/transactions`,
        '2020-05-01', userToken
      );
      if ((res.value || []).length > 0) return cacheSet(cacheKey, res.value);
    } catch(e) { /* try next */ }
  }

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
