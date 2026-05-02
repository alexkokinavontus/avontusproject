import { useState, useEffect, useCallback, useMemo } from "react";
import { fetchAllData } from "./api/azure";

// ── Constants ─────────────────────────────────────────────────────────────────
const BUDGET = 155000;
const COLORS = ["#60a5fa","#a78bfa","#34d399","#f59e0b","#f87171","#38bdf8","#fb923c","#e879f9","#4ade80","#facc15","#818cf8","#2dd4bf"];
const fmt = (n, d = 0) => n == null ? "$0" : "$" + Number(n).toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });
const fmtDate = d => new Date(d).toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });

function getDefaultDates() {
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), 1);
  return {
    start: start.toISOString().split("T")[0],
    end: now.toISOString().split("T")[0],
  };
}

// ── Mini Components ───────────────────────────────────────────────────────────
function Pill({ label, color = "#60a5fa" }) {
  return <span className="pill" style={{ background: color + "22", color, borderColor: color + "44" }}>{label}</span>;
}

function MiniBar({ value, max, color }) {
  return (
    <div className="mini-bar-track">
      <div className="mini-bar-fill" style={{ width: `${Math.min((value / max) * 100, 100)}%`, background: color }} />
    </div>
  );
}

function Sparkline({ values = [], color = "#60a5fa", w = 100, h = 28 }) {
  if (values.length < 2) return <span style={{ color: "var(--t3)", fontSize: 10 }}>—</span>;
  const max = Math.max(...values, 1);
  const pts = values.map((v, i) => `${(i / (values.length - 1)) * w},${h - (v / max) * h}`).join(" ");
  return (
    <svg width={w} height={h} style={{ overflow: "visible", display: "block" }}>
      <polyline points={pts} fill="none" stroke={color} strokeWidth="1.5" strokeLinejoin="round" strokeLinecap="round" />
    </svg>
  );
}

function TrendBadge({ current, previous }) {
  if (!previous) return null;
  const pct = ((current - previous) / previous) * 100;
  const up = pct > 0;
  return (
    <span className={`trend-badge ${up ? "up" : "down"}`}>
      {up ? "▲" : "▼"} {Math.abs(pct).toFixed(1)}%
    </span>
  );
}

// ── Bar Chart ─────────────────────────────────────────────────────────────────
function StackedBar({ months, services, colorMap }) {
  if (!months.length) return <div className="empty-state">No trend data for this period</div>;
  const max = Math.max(...months.map(m => m.total), 1);
  const topServices = services.slice(0, 8);
  return (
    <div className="chart-outer">
      <div className="chart-yaxis">
        {[1, 0.75, 0.5, 0.25, 0].map(p => (
          <div key={p} className="y-label">{fmt(max * p, 0)}</div>
        ))}
      </div>
      <div className="chart-body">
        <div className="chart-gridlines">
          {[0.25, 0.5, 0.75, 1].map(p => (
            <div key={p} className="gridline" style={{ bottom: `${p * 100}%` }} />
          ))}
        </div>
        <div className="bar-row">
          {months.map((m, i) => (
            <div key={i} className="bar-col">
              <div className="bar-stack" style={{ height: `${(m.total / max) * 100}%` }}>
                {topServices.map(svc => {
                  const val = m[svc] || 0;
                  const pct = (val / m.total) * 100;
                  return pct > 0 ? (
                    <div key={svc} className="bar-seg" style={{ height: `${pct}%`, background: colorMap[svc] || "#60a5fa" }}
                      title={`${svc}: ${fmt(val)}`} />
                  ) : null;
                })}
              </div>
              <div className="bar-label">{m.label}</div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ── Daily Trend Line ──────────────────────────────────────────────────────────
function DailyLine({ days }) {
  if (days.length < 2) return <div className="empty-state">No daily data</div>;
  const max = Math.max(...days.map(d => d.total), 1);
  const W = 800, H = 100;
  const pts = days.map((d, i) => `${(i / (days.length - 1)) * W},${H - (d.total / max) * H}`).join(" ");
  const area = `0,${H} ${pts} ${W},${H}`;
  return (
    <div className="daily-chart">
      <svg viewBox={`0 0 ${W} ${H}`} preserveAspectRatio="none" style={{ width: "100%", height: 80 }}>
        <defs>
          <linearGradient id="lg" x1="0" y1="0" x2="0" y2="1">
            <stop offset="0%" stopColor="#60a5fa" stopOpacity="0.3" />
            <stop offset="100%" stopColor="#60a5fa" stopOpacity="0" />
          </linearGradient>
        </defs>
        <polygon points={area} fill="url(#lg)" />
        <polyline points={pts} fill="none" stroke="#60a5fa" strokeWidth="2" strokeLinejoin="round" />
      </svg>
      <div className="daily-labels">
        {days.filter((_, i) => i % Math.ceil(days.length / 8) === 0).map((d, i) => (
          <span key={i} className="daily-label">{new Date(d.date).toLocaleDateString("en-US", { month: "short", day: "numeric" })}</span>
        ))}
      </div>
    </div>
  );
}

// ── Date Range Picker ─────────────────────────────────────────────────────────
const PRESETS = [
  { label: "MTD", getRange: () => { const n = new Date(); return { start: new Date(n.getFullYear(), n.getMonth(), 1).toISOString().split("T")[0], end: n.toISOString().split("T")[0] }; } },
  { label: "Last 30d", getRange: () => { const n = new Date(); const s = new Date(n); s.setDate(s.getDate() - 30); return { start: s.toISOString().split("T")[0], end: n.toISOString().split("T")[0] }; } },
  { label: "Last 90d", getRange: () => { const n = new Date(); const s = new Date(n); s.setDate(s.getDate() - 90); return { start: s.toISOString().split("T")[0], end: n.toISOString().split("T")[0] }; } },
  { label: "Last 6m", getRange: () => { const n = new Date(); const s = new Date(n); s.setMonth(s.getMonth() - 6); return { start: s.toISOString().split("T")[0], end: n.toISOString().split("T")[0] }; } },
  { label: "YTD", getRange: () => { const n = new Date(); return { start: `${n.getFullYear()}-01-01`, end: n.toISOString().split("T")[0] }; } },
  { label: "Last 12m", getRange: () => { const n = new Date(); const s = new Date(n); s.setFullYear(s.getFullYear() - 1); return { start: s.toISOString().split("T")[0], end: n.toISOString().split("T")[0] }; } },
];

function DateRangePicker({ start, end, onChange, activePreset, onPreset }) {
  return (
    <div className="date-picker">
      <div className="preset-btns">
        {PRESETS.map(p => (
          <button key={p.label} className={`preset-btn ${activePreset === p.label ? "active" : ""}`}
            onClick={() => { onPreset(p.label); onChange(p.getRange()); }}>
            {p.label}
          </button>
        ))}
      </div>
      <div className="date-inputs">
        <input type="date" value={start} max={end} onChange={e => { onPreset(null); onChange({ start: e.target.value, end }); }} className="date-input" />
        <span className="date-sep">→</span>
        <input type="date" value={end} min={start} onChange={e => { onPreset(null); onChange({ start, end: e.target.value }); }} className="date-input" />
      </div>
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const defaults = getDefaultDates();
  const [dateRange, setDateRange] = useState(defaults);
  const [activePreset, setActivePreset] = useState("MTD");
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [lastRefresh, setLastRefresh] = useState(null);
  const [activeView, setActiveView] = useState("Overview");
  const [expandedRows, setExpandedRows] = useState({});
  const [searchTerm, setSearchTerm] = useState("");
  const [sortBy, setSortBy] = useState("cost");
  const [selectedSub, setSelectedSub] = useState("all");
  const [expandedServices, setExpandedServices] = useState({});

  const load = useCallback(async (range = dateRange) => {
    setLoading(true); setError(null);
    try {
      const result = await fetchAllData(range.start, range.end);
      setData(result);
      setLastRefresh(new Date());
    } catch (e) { setError(e.message); }
    finally { setLoading(false); }
  }, [dateRange]);

  useEffect(() => { load(); }, []);

  const handleDateChange = useCallback((range) => {
    setDateRange(range);
    load(range);
  }, [load]);

  // ── Process cost data ──
  const processed = useMemo(() => {
    if (!data) return null;

    // Flat rows: { sub, service, rg, resourceId, resourceType, cost }
    const rows = [];
    (data.costs || []).forEach(c => {
      const cols = c.data?.properties?.columns?.map(col => col.name) || [];
      const costIdx = cols.findIndex(n => n === "Cost" || n === "PreTaxCost");
      const svcIdx = cols.findIndex(n => n === "ServiceName");
      const rgIdx = cols.findIndex(n => n === "ResourceGroupName");
      const ridIdx = cols.findIndex(n => n === "ResourceId");
      const rtIdx = cols.findIndex(n => n === "ResourceType");

      (c.data?.properties?.rows || []).forEach(r => {
        const cost = costIdx >= 0 ? (r[costIdx] || 0) : 0;
        if (cost < 0.001) return;
        rows.push({
          sub: c.subscription?.displayName || c.subscription?.subscriptionId,
          subId: c.subscription?.subscriptionId,
          service: svcIdx >= 0 ? (r[svcIdx] || "Unknown") : "Unknown",
          rg: rgIdx >= 0 ? (r[rgIdx] || "Unknown") : "Unknown",
          resourceId: ridIdx >= 0 ? (r[ridIdx] || "") : "",
          resourceType: rtIdx >= 0 ? (r[rtIdx] || "") : "",
          cost,
        });
      });
    });

    // Monthly trend
    const monthMap = {};
    const allServices = new Set();
    (data.trends || []).forEach(t => {
      const cols = t.data?.properties?.columns?.map(c => c.name) || [];
      const costIdx = cols.findIndex(n => n === "Cost" || n === "PreTaxCost");
      const dateIdx = cols.findIndex(n => n === "BillingMonth" || n === "UsageDate");
      const svcIdx = cols.findIndex(n => n === "ServiceName");
      (t.data?.properties?.rows || []).forEach(r => {
        const cost = costIdx >= 0 ? (r[costIdx] || 0) : 0;
        if (!cost) return;
        const rawDate = dateIdx >= 0 ? r[dateIdx] : null;
        if (!rawDate) return;
        const d = new Date(rawDate);
        const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
        const label = d.toLocaleString("default", { month: "short", year: "2-digit" });
        if (!monthMap[key]) monthMap[key] = { label, total: 0 };
        const svc = svcIdx >= 0 ? (r[svcIdx] || "Other") : "Other";
        allServices.add(svc);
        monthMap[key][svc] = (monthMap[key][svc] || 0) + cost;
        monthMap[key].total += cost;
      });
    });
    const months = Object.values(monthMap).sort((a, b) => a.label.localeCompare(b.label));

    // Daily trend
    const dayMap = {};
    (data.daily || []).forEach(t => {
      const cols = t.data?.properties?.columns?.map(c => c.name) || [];
      const costIdx = cols.findIndex(n => n === "Cost" || n === "PreTaxCost");
      const dateIdx = cols.findIndex(n => n === "UsageDate" || n === "BillingMonth");
      (t.data?.properties?.rows || []).forEach(r => {
        const cost = costIdx >= 0 ? (r[costIdx] || 0) : 0;
        if (!cost) return;
        const date = dateIdx >= 0 ? String(r[dateIdx]) : null;
        if (!date) return;
        const key = date.slice(0, 10);
        if (!dayMap[key]) dayMap[key] = { date: key, total: 0 };
        dayMap[key].total += cost;
      });
    });
    const days = Object.values(dayMap).sort((a, b) => a.date.localeCompare(b.date));

    // Service totals
    const svcMap = {};
    rows.forEach(r => {
      if (!svcMap[r.service]) svcMap[r.service] = { name: r.service, total: 0, subs: {}, rgs: {}, resources: [] };
      svcMap[r.service].total += r.cost;
      svcMap[r.service].subs[r.sub] = (svcMap[r.service].subs[r.sub] || 0) + r.cost;
      svcMap[r.service].rgs[r.rg] = (svcMap[r.service].rgs[r.rg] || 0) + r.cost;
      svcMap[r.service].resources.push(r);
    });
    const services = Object.values(svcMap).sort((a, b) => b.total - a.total);

    // Color map
    const colorMap = {};
    services.forEach((s, i) => { colorMap[s.name] = COLORS[i % COLORS.length]; });
    Array.from(allServices).forEach((s, i) => { if (!colorMap[s]) colorMap[s] = COLORS[i % COLORS.length]; });

    // Sub totals
    const subMap = {};
    rows.forEach(r => {
      if (!subMap[r.subId]) subMap[r.subId] = { id: r.subId, name: r.sub, total: 0, services: {} };
      subMap[r.subId].total += r.cost;
      subMap[r.subId].services[r.service] = (subMap[r.subId].services[r.service] || 0) + r.cost;
    });
    const subs = Object.values(subMap).sort((a, b) => b.total - a.total);

    const grandTotal = rows.reduce((s, r) => s + r.cost, 0);

    // App Services specifically
    const appServices = rows.filter(r =>
      r.service?.toLowerCase().includes("app service") ||
      r.resourceType?.toLowerCase().includes("microsoft.web/sites") ||
      r.resourceType?.toLowerCase().includes("microsoft.web/serverfarms") ||
      r.service?.toLowerCase().includes("azure app service")
    );

    return { rows, months, days, services, colorMap, subs, grandTotal, appServices };
  }, [data]);

  // ── Filter rows ──
  const filteredRows = useMemo(() => {
    if (!processed) return [];
    let rows = processed.rows;
    if (selectedSub !== "all") rows = rows.filter(r => r.subId === selectedSub);
    if (searchTerm) {
      const q = searchTerm.toLowerCase();
      rows = rows.filter(r => r.service?.toLowerCase().includes(q) || r.rg?.toLowerCase().includes(q) || r.sub?.toLowerCase().includes(q));
    }
    if (sortBy === "cost") rows = [...rows].sort((a, b) => b.cost - a.cost);
    else if (sortBy === "service") rows = [...rows].sort((a, b) => a.service.localeCompare(b.service));
    else if (sortBy === "rg") rows = [...rows].sort((a, b) => a.rg.localeCompare(b.rg));
    return rows;
  }, [processed, selectedSub, searchTerm, sortBy]);

  const filteredServices = useMemo(() => {
    if (!processed) return [];
    let svcs = processed.services;
    if (selectedSub !== "all") {
      const subName = data?.subscriptions?.find(s => s.subscriptionId === selectedSub)?.displayName;
      svcs = svcs.map(s => ({
        ...s,
        total: s.subs[subName] || 0,
        resources: s.resources.filter(r => r.subId === selectedSub),
      })).filter(s => s.total > 0).sort((a, b) => b.total - a.total);
    }
    if (searchTerm) {
      const q = searchTerm.toLowerCase();
      svcs = svcs.filter(s => s.name.toLowerCase().includes(q));
    }
    return svcs;
  }, [processed, selectedSub, searchTerm, data]);

  const grandTotal = processed?.grandTotal || 0;
  const now = new Date();
  const daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();

  if (loading) return (
    <div className="splash">
      <div className="splash-spinner"><div className="sp-ring" /><div className="sp-inner">A</div></div>
      <div className="splash-msg">Fetching cost data…</div>
      <div className="splash-range">{fmtDate(dateRange.start)} → {fmtDate(dateRange.end)}</div>
    </div>
  );

  if (error) return (
    <div className="splash">
      <div className="err-card">
        <div className="err-ico">⚠</div>
        <div className="err-ttl">Failed to load</div>
        <div className="err-msg">{error}</div>
        <button className="btn-p" onClick={() => load()}>↻ Retry</button>
      </div>
    </div>
  );

  return (
    <div className="shell">
      {/* Sidebar */}
      <aside className="sidebar">
        <div className="sb-brand">
          <div className="sb-mark">A</div>
          <div><div className="sb-title">AzureReader</div><div className="sb-domain">avontus.com</div></div>
        </div>
        <nav className="sb-nav">
          {[
            { id: "Overview", icon: "▣", label: "Overview" },
            { id: "Services", icon: "⬡", label: "By Service" },
            { id: "AppServices", icon: "⬢", label: "App Services" },
            { id: "Subscriptions", icon: "◈", label: "Subscriptions" },
            { id: "Breakdown", icon: "≡", label: "Full Breakdown" },
          ].map(v => (
            <button key={v.id} className={`sb-item ${activeView === v.id ? "active" : ""}`} onClick={() => setActiveView(v.id)}>
              <span className="sb-ico">{v.icon}</span>{v.label}
            </button>
          ))}
        </nav>
        <div className="sb-divider">SUBSCRIPTIONS</div>
        <div className="sb-subs">
          <button className={`sb-sub-item ${selectedSub === "all" ? "active" : ""}`} onClick={() => setSelectedSub("all")}>
            <span className="sb-sub-dot" style={{ background: "#60a5fa" }} />All subscriptions
          </button>
          {(data?.subscriptions || []).map((s, i) => (
            <button key={s.subscriptionId} className={`sb-sub-item ${selectedSub === s.subscriptionId ? "active" : ""}`}
              onClick={() => setSelectedSub(s.subscriptionId)}>
              <span className="sb-sub-dot" style={{ background: COLORS[i % COLORS.length] }} />
              <span className="sb-sub-name">{s.displayName}</span>
            </button>
          ))}
        </div>
        <div className="sb-foot">
          <div>bd98204b · AzureReader</div>
          {lastRefresh && <div>{lastRefresh.toLocaleTimeString()}</div>}
        </div>
      </aside>

      {/* Main */}
      <div className="main">
        {/* Top control bar */}
        <div className="control-bar">
          <div className="ctrl-left">
            <div className="ctrl-title">{activeView === "AppServices" ? "App Services Cost Analysis" : activeView === "Services" ? "Cost by Service" : activeView === "Breakdown" ? "Full Cost Breakdown" : activeView === "Subscriptions" ? "Subscriptions" : "Cost Overview"}</div>
            <div className="ctrl-range">{fmtDate(dateRange.start)} — {fmtDate(dateRange.end)}</div>
          </div>
          <div className="ctrl-right">
            <DateRangePicker start={dateRange.start} end={dateRange.end} onChange={handleDateChange} activePreset={activePreset} onPreset={setActivePreset} />
            <button className="ctrl-refresh" onClick={() => load()} title="Refresh">↻</button>
          </div>
        </div>

        {/* Filter bar */}
        {(activeView === "Services" || activeView === "Breakdown" || activeView === "AppServices") && (
          <div className="filter-bar">
            <div className="search-wrap">
              <span className="search-ico">⌕</span>
              <input className="search-input" placeholder="Search services, resource groups…" value={searchTerm} onChange={e => setSearchTerm(e.target.value)} />
              {searchTerm && <button className="search-clear" onClick={() => setSearchTerm("")}>×</button>}
            </div>
            <select className="sort-select" value={sortBy} onChange={e => setSortBy(e.target.value)}>
              <option value="cost">Sort: Highest cost</option>
              <option value="service">Sort: Service name</option>
              <option value="rg">Sort: Resource group</option>
            </select>
            <div className="result-count">{filteredServices.length} services · {fmt(grandTotal)}</div>
          </div>
        )}

        <div className="content">
          {/* ── OVERVIEW ── */}
          {activeView === "Overview" && processed && (
            <>
              {/* KPI row */}
              <div className="kpi-row">
                <div className="kpi-card accent-blue">
                  <div className="kpi-label">Total Spend</div>
                  <div className="kpi-val">{fmt(grandTotal)}</div>
                  <div className="kpi-sub">{fmtDate(dateRange.start)} – {fmtDate(dateRange.end)}</div>
                </div>
                <div className="kpi-card accent-purple">
                  <div className="kpi-label">Budget ({fmt(BUDGET)}/mo)</div>
                  <div className="kpi-val" style={{ color: grandTotal > BUDGET ? "var(--red)" : "var(--green)" }}>
                    {grandTotal > BUDGET ? `+${fmt(grandTotal - BUDGET)} over` : `${fmt(BUDGET - grandTotal)} under`}
                  </div>
                  <div className="kpi-progress"><div className="kpi-prog-fill" style={{ width: `${Math.min(grandTotal / BUDGET * 100, 100)}%`, background: grandTotal > BUDGET ? "var(--red)" : "var(--green)" }} /></div>
                </div>
                <div className="kpi-card accent-green">
                  <div className="kpi-label">Subscriptions</div>
                  <div className="kpi-val">{processed.subs.length}</div>
                  <div className="kpi-sub">Active in avontus.com</div>
                </div>
                <div className="kpi-card accent-amber">
                  <div className="kpi-label">Top Service</div>
                  <div className="kpi-val" style={{ fontSize: 16 }}>{processed.services[0]?.name?.split(" ").slice(0, 3).join(" ") || "—"}</div>
                  <div className="kpi-sub">{fmt(processed.services[0]?.total)} · {processed.services[0] ? ((processed.services[0].total / grandTotal) * 100).toFixed(1) : 0}%</div>
                </div>
              </div>

              {/* Trend + daily */}
              <div className="two-col">
                <div className="panel">
                  <div className="panel-hdr"><span className="panel-ttl">Monthly Cost Trend</span><span className="panel-sub">{processed.months.length} months</span></div>
                  <StackedBar months={processed.months} services={processed.services.slice(0, 8).map(s => s.name)} colorMap={processed.colorMap} />
                  <div className="legend-row">
                    {processed.services.slice(0, 6).map((s, i) => (
                      <div key={s.name} className="leg-item">
                        <span className="leg-dot" style={{ background: processed.colorMap[s.name] }} />
                        {s.name.replace("Azure ", "").replace("Microsoft ", "")}
                      </div>
                    ))}
                  </div>
                </div>
                <div className="panel">
                  <div className="panel-hdr"><span className="panel-ttl">Daily Spend</span><span className="panel-sub">{processed.days.length} days</span></div>
                  <DailyLine days={processed.days} />
                  <div className="daily-stats">
                    <div className="ds-item"><span>Avg/day</span><strong>{fmt(processed.days.reduce((s, d) => s + d.total, 0) / (processed.days.length || 1))}</strong></div>
                    <div className="ds-item"><span>Peak day</span><strong>{fmt(Math.max(...processed.days.map(d => d.total)))}</strong></div>
                    <div className="ds-item"><span>Days tracked</span><strong>{processed.days.length}</strong></div>
                  </div>
                </div>
              </div>

              {/* Top services quick view */}
              <div className="panel">
                <div className="panel-hdr"><span className="panel-ttl">Top Services by Cost</span><button className="view-all-btn" onClick={() => setActiveView("Services")}>View all →</button></div>
                <table className="tbl">
                  <thead><tr><th>Service</th><th>Subscriptions</th><th className="r">Cost</th><th className="r">Share</th><th>Distribution</th></tr></thead>
                  <tbody>
                    {processed.services.slice(0, 10).map((s, i) => (
                      <tr key={s.name}>
                        <td><span className="svc-dot" style={{ background: processed.colorMap[s.name] }} /><strong>{s.name}</strong></td>
                        <td>{Object.keys(s.subs).length} sub{Object.keys(s.subs).length !== 1 ? "s" : ""}</td>
                        <td className="r mono">{fmt(s.total, 2)}</td>
                        <td className="r mono">{((s.total / grandTotal) * 100).toFixed(1)}%</td>
                        <td><MiniBar value={s.total} max={processed.services[0]?.total || 1} color={processed.colorMap[s.name]} /></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </>
          )}

          {/* ── BY SERVICE ── */}
          {activeView === "Services" && processed && (
            <div className="panel">
              <div className="panel-hdr">
                <span className="panel-ttl">Cost by Service — Drill Down</span>
                <span className="panel-sub">{filteredServices.length} services</span>
              </div>
              <table className="tbl">
                <thead>
                  <tr>
                    <th style={{ width: 32 }} />
                    <th>Service</th>
                    <th className="r">Cost</th>
                    <th className="r">% Total</th>
                    <th>Distribution</th>
                    <th className="r">Resource Groups</th>
                    <th className="r">Resources</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredServices.map((s, i) => (
                    <>
                      <tr key={s.name} className={`svc-row ${expandedServices[s.name] ? "expanded" : ""}`}
                        onClick={() => setExpandedServices(p => ({ ...p, [s.name]: !p[s.name] }))}>
                        <td><span className="expand-btn">{expandedServices[s.name] ? "▾" : "▸"}</span></td>
                        <td>
                          <div className="svc-name-cell">
                            <span className="svc-dot" style={{ background: processed.colorMap[s.name] }} />
                            <strong>{s.name}</strong>
                          </div>
                        </td>
                        <td className="r mono cost-hi">{fmt(s.total, 2)}</td>
                        <td className="r mono">{((s.total / grandTotal) * 100).toFixed(2)}%</td>
                        <td><MiniBar value={s.total} max={filteredServices[0]?.total || 1} color={processed.colorMap[s.name]} /></td>
                        <td className="r">{Object.keys(s.rgs).length}</td>
                        <td className="r">{s.resources.length}</td>
                      </tr>
                      {expandedServices[s.name] && (
                        <>
                          {/* Sub breakdown */}
                          <tr className="drill-header">
                            <td /><td colSpan={6}><span className="drill-label">By Subscription</span></td>
                          </tr>
                          {Object.entries(s.subs).sort((a, b) => b[1] - a[1]).map(([sub, cost]) => (
                            <tr key={sub} className="drill-sub-row">
                              <td /><td style={{ paddingLeft: 28 }}><span className="drill-arrow">↳</span>{sub}</td>
                              <td className="r mono">{fmt(cost, 2)}</td>
                              <td className="r mono muted">{((cost / s.total) * 100).toFixed(1)}%</td>
                              <td><MiniBar value={cost} max={s.total} color={processed.colorMap[s.name] + "88"} /></td>
                              <td /><td />
                            </tr>
                          ))}
                          {/* RG breakdown */}
                          <tr className="drill-header">
                            <td /><td colSpan={6}><span className="drill-label">By Resource Group</span></td>
                          </tr>
                          {Object.entries(s.rgs).sort((a, b) => b[1] - a[1]).slice(0, 10).map(([rg, cost]) => (
                            <tr key={rg} className="drill-rg-row">
                              <td /><td style={{ paddingLeft: 28 }}><span className="drill-arrow">↳</span><span className="rg-tag">{rg}</span></td>
                              <td className="r mono">{fmt(cost, 2)}</td>
                              <td className="r mono muted">{((cost / s.total) * 100).toFixed(1)}%</td>
                              <td><MiniBar value={cost} max={s.total} color={processed.colorMap[s.name] + "66"} /></td>
                              <td /><td />
                            </tr>
                          ))}
                        </>
                      )}
                    </>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {/* ── APP SERVICES ── */}
          {activeView === "AppServices" && processed && (
            <>
              <div className="kpi-row">
                <div className="kpi-card accent-blue">
                  <div className="kpi-label">App Service Total</div>
                  <div className="kpi-val">{fmt(processed.appServices.reduce((s, r) => s + r.cost, 0))}</div>
                  <div className="kpi-sub">{((processed.appServices.reduce((s, r) => s + r.cost, 0) / grandTotal) * 100).toFixed(1)}% of total spend</div>
                </div>
                <div className="kpi-card accent-purple">
                  <div className="kpi-label">App Service Resources</div>
                  <div className="kpi-val">{processed.appServices.length}</div>
                  <div className="kpi-sub">Across {new Set(processed.appServices.map(r => r.sub)).size} subscriptions</div>
                </div>
                <div className="kpi-card accent-green">
                  <div className="kpi-label">Resource Groups</div>
                  <div className="kpi-val">{new Set(processed.appServices.map(r => r.rg)).size}</div>
                  <div className="kpi-sub">Containing App Services</div>
                </div>
              </div>
              <div className="panel">
                <div className="panel-hdr"><span className="panel-ttl">App Services — Full Cost Drill Down</span></div>
                {processed.appServices.length === 0 ? (
                  <div className="empty-state-lg">
                    <div>No App Service costs found for this period.</div>
                    <div className="muted">Try expanding the date range or check that the AzureReader app has Reader access on all subscriptions.</div>
                  </div>
                ) : (
                  <table className="tbl">
                    <thead>
                      <tr>
                        <th style={{ width: 32 }} />
                        <th>Resource / App</th>
                        <th>Subscription</th>
                        <th>Resource Group</th>
                        <th>Type</th>
                        <th className="r">Cost</th>
                        <th className="r">% of AS Total</th>
                      </tr>
                    </thead>
                    <tbody>
                      {(() => {
                        const asTotal = processed.appServices.reduce((s, r) => s + r.cost, 0);
                        const byRg = {};
                        processed.appServices.forEach(r => {
                          const key = `${r.subId}::${r.rg}`;
                          if (!byRg[key]) byRg[key] = { sub: r.sub, rg: r.rg, total: 0, items: [] };
                          byRg[key].total += r.cost;
                          byRg[key].items.push(r);
                        });
                        return Object.entries(byRg).sort((a, b) => b[1].total - a[1].total).flatMap(([key, group]) => {
                          const isExp = expandedRows[key];
                          return [
                            <tr key={key} className={`rg-group-row ${isExp ? "expanded" : ""}`}
                              onClick={() => setExpandedRows(p => ({ ...p, [key]: !p[key] }))}>
                              <td><span className="expand-btn">{isExp ? "▾" : "▸"}</span></td>
                              <td colSpan={3}>
                                <span className="rg-tag">{group.rg}</span>
                                <span className="sub-badge">{group.sub}</span>
                              </td>
                              <td><Pill label="Resource Group" color="#60a5fa" /></td>
                              <td className="r mono cost-hi">{fmt(group.total, 2)}</td>
                              <td className="r mono">{((group.total / asTotal) * 100).toFixed(1)}%</td>
                            </tr>,
                            ...(isExp ? group.items.sort((a, b) => b.cost - a.cost).map((r, ri) => (
                              <tr key={ri} className="resource-row">
                                <td />
                                <td style={{ paddingLeft: 24 }}>
                                  <span className="drill-arrow">↳</span>
                                  <span className="res-name">{r.resourceId?.split("/").pop() || r.resourceId || "—"}</span>
                                </td>
                                <td className="muted">{r.sub}</td>
                                <td className="muted">{r.rg}</td>
                                <td><span className="type-tag">{r.resourceType?.split("/").pop() || r.service}</span></td>
                                <td className="r mono">{fmt(r.cost, 2)}</td>
                                <td className="r mono muted">{((r.cost / asTotal) * 100).toFixed(2)}%</td>
                              </tr>
                            )) : [])
                          ];
                        });
                      })()}
                    </tbody>
                  </table>
                )}
              </div>
            </>
          )}

          {/* ── SUBSCRIPTIONS ── */}
          {activeView === "Subscriptions" && processed && (
            <div className="sub-grid">
              {processed.subs.map((sub, i) => (
                <div key={sub.id} className="sub-card">
                  <div className="sc-accent" style={{ background: COLORS[i % COLORS.length] }} />
                  <div className="sc-head">
                    <div><div className="sc-name">{sub.name}</div><div className="sc-id">{sub.id?.slice(0, 8)}…</div></div>
                    <div className="sc-cost">{fmt(sub.total)}</div>
                  </div>
                  <MiniBar value={sub.total} max={processed.subs[0]?.total || 1} color={COLORS[i % COLORS.length]} />
                  <div className="sc-share">{((sub.total / grandTotal) * 100).toFixed(1)}% of total</div>
                  <div className="sc-services">
                    {Object.entries(sub.services).sort((a, b) => b[1] - a[1]).slice(0, 6).map(([svc, cost]) => (
                      <div key={svc} className="sc-svc">
                        <span className="sc-svc-dot" style={{ background: processed.colorMap[svc] }} />
                        <span className="sc-svc-name">{svc.replace("Azure ", "").replace("Microsoft ", "")}</span>
                        <span className="sc-svc-cost">{fmt(cost)}</span>
                      </div>
                    ))}
                  </div>
                  <Sparkline values={Object.values(sub.services)} color={COLORS[i % COLORS.length]} w={200} h={30} />
                </div>
              ))}
            </div>
          )}

          {/* ── FULL BREAKDOWN ── */}
          {activeView === "Breakdown" && (
            <div className="panel">
              <div className="panel-hdr">
                <span className="panel-ttl">Full Cost Breakdown — All Resources</span>
                <span className="panel-sub">{filteredRows.length} line items</span>
              </div>
              <table className="tbl">
                <thead>
                  <tr>
                    <th>Service</th>
                    <th>Subscription</th>
                    <th>Resource Group</th>
                    <th>Resource</th>
                    <th>Type</th>
                    <th className="r">Cost</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredRows.slice(0, 500).map((r, i) => (
                    <tr key={i}>
                      <td><span className="svc-dot" style={{ background: processed?.colorMap[r.service] || "#60a5fa" }} />{r.service}</td>
                      <td className="muted">{r.sub}</td>
                      <td><span className="rg-tag-sm">{r.rg}</span></td>
                      <td className="muted res-id">{r.resourceId?.split("/").pop() || "—"}</td>
                      <td className="muted">{r.resourceType?.split("/").pop() || "—"}</td>
                      <td className="r mono cost-hi">{fmt(r.cost, 4)}</td>
                    </tr>
                  ))}
                  {filteredRows.length > 500 && (
                    <tr><td colSpan={6} className="muted" style={{ textAlign: "center", padding: "12px" }}>Showing 500 of {filteredRows.length} rows. Use filters to narrow results.</td></tr>
                  )}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
