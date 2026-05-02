import { useState, useEffect, useCallback, useMemo } from "react";
import { fetchAllData } from "./api/azure";

// ── Constants ─────────────────────────────────────────────────────────────────
const BUDGET = 155000;
const COLORS = ["#60a5fa","#a78bfa","#34d399","#fb923c","#f87171","#2dd4bf","#facc15","#e879f9","#4ade80","#f59e0b","#818cf8","#38bdf8"];
const fmt = (n, d = 0) => !n ? "$0" : "$" + Number(n).toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });
const fmtDate = s => new Date(s + "T12:00:00").toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });

// App Service resource types
const AS_TYPES = ["microsoft.web/sites","microsoft.web/serverfarms","microsoft.web/hostingenvironments","microsoft.web/staticsites","microsoft.web/certificates"];
const AS_SERVICES = ["app service","azure app service","app service plan","web apps"];
const isAppService = r =>
  AS_TYPES.some(t => r.resourceType?.toLowerCase().includes(t)) ||
  AS_SERVICES.some(s => r.service?.toLowerCase().includes(s));

function getDefaultDates() {
  const n = new Date();
  return { start: new Date(n.getFullYear(), n.getMonth(), 1).toISOString().split("T")[0], end: n.toISOString().split("T")[0] };
}

const PRESETS = [
  { label: "MTD",      fn: () => { const n=new Date(); return { start:new Date(n.getFullYear(),n.getMonth(),1).toISOString().split("T")[0], end:n.toISOString().split("T")[0] }; } },
  { label: "Last 30d", fn: () => { const n=new Date(),s=new Date(n); s.setDate(s.getDate()-30); return { start:s.toISOString().split("T")[0], end:n.toISOString().split("T")[0] }; } },
  { label: "Last 90d", fn: () => { const n=new Date(),s=new Date(n); s.setDate(s.getDate()-90); return { start:s.toISOString().split("T")[0], end:n.toISOString().split("T")[0] }; } },
  { label: "Last 6m",  fn: () => { const n=new Date(),s=new Date(n); s.setMonth(s.getMonth()-6); return { start:s.toISOString().split("T")[0], end:n.toISOString().split("T")[0] }; } },
  { label: "YTD",      fn: () => { const n=new Date(); return { start:`${n.getFullYear()}-01-01`, end:n.toISOString().split("T")[0] }; } },
  { label: "Last 12m", fn: () => { const n=new Date(),s=new Date(n); s.setFullYear(s.getFullYear()-1); return { start:s.toISOString().split("T")[0], end:n.toISOString().split("T")[0] }; } },
];

// ── UI Components ─────────────────────────────────────────────────────────────
function Spark({ vals=[], color="#60a5fa", w=90, h=26 }) {
  if (vals.length < 2) return null;
  const max = Math.max(...vals, 1);
  const pts = vals.map((v,i) => `${(i/(vals.length-1))*w},${h-(v/max)*h}`).join(" ");
  return <svg width={w} height={h} style={{overflow:"visible",display:"block"}}><polyline points={pts} fill="none" stroke={color} strokeWidth="1.5" strokeLinejoin="round" strokeLinecap="round"/></svg>;
}

function Bar({ pct, color }) {
  return <div className="bar-track"><div className="bar-fill" style={{width:`${Math.min(pct,100)}%`,background:color}}/></div>;
}

function StackChart({ months, services, colorMap }) {
  if (!months.length) return <div className="empty-msg">No trend data for selected period</div>;
  const max = Math.max(...months.map(m => m.total), 1);
  return (
    <div className="stack-chart">
      <div className="sc-bars">
        {months.map((m, i) => (
          <div key={i} className="sc-col">
            <div className="sc-stack" style={{ height:`${(m.total/max)*100}%` }}>
              {services.map(svc => {
                const v = m[svc] || 0;
                if (!v) return null;
                return <div key={svc} className="sc-seg" style={{ height:`${(v/m.total)*100}%`, background:colorMap[svc]||"#60a5fa" }} title={`${svc}: ${fmt(v)}`}/>;
              })}
            </div>
            <div className="sc-lbl">{m.label}</div>
          </div>
        ))}
      </div>
    </div>
  );
}

function DailyChart({ days }) {
  if (days.length < 2) return <div className="empty-msg">No daily data</div>;
  const vals = days.map(d => d.total);
  const max = Math.max(...vals, 1);
  const W = 600, H = 80;
  const pts = vals.map((v,i) => `${(i/(vals.length-1))*W},${H-(v/max)*H}`).join(" ");
  const area = `0,${H} ${pts} ${W},${H}`;
  return (
    <div className="daily-wrap">
      <svg viewBox={`0 0 ${W} ${H}`} preserveAspectRatio="none" style={{width:"100%",height:70}}>
        <defs><linearGradient id="dg" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#60a5fa" stopOpacity=".25"/><stop offset="100%" stopColor="#60a5fa" stopOpacity="0"/></linearGradient></defs>
        <polygon points={area} fill="url(#dg)"/>
        <polyline points={pts} fill="none" stroke="#60a5fa" strokeWidth="2" strokeLinejoin="round"/>
      </svg>
      <div className="daily-lbls">
        {days.filter((_,i) => i % Math.max(1,Math.floor(days.length/7)) === 0).map((d,i) => (
          <span key={i}>{new Date(d.date+"T12:00:00").toLocaleDateString("en-US",{month:"short",day:"numeric"})}</span>
        ))}
      </div>
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [dateRange, setDateRange] = useState(getDefaultDates);
  const [activePreset, setActivePreset] = useState("MTD");
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [refreshedAt, setRefreshedAt] = useState(null);
  const [view, setView] = useState("overview");
  const [selectedSubId, setSelectedSubId] = useState(null); // null = all
  const [expandedSvc, setExpandedSvc] = useState({});
  const [expandedRg, setExpandedRg] = useState({});
  const [search, setSearch] = useState("");
  const [sortDir, setSortDir] = useState("desc");
  const [debugMode, setDebugMode] = useState(false);

  const load = useCallback(async (range) => {
    setLoading(true); setError(null); setExpandedSvc({}); setExpandedRg({});
    try {
      const result = await fetchAllData(range.start, range.end);
      setData(result);
      setRefreshedAt(new Date());
    } catch(e) { setError(e.message); }
    finally { setLoading(false); }
  }, []);

  useEffect(() => { load(dateRange); }, []);

  const applyPreset = (p) => {
    const range = p.fn();
    setActivePreset(p.label);
    setDateRange(range);
    load(range);
  };

  const applyCustomDate = (field, val) => {
    const range = { ...dateRange, [field]: val };
    setActivePreset(null);
    setDateRange(range);
    load(range);
  };

  // ── Processed data ──────────────────────────────────────────────────────────
  const P = useMemo(() => {
    if (!data) return null;

    // Filter by selected subscription
    const rows = selectedSubId
      ? data.allDetailed.filter(r => r.subId === selectedSubId)
      : data.allDetailed;

    const monthlyRows = selectedSubId
      ? data.allMonthly.filter(r => r.subId === selectedSubId)
      : data.allMonthly;

    const dailyRows = selectedSubId
      ? data.allDaily.filter(r => r.subId === selectedSubId)
      : data.allDaily;

    // Apply search
    const q = search.toLowerCase();
    const filtered = q
      ? rows.filter(r => r.service?.toLowerCase().includes(q) || r.rg?.toLowerCase().includes(q) || r.resourceId?.toLowerCase().includes(q) || r.resourceType?.toLowerCase().includes(q))
      : rows;

    // Grand total
    const grandTotal = rows.reduce((s, r) => s + r.cost, 0);

    // Services map
    const svcMap = {};
    filtered.forEach(r => {
      if (!svcMap[r.service]) svcMap[r.service] = { name: r.service, total: 0, subs: {}, rgs: {}, resources: [] };
      svcMap[r.service].total += r.cost;
      svcMap[r.service].subs[r.sub] = (svcMap[r.service].subs[r.sub] || 0) + r.cost;
      const rgKey = r.rg || "(no resource group)";
      svcMap[r.service].rgs[rgKey] = (svcMap[r.service].rgs[rgKey] || 0) + r.cost;
      svcMap[r.service].resources.push(r);
    });
    const services = Object.values(svcMap).sort((a,b) => sortDir==="desc" ? b.total-a.total : a.total-b.total);

    // Color map
    const colorMap = {};
    [...services].sort((a,b)=>b.total-a.total).forEach((s,i) => { colorMap[s.name] = COLORS[i%COLORS.length]; });

    // Monthly trend
    const monthMap = {};
    monthlyRows.forEach(r => {
      const ds = r.date.length >= 6 ? r.date.slice(0,7) : r.date;
      const key = ds.replace(/(\d{4})(\d{2})/, "$1-$2");
      const d = new Date(key + "-01T12:00:00");
      const label = d.toLocaleString("default",{month:"short",year:"2-digit"});
      if (!monthMap[key]) monthMap[key] = { label, total:0 };
      monthMap[key][r.service] = (monthMap[key][r.service]||0) + r.cost;
      monthMap[key].total += r.cost;
    });
    const months = Object.values(monthMap).sort((a,b)=>a.label.localeCompare(b.label));

    // Daily trend
    const dayMap = {};
    dailyRows.forEach(r => {
      const ds = r.date.length >= 8 ? (r.date.slice(0,4)+"-"+r.date.slice(4,6)+"-"+r.date.slice(6,8)) : r.date.slice(0,10);
      if (!dayMap[ds]) dayMap[ds] = { date:ds, total:0 };
      dayMap[ds].total += r.cost;
    });
    const days = Object.values(dayMap).sort((a,b)=>a.date.localeCompare(b.date));

    // Subscription totals
    const subMap = {};
    data.allDetailed.forEach(r => {
      if (!subMap[r.subId]) subMap[r.subId] = { id:r.subId, name:r.sub, total:0, services:{}, appSvcTotal:0 };
      subMap[r.subId].total += r.cost;
      subMap[r.subId].services[r.service] = (subMap[r.subId].services[r.service]||0)+r.cost;
      if (isAppService(r)) subMap[r.subId].appSvcTotal += r.cost;
    });
    const subs = Object.values(subMap).sort((a,b)=>b.total-a.total);

    // App Services
    const appSvcRows = filtered.filter(isAppService);
    const appSvcByRg = {};
    appSvcRows.forEach(r => {
      const rg = r.rg || "(no resource group)";
      if (!appSvcByRg[rg]) appSvcByRg[rg] = { rg, sub:r.sub, subId:r.subId, total:0, resources:[] };
      appSvcByRg[rg].total += r.cost;
      appSvcByRg[rg].resources.push(r);
    });
    const appSvcGroups = Object.values(appSvcByRg).sort((a,b)=>b.total-a.total);
    const appSvcTotal = appSvcRows.reduce((s,r)=>s+r.cost,0);

    return { rows:filtered, grandTotal, services, colorMap, months, days, subs, appSvcRows, appSvcGroups, appSvcTotal };
  }, [data, selectedSubId, search, sortDir]);

  const selectedSub = data?.subscriptions?.find(s => s.subscriptionId === selectedSubId);
  const now = new Date();
  const daysInMonth = new Date(now.getFullYear(), now.getMonth()+1, 0).getDate();
  const dayOfMonth = now.getDate();

  if (loading) return (
    <div className="splash">
      <div className="sp-ring-wrap"><div className="sp-ring"/><div className="sp-logo">A</div></div>
      <div className="sp-title">Loading Azure Cost Data</div>
      <div className="sp-range">{fmtDate(dateRange.start)} → {fmtDate(dateRange.end)}</div>
      <div className="sp-subs">{data?.subscriptions?.length || "..."} subscriptions</div>
    </div>
  );

  if (error) return (
    <div className="splash">
      <div className="err-card">
        <div className="err-icon">⚠</div>
        <h3>Connection Error</h3>
        <p className="err-msg">{error}</p>
        <p className="err-hint">Check that the Azure Function proxy is running and the AzureReader app has Cost Management Reader + Reader roles on all subscriptions.</p>
        <button className="btn-primary" onClick={() => load(dateRange)}>↻ Retry</button>
      </div>
    </div>
  );

  if (!P) return null;

  return (
    <div className="shell">
      {/* ── Sidebar ── */}
      <aside className="sidebar">
        <div className="sb-brand">
          <div className="sb-mark">A</div>
          <div><div className="sb-name">AzureReader</div><div className="sb-tenant">avontus.com</div></div>
        </div>

        <nav className="sb-nav">
          {[
            {id:"overview",  icon:"▣", label:"Overview"},
            {id:"services",  icon:"⬡", label:"By Service"},
            {id:"appsvcs",   icon:"⬢", label:"App Services"},
            {id:"subs",      icon:"◈", label:"Subscriptions"},
            {id:"breakdown", icon:"≡", label:"Full Breakdown"},
          ].map(v => (
            <button key={v.id} className={`sb-item${view===v.id?" active":""}`} onClick={()=>setView(v.id)}>
              <span className="sb-ico">{v.icon}</span>{v.label}
            </button>
          ))}
        </nav>

        <div className="sb-divider">SUBSCRIPTIONS</div>
        <div className="sb-subs-list">
          <button className={`sb-sub${!selectedSubId?" sel":""}`} onClick={()=>setSelectedSubId(null)}>
            <span className="sb-dot" style={{background:"#60a5fa"}}/>
            <span className="sb-sub-label">All subscriptions</span>
            <span className="sb-sub-cost">{fmt(P.grandTotal)}</span>
          </button>
          {P.subs.map((s,i) => (
            <button key={s.id} className={`sb-sub${selectedSubId===s.id?" sel":""}`} onClick={()=>setSelectedSubId(s.id)}>
              <span className="sb-dot" style={{background:COLORS[i%COLORS.length]}}/>
              <span className="sb-sub-label">{s.name}</span>
              <span className="sb-sub-cost">{fmt(s.total)}</span>
            </button>
          ))}
        </div>

        <div className="sb-footer">
          <div>Tenant: bd98204b</div>
          {refreshedAt && <div>{refreshedAt.toLocaleTimeString()}</div>}
          {data?.errors?.length > 0 && <div className="sb-warn">⚠ {data.errors.length} errors</div>}
        </div>
      </aside>

      {/* ── Main ── */}
      <div className="main">
        {/* Control bar */}
        <div className="ctrl-bar">
          <div className="ctrl-info">
            <span className="ctrl-title">
              {!selectedSubId ? "All Subscriptions" : selectedSub?.displayName}
            </span>
            <span className="ctrl-range">{fmtDate(dateRange.start)} — {fmtDate(dateRange.end)}</span>
          </div>
          <div className="ctrl-filters">
            <div className="presets">
              {PRESETS.map(p => (
                <button key={p.label} className={`preset${activePreset===p.label?" active":""}`} onClick={()=>applyPreset(p)}>{p.label}</button>
              ))}
            </div>
            <div className="date-range">
              <input type="date" className="date-in" value={dateRange.start} max={dateRange.end}
                onChange={e=>applyCustomDate("start",e.target.value)}/>
              <span className="date-arrow">→</span>
              <input type="date" className="date-in" value={dateRange.end} min={dateRange.start}
                onChange={e=>applyCustomDate("end",e.target.value)}/>
            </div>
            <button className="refresh-btn" onClick={()=>load(dateRange)} title="Refresh">↻</button>
          </div>
        </div>

        {/* Search bar for relevant views */}
        {["services","appsvcs","breakdown"].includes(view) && (
          <div className="search-bar">
            <span className="si">⌕</span>
            <input className="si-input" placeholder="Search services, resource groups, resource IDs…"
              value={search} onChange={e=>setSearch(e.target.value)}/>
            {search && <button className="si-clear" onClick={()=>setSearch("")}>×</button>}
            <select className="si-sort" value={sortDir} onChange={e=>setSortDir(e.target.value)}>
              <option value="desc">Highest cost first</option>
              <option value="asc">Lowest cost first</option>
            </select>
            <span className="si-count">{P.services.length} services · {P.rows.length} line items · {fmt(P.grandTotal)}</span>
          </div>
        )}

        <div className="content">

          {/* ═══════════════════════════════ OVERVIEW ═══════════════════════════════ */}
          {view==="overview" && (
            <>
              <div className="kpis">
                <div className="kpi blue">
                  <div className="kpi-lbl">Total Spend</div>
                  <div className="kpi-val">{fmt(P.grandTotal)}</div>
                  <div className="kpi-sub">{fmtDate(dateRange.start)} – {fmtDate(dateRange.end)}</div>
                </div>
                <div className="kpi" style={{"--ac": P.grandTotal>BUDGET?"var(--red)":"var(--green)"}}>
                  <div className="kpi-lbl">vs Budget ({fmt(BUDGET)})</div>
                  <div className="kpi-val" style={{color:P.grandTotal>BUDGET?"var(--red)":"var(--green)"}}>
                    {P.grandTotal>BUDGET?`+${fmt(P.grandTotal-BUDGET)} over`:`${fmt(BUDGET-P.grandTotal)} under`}
                  </div>
                  <Bar pct={P.grandTotal/BUDGET*100} color={P.grandTotal>BUDGET?"var(--red)":"var(--green)"}/>
                </div>
                <div className="kpi purple">
                  <div className="kpi-lbl">App Services</div>
                  <div className="kpi-val">{fmt(P.appSvcTotal)}</div>
                  <div className="kpi-sub">{P.appSvcRows.length} resources · {P.grandTotal?(P.appSvcTotal/P.grandTotal*100).toFixed(1):0}% of total</div>
                </div>
                <div className="kpi amber">
                  <div className="kpi-lbl">Subscriptions</div>
                  <div className="kpi-val">{P.subs.length}</div>
                  <div className="kpi-sub">{P.services.length} services used</div>
                </div>
              </div>

              <div className="two-col">
                <div className="panel">
                  <div className="phdr"><span className="ptitle">Monthly Trend</span><span className="psub">{P.months.length} months</span></div>
                  <StackChart months={P.months} services={P.services.slice(0,8).map(s=>s.name)} colorMap={P.colorMap}/>
                  <div className="leg">
                    {P.services.slice(0,6).map(s=>(
                      <div key={s.name} className="leg-i"><span className="leg-dot" style={{background:P.colorMap[s.name]}}/>{s.name.replace("Azure ","")}</div>
                    ))}
                  </div>
                </div>
                <div className="panel">
                  <div className="phdr"><span className="ptitle">Daily Spend</span><span className="psub">{P.days.length} days</span></div>
                  <DailyChart days={P.days}/>
                  <div className="day-stats">
                    <div className="ds"><span>Avg/day</span><strong>{fmt(P.days.reduce((s,d)=>s+d.total,0)/(P.days.length||1))}</strong></div>
                    <div className="ds"><span>Peak</span><strong>{fmt(Math.max(...P.days.map(d=>d.total),0))}</strong></div>
                    <div className="ds"><span>Days</span><strong>{P.days.length}</strong></div>
                  </div>
                </div>
              </div>

              <div className="panel">
                <div className="phdr">
                  <span className="ptitle">Top Services</span>
                  <button className="view-more" onClick={()=>setView("services")}>View all with drill-down →</button>
                </div>
                <table className="tbl">
                  <thead><tr><th>Service</th><th className="r">Cost</th><th className="r">% Share</th><th>Distribution</th><th className="r">Subscriptions</th></tr></thead>
                  <tbody>
                    {P.services.slice(0,12).map((s,i) => (
                      <tr key={s.name} className="clickable" onClick={()=>{setView("services");setTimeout(()=>setExpandedSvc({[s.name]:true}),100);}}>
                        <td><span className="dot" style={{background:P.colorMap[s.name]}}/><strong>{s.name}</strong></td>
                        <td className="r mono strong">{fmt(s.total,2)}</td>
                        <td className="r mono">{P.grandTotal?(s.total/P.grandTotal*100).toFixed(2):0}%</td>
                        <td><Bar pct={P.services[0]?s.total/P.services[0].total*100:0} color={P.colorMap[s.name]}/></td>
                        <td className="r">{Object.keys(s.subs).length}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </>
          )}

          {/* ═══════════════════════════════ BY SERVICE ══════════════════════════════ */}
          {view==="services" && (
            <div className="panel">
              <div className="phdr"><span className="ptitle">Cost by Service — Click to Expand</span><span className="psub">{P.services.length} services · {fmt(P.grandTotal)}</span></div>
              <table className="tbl">
                <thead>
                  <tr>
                    <th style={{width:28}}/>
                    <th>Service</th>
                    <th className="r">Total Cost</th>
                    <th className="r">% of Total</th>
                    <th>Share</th>
                    <th className="r">RGs</th>
                    <th className="r">Resources</th>
                  </tr>
                </thead>
                <tbody>
                  {P.services.map(s => {
                    const exp = expandedSvc[s.name];
                    return (
                      <>
                        <tr key={s.name} className={`svc-row${exp?" exp":""}`}
                          onClick={()=>setExpandedSvc(p=>({...p,[s.name]:!p[s.name]}))}>
                          <td className="exp-td">{exp?"▾":"▸"}</td>
                          <td><span className="dot" style={{background:P.colorMap[s.name]}}/><strong>{s.name}</strong></td>
                          <td className="r mono strong">{fmt(s.total,2)}</td>
                          <td className="r mono">{P.grandTotal?(s.total/P.grandTotal*100).toFixed(2):0}%</td>
                          <td><Bar pct={P.services[0]?s.total/P.services[0].total*100:0} color={P.colorMap[s.name]}/></td>
                          <td className="r">{Object.keys(s.rgs).length}</td>
                          <td className="r">{s.resources.length}</td>
                        </tr>
                        {exp && <>
                          {/* By Subscription */}
                          <tr className="drill-hdr"><td/><td colSpan={6}><span className="dlbl">📋 By Subscription</span></td></tr>
                          {Object.entries(s.subs).sort((a,b)=>b[1]-a[1]).map(([sub,cost])=>(
                            <tr key={"sub_"+sub} className="drill-sub">
                              <td/>
                              <td style={{paddingLeft:30}}><span className="darrow">↳</span><span className="sub-pill">{sub}</span></td>
                              <td className="r mono">{fmt(cost,2)}</td>
                              <td className="r mono dim">{s.total?(cost/s.total*100).toFixed(1):0}%</td>
                              <td><Bar pct={s.total?cost/s.total*100:0} color={P.colorMap[s.name]+"99"}/></td>
                              <td/><td/>
                            </tr>
                          ))}
                          {/* By Resource Group */}
                          <tr className="drill-hdr"><td/><td colSpan={6}><span className="dlbl">📁 By Resource Group</span></td></tr>
                          {Object.entries(s.rgs).sort((a,b)=>b[1]-a[1]).slice(0,20).map(([rg,cost])=>{
                            const rgKey = s.name+"::"+rg;
                            const rgExp = expandedRg[rgKey];
                            const rgResources = s.resources.filter(r=>(r.rg||"(no resource group)")===rg);
                            return (
                              <>
                                <tr key={"rg_"+rg} className={`drill-rg${rgExp?" exp":""}`}
                                  onClick={e=>{e.stopPropagation();setExpandedRg(p=>({...p,[rgKey]:!p[rgKey]}));}}>
                                  <td/>
                                  <td style={{paddingLeft:30}}>
                                    <span className="darrow">{rgExp?"▾":"▸"}</span>
                                    <span className="rg-tag">{rg}</span>
                                  </td>
                                  <td className="r mono">{fmt(cost,2)}</td>
                                  <td className="r mono dim">{s.total?(cost/s.total*100).toFixed(1):0}%</td>
                                  <td><Bar pct={s.total?cost/s.total*100:0} color={P.colorMap[s.name]+"77"}/></td>
                                  <td/><td className="r dim">{rgResources.length}</td>
                                </tr>
                                {rgExp && rgResources.sort((a,b)=>b.cost-a.cost).map((r,ri)=>(
                                  <tr key={"res_"+ri} className="drill-res">
                                    <td/>
                                    <td style={{paddingLeft:52}}>
                                      <span className="darrow">↳</span>
                                      <span className="res-name">{r.resourceId?.split("/").pop() || r.resourceName || r.resourceId || "—"}</span>
                                      {r.resourceType && <span className="res-type">{r.resourceType.split("/").pop()}</span>}
                                    </td>
                                    <td className="r mono">{fmt(r.cost,4)}</td>
                                    <td className="r mono dim">{s.total?(r.cost/s.total*100).toFixed(2):0}%</td>
                                    <td/>
                                    <td className="r dim" style={{fontSize:10}}>{r.sub}</td>
                                    <td/>
                                  </tr>
                                ))}
                              </>
                            );
                          })}
                        </>}
                      </>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}

          {/* ═══════════════════════════════ APP SERVICES ════════════════════════════ */}
          {view==="appsvcs" && (
            <>
              <div className="kpis">
                <div className="kpi blue">
                  <div className="kpi-lbl">App Service Total</div>
                  <div className="kpi-val">{fmt(P.appSvcTotal)}</div>
                  <div className="kpi-sub">{P.grandTotal?(P.appSvcTotal/P.grandTotal*100).toFixed(1):0}% of all spend</div>
                </div>
                <div className="kpi purple">
                  <div className="kpi-lbl">Resource Groups</div>
                  <div className="kpi-val">{P.appSvcGroups.length}</div>
                  <div className="kpi-sub">With App Service resources</div>
                </div>
                <div className="kpi green">
                  <div className="kpi-lbl">Total Resources</div>
                  <div className="kpi-val">{P.appSvcRows.length}</div>
                  <div className="kpi-sub">Web apps, plans, slots</div>
                </div>
                <div className="kpi amber">
                  <div className="kpi-lbl">Subscriptions</div>
                  <div className="kpi-val">{new Set(P.appSvcRows.map(r=>r.subId)).size}</div>
                  <div className="kpi-sub">With App Services</div>
                </div>
              </div>

              {P.appSvcRows.length === 0 ? (
                <div className="panel">
                  <div className="empty-lg">
                    <div className="empty-icon">⬢</div>
                    <div className="empty-title">No App Service costs found</div>
                    <div className="empty-sub">This could mean: no App Services exist in this period, or costs appear under a different service name. Check the "By Service" view and look for "App Service", "Web Apps", or "Azure App Service".</div>
                    <button className="btn-primary" onClick={()=>setView("services")}>View all services →</button>
                  </div>
                </div>
              ) : (
                <div className="panel">
                  <div className="phdr">
                    <span className="ptitle">App Services — Grouped by Resource Group</span>
                    <span className="psub">Click to expand individual resources</span>
                  </div>
                  <table className="tbl">
                    <thead>
                      <tr>
                        <th style={{width:28}}/>
                        <th>Resource Group / Resource</th>
                        <th>Subscription</th>
                        <th>Type</th>
                        <th className="r">Cost</th>
                        <th className="r">% of AS</th>
                        <th>Share</th>
                      </tr>
                    </thead>
                    <tbody>
                      {P.appSvcGroups.map(group => {
                        const key = group.subId+"::"+group.rg;
                        const exp = expandedRg[key];
                        return (
                          <>
                            <tr key={key} className={`rg-row${exp?" exp":""}`}
                              onClick={()=>setExpandedRg(p=>({...p,[key]:!p[key]}))}>
                              <td className="exp-td">{exp?"▾":"▸"}</td>
                              <td><span className="rg-tag">{group.rg}</span></td>
                              <td><span className="sub-pill">{group.sub}</span></td>
                              <td><span className="type-pill">Resource Group</span></td>
                              <td className="r mono strong">{fmt(group.total,2)}</td>
                              <td className="r mono">{P.appSvcTotal?(group.total/P.appSvcTotal*100).toFixed(1):0}%</td>
                              <td><Bar pct={P.appSvcGroups[0]?group.total/P.appSvcGroups[0].total*100:0} color="#60a5fa"/></td>
                            </tr>
                            {exp && group.resources.sort((a,b)=>b.cost-a.cost).map((r,ri)=>{
                              const name = r.resourceId?.split("/").pop() || r.resourceName || r.resourceId || "Unknown resource";
                              const type = r.resourceType?.split("/").pop() || r.service;
                              return (
                                <tr key={"as_"+ri} className="drill-res">
                                  <td/>
                                  <td style={{paddingLeft:26}}>
                                    <span className="darrow">↳</span>
                                    <span className="res-name">{name}</span>
                                  </td>
                                  <td className="dim">{r.sub}</td>
                                  <td><span className="type-pill">{type}</span></td>
                                  <td className="r mono">{fmt(r.cost,4)}</td>
                                  <td className="r mono dim">{P.appSvcTotal?(r.cost/P.appSvcTotal*100).toFixed(2):0}%</td>
                                  <td><Bar pct={group.total?r.cost/group.total*100:0} color="#a78bfa"/></td>
                                </tr>
                              );
                            })}
                          </>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </>
          )}

          {/* ═══════════════════════════════ SUBSCRIPTIONS ═══════════════════════════ */}
          {view==="subs" && (
            <>
              <div className="panel">
                <div className="phdr"><span className="ptitle">All Subscriptions — Click to drill in</span></div>
                <table className="tbl">
                  <thead>
                    <tr>
                      <th>Subscription</th>
                      <th className="r">Total Cost</th>
                      <th className="r">App Services</th>
                      <th className="r">% of Total</th>
                      <th>Distribution</th>
                      <th className="r">Top Service</th>
                    </tr>
                  </thead>
                  <tbody>
                    {P.subs.map((s,i) => {
                      const topSvc = Object.entries(s.services).sort((a,b)=>b[1]-a[1])[0];
                      return (
                        <tr key={s.id} className="clickable"
                          onClick={()=>{ setSelectedSubId(s.id); setView("services"); }}>
                          <td>
                            <span className="dot" style={{background:COLORS[i%COLORS.length]}}/>
                            <strong>{s.name}</strong>
                            <div className="sub-id">{s.id?.slice(0,8)}…</div>
                          </td>
                          <td className="r mono strong">{fmt(s.total,2)}</td>
                          <td className="r mono">{fmt(s.appSvcTotal,2)}</td>
                          <td className="r mono">{P.grandTotal?(s.total/P.grandTotal*100).toFixed(1):0}%</td>
                          <td><Bar pct={P.subs[0]?s.total/P.subs[0].total*100:0} color={COLORS[i%COLORS.length]}/></td>
                          <td className="r dim" style={{fontSize:11}}>{topSvc?`${topSvc[0].replace("Azure ","")} (${fmt(topSvc[1])})`:""}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              {/* Sub detail cards */}
              <div className="sub-grid">
                {P.subs.map((s,i)=>(
                  <div key={s.id} className="sub-card" onClick={()=>{setSelectedSubId(s.id);setView("services");}}>
                    <div className="sc-stripe" style={{background:COLORS[i%COLORS.length]}}/>
                    <div className="sc-top">
                      <div className="sc-name">{s.name}</div>
                      <div className="sc-cost">{fmt(s.total)}</div>
                    </div>
                    <div className="sc-id">{s.id}</div>
                    <Bar pct={P.subs[0]?s.total/P.subs[0].total*100:0} color={COLORS[i%COLORS.length]}/>
                    <div className="sc-appsvcs">
                      <span>App Services:</span><strong>{fmt(s.appSvcTotal)}</strong>
                    </div>
                    <div className="sc-services">
                      {Object.entries(s.services).sort((a,b)=>b[1]-a[1]).slice(0,5).map(([svc,cost])=>(
                        <div key={svc} className="sc-svc">
                          <span className="sc-dot" style={{background:P.colorMap[svc]||"#60a5fa"}}/>
                          <span className="sc-svc-name">{svc.replace("Azure ","")}</span>
                          <span className="sc-svc-cost">{fmt(cost)}</span>
                        </div>
                      ))}
                    </div>
                    <div className="sc-cta">Click to drill into services →</div>
                  </div>
                ))}
              </div>
            </>
          )}

          {/* ═══════════════════════════════ BREAKDOWN ═══════════════════════════════ */}
          {view==="breakdown" && (
            <div className="panel">
              <div className="phdr">
                <span className="ptitle">Full Cost Breakdown — Every Line Item</span>
                <span className="psub">{P.rows.length} records · {fmt(P.grandTotal)}</span>
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
                  {P.rows.sort((a,b)=>sortDir==="desc"?b.cost-a.cost:a.cost-b.cost).slice(0,1000).map((r,i)=>(
                    <tr key={i}>
                      <td><span className="dot" style={{background:P.colorMap[r.service]||"#60a5fa"}}/>{r.service}</td>
                      <td className="dim">{r.sub}</td>
                      <td><span className="rg-tag-sm">{r.rg||"—"}</span></td>
                      <td className="mono-sm">{r.resourceId?.split("/").pop()||r.resourceName||"—"}</td>
                      <td className="dim" style={{fontSize:11}}>{r.resourceType?.split("/").pop()||"—"}</td>
                      <td className="r mono strong">{fmt(r.cost,4)}</td>
                    </tr>
                  ))}
                  {P.rows.length>1000&&<tr><td colSpan={6} className="dim" style={{textAlign:"center",padding:12}}>Showing 1,000 of {P.rows.length} rows — use search to narrow</td></tr>}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
