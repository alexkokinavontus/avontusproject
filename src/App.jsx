import { useState, useEffect, useCallback } from "react";
import { fetchAllData } from "./api/azure";

const fmt = (n, d = 0) => n == null ? "$0" : "$" + Number(n).toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });
const COLORS = ["#5b8af5","#a78bfa","#f59e0b","#34d399","#f87171","#38bdf8","#fb923c","#e879f9","#4ade80","#facc15"];

function Sparkline({ values = [], color = "#5b8af5", width = 120, height = 32 }) {
  if (!values.length) return null;
  const max = Math.max(...values, 1);
  const pts = values.map((v, i) => `${(i/(values.length-1))*width},${height-(v/max)*height}`).join(" ");
  return <svg width={width} height={height} style={{overflow:"visible"}}><polyline points={pts} fill="none" stroke={color} strokeWidth="1.5" strokeLinejoin="round" strokeLinecap="round"/></svg>;
}

function BarChart({ months, subscriptions }) {
  if (!months.length) return <div className="no-data">No trend data</div>;
  const max = Math.max(...months.map(m => m.total), 1);
  return (
    <div className="bar-chart">
      {months.map((m, i) => (
        <div key={i} className="bar-col">
          <div className="bar-stack">
            {subscriptions.map((sub, si) => {
              const val = m[sub] || 0;
              const h = (val / max) * 100;
              return h > 0 ? <div key={sub} className="bar-seg" style={{ height: `${h}%`, background: COLORS[si % COLORS.length] }} title={`${sub}: ${fmt(val)}`} /> : null;
            })}
          </div>
          <div className="bar-label">{m.label}</div>
        </div>
      ))}
    </div>
  );
}

function ProgressBar({ value, max, color = "#5b8af5" }) {
  const pct = Math.min((value / max) * 100, 100);
  return <div className="progress-track"><div className="progress-fill" style={{ width: `${pct}%`, background: color }} /></div>;
}

export default function App() {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [lastRefresh, setLastRefresh] = useState(null);
  const [activeView, setActiveView] = useState("Overview");
  const [expandedSubs, setExpandedSubs] = useState({});
  const [showAllMonths, setShowAllMonths] = useState(false);

  const load = useCallback(async () => {
    setLoading(true); setError(null);
    try {
      const result = await fetchAllData();
      setData(result);
      setLastRefresh(new Date());
    } catch (e) { setError(e.message); }
    finally { setLoading(false); }
  }, []);

  useEffect(() => { load(); }, []);

  const totalCost = data?.costs.reduce((s,c) => s+(c.data?.properties?.rows||[]).reduce((a,r)=>a+(r[0]||0),0),0)||0;
  const BUDGET = 155000;
  const overBudget = totalCost > BUDGET;
  const now = new Date();
  const daysInMonth = new Date(now.getFullYear(), now.getMonth()+1, 0).getDate();
  const dayOfMonth = now.getDate();
  const forecast = totalCost > 0 ? (totalCost/dayOfMonth)*daysInMonth : 0;

  const monthMap = {};
  const subNames = [];
  (data?.trends||[]).forEach((t,idx) => {
    const name = t.subscription?.displayName||`Sub${idx}`;
    if (!subNames.includes(name)) subNames.push(name);
    (t.data?.properties?.rows||[]).forEach(r => {
      if (!r[1]) return;
      const d = new Date(r[1]);
      const key = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
      const label = d.toLocaleString("default",{month:"short",year:"2-digit"});
      if (!monthMap[key]) monthMap[key] = { label, total:0 };
      monthMap[key][name] = (monthMap[key][name]||0)+(r[0]||0);
      monthMap[key].total += r[0]||0;
    });
  });
  const allMonths = Object.values(monthMap).sort((a,b)=>a.label.localeCompare(b.label));
  const months = showAllMonths ? allMonths : allMonths.slice(-6);

  const subRows = (data?.costs||[]).map((c,i) => {
    const rows = c.data?.properties?.rows||[];
    const total = rows.reduce((s,r)=>s+(r[0]||0),0);
    const services = {};
    rows.forEach(r => { const svc=r[1]||"Other"; services[svc]=(services[svc]||0)+(r[0]||0); });
    return { id:c.subscription?.subscriptionId, name:c.subscription?.displayName||c.subscription?.subscriptionId, total, color:COLORS[i%COLORS.length], services:Object.entries(services).sort((a,b)=>b[1]-a[1]), state:c.subscription?.state };
  }).sort((a,b)=>b.total-a.total);

  const grandTotal = subRows.reduce((s,r)=>s+r.total,0);
  const subSparklines = {};
  (data?.trends||[]).forEach(t => { subSparklines[t.subscription?.displayName]=(t.data?.properties?.rows||[]).map(r=>r[0]||0); });

  if (loading) return (
    <div className="splash">
      <div className="splash-ring"/>
      <div className="splash-logo">A</div>
      <div className="splash-text">Loading Azure data…</div>
    </div>
  );

  if (error) return (
    <div className="splash">
      <div className="error-box">
        <div className="err-icon">⚠</div>
        <div className="err-title">Connection Error</div>
        <div className="err-msg">{error}</div>
        <button className="btn-primary" onClick={load}>Retry</button>
      </div>
    </div>
  );

  return (
    <div className="shell">
      <aside className="sidebar">
        <div className="sb-logo"><div className="sb-mark">A</div><div><div className="sb-name">AzureReader</div><div className="sb-tenant">avontus.com</div></div></div>
        <nav className="sb-nav">
          {["Overview","Costs","Resources","Subscriptions"].map(v=>(
            <button key={v} className={`sb-item ${activeView===v?"active":""}`} onClick={()=>setActiveView(v)}>
              <span className="sb-icon">{v==="Overview"?"◉":v==="Costs"?"💳":v==="Resources"?"🗄":"📋"}</span>{v}
            </button>
          ))}
        </nav>
        <div className="sb-section">SUBSCRIPTIONS</div>
        <div className="sb-subs">
          {(data?.subscriptions||[]).slice(0,8).map((s,i)=>(
            <div key={s.subscriptionId} className="sb-sub">
              <span className="sb-dot" style={{background:COLORS[i%COLORS.length]}}/>
              <span className="sb-sub-name">{s.displayName}</span>
            </div>
          ))}
          {(data?.subscriptions||[]).length>8&&<div className="sb-more">+{data.subscriptions.length-8} more</div>}
        </div>
        <div className="sb-footer"><div>Tenant: bd98204b</div><div>App: AzureReader</div></div>
      </aside>

      <div className="main">
        <header className="topbar">
          <div className="tb-left">
            <div>
              <div className="tb-month">April 2026</div>
              {lastRefresh&&<div className="tb-time">Live data as of {lastRefresh.toLocaleDateString("en-US",{month:"short",day:"numeric"})}, {lastRefresh.toLocaleTimeString()}</div>}
            </div>
          </div>
          <div className="tb-right">
            <div className="tb-badge live">● Live</div>
            <div className="tb-badge day">Day {dayOfMonth} of {daysInMonth}</div>
            <button className="tb-btn" onClick={load} title="Refresh">↻</button>
          </div>
        </header>

        <div className="content">
          {activeView==="Overview"&&<>
            <div className="hero-row">
              <div className="hero-card">
                <div className="hc-label">$ Actual Spend</div>
                <div className="hc-value blue">{fmt(totalCost)}</div>
                <div className="hc-trend" style={{color:overBudget?"#f87171":"#34d399"}}>
                  {overBudget?`▲ ${((totalCost/BUDGET-1)*100).toFixed(1)}% over budget`:`▼ ${(100-totalCost/BUDGET*100).toFixed(1)}% under budget`}
                </div>
                <Sparkline values={allMonths.map(m=>m.total)} color="#5b8af5"/>
              </div>
              <div className="hero-card">
                <div className="hc-label">◎ Month-End Forecast</div>
                <div className="hc-value purple">{fmt(forecast)}</div>
                <div className="hc-trend" style={{color:forecast>BUDGET?"#f87171":"#6b7280"}}>
                  {forecast>BUDGET?`+${fmt(forecast-BUDGET)} over budget`:`${fmt(BUDGET-forecast)} remaining`}
                </div>
                <Sparkline values={allMonths.map(m=>m.total)} color="#a78bfa"/>
              </div>
            </div>

            <div className="section-card">
              <div className="sc-row"><span className="sc-label">Month Progress</span><span className="sc-pct">{((dayOfMonth/daysInMonth)*100).toFixed(0)}%</span></div>
              <ProgressBar value={dayOfMonth} max={daysInMonth} color="#5b8af5"/>
            </div>

            <div className="panel">
              <div className="panel-header">
                <span className="panel-title">Monthly Cost Trend <span className="live-tag">Live</span></span>
                <button className="toggle-btn" onClick={()=>setShowAllMonths(!showAllMonths)}>
                  {showAllMonths?"▸ Show 6 months":"▸ Show all months"}
                </button>
              </div>
              <BarChart months={months} subscriptions={subNames}/>
              <div className="legend">
                {subNames.slice(0,6).map((n,i)=>(
                  <div key={n} className="legend-item"><span className="legend-dot" style={{background:COLORS[i%COLORS.length]}}/>{n.split(" ").slice(0,2).join(" ")}</div>
                ))}
              </div>
            </div>

            <div className="panel">
              <div className="panel-header">
                <div>
                  <div className="panel-title">Subscription Cost Overview</div>
                  <div className="panel-sub">{data?.subscriptions?.length} subscriptions across avontus.com</div>
                </div>
                <div className="panel-total">{fmt(grandTotal)}<div className="panel-total-sub">this month</div></div>
              </div>
              <table className="cost-table">
                <thead>
                  <tr>
                    <th>PRODUCT</th>
                    <th className="r">APR 26</th>
                    <th className="r">SHARE</th>
                    <th className="r">TREND</th>
                  </tr>
                </thead>
                <tbody>
                  <tr className="total-row">
                    <td>▸ All Products Total</td>
                    <td className="r"><strong>{fmt(grandTotal)}</strong> <span className="live-tag">Live</span></td>
                    <td className="r">100%</td>
                    <td className="r"/>
                  </tr>
                  <tr className="budget-row">
                    <td>Total Budget</td>
                    <td className="r budget-val">{fmt(BUDGET)}</td>
                    <td className="r"/>
                    <td className="r"/>
                  </tr>
                  {subRows.map(sub=>(
                    <>
                      <tr key={sub.id} className="sub-row" onClick={()=>setExpandedSubs(p=>({...p,[sub.id]:!p[sub.id]}))}>
                        <td>
                          <span className="expand-icon">{expandedSubs[sub.id]?"▾":"▸"}</span>
                          <span className="sub-dot-sm" style={{background:sub.color}}/>
                          {sub.name}
                        </td>
                        <td className="r cost-val">{fmt(sub.total)}</td>
                        <td className="r">
                          <div className="inline-bar-wrap">
                            <div className="inline-bar" style={{width:`${grandTotal?(sub.total/grandTotal*100):0}%`,background:sub.color}}/>
                            <span>{grandTotal?(sub.total/grandTotal*100).toFixed(1):0}%</span>
                          </div>
                        </td>
                        <td className="r"><Sparkline values={subSparklines[sub.name]||[]} color={sub.color} width={80} height={24}/></td>
                      </tr>
                      {expandedSubs[sub.id]&&sub.services.map(([svc,cost])=>(
                        <tr key={svc} className="svc-row">
                          <td style={{paddingLeft:48}}>{svc}</td>
                          <td className="r">{fmt(cost,2)}</td>
                          <td className="r">{sub.total?(cost/sub.total*100).toFixed(1):0}%</td>
                          <td className="r"/>
                        </tr>
                      ))}
                    </>
                  ))}
                </tbody>
              </table>
            </div>
          </>}

          {activeView==="Costs"&&(
            <div className="panel">
              <div className="panel-header"><span className="panel-title">Full Cost Breakdown — All Services</span></div>
              <table className="cost-table">
                <thead><tr><th>SUBSCRIPTION</th><th>SERVICE</th><th>RESOURCE GROUP</th><th className="r">COST (MTD)</th></tr></thead>
                <tbody>
                  {(data?.costs||[]).flatMap((c,ci)=>
                    (c.data?.properties?.rows||[]).map(r=>({sub:c.subscription?.displayName,svc:r[1]||"Unknown",rg:r[2]||"Unknown",cost:r[0]||0,color:COLORS[ci%COLORS.length]}))
                  ).sort((a,b)=>b.cost-a.cost).map((r,i)=>(
                    <tr key={i}><td><span className="sub-dot-sm" style={{background:r.color}}/>{r.sub}</td><td>{r.svc}</td><td>{r.rg}</td><td className="r cost-val">{fmt(r.cost,2)}</td></tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}

          {activeView==="Resources"&&(
            <div className="panel">
              <div className="panel-header">
                <span className="panel-title">Resources by Type</span>
                <span className="panel-meta">{data?.resourcesByType?.data?.rows?.reduce((s,r)=>s+(r[1]||0),0)||0} total resources</span>
              </div>
              {!(data?.resourcesByType?.data?.rows||[]).length ? (
                <div className="no-data">No resource data — ensure Reader role is assigned to AzureReader service principal on all subscriptions.</div>
              ) : (
                <table className="cost-table">
                  <thead><tr><th>RESOURCE TYPE</th><th className="r">COUNT</th></tr></thead>
                  <tbody>
                    {(data?.resourcesByType?.data?.rows||[]).map((r,i)=>(
                      <tr key={i}><td><span className="sub-dot-sm" style={{background:COLORS[i%COLORS.length]}}/>{r[0]?.split("/").pop()}</td><td className="r">{r[1]?.toLocaleString()}</td></tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          )}

          {activeView==="Subscriptions"&&(
            <div className="sub-cards">
              {subRows.map((sub,i)=>(
                <div key={sub.id} className="sub-card">
                  <div className="scard-accent" style={{background:sub.color}}/>
                  <div className="scard-top">
                    <div><div className="scard-name">{sub.name}</div><div className="scard-id">{sub.id?.slice(0,8)}…</div></div>
                    <span className={`status-badge ${sub.state==="Enabled"?"green":"red"}`}>{sub.state||"Unknown"}</span>
                  </div>
                  <div className="scard-cost">{fmt(sub.total)}</div>
                  <div className="scard-cost-label">Month-to-date</div>
                  <ProgressBar value={sub.total} max={Math.max(grandTotal/subRows.length*2,1)} color={sub.color}/>
                  <div className="scard-services">
                    {sub.services.slice(0,5).map(([svc,cost])=>(
                      <div key={svc} className="scard-svc-row"><span>{svc}</span><span>{fmt(cost)}</span></div>
                    ))}
                  </div>
                  <Sparkline values={subSparklines[sub.name]||[]} color={sub.color} width={220} height={36}/>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
