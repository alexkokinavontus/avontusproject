import { useState, useEffect, useCallback, useMemo } from "react";
import { fetchAllData } from "./api/azure";

const BUDGET = 155000;
const COLORS = ["#60a5fa","#a78bfa","#34d399","#fb923c","#f87171","#2dd4bf","#facc15","#e879f9","#4ade80","#f59e0b","#818cf8","#38bdf8"];
const fmt = (n, d = 0) => !n ? "$0" : "$" + Number(n).toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });
const fmtD = s => new Date(s + "T12:00:00Z").toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });

// Parse date from Azure — handles 20260401 (number) or "2026-04-01" (string)
function parseAzureDate(raw) {
  const s = String(raw || "");
  if (!s || s === "null") return null;
  // Numeric like 20260401
  if (/^\d{8}$/.test(s)) return `${s.slice(0,4)}-${s.slice(4,6)}-${s.slice(6,8)}`;
  // Already ISO-ish
  if (s.includes("-")) return s.slice(0,10);
  return null;
}

const AS_TYPES = ["microsoft.web/sites","microsoft.web/serverfarms","microsoft.web/hostingenvironments","microsoft.web/staticsites"];
const AS_SVCS  = ["azure app service","app service","web apps","app service plan"];
const isAppSvc = r => AS_TYPES.some(t => r.resourceType?.toLowerCase().includes(t)) || AS_SVCS.some(s => r.service?.toLowerCase().includes(s));

const PRESETS = [
  { label:"MTD",      fn:()=>{ const n=new Date(); return {start:new Date(n.getFullYear(),n.getMonth(),1).toISOString().slice(0,10),end:n.toISOString().slice(0,10)}; }},
  { label:"Last 30d", fn:()=>{ const n=new Date(),s=new Date(n); s.setDate(s.getDate()-30); return {start:s.toISOString().slice(0,10),end:n.toISOString().slice(0,10)}; }},
  { label:"Last 90d", fn:()=>{ const n=new Date(),s=new Date(n); s.setDate(s.getDate()-90); return {start:s.toISOString().slice(0,10),end:n.toISOString().slice(0,10)}; }},
  { label:"Last 6m",  fn:()=>{ const n=new Date(),s=new Date(n); s.setMonth(s.getMonth()-6); return {start:s.toISOString().slice(0,10),end:n.toISOString().slice(0,10)}; }},
  { label:"YTD",      fn:()=>{ const n=new Date(); return {start:`${n.getFullYear()}-01-01`,end:n.toISOString().slice(0,10)}; }},
  { label:"Last 12m", fn:()=>{ const n=new Date(),s=new Date(n); s.setFullYear(s.getFullYear()-1); return {start:s.toISOString().slice(0,10),end:n.toISOString().slice(0,10)}; }},
];

// ── Mini Components ───────────────────────────────────────────────────────────
function Spark({ vals=[], color="#60a5fa", w=90, h=26 }) {
  if (!vals || vals.length < 2) return null;
  const max = Math.max(...vals, 1);
  const pts = vals.map((v,i)=>`${(i/(vals.length-1))*w},${h-(v/max)*h}`).join(" ");
  return <svg width={w} height={h} style={{overflow:"visible",display:"block"}}><polyline points={pts} fill="none" stroke={color} strokeWidth="1.5" strokeLinejoin="round" strokeLinecap="round"/></svg>;
}

function Bar({ pct, color="#60a5fa" }) {
  return <div className="bar-track"><div className="bar-fill" style={{width:`${Math.min(Math.max(pct,0),100)}%`,background:color}}/></div>;
}

function StackChart({ months, topServices, colorMap }) {
  if (!months.length) return <div className="empty-msg">No trend data</div>;
  const max = Math.max(...months.map(m=>m.total),1);
  return (
    <div className="sc-wrap">
      <div className="sc-bars">
        {months.map((m,i)=>(
          <div key={i} className="sc-col">
            <div className="sc-stack" style={{height:`${(m.total/max)*100}%`}}>
              {topServices.map(svc=>{
                const v=m[svc]||0;
                if (!v) return null;
                return <div key={svc} className="sc-seg" style={{height:`${(v/m.total)*100}%`,background:colorMap[svc]||"#60a5fa"}} title={`${svc}: ${fmt(v)}`}/>;
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
  if (!days || days.length < 2) return <div className="empty-msg">No daily data</div>;
  const vals = days.map(d=>d.total);
  const max = Math.max(...vals,1);
  const W=600,H=70;
  const pts = vals.map((v,i)=>`${(i/(vals.length-1))*W},${H-(v/max)*H}`).join(" ");
  return (
    <div className="dc-wrap">
      <svg viewBox={`0 0 ${W} ${H}`} preserveAspectRatio="none" style={{width:"100%",height:60}}>
        <defs><linearGradient id="dg" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#60a5fa" stopOpacity=".2"/><stop offset="100%" stopColor="#60a5fa" stopOpacity="0"/></linearGradient></defs>
        <polygon points={`0,${H} ${pts} ${W},${H}`} fill="url(#dg)"/>
        <polyline points={pts} fill="none" stroke="#60a5fa" strokeWidth="2" strokeLinejoin="round"/>
      </svg>
      <div className="dc-lbls">
        {days.filter((_,i)=>i%Math.max(1,Math.floor(days.length/6))===0).map((d,i)=>(
          <span key={i}>{new Date(d.date+"T12:00:00Z").toLocaleDateString("en-US",{month:"short",day:"numeric"})}</span>
        ))}
      </div>
    </div>
  );
}

// ── App ───────────────────────────────────────────────────────────────────────
export default function App() {
  const initDates = PRESETS[0].fn();
  const [dates, setDates] = useState(initDates);
  const [preset, setPreset] = useState("MTD");
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [refreshedAt, setRefreshedAt] = useState(null);
  const [view, setView] = useState("overview");
  const [selSubId, setSelSubId] = useState(null);
  const [expSvc, setExpSvc] = useState({});
  const [expRg, setExpRg] = useState({});
  const [search, setSearch] = useState("");
  const [sortDesc, setSortDesc] = useState(true);

  const load = useCallback(async (range) => {
    setLoading(true); setError(null); setExpSvc({}); setExpRg({});
    try {
      const r = await fetchAllData(range.start, range.end);
      setData(r);
      setRefreshedAt(new Date());
    } catch(e) { setError(e.message); }
    finally { setLoading(false); }
  }, []);

  useEffect(() => { load(initDates); }, []);

  const applyPreset = p => { const r=p.fn(); setPreset(p.label); setDates(r); load(r); };
  const applyDate = (k,v) => { const r={...dates,[k]:v}; setPreset(null); setDates(r); load(r); };

  // ── Process ────────────────────────────────────────────────────────────────
  const P = useMemo(() => {
    if (!data) return null;

    const baseRows = selSubId ? data.allDetailed.filter(r=>r.subId===selSubId) : data.allDetailed;
    const q = search.toLowerCase();
    const rows = q ? baseRows.filter(r=>r.service?.toLowerCase().includes(q)||r.rg?.toLowerCase().includes(q)||r.resourceType?.toLowerCase().includes(q)) : baseRows;

    const grandTotal = baseRows.reduce((s,r)=>s+r.cost,0);

    // Services
    const svcMap = {};
    rows.forEach(r => {
      if (!svcMap[r.service]) svcMap[r.service] = { name:r.service, total:0, subs:{}, rgs:{}, resources:[] };
      svcMap[r.service].total += r.cost;
      svcMap[r.service].subs[r.sub] = (svcMap[r.service].subs[r.sub]||0) + r.cost;
      const rg = r.rg || "(no resource group)";
      if (!svcMap[r.service].rgs[rg]) svcMap[r.service].rgs[rg] = { total:0, types:{} };
      svcMap[r.service].rgs[rg].total += r.cost;
      const rt = r.resourceType || "unknown";
      svcMap[r.service].rgs[rg].types[rt] = (svcMap[r.service].rgs[rg].types[rt]||0) + r.cost;
      svcMap[r.service].resources.push(r);
    });
    const services = Object.values(svcMap).sort((a,b)=>sortDesc?b.total-a.total:a.total-b.total);

    // Color map
    const colorMap = {};
    Object.values(svcMap).sort((a,b)=>b.total-a.total).forEach((s,i)=>{ colorMap[s.name]=COLORS[i%COLORS.length]; });

    // Monthly trend
    const monthMap = {};
    const monthSrc = selSubId ? data.allMonthly.filter(r=>r.subId===selSubId) : data.allMonthly;
    monthSrc.forEach(r => {
      const ds = parseAzureDate(r.date);
      if (!ds) return;
      const key = ds.slice(0,7);
      const d = new Date(key+"-01T12:00:00Z");
      const label = d.toLocaleString("default",{month:"short",year:"2-digit"});
      if (!monthMap[key]) monthMap[key]={label,total:0};
      monthMap[key][r.service] = (monthMap[key][r.service]||0)+r.cost;
      monthMap[key].total += r.cost;
    });
    const months = Object.values(monthMap).sort((a,b)=>a.label.localeCompare(b.label));

    // Daily
    const dayMap = {};
    const daySrc = selSubId ? data.allDaily.filter(r=>r.subId===selSubId) : data.allDaily;
    daySrc.forEach(r => {
      const ds = parseAzureDate(r.date);
      if (!ds) return;
      const key = ds.slice(0,10);
      if (!dayMap[key]) dayMap[key]={date:key,total:0};
      dayMap[key].total += r.cost;
    });
    const days = Object.values(dayMap).sort((a,b)=>a.date.localeCompare(b.date));

    // Subscriptions
    const subMap = {};
    data.allDetailed.forEach(r => {
      if (!subMap[r.subId]) subMap[r.subId]={id:r.subId,name:r.sub,total:0,svcs:{},asSvcTotal:0};
      subMap[r.subId].total += r.cost;
      subMap[r.subId].svcs[r.service] = (subMap[r.subId].svcs[r.service]||0)+r.cost;
      if (isAppSvc(r)) subMap[r.subId].asSvcTotal += r.cost;
    });
    const subs = Object.values(subMap).sort((a,b)=>b.total-a.total);

    // App Services
    const asRows = rows.filter(isAppSvc);
    const asRgMap = {};
    asRows.forEach(r => {
      const rg = r.rg || "(no resource group)";
      const key = r.subId+"::"+rg;
      if (!asRgMap[key]) asRgMap[key]={rg,sub:r.sub,subId:r.subId,total:0,resources:[]};
      asRgMap[key].total += r.cost;
      asRgMap[key].resources.push(r);
    });
    const asGroups = Object.values(asRgMap).sort((a,b)=>b.total-a.total);
    const asTotal = asRows.reduce((s,r)=>s+r.cost,0);

    return { rows, grandTotal, services, colorMap, months, days, subs, asRows, asGroups, asTotal };
  }, [data, selSubId, search, sortDesc]);

  const selSub = data?.subscriptions?.find(s=>s.subscriptionId===selSubId);
  const now = new Date();
  const dayOfMonth = now.getDate();
  const daysInMonth = new Date(now.getFullYear(),now.getMonth()+1,0).getDate();

  if (loading) return (
    <div className="splash">
      <div className="sp-wrap"><div className="sp-ring"/><div className="sp-logo">A</div></div>
      <div className="sp-title">Loading Azure Cost Data</div>
      <div className="sp-sub">{fmtD(dates.start)} → {fmtD(dates.end)}</div>
    </div>
  );

  if (error) return (
    <div className="splash">
      <div className="err-card">
        <div>⚠</div>
        <h3>Connection Error</h3>
        <p className="err-msg">{error}</p>
        <button className="btn-p" onClick={()=>load(dates)}>↻ Retry</button>
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
          <div><div className="sb-name">AzureReader</div><div className="sb-ten">avontus.com</div></div>
        </div>
        <nav className="sb-nav">
          {[{id:"overview",ico:"▣",lbl:"Overview"},{id:"services",ico:"⬡",lbl:"By Service"},{id:"appsvcs",ico:"⬢",lbl:"App Services"},{id:"subs",ico:"◈",lbl:"Subscriptions"},{id:"breakdown",ico:"≡",lbl:"Full Breakdown"}].map(v=>(
            <button key={v.id} className={`sb-item${view===v.id?" on":""}`} onClick={()=>setView(v.id)}>
              <span className="sb-ico">{v.ico}</span>{v.lbl}
            </button>
          ))}
        </nav>
        <div className="sb-div">SUBSCRIPTIONS</div>
        <div className="sb-subs">
          <button className={`sb-sub${!selSubId?" sel":""}`} onClick={()=>setSelSubId(null)}>
            <span className="dot" style={{background:"#60a5fa",width:6,height:6,borderRadius:"50%",flexShrink:0}}/>
            <span className="sb-sn">All subscriptions</span>
            <span className="sb-sc">{fmt(P.grandTotal)}</span>
          </button>
          {P.subs.map((s,i)=>(
            <button key={s.id} className={`sb-sub${selSubId===s.id?" sel":""}`} onClick={()=>setSelSubId(s.id)}>
              <span className="dot" style={{background:COLORS[i%COLORS.length],width:6,height:6,borderRadius:"50%",flexShrink:0}}/>
              <span className="sb-sn">{s.name}</span>
              <span className="sb-sc">{fmt(s.total)}</span>
            </button>
          ))}
        </div>
        <div className="sb-foot">
          <div>bd98204b · AzureReader</div>
          {refreshedAt&&<div>{refreshedAt.toLocaleTimeString()}</div>}
          {data?.errors?.length>0&&<div style={{color:"var(--amber)"}}>⚠ {data.errors.length} warn</div>}
        </div>
      </aside>

      {/* ── Main ── */}
      <div className="main">
        {/* Control bar */}
        <div className="cbar">
          <div className="cbar-l">
            <span className="cbar-title">{selSubId ? selSub?.displayName : "All Subscriptions"}</span>
            <span className="cbar-range">{fmtD(dates.start)} — {fmtD(dates.end)}</span>
          </div>
          <div className="cbar-r">
            <div className="presets">
              {PRESETS.map(p=>(
                <button key={p.label} className={`ps${preset===p.label?" on":""}`} onClick={()=>applyPreset(p)}>{p.label}</button>
              ))}
            </div>
            <div className="dr">
              <input type="date" className="di" value={dates.start} max={dates.end} onChange={e=>applyDate("start",e.target.value)}/>
              <span className="da">→</span>
              <input type="date" className="di" value={dates.end} min={dates.start} onChange={e=>applyDate("end",e.target.value)}/>
            </div>
            <button className="rb" onClick={()=>load(dates)}>↻</button>
          </div>
        </div>

        {/* Search bar */}
        {["services","appsvcs","breakdown"].includes(view)&&(
          <div className="sbar">
            <span className="si">⌕</span>
            <input className="si-in" placeholder="Filter by service, resource group, type…" value={search} onChange={e=>setSearch(e.target.value)}/>
            {search&&<button className="si-x" onClick={()=>setSearch("")}>×</button>}
            <select className="si-sort" value={sortDesc?"desc":"asc"} onChange={e=>setSortDesc(e.target.value==="desc")}>
              <option value="desc">Highest cost first</option>
              <option value="asc">Lowest cost first</option>
            </select>
            <span className="si-info">{P.services.length} services · {P.rows.length} records · {fmt(P.grandTotal)}</span>
          </div>
        )}

        <div className="content">

          {/* ══ OVERVIEW ══ */}
          {view==="overview"&&P&&(
            <>
              <div className="kpis">
                <div className="kpi blue">
                  <div className="kpi-l">Total Spend</div>
                  <div className="kpi-v">{fmt(P.grandTotal)}</div>
                  <div className="kpi-s">{fmtD(dates.start)} – {fmtD(dates.end)}</div>
                </div>
                <div className="kpi" style={{"--ac":P.grandTotal>BUDGET?"var(--red)":"var(--green)"}}>
                  <div className="kpi-l">Budget ({fmt(BUDGET)})</div>
                  <div className="kpi-v" style={{color:P.grandTotal>BUDGET?"var(--red)":"var(--green)"}}>{P.grandTotal>BUDGET?`+${fmt(P.grandTotal-BUDGET)} over`:`${fmt(BUDGET-P.grandTotal)} under`}</div>
                  <Bar pct={P.grandTotal/BUDGET*100} color={P.grandTotal>BUDGET?"var(--red)":"var(--green)"}/>
                </div>
                <div className="kpi purple">
                  <div className="kpi-l">App Services</div>
                  <div className="kpi-v">{fmt(P.asTotal)}</div>
                  <div className="kpi-s">{P.grandTotal?(P.asTotal/P.grandTotal*100).toFixed(1):0}% of total · {P.asRows.length} resources</div>
                </div>
                <div className="kpi amber">
                  <div className="kpi-l">Month Progress</div>
                  <div className="kpi-v">{((dayOfMonth/daysInMonth)*100).toFixed(0)}%</div>
                  <Bar pct={dayOfMonth/daysInMonth*100} color="var(--amber)"/>
                </div>
              </div>

              <div className="two-col">
                <div className="panel">
                  <div className="phdr"><span className="pt">Monthly Cost Trend</span><span className="ps">{P.months.length} months</span></div>
                  <StackChart months={P.months} topServices={P.services.slice(0,8).map(s=>s.name)} colorMap={P.colorMap}/>
                  <div className="leg">{P.services.slice(0,6).map(s=><div key={s.name} className="leg-i"><span style={{width:7,height:7,borderRadius:"50%",background:P.colorMap[s.name],display:"inline-block",marginRight:5,flexShrink:0}}/>{s.name.replace("Azure ","")}</div>)}</div>
                </div>
                <div className="panel">
                  <div className="phdr"><span className="pt">Daily Spend</span><span className="ps">{P.days.length} days</span></div>
                  <DailyChart days={P.days}/>
                  <div className="dstats">
                    <div className="ds"><span>Avg/day</span><strong>{fmt(P.days.reduce((s,d)=>s+d.total,0)/(P.days.length||1))}</strong></div>
                    <div className="ds"><span>Peak day</span><strong>{fmt(Math.max(...(P.days.map(d=>d.total)||[0]),0))}</strong></div>
                    <div className="ds"><span>Tracked</span><strong>{P.days.length}d</strong></div>
                  </div>
                </div>
              </div>

              <div className="panel">
                <div className="phdr"><span className="pt">Top Services</span><button className="vm" onClick={()=>setView("services")}>Drill down →</button></div>
                <table className="tbl">
                  <thead><tr><th>Service</th><th className="r">Cost</th><th className="r">% Share</th><th>Distribution</th><th className="r">Subscriptions</th></tr></thead>
                  <tbody>
                    {P.services.slice(0,12).map((s,i)=>(
                      <tr key={s.name} className="click" onClick={()=>{setView("services");setTimeout(()=>setExpSvc({[s.name]:true}),50);}}>
                        <td><span className="dot" style={{background:P.colorMap[s.name]}}/><strong>{s.name}</strong></td>
                        <td className="r mono hi">{fmt(s.total,2)}</td>
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

          {/* ══ BY SERVICE ══ */}
          {view==="services"&&P&&(
            <div className="panel">
              <div className="phdr"><span className="pt">Cost by Service — Click to Expand</span><span className="ps">{P.services.length} services · {fmt(P.grandTotal)}</span></div>
              <table className="tbl">
                <thead>
                  <tr><th style={{width:28}}/><th>Service</th><th className="r">Cost</th><th className="r">% Total</th><th>Share</th><th className="r">Res. Groups</th><th className="r">Records</th></tr>
                </thead>
                <tbody>
                  {P.services.map(s=>{
                    const ex=expSvc[s.name];
                    return (<>
                      <tr key={s.name} className={`svc-row${ex?" ex":""}`} onClick={()=>setExpSvc(p=>({...p,[s.name]:!p[s.name]}))}>
                        <td className="ex-td">{ex?"▾":"▸"}</td>
                        <td><span className="dot" style={{background:P.colorMap[s.name]}}/><strong>{s.name}</strong></td>
                        <td className="r mono hi">{fmt(s.total,2)}</td>
                        <td className="r mono">{P.grandTotal?(s.total/P.grandTotal*100).toFixed(2):0}%</td>
                        <td><Bar pct={P.services[0]?s.total/P.services[0].total*100:0} color={P.colorMap[s.name]}/></td>
                        <td className="r">{Object.keys(s.rgs).length}</td>
                        <td className="r">{s.resources.length}</td>
                      </tr>
                      {ex&&<>
                        {/* By Subscription */}
                        <tr className="dhdr"><td/><td colSpan={6}><span className="dl">📋 By Subscription</span></td></tr>
                        {Object.entries(s.subs).sort((a,b)=>b[1]-a[1]).map(([sub,cost])=>(
                          <tr key={"s_"+sub} className="dsub">
                            <td/><td style={{paddingLeft:30}}><span className="da">↳</span><span className="spill">{sub}</span></td>
                            <td className="r mono">{fmt(cost,2)}</td>
                            <td className="r mono dim">{s.total?(cost/s.total*100).toFixed(1):0}%</td>
                            <td><Bar pct={s.total?cost/s.total*100:0} color={P.colorMap[s.name]+"99"}/></td>
                            <td/><td/>
                          </tr>
                        ))}
                        {/* By Resource Group */}
                        <tr className="dhdr"><td/><td colSpan={6}><span className="dl">📁 By Resource Group</span></td></tr>
                        {Object.entries(s.rgs).sort((a,b)=>b[1].total-a[1].total).slice(0,30).map(([rg,rgData])=>{
                          const rk=s.name+"::"+rg;
                          const rex=expRg[rk];
                          return (<>
                            <tr key={"rg_"+rg} className={`drg${rex?" ex":""}`} onClick={e=>{e.stopPropagation();setExpRg(p=>({...p,[rk]:!p[rk]}));}}>
                              <td/><td style={{paddingLeft:30}}><span className="da">{rex?"▾":"▸"}</span><span className="rgt">{rg}</span></td>
                              <td className="r mono">{fmt(rgData.total,2)}</td>
                              <td className="r mono dim">{s.total?(rgData.total/s.total*100).toFixed(1):0}%</td>
                              <td><Bar pct={s.total?rgData.total/s.total*100:0} color={P.colorMap[s.name]+"77"}/></td>
                              <td/><td className="r dim">{Object.keys(rgData.types).length} types</td>
                            </tr>
                            {rex&&Object.entries(rgData.types).sort((a,b)=>b[1]-a[1]).map(([rt,cost])=>(
                              <tr key={"rt_"+rt} className="dres">
                                <td/><td style={{paddingLeft:52}}><span className="da">↳</span><span className="rtt">{rt.split("/").pop()}</span><span className="rtfull">{rt}</span></td>
                                <td className="r mono">{fmt(cost,2)}</td>
                                <td className="r mono dim">{s.total?(cost/s.total*100).toFixed(2):0}%</td>
                                <td/><td/><td/>
                              </tr>
                            ))}
                          </>);
                        })}
                      </>}
                    </>);
                  })}
                </tbody>
              </table>
            </div>
          )}

          {/* ══ APP SERVICES ══ */}
          {view==="appsvcs"&&P&&(
            <>
              <div className="kpis">
                <div className="kpi blue"><div className="kpi-l">App Service Spend</div><div className="kpi-v">{fmt(P.asTotal)}</div><div className="kpi-s">{P.grandTotal?(P.asTotal/P.grandTotal*100).toFixed(1):0}% of total</div></div>
                <div className="kpi purple"><div className="kpi-l">Resource Groups</div><div className="kpi-v">{P.asGroups.length}</div><div className="kpi-s">With App Services</div></div>
                <div className="kpi green"><div className="kpi-l">Resources</div><div className="kpi-v">{P.asRows.length}</div><div className="kpi-s">Web apps, plans, slots</div></div>
                <div className="kpi amber"><div className="kpi-l">Subscriptions</div><div className="kpi-v">{new Set(P.asRows.map(r=>r.subId)).size}</div><div className="kpi-s">With App Services</div></div>
              </div>
              {P.asGroups.length===0?(
                <div className="panel"><div className="empty-lg">
                  <div style={{fontSize:28}}>⬢</div>
                  <div style={{fontFamily:"'Syne',sans-serif",fontSize:16,fontWeight:700}}>No App Service costs found</div>
                  <div style={{fontSize:12,color:"var(--t3)",maxWidth:400,lineHeight:1.6}}>App Service costs may appear under "Azure App Service" in the By Service view. Try expanding the date range or check that subscriptions are included.</div>
                  <button className="btn-p" onClick={()=>setView("services")}>Check By Service →</button>
                </div></div>
              ):(
                <div className="panel">
                  <div className="phdr"><span className="pt">App Services — By Resource Group → Resource Type</span><span className="ps">Click to expand</span></div>
                  <table className="tbl">
                    <thead><tr><th style={{width:28}}/><th>Resource Group / Resource</th><th>Subscription</th><th>Type</th><th className="r">Cost</th><th className="r">% of AS</th><th>Share</th></tr></thead>
                    <tbody>
                      {P.asGroups.map(g=>{
                        const gk=g.subId+"::"+g.rg;
                        const gex=expRg[gk];
                        return (<>
                          <tr key={gk} className={`rg-row${gex?" ex":""}`} onClick={()=>setExpRg(p=>({...p,[gk]:!p[gk]}))}>
                            <td className="ex-td">{gex?"▾":"▸"}</td>
                            <td><span className="rgt">{g.rg}</span></td>
                            <td><span className="spill">{g.sub}</span></td>
                            <td><span className="tpill">Resource Group</span></td>
                            <td className="r mono hi">{fmt(g.total,2)}</td>
                            <td className="r mono">{P.asTotal?(g.total/P.asTotal*100).toFixed(1):0}%</td>
                            <td><Bar pct={P.asGroups[0]?g.total/P.asGroups[0].total*100:0} color="#60a5fa"/></td>
                          </tr>
                          {gex&&g.resources.sort((a,b)=>b.cost-a.cost).map((r,ri)=>{
                            const nm = r.resourceType?.split("/").pop() || r.service;
                            return (
                              <tr key={"as_"+ri} className="dres">
                                <td/>
                                <td style={{paddingLeft:26}}><span className="da">↳</span><span className="rtt">{nm}</span></td>
                                <td className="dim">{r.sub}</td>
                                <td><span className="tpill">{r.resourceType?.split("/").pop()||"—"}</span></td>
                                <td className="r mono">{fmt(r.cost,2)}</td>
                                <td className="r mono dim">{P.asTotal?(r.cost/P.asTotal*100).toFixed(2):0}%</td>
                                <td><Bar pct={g.total?r.cost/g.total*100:0} color="#a78bfa"/></td>
                              </tr>
                            );
                          })}
                        </>);
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </>
          )}

          {/* ══ SUBSCRIPTIONS ══ */}
          {view==="subs"&&P&&(
            <>
              <div className="panel">
                <div className="phdr"><span className="pt">All Subscriptions</span><span className="ps">Click to filter all views</span></div>
                <table className="tbl">
                  <thead><tr><th>Subscription</th><th className="r">Total Cost</th><th className="r">App Services</th><th className="r">% of Total</th><th>Share</th><th className="r">Top Service</th></tr></thead>
                  <tbody>
                    {P.subs.map((s,i)=>{
                      const top=Object.entries(s.svcs).sort((a,b)=>b[1]-a[1])[0];
                      return (
                        <tr key={s.id} className="click" onClick={()=>{setSelSubId(s.id);setView("services");}}>
                          <td><span className="dot" style={{background:COLORS[i%COLORS.length]}}/><strong>{s.name}</strong><div style={{fontFamily:"'DM Mono',monospace",fontSize:9,color:"var(--t3)",marginTop:1}}>{s.id?.slice(0,8)}…</div></td>
                          <td className="r mono hi">{fmt(s.total,2)}</td>
                          <td className="r mono" style={{color:"var(--purple)"}}>{fmt(s.asSvcTotal,2)}</td>
                          <td className="r mono">{P.grandTotal?(s.total/P.grandTotal*100).toFixed(1):0}%</td>
                          <td><Bar pct={P.subs[0]?s.total/P.subs[0].total*100:0} color={COLORS[i%COLORS.length]}/></td>
                          <td className="r dim" style={{fontSize:11}}>{top?`${top[0].replace("Azure ","")} · ${fmt(top[1])}`:""}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              <div className="sub-grid">
                {P.subs.map((s,i)=>(
                  <div key={s.id} className="sub-card" onClick={()=>{setSelSubId(s.id);setView("appsvcs");}}>
                    <div className="sc-str" style={{background:COLORS[i%COLORS.length]}}/>
                    <div className="sc-top"><div className="sc-nm">{s.name}</div><div className="sc-cost">{fmt(s.total)}</div></div>
                    <div className="sc-id">{s.id?.slice(0,36)}</div>
                    <Bar pct={P.subs[0]?s.total/P.subs[0].total*100:0} color={COLORS[i%COLORS.length]}/>
                    <div className="sc-as">App Services: <strong>{fmt(s.asSvcTotal)}</strong></div>
                    <div className="sc-svcs">
                      {Object.entries(s.svcs).sort((a,b)=>b[1]-a[1]).slice(0,5).map(([svc,cost])=>(
                        <div key={svc} className="sc-sv">
                          <span className="dot" style={{background:P.colorMap[svc]||"#60a5fa",width:5,height:5,borderRadius:"50%",flexShrink:0}}/>
                          <span className="sc-svn">{svc.replace("Azure ","")}</span>
                          <span className="sc-svc">{fmt(cost)}</span>
                        </div>
                      ))}
                    </div>
                    <div className="sc-cta">Click → App Services breakdown</div>
                  </div>
                ))}
              </div>
            </>
          )}

          {/* ══ BREAKDOWN ══ */}
          {view==="breakdown"&&P&&(
            <div className="panel">
              <div className="phdr"><span className="pt">Full Cost Breakdown — Every Record</span><span className="ps">{P.rows.length} records · {fmt(P.grandTotal)}</span></div>
              <table className="tbl">
                <thead><tr><th>Service</th><th>Subscription</th><th>Resource Group</th><th>Resource Type</th><th className="r">Cost</th></tr></thead>
                <tbody>
                  {P.rows.sort((a,b)=>sortDesc?b.cost-a.cost:a.cost-b.cost).slice(0,1000).map((r,i)=>(
                    <tr key={i}>
                      <td><span className="dot" style={{background:P.colorMap[r.service]||"#60a5fa"}}/>{r.service}</td>
                      <td className="dim">{r.sub}</td>
                      <td><span className="rgt-sm">{r.rg||"—"}</span></td>
                      <td className="dim" style={{fontSize:11}}>{r.resourceType?.split("/").pop()||"—"}</td>
                      <td className="r mono hi">{fmt(r.cost,4)}</td>
                    </tr>
                  ))}
                  {P.rows.length>1000&&<tr><td colSpan={5} className="dim" style={{textAlign:"center",padding:12}}>Showing 1,000 of {P.rows.length} — use filters to narrow</td></tr>}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
