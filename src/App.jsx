import { useState, useEffect, useCallback, useMemo } from "react";
import { fetchAllData, parseAzureDate, TENANTS, fetchAllInvoices, fetchInvoiceTransactions } from "./api/azure";
import { getStoredUser, signIn, signOut } from "./auth/msal";

const BUDGET_TOTAL = 155000;
const SVC_COLORS = ["#60a5fa","#a78bfa","#34d399","#fb923c","#f87171","#2dd4bf","#facc15","#e879f9","#4ade80","#f59e0b","#818cf8","#38bdf8","#f472b6","#84cc16","#22d3ee"];
const fmt = (n,d=0) => !n?"$0":"$"+Number(n).toLocaleString("en-US",{minimumFractionDigits:d,maximumFractionDigits:d});
const fmtD = s => new Date(s+"T12:00:00Z").toLocaleDateString("en-US",{month:"short",day:"numeric",year:"numeric"});

const AS_TYPES = ["microsoft.web/sites","microsoft.web/serverfarms","microsoft.web/hostingenvironments","microsoft.web/staticsites"];
const AS_SVCS  = ["azure app service","app service","web apps","app service plan"];
const isAS = r => AS_TYPES.some(t=>r.resourceType?.toLowerCase().includes(t))||AS_SVCS.some(s=>r.service?.toLowerCase().includes(s));

const PRESETS = [
  {label:"MTD",      fn:()=>{const n=new Date();return{start:new Date(n.getFullYear(),n.getMonth(),1).toISOString().slice(0,10),end:n.toISOString().slice(0,10)};}},
  {label:"Last 30d", fn:()=>{const n=new Date(),s=new Date(n);s.setDate(s.getDate()-30);return{start:s.toISOString().slice(0,10),end:n.toISOString().slice(0,10)};}},
  {label:"Last 90d", fn:()=>{const n=new Date(),s=new Date(n);s.setDate(s.getDate()-90);return{start:s.toISOString().slice(0,10),end:n.toISOString().slice(0,10)};}},
  {label:"Last 6m",  fn:()=>{const n=new Date(),s=new Date(n);s.setMonth(s.getMonth()-6);return{start:s.toISOString().slice(0,10),end:n.toISOString().slice(0,10)};}},
  {label:"YTD",      fn:()=>{const n=new Date();return{start:`${n.getFullYear()}-01-01`,end:n.toISOString().slice(0,10)};}},
  {label:"Last 12m", fn:()=>{const n=new Date(),s=new Date(n);s.setFullYear(s.getFullYear()-1);return{start:s.toISOString().slice(0,10),end:n.toISOString().slice(0,10)};}},
];

function Bar({pct,color="#60a5fa"}){return <div className="bt"><div className="bf" style={{width:`${Math.min(Math.max(pct,0),100)}%`,background:color}}/></div>;}
function Spark({vals=[],color="#60a5fa",w=90,h=26}){
  if(!vals||vals.length<2)return null;
  const max=Math.max(...vals,1);
  const pts=vals.map((v,i)=>`${(i/(vals.length-1))*w},${h-(v/max)*h}`).join(" ");
  return <svg width={w} height={h} style={{overflow:"visible",display:"block"}}><polyline points={pts} fill="none" stroke={color} strokeWidth="1.5" strokeLinejoin="round" strokeLinecap="round"/></svg>;
}

function StackChart({months,topSvcs,colorMap}){
  if(!months.length)return <div className="empty-m">No trend data</div>;
  const max=Math.max(...months.map(m=>m.total),1);
  return(
    <div className="sch">
      {months.map((m,i)=>(
        <div key={i} className="sch-col">
          <div className="sch-stack" style={{height:`${(m.total/max)*100}%`}}>
            {topSvcs.map(svc=>{const v=m[svc]||0;if(!v)return null;return<div key={svc} className="sch-seg" style={{height:`${(v/m.total)*100}%`,background:colorMap[svc]||"#60a5fa"}} title={`${svc}: ${fmt(v)}`}/>;} )}
          </div>
          <div className="sch-lbl">{m.label}</div>
        </div>
      ))}
    </div>
  );
}

function DailyLine({days}){
  if(!days||days.length<2)return <div className="empty-m">No daily data</div>;
  const vals=days.map(d=>d.total),max=Math.max(...vals,1),W=600,H=70;
  const pts=vals.map((v,i)=>`${(i/(vals.length-1))*W},${H-(v/max)*H}`).join(" ");
  return(
    <div className="dl-wrap">
      <svg viewBox={`0 0 ${W} ${H}`} preserveAspectRatio="none" style={{width:"100%",height:60}}>
        <defs><linearGradient id="dlg" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#60a5fa" stopOpacity=".2"/><stop offset="100%" stopColor="#60a5fa" stopOpacity="0"/></linearGradient></defs>
        <polygon points={`0,${H} ${pts} ${W},${H}`} fill="url(#dlg)"/>
        <polyline points={pts} fill="none" stroke="#60a5fa" strokeWidth="2" strokeLinejoin="round"/>
      </svg>
      <div className="dl-lbls">
        {days.filter((_,i)=>i%Math.max(1,Math.floor(days.length/6))===0).map((d,i)=>(
          <span key={i}>{new Date(d.date+"T12:00:00Z").toLocaleDateString("en-US",{month:"short",day:"numeric"})}</span>
        ))}
      </div>
    </div>
  );
}

function Pagination({total, page, pageSize, onPage}) {
  const totalPages = Math.ceil(total / pageSize);
  if (totalPages <= 1) return null;
  const start = (page-1)*pageSize+1;
  const end = Math.min(page*pageSize, total);
  const pages = [];
  // Show first, last, current ±2
  const add = p => { if (p>=1&&p<=totalPages&&!pages.includes(p)) pages.push(p); };
  [1,2].forEach(add);
  [page-2,page-1,page,page+1,page+2].forEach(add);
  [totalPages-1,totalPages].forEach(add);
  pages.sort((a,b)=>a-b);
  return (
    <div className="pgn">
      <span className="pgn-info">{start}–{end} of {total}</span>
      <div className="pgn-btns">
        <button className="pgn-btn" disabled={page===1} onClick={()=>onPage(1)}>«</button>
        <button className="pgn-btn" disabled={page===1} onClick={()=>onPage(page-1)}>‹</button>
        {pages.map((p,i) => {
          const prev = pages[i-1];
          return [
            prev && p-prev>1 ? <span key={"e"+p} className="pgn-ellipsis">…</span> : null,
            <button key={p} className={`pgn-btn${page===p?" pgn-active":""}`} onClick={()=>onPage(p)}>{p}</button>
          ];
        })}
        <button className="pgn-btn" disabled={page===totalPages} onClick={()=>onPage(page+1)}>›</button>
        <button className="pgn-btn" disabled={page===totalPages} onClick={()=>onPage(totalPages)}>»</button>
      </div>
      <select className="pgn-jump" value={page} onChange={e=>onPage(Number(e.target.value))}>
        {Array.from({length:totalPages},(_,i)=>i+1).map(p=><option key={p} value={p}>Page {p}</option>)}
      </select>
    </div>
  );
}

function TenantBadge({name,color}){
  return <span className="tbadge" style={{background:color+"22",color,borderColor:color+"44"}}>{name}</span>;
}

export default function App(){
  const initD=PRESETS[0].fn();
  const [dates,setDates]=useState(initD);
  const [preset,setPreset]=useState("MTD");
  const [data,setData]=useState(null);
  const [loading,setLoading]=useState(true);
  const [error,setError]=useState(null);
  const [refreshedAt,setRefreshedAt]=useState(null);
  const [view,setView]=useState("overview");
  const [selTenantId,setSelTenantId]=useState(null);
  const [selSubId,setSelSubId]=useState(null);
  const [expSvc,setExpSvc]=useState({});
  const [expRg,setExpRg]=useState({});
  const [search,setSearch]=useState("");
  const [sortDesc,setSortDesc]=useState(true);
  const [progress,setProgress]=useState({i:0,total:0,name:""});
  const [pages,setPages]=useState({});  // {viewKey: pageNumber}
  const PAGE_SIZE = 50;
  const getPage = (key) => pages[key] || 1;
  const setPage = (key, p) => setPages(prev => ({...prev, [key]: p}));
  const [invoices,setInvoices]=useState(null);
  const [invoicesLoading,setInvoicesLoading]=useState(false);
  const [invoiceFilter,setInvoiceFilter]=useState("all");
  const [invoiceSearch,setInvoiceSearch]=useState("");
  const [selectedInvoice,setSelectedInvoice]=useState(null);
  const [invoiceTxns,setInvoiceTxns]=useState([]);
  const [txnsLoading,setTxnsLoading]=useState(false);
  const [theme,setTheme]=useState(()=>localStorage.getItem("azr-theme")||"dark");
  const [user,setUser]=useState(null);
  const [authLoading,setAuthLoading]=useState(true);

  useEffect(()=>{
    document.documentElement.setAttribute("data-theme",theme);
    localStorage.setItem("azr-theme",theme);
  },[theme]);

  useEffect(()=>{
    getStoredUser().then(u=>{setUser(u);setAuthLoading(false);});
  },[]);

  const handleSignIn=async()=>{
    setAuthLoading(true);
    try{ const u=await signIn(); if(u)setUser(u); }
    catch(e){ console.error(e); }
    finally{ setAuthLoading(false); }
  };

  const handleSignOut=async()=>{
    await signOut();
    setUser(null);
  };

  const load=useCallback(async(range)=>{
    setLoading(true);setError(null);setExpSvc({});setExpRg({});
    setProgress({i:0,total:0,name:"Connecting..."});
    try{
      const r=await fetchAllData(range.start,range.end,(i,total,name)=>setProgress({i:i+1,total,name}));
      setData(r);setRefreshedAt(new Date());
    }catch(e){setError(e.message);}
    finally{setLoading(false);}
  },[]);

  useEffect(()=>{load(initD);},[]);

  const applyPreset=p=>{const r=p.fn();setPreset(p.label);setDates(r);load(r);};

  const openInvoice = async (inv) => {
    setSelectedInvoice(inv);
    setInvoiceTxns([]);
    setTxnsLoading(true);
    try {
      const userToken = await import("./auth/msal").then(m => m.getUserMgmtToken());
      const txns = await fetchInvoiceTransactions(inv, userToken);
      setInvoiceTxns(txns);
    } catch(e) { console.warn('txns failed:', e); }
    finally { setTxnsLoading(false); }
  };

  const loadInvoices=useCallback(async()=>{
    if(!data)return;
    setInvoicesLoading(true);
    try{
      const invs=await fetchAllInvoices(data.subscriptions, TENANTS);
      setInvoices(invs);
    }catch(e){console.error('Invoice load failed:',e);}
    finally{setInvoicesLoading(false);}
  },[data]);

  useEffect(()=>{
    if(view==="invoices"&&data&&!invoices&&!invoicesLoading)loadInvoices();
  },[view,data]);
  const applyDate=(k,v)=>{const r={...dates,[k]:v};setPreset(null);setDates(r);load(r);};

  const exportInvoicePDF=async(inv)=>{
    // Dynamic import jsPDF
    const {jsPDF} = await import("https://cdn.jsdelivr.net/npm/jspdf@2.5.1/+esm");
    const doc = new jsPDF({orientation:"portrait",unit:"mm",format:"a4"});
    const W=210, margin=20;
    let y=margin;

    // Header
    doc.setFillColor(7,9,15);
    doc.rect(0,0,W,40,"F");
    doc.setTextColor(255,255,255);
    doc.setFontSize(20);doc.setFont("helvetica","bold");
    doc.text("AzureReader",margin,18);
    doc.setFontSize(10);doc.setFont("helvetica","normal");
    doc.text("Azure Cost Statement",margin,26);
    doc.setFontSize(14);doc.setFont("helvetica","bold");
    doc.text(inv.label||inv.name,W-margin,18,{align:"right"});
    doc.setFontSize(9);doc.setFont("helvetica","normal");
    doc.text(inv.periodStart+" – "+inv.periodEnd,W-margin,26,{align:"right"});
    y=50;

    // Statement info box
    doc.setFillColor(240,244,250);
    doc.rect(margin,y,W-margin*2,28,"F");
    doc.setTextColor(15,22,41);
    doc.setFontSize(9);doc.setFont("helvetica","bold");
    doc.text("BILLING PERIOD",margin+4,y+8);
    doc.text("STATUS",margin+60,y+8);
    doc.text("TOTAL AMOUNT",margin+110,y+8);
    doc.text("SUBSCRIPTIONS",margin+155,y+8);
    doc.setFont("helvetica","normal");doc.setFontSize(11);
    doc.text(inv.label||inv.periodStart,margin+4,y+18);
    doc.setTextColor(5,150,105);
    doc.text(inv.status,margin+60,y+18);
    doc.setTextColor(15,22,41);doc.setFont("helvetica","bold");doc.setFontSize(13);
    doc.text("$"+Number(inv.amount).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2}),margin+110,y+18);
    doc.setFont("helvetica","normal");doc.setFontSize(11);
    doc.text(String(Object.keys(inv.bySub||{}).length||"—"),margin+155,y+18);
    y+=36;

    // By Tenant
    if(inv.byTenant&&Object.keys(inv.byTenant).length>0){
      doc.setFontSize(10);doc.setFont("helvetica","bold");doc.setTextColor(15,22,41);
      doc.text("Cost by Tenant",margin,y);y+=6;
      doc.setLineWidth(0.3);doc.setDrawColor(200,210,230);
      doc.line(margin,y,W-margin,y);y+=5;
      Object.entries(inv.byTenant).sort((a,b)=>b[1]-a[1]).forEach(([t,c])=>{
        doc.setFont("helvetica","normal");doc.setFontSize(9);doc.setTextColor(61,79,114);
        doc.text(t,margin+2,y);
        doc.setTextColor(15,22,41);doc.setFont("helvetica","bold");
        doc.text("$"+Number(c).toLocaleString("en-US",{minimumFractionDigits:2}),W-margin,y,{align:"right"});
        y+=6;
      });
      y+=4;
    }

    // By Subscription
    if(inv.bySub&&Object.keys(inv.bySub).length>0){
      doc.setFontSize(10);doc.setFont("helvetica","bold");doc.setTextColor(15,22,41);
      doc.text("Cost by Subscription",margin,y);y+=6;
      doc.setLineWidth(0.3);doc.setDrawColor(200,210,230);
      doc.line(margin,y,W-margin,y);y+=2;

      // Table header
      doc.setFillColor(240,244,250);
      doc.rect(margin,y,W-margin*2,7,"F");
      doc.setFontSize(8);doc.setFont("helvetica","bold");doc.setTextColor(100,120,150);
      doc.text("SUBSCRIPTION",margin+2,y+5);
      doc.text("AMOUNT",W-margin-2,y+5,{align:"right"});
      y+=9;

      let rowIdx=0;
      Object.entries(inv.bySub).sort((a,b)=>b[1]-a[1]).forEach(([sub,cost])=>{
        if(rowIdx%2===0){doc.setFillColor(248,250,252);doc.rect(margin,y-1,W-margin*2,7,"F");}
        doc.setFont("helvetica","normal");doc.setFontSize(9);doc.setTextColor(61,79,114);
        doc.text(sub.length>55?sub.slice(0,52)+"...":sub,margin+2,y+4);
        doc.setTextColor(15,22,41);doc.setFont("helvetica","bold");
        doc.text("$"+Number(cost).toLocaleString("en-US",{minimumFractionDigits:2}),W-margin-2,y+4,{align:"right"});
        y+=7;rowIdx++;
        if(y>265){doc.addPage();y=20;}
      });
      y+=4;
    }

    // Total line
    doc.setLineWidth(0.5);doc.setDrawColor(29,111,216);
    doc.line(margin,y,W-margin,y);y+=6;
    doc.setFontSize(12);doc.setFont("helvetica","bold");doc.setTextColor(15,22,41);
    doc.text("Total",margin,y);
    doc.setTextColor(29,111,216);
    doc.text("$"+Number(inv.amount).toLocaleString("en-US",{minimumFractionDigits:2,maximumFractionDigits:2}),W-margin,y,{align:"right"});
    y+=10;

    // Footer
    doc.setFontSize(7);doc.setFont("helvetica","normal");doc.setTextColor(150,160,180);
    doc.text("Generated by AzureReader · avontus.com · "+new Date().toLocaleDateString("en-US",{month:"long",day:"numeric",year:"numeric"}),margin,285);
    doc.text("Data sourced from Azure Cost Management API",W-margin,285,{align:"right"});

    doc.save((inv.name||"statement")+".pdf");
  };

  // ── Process ──────────────────────────────────────────────────────────────
  const P=useMemo(()=>{
    if(!data)return null;

    let rows=data.allDetailed;
    if(selTenantId) rows=rows.filter(r=>r.tenantId===selTenantId);
    if(selSubId)    rows=rows.filter(r=>r.subId===selSubId);
    const q=search.toLowerCase();
    if(q) rows=rows.filter(r=>r.service?.toLowerCase().includes(q)||r.rg?.toLowerCase().includes(q)||r.resourceType?.toLowerCase().includes(q)||r.tenant?.toLowerCase().includes(q));

    const grandTotal=rows.reduce((s,r)=>s+r.cost,0);
    const allTotal=data.allDetailed.reduce((s,r)=>s+r.cost,0);

    // Services
    const svcMap={};
    rows.forEach(r=>{
      if(!svcMap[r.service])svcMap[r.service]={name:r.service,total:0,tenants:{},subs:{},rgs:{},resources:[]};
      svcMap[r.service].total+=r.cost;
      svcMap[r.service].tenants[r.tenant]=(svcMap[r.service].tenants[r.tenant]||0)+r.cost;
      svcMap[r.service].subs[r.sub]=(svcMap[r.service].subs[r.sub]||0)+r.cost;
      const rg=r.rg||"(no resource group)";
      if(!svcMap[r.service].rgs[rg])svcMap[r.service].rgs[rg]={total:0,types:{}};
      svcMap[r.service].rgs[rg].total+=r.cost;
      const rt=r.resourceType||"unknown";
      svcMap[r.service].rgs[rg].types[rt]=(svcMap[r.service].rgs[rg].types[rt]||0)+r.cost;
      svcMap[r.service].resources.push(r);
    });
    const services=Object.values(svcMap).sort((a,b)=>sortDesc?b.total-a.total:a.total-b.total);
    const colorMap={};
    [...services].sort((a,b)=>b.total-a.total).forEach((s,i)=>{colorMap[s.name]=SVC_COLORS[i%SVC_COLORS.length];});

    // Tenants
    const tenantMap={};
    data.allDetailed.forEach(r=>{
      if(!tenantMap[r.tenantId])tenantMap[r.tenantId]={id:r.tenantId,name:r.tenant,color:r.tenantColor,total:0,subs:{},svcs:{}};
      tenantMap[r.tenantId].total+=r.cost;
      tenantMap[r.tenantId].subs[r.subId]=r.sub;
      tenantMap[r.tenantId].svcs[r.service]=(tenantMap[r.tenantId].svcs[r.service]||0)+r.cost;
    });
    const tenants=Object.values(tenantMap).sort((a,b)=>b.total-a.total);

    // Subscriptions
    const subMap={};
    data.allDetailed.forEach(r=>{
      if(!subMap[r.subId])subMap[r.subId]={id:r.subId,name:r.sub,tenant:r.tenant,tenantId:r.tenantId,tenantColor:r.tenantColor,total:0,svcs:{},asTotal:0};
      subMap[r.subId].total+=r.cost;
      subMap[r.subId].svcs[r.service]=(subMap[r.subId].svcs[r.service]||0)+r.cost;
      if(isAS(r))subMap[r.subId].asTotal+=r.cost;
    });
    const subs=Object.values(subMap).sort((a,b)=>b.total-a.total);

    // Monthly trend
    const mMap={};
    const mSrc=selTenantId?data.allMonthly.filter(r=>r.tenantId===selTenantId):selSubId?data.allMonthly.filter(r=>r.subId===selSubId):data.allMonthly;
    mSrc.forEach(r=>{
      const ds=parseAzureDate(r.date);if(!ds)return;
      const key=ds.slice(0,7);
      const lbl=new Date(key+"-01T12:00:00Z").toLocaleString("default",{month:"short",year:"2-digit"});
      if(!mMap[key])mMap[key]={label:lbl,total:0};
      mMap[key][r.service]=(mMap[key][r.service]||0)+r.cost;
      mMap[key].total+=r.cost;
    });
    const months=Object.values(mMap).sort((a,b)=>a.label.localeCompare(b.label));

    // Daily
    const dMap={};
    const dSrc=selTenantId?data.allDaily.filter(r=>r.tenantId===selTenantId):selSubId?data.allDaily.filter(r=>r.subId===selSubId):data.allDaily;
    dSrc.forEach(r=>{
      const ds=parseAzureDate(r.date);if(!ds)return;
      const key=ds.slice(0,10);
      if(!dMap[key])dMap[key]={date:key,total:0};
      dMap[key].total+=r.cost;
    });
    const days=Object.values(dMap).sort((a,b)=>a.date.localeCompare(b.date));

    // App Services
    const asRows=rows.filter(isAS);
    const asRgMap={};
    asRows.forEach(r=>{
      const rg=r.rg||"(no rg)";
      const key=r.subId+"::"+rg;
      if(!asRgMap[key])asRgMap[key]={rg,sub:r.sub,subId:r.subId,tenant:r.tenant,tenantId:r.tenantId,tenantColor:r.tenantColor,total:0,resources:[]};
      asRgMap[key].total+=r.cost;
      asRgMap[key].resources.push(r);
    });
    const asGroups=Object.values(asRgMap).sort((a,b)=>b.total-a.total);
    const asTotal=asRows.reduce((s,r)=>s+r.cost,0);

    return{rows,grandTotal,allTotal,services,colorMap,tenants,subs,months,days,asRows,asGroups,asTotal};
  },[data,selTenantId,selSubId,search,sortDesc]);

  const now=new Date();
  const dayOfMonth=now.getDate();
  const daysInMonth=new Date(now.getFullYear(),now.getMonth()+1,0).getDate();
  const selTenant=TENANTS.find(t=>t.id===selTenantId);
  const selSub=data?.subscriptions?.find(s=>s.subscriptionId===selSubId);

  // Auth gate
  if(authLoading)return(
    <div className="splash">
      <div className="sp-wrap"><div className="sp-ring"/><div className="sp-logo">A</div></div>
      <div className="sp-title">AzureReader</div>
      <div className="sp-sub">Checking authentication…</div>
    </div>
  );

  if(!user)return(
    <div className="splash">
      <div className="login-card">
        <div className="lc-logo"><div className="sp-logo" style={{width:52,height:52,fontSize:26}}>A</div></div>
        <div className="lc-title">AzureReader</div>
        <div className="lc-sub">Sign in with your Microsoft account to access the Azure Cost Dashboard</div>
        <button className="lc-btn" onClick={handleSignIn}>
          <svg width="20" height="20" viewBox="0 0 23 23" style={{flexShrink:0}}><path fill="#f25022" d="M0 0h11v11H0z"/><path fill="#00a4ef" d="M12 0h11v11H12z"/><path fill="#7fba00" d="M0 12h11v11H0z"/><path fill="#ffb900" d="M12 12h11v11H12z"/></svg>
          Sign in with Microsoft
        </button>
        <div className="lc-hint">Requires access to the Avontus Software tenant</div>
        <div className="lc-theme">
          <button className="rb" onClick={()=>setTheme(t=>t==="dark"?"light":"dark")}>{theme==="dark"?"☀ Light":"🌙 Dark"}</button>
        </div>
      </div>
    </div>
  );

  if(loading)return(
    <div className="splash">
      <div className="sp-wrap"><div className="sp-ring"/><div className="sp-logo">A</div></div>
      <div className="sp-title">Loading All Tenants</div>
      <div className="sp-sub">{fmtD(dates.start)} → {fmtD(dates.end)}</div>
      {progress.total>0&&(
        <div className="sp-prog">
          <div className="sp-pb"><div className="sp-pf" style={{width:`${(progress.i/progress.total)*100}%`}}/></div>
          <div className="sp-pl">{progress.i}/{progress.total} — {progress.name}</div>
        </div>
      )}
    </div>
  );

  if(error)return(
    <div className="splash"><div className="err-card">
      <div>⚠</div><h3>Error</h3>
      <p className="err-msg">{error}</p>
      <button className="btn-p" onClick={()=>load(dates)}>↻ Retry</button>
    </div></div>
  );

  if(!P)return null;

  return(
    <div className="shell">
      {/* ── Sidebar ── */}
      <aside className="sidebar">
        <div className="sb-brand">
          <div className="sb-mark">A</div>
          <div><div className="sb-name">AzureReader</div><div className="sb-ten">4 tenants · 14 subs</div></div>
        </div>
        <nav className="sb-nav">
          {[{id:"overview",ico:"▣",lbl:"Overview"},{id:"tenants",ico:"◈",lbl:"Tenants"},{id:"services",ico:"⬡",lbl:"By Service"},{id:"appsvcs",ico:"⬢",lbl:"App Services"},{id:"subs",ico:"≡",lbl:"Subscriptions"},{id:"breakdown",ico:"⊞",lbl:"Full Breakdown"},{id:"invoices",ico:"🧾",lbl:"Invoices"}].map(v=>(
            <button key={v.id} className={`sb-item${view===v.id?" on":""}`} onClick={()=>{setView(v.id);setPages({})}}>
              <span className="sb-ico">{v.ico}</span>{v.lbl}
            </button>
          ))}
        </nav>

        <div className="sb-div">TENANTS</div>
        <div className="sb-subs">
          <button className={`sb-sub${!selTenantId&&!selSubId?" sel":""}`} onClick={()=>{setSelTenantId(null);setSelSubId(null);}}>
            <span className="sdot" style={{background:"#60a5fa"}}/>
            <span className="ssn">All tenants</span>
            <span className="ssc">{fmt(P.allTotal)}</span>
          </button>
          {TENANTS.map(t=>{
            const td=P.tenants.find(x=>x.id===t.id);
            const isSelT=selTenantId===t.id&&!selSubId;
            return(
              <button key={t.id} className={`sb-sub${isSelT?" sel":""}`} onClick={()=>{setSelTenantId(t.id);setSelSubId(null);}}>
                <span className="sdot" style={{background:t.color}}/>
                <span className="ssn">{t.name}</span>
                <span className="ssc">{fmt(td?.total||0)}</span>
              </button>
            );
          })}
        </div>

        <div className="sb-div">SUBSCRIPTIONS</div>
        <div className="sb-subs">
          {P.subs.map(s=>(
            <button key={s.id} className={`sb-sub${selSubId===s.id?" sel":""}`}
              onClick={()=>{setSelSubId(s.id);setSelTenantId(null);}}>
              <span className="sdot" style={{background:s.tenantColor}}/>
              <span className="ssn">{s.name}</span>
              <span className="ssc">{fmt(s.total)}</span>
            </button>
          ))}
        </div>

        <div className="sb-foot">
          <div>4 tenants · 14 subscriptions</div>
          {refreshedAt&&<div>{refreshedAt.toLocaleTimeString()}</div>}
          {data?.errors?.length>0&&<div style={{color:"var(--amber)"}}>⚠ {data.errors.length} warn</div>}
        </div>
      </aside>

      {/* ── Main ── */}
      <div className="main">
        <div className="cbar">
          <div className="cbar-l">
            <span className="cbar-title">
              {selSubId?selSub?.displayName:selTenantId?selTenant?.name:"All Tenants"}
            </span>
            <span className="cbar-range">{fmtD(dates.start)} — {fmtD(dates.end)}</span>
          </div>
          <div className="cbar-r">
            <div className="presets">{PRESETS.map(p=><button key={p.label} className={`ps${preset===p.label?" on":""}`} onClick={()=>applyPreset(p)}>{p.label}</button>)}</div>
            <div className="dr">
              <input type="date" className="di" value={dates.start} max={dates.end} onChange={e=>applyDate("start",e.target.value)}/>
              <span style={{color:"var(--t3)",fontSize:11}}>→</span>
              <input type="date" className="di" value={dates.end} min={dates.start} onChange={e=>applyDate("end",e.target.value)}/>
            </div>
            <button className="rb" onClick={()=>load(dates)}>↻</button>
            <button className="rb theme-btn" onClick={()=>setTheme(t=>t==="dark"?"light":"dark")} title="Toggle theme">
              {theme==="dark"?"☀":"🌙"}
            </button>
            <div className="user-menu">
              <div className="user-avatar" title={user?.email}>{user?.initials||"?"}</div>
              <div className="user-drop">
                <div className="ud-name">{user?.name}</div>
                <div className="ud-email">{user?.email}</div>
                <button className="ud-signout" onClick={handleSignOut}>Sign out</button>
              </div>
            </div>
          </div>
        </div>

        {["services","appsvcs","breakdown"].includes(view)&&(
          <div className="sbar">
            <span style={{color:"var(--t3)",fontSize:14,flexShrink:0}}>⌕</span>
            <input className="si-in" placeholder="Filter service, resource group, tenant…" value={search} onChange={e=>setSearch(e.target.value)}/>
            {search&&<button className="si-x" onClick={()=>setSearch("")}>×</button>}
            <select className="si-sort" value={sortDesc?"desc":"asc"} onChange={e=>setSortDesc(e.target.value==="desc")}>
              <option value="desc">Highest cost first</option>
              <option value="asc">Lowest cost first</option>
            </select>
            <span className="si-info">{P.services.length} services · {P.rows.length} records · {fmt(P.grandTotal)}</span>
          </div>
        )}

        <div className="content">

          {/* ═══ OVERVIEW ═══ */}
          {view==="overview"&&(
            <>
              <div className="kpis">
                <div className="kpi blue"><div className="kl">Total Spend</div><div className="kv">{fmt(P.grandTotal)}</div><div className="ks">{data?.subscriptions?.length} subscriptions · {P.tenants.length} tenants</div></div>

                <div className="kpi purple"><div className="kl">App Services</div><div className="kv">{fmt(P.asTotal)}</div><div className="ks">{P.grandTotal?(P.asTotal/P.grandTotal*100).toFixed(1):0}% of total</div></div>
                <div className="kpi amber"><div className="kl">Month Progress</div><div className="kv">{((dayOfMonth/daysInMonth)*100).toFixed(0)}%</div><Bar pct={dayOfMonth/daysInMonth*100} color="var(--amber)"/></div>
              </div>

              {/* Tenant cost cards */}
              <div className="tenant-cards">
                {P.tenants.map(t=>{
                  const tConf=TENANTS.find(x=>x.id===t.id);
                  const subCount=Object.keys(t.subs).length;
                  const topSvc=Object.entries(t.svcs).sort((a,b)=>b[1]-a[1])[0];
                  return(
                    <div key={t.id} className="tc" onClick={()=>{setSelTenantId(t.id);setSelSubId(null);setView("services");}}>
                      <div className="tc-stripe" style={{background:tConf?.color||"#60a5fa"}}/>
                      <div className="tc-name">{t.name}</div>
                      <div className="tc-cost">{fmt(t.total)}</div>
                      <Bar pct={P.tenants[0]?t.total/P.tenants[0].total*100:0} color={tConf?.color||"#60a5fa"}/>
                      <div className="tc-meta">{subCount} subscription{subCount!==1?"s":""} · {((t.total/P.allTotal)*100).toFixed(1)}% of total</div>
                      {topSvc&&<div className="tc-top">Top: {topSvc[0].replace("Azure ","")} · {fmt(topSvc[1])}</div>}
                      <div className="tc-cta">View services →</div>
                    </div>
                  );
                })}
              </div>

              <div className="two-col">
                <div className="panel">
                  <div className="phdr"><span className="pt">Monthly Cost Trend</span><span className="ps">{P.months.length} months</span></div>
                  <StackChart months={P.months} topSvcs={P.services.slice(0,8).map(s=>s.name)} colorMap={P.colorMap}/>
                  <div className="leg">{P.services.slice(0,6).map(s=><div key={s.name} className="leg-i"><span style={{width:7,height:7,borderRadius:"50%",background:P.colorMap[s.name],display:"inline-block",marginRight:5,flexShrink:0}}/>{s.name.replace("Azure ","")}</div>)}</div>
                </div>
                <div className="panel">
                  <div className="phdr"><span className="pt">Daily Spend</span><span className="ps">{P.days.length} days</span></div>
                  <DailyLine days={P.days}/>
                  <div className="dstats">
                    <div className="ds"><span>Avg/day</span><strong>{fmt(P.days.reduce((s,d)=>s+d.total,0)/(P.days.length||1))}</strong></div>
                    <div className="ds"><span>Peak</span><strong>{fmt(Math.max(...(P.days.map(d=>d.total)||[0]),0))}</strong></div>
                    <div className="ds"><span>Days</span><strong>{P.days.length}</strong></div>
                  </div>
                </div>
              </div>

              <div className="panel">
                <div className="phdr"><span className="pt">Top Services — All Tenants</span><button className="vm" onClick={()=>setView("services")}>Full drill-down →</button></div>
                <table className="tbl">
                  <thead><tr><th>Service</th><th>Tenants</th><th className="r">Cost</th><th className="r">% Share</th><th>Distribution</th></tr></thead>
                  <tbody>
                    {P.services.slice(0,12).map(s=>(
                      <tr key={s.name} className="click" onClick={()=>{setView("services");setTimeout(()=>setExpSvc({[s.name]:true}),50);}}>
                        <td><span className="dot" style={{background:P.colorMap[s.name]}}/><strong>{s.name}</strong></td>
                        <td>
                          <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>
                            {Object.entries(s.tenants).slice(0,3).map(([tn])=>{
                              const tc=TENANTS.find(t=>t.name===tn);
                              return <TenantBadge key={tn} name={tn.split(" ")[0]} color={tc?.color||"#60a5fa"}/>;
                            })}
                          </div>
                        </td>
                        <td className="r mono hi">{fmt(s.total,2)}</td>
                        <td className="r mono">{P.grandTotal?(s.total/P.grandTotal*100).toFixed(2):0}%</td>
                        <td><Bar pct={P.services[0]?s.total/P.services[0].total*100:0} color={P.colorMap[s.name]}/></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </>
          )}

          {/* ═══ TENANTS ═══ */}
          {view==="tenants"&&(
            <>
              <div className="panel">
                <div className="phdr"><span className="pt">All Tenants — Cost Breakdown</span></div>
                <table className="tbl">
                  <thead><tr><th>Tenant</th><th className="r">Total Cost</th><th className="r">App Services</th><th className="r">% of Total</th><th>Share</th><th className="r">Subscriptions</th><th className="r">Top Service</th></tr></thead>
                  <tbody>
                    {P.tenants.map(t=>{
                      const tc=TENANTS.find(x=>x.id===t.id);
                      const asT=data.allDetailed.filter(r=>r.tenantId===t.id&&isAS(r)).reduce((s,r)=>s+r.cost,0);
                      const top=Object.entries(t.svcs).sort((a,b)=>b[1]-a[1])[0];
                      return(
                        <tr key={t.id} className="click" onClick={()=>{setSelTenantId(t.id);setSelSubId(null);setView("services");}}>
                          <td><span className="dot" style={{background:tc?.color||"#60a5fa"}}/><strong>{t.name}</strong></td>
                          <td className="r mono hi">{fmt(t.total,2)}</td>
                          <td className="r mono" style={{color:"var(--purple)"}}>{fmt(asT,2)}</td>
                          <td className="r mono">{P.allTotal?(t.total/P.allTotal*100).toFixed(1):0}%</td>
                          <td><Bar pct={P.tenants[0]?t.total/P.tenants[0].total*100:0} color={tc?.color||"#60a5fa"}/></td>
                          <td className="r">{Object.keys(t.subs).length}</td>
                          <td className="r dim" style={{fontSize:11}}>{top?`${top[0].replace("Azure ","")} · ${fmt(top[1])}`:""}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              <div className="tenant-cards">
                {P.tenants.map(t=>{
                  const tc=TENANTS.find(x=>x.id===t.id);
                  const subsInTenant=P.subs.filter(s=>s.tenantId===t.id);
                  return(
                    <div key={t.id} className="tc">
                      <div className="tc-stripe" style={{background:tc?.color||"#60a5fa"}}/>
                      <div className="tc-name">{t.name}</div>
                      <div className="tc-cost">{fmt(t.total)}</div>
                      <Bar pct={P.tenants[0]?t.total/P.tenants[0].total*100:0} color={tc?.color||"#60a5fa"}/>
                      <div className="tc-svcs">
                        {Object.entries(t.svcs).sort((a,b)=>b[1]-a[1]).slice(0,5).map(([svc,cost])=>(
                          <div key={svc} className="tc-svc">
                            <span className="dot" style={{background:P.colorMap[svc]||"#60a5fa",width:6,height:6,borderRadius:"50%",flexShrink:0}}/>
                            <span style={{flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"var(--t2)",fontSize:11}}>{svc.replace("Azure ","")}</span>
                            <span style={{fontFamily:"'DM Mono',monospace",fontSize:11,color:"var(--text)"}}>{fmt(cost)}</span>
                          </div>
                        ))}
                      </div>
                      <div className="sb-div" style={{padding:"10px 0 4px",fontSize:9}}>SUBSCRIPTIONS</div>
                      {subsInTenant.map(s=>(
                        <div key={s.id} className="tc-svc" style={{cursor:"pointer"}} onClick={()=>{setSelSubId(s.id);setView("services");}}>
                          <span className="sdot" style={{background:tc?.color||"#60a5fa"}}/>
                          <span style={{flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"var(--t2)",fontSize:11}}>{s.name}</span>
                          <span style={{fontFamily:"'DM Mono',monospace",fontSize:11,color:"var(--text)"}}>{fmt(s.total)}</span>
                        </div>
                      ))}
                    </div>
                  );
                })}
              </div>
            </>
          )}

          {/* ═══ BY SERVICE ═══ */}
          {view==="services"&&(
            <div className="panel">
              <div className="phdr"><span className="pt">Cost by Service{selTenantId?` — ${selTenant?.name}`:selSubId?` — ${selSub?.displayName}`:""}</span><span className="ps">{P.services.length} services · {fmt(P.grandTotal)}</span></div>
              <table className="tbl">
                <thead><tr><th style={{width:28}}/><th>Service</th><th className="r">Cost</th><th className="r">% Total</th><th>Share</th><th className="r">Tenants</th><th className="r">RGs</th></tr></thead>
                <tbody>
                  {P.services.slice((getPage("svc")-1)*PAGE_SIZE, getPage("svc")*PAGE_SIZE).map(s=>{
                    const ex=expSvc[s.name];
                    return(<>
                      <tr key={s.name} className={`svc-row${ex?" ex":""}`} onClick={()=>setExpSvc(p=>({...p,[s.name]:!p[s.name]}))}>
                        <td className="ex-td">{ex?"▾":"▸"}</td>
                        <td><span className="dot" style={{background:P.colorMap[s.name]}}/><strong>{s.name}</strong></td>
                        <td className="r mono hi">{fmt(s.total,2)}</td>
                        <td className="r mono">{P.grandTotal?(s.total/P.grandTotal*100).toFixed(2):0}%</td>
                        <td><Bar pct={P.services[0]?s.total/P.services[0].total*100:0} color={P.colorMap[s.name]}/></td>
                        <td className="r">
                          <div style={{display:"flex",gap:3,justifyContent:"flex-end",flexWrap:"wrap"}}>
                            {Object.entries(s.tenants).sort((a,b)=>b[1]-a[1]).map(([tn,cost])=>{const tc=TENANTS.find(t=>t.name===tn);return <TenantBadge key={tn} name={tn.split(" ")[0]} color={tc?.color||"#60a5fa"}/>;})}
                          </div>
                        </td>
                        <td className="r">{Object.keys(s.rgs).length}</td>
                      </tr>
                      {ex&&<>
                        {/* By Tenant */}
                        {Object.keys(s.tenants).length>1&&<>
                          <tr className="dhdr"><td/><td colSpan={6}><span className="dl">🏢 By Tenant</span></td></tr>
                          {Object.entries(s.tenants).sort((a,b)=>b[1]-a[1]).map(([tn,cost])=>{
                            const tc=TENANTS.find(t=>t.name===tn);
                            return<tr key={"t_"+tn} className="dsub"><td/><td style={{paddingLeft:30}}><span className="da">↳</span><TenantBadge name={tn} color={tc?.color||"#60a5fa"}/></td><td className="r mono">{fmt(cost,2)}</td><td className="r mono dim">{s.total?(cost/s.total*100).toFixed(1):0}%</td><td><Bar pct={s.total?cost/s.total*100:0} color={tc?.color||"#60a5fa"}/></td><td/><td/></tr>;
                          })}
                        </>}
                        {/* By Sub */}
                        <tr className="dhdr"><td/><td colSpan={6}><span className="dl">📋 By Subscription</span></td></tr>
                        {Object.entries(s.subs).sort((a,b)=>b[1]-a[1]).map(([sub,cost])=>(
                          <tr key={"s_"+sub} className="dsub"><td/><td style={{paddingLeft:30}}><span className="da">↳</span><span className="spill">{sub}</span></td><td className="r mono">{fmt(cost,2)}</td><td className="r mono dim">{s.total?(cost/s.total*100).toFixed(1):0}%</td><td><Bar pct={s.total?cost/s.total*100:0} color={P.colorMap[s.name]+"99"}/></td><td/><td/></tr>
                        ))}
                        {/* By RG */}
                        <tr className="dhdr"><td/><td colSpan={6}><span className="dl">📁 By Resource Group</span></td></tr>
                        {Object.entries(s.rgs).sort((a,b)=>b[1].total-a[1].total).slice(0,20).map(([rg,rgd])=>{
                          const rk=s.name+"::"+rg;const rex=expRg[rk];
                          return(<>
                            <tr key={"rg_"+rg} className={`drg${rex?" ex":""}`} onClick={e=>{e.stopPropagation();setExpRg(p=>({...p,[rk]:!p[rk]}));}}>
                              <td/><td style={{paddingLeft:30}}><span className="da">{rex?"▾":"▸"}</span><span className="rgt">{rg}</span></td>
                              <td className="r mono">{fmt(rgd.total,2)}</td><td className="r mono dim">{s.total?(rgd.total/s.total*100).toFixed(1):0}%</td>
                              <td><Bar pct={s.total?rgd.total/s.total*100:0} color={P.colorMap[s.name]+"77"}/></td><td/><td className="r dim">{Object.keys(rgd.types).length} types</td>
                            </tr>
                            {rex&&Object.entries(rgd.types).sort((a,b)=>b[1]-a[1]).map(([rt,cost])=>(
                              <tr key={"rt_"+rt} className="dres"><td/><td style={{paddingLeft:52}}><span className="da">↳</span><span className="rtt">{rt.split("/").pop()}</span><span className="rtf">{rt}</span></td><td className="r mono">{fmt(cost,2)}</td><td className="r mono dim">{s.total?(cost/s.total*100).toFixed(2):0}%</td><td/><td/><td/></tr>
                            ))}
                          </>);
                        })}
                      </>}
                    </>);
                  })}
                </tbody>
              </table>
              <Pagination total={P.services.length} page={getPage("svc")} pageSize={PAGE_SIZE} onPage={p=>setPage("svc",p)}/>
            </div>
          )}

          {/* ═══ APP SERVICES ═══ */}
          {view==="appsvcs"&&(
            <>
              <div className="kpis">
                <div className="kpi blue"><div className="kl">App Service Total</div><div className="kv">{fmt(P.asTotal)}</div><div className="ks">{P.grandTotal?(P.asTotal/P.grandTotal*100).toFixed(1):0}% of spend</div></div>
                <div className="kpi purple"><div className="kl">Resource Groups</div><div className="kv">{P.asGroups.length}</div></div>
                <div className="kpi green"><div className="kl">Resources</div><div className="kv">{P.asRows.length}</div><div className="ks">Plans, apps, slots</div></div>
                <div className="kpi amber"><div className="kl">Tenants with AS</div><div className="kv">{new Set(P.asRows.map(r=>r.tenantId)).size}</div></div>
              </div>
              <div className="panel">
                <div className="phdr"><span className="pt">App Services — All Tenants</span><span className="ps">Click to expand</span></div>
                {P.asGroups.length===0?<div className="empty-m">No App Service costs found for this period</div>:(
                  <table className="tbl">
                    <thead><tr><th style={{width:28}}/><th>Resource Group / Resource</th><th>Tenant</th><th>Subscription</th><th>Type</th><th className="r">Cost</th><th className="r">% of AS</th></tr></thead>
                    <tbody>
                      {P.asGroups.slice((getPage('as')-1)*PAGE_SIZE, getPage('as')*PAGE_SIZE).map(g=>{
                        const gk=g.subId+"::"+g.rg;const gex=expRg[gk];
                        return(<>
                          <tr key={gk} className={`rg-row${gex?" ex":""}`} onClick={()=>setExpRg(p=>({...p,[gk]:!p[gk]}))}>
                            <td className="ex-td">{gex?"▾":"▸"}</td>
                            <td><span className="rgt">{g.rg}</span></td>
                            <td><TenantBadge name={g.tenant} color={g.tenantColor}/></td>
                            <td><span className="spill">{g.sub}</span></td>
                            <td><span className="tpill">RG</span></td>
                            <td className="r mono hi">{fmt(g.total,2)}</td>
                            <td className="r mono">{P.asTotal?(g.total/P.asTotal*100).toFixed(1):0}%</td>
                          </tr>
                          {gex&&g.resources.sort((a,b)=>b.cost-a.cost).map((r,ri)=>(
                            <tr key={ri} className="dres">
                              <td/><td style={{paddingLeft:26}}><span className="da">↳</span><span className="rtt">{r.resourceType?.split("/").pop()||r.service}</span></td>
                              <td><TenantBadge name={r.tenant.split(" ")[0]} color={r.tenantColor}/></td>
                              <td className="dim">{r.sub}</td>
                              <td><span className="tpill">{r.resourceType?.split("/").pop()||"—"}</span></td>
                              <td className="r mono">{fmt(r.cost,2)}</td>
                              <td className="r mono dim">{P.asTotal?(r.cost/P.asTotal*100).toFixed(2):0}%</td>
                            </tr>
                          ))}
                        </>);
                      })}
                    </tbody>
                  </table>
                  <Pagination total={P.asGroups.length} page={getPage("as")} pageSize={PAGE_SIZE} onPage={p=>setPage("as",p)}/>
                )}
              </div>
            </>
          )}

          {/* ═══ SUBSCRIPTIONS ═══ */}
          {view==="subs"&&(
            <>
              <div className="panel">
                <div className="phdr"><span className="pt">All 14 Subscriptions</span><span className="ps">Across 4 tenants</span></div>
                <table className="tbl">
                  <thead><tr><th>Subscription</th><th>Tenant</th><th className="r">Cost</th><th className="r">App Svcs</th><th className="r">% Total</th><th>Share</th><th className="r">Top Service</th></tr></thead>
                  <tbody>
                    {P.subs.slice((getPage('subs')-1)*PAGE_SIZE, getPage('subs')*PAGE_SIZE).map((s,i)=>{
                      const top=Object.entries(s.svcs).sort((a,b)=>b[1]-a[1])[0];
                      return(
                        <tr key={s.id} className="click" onClick={()=>{setSelSubId(s.id);setView("services");}}>
                          <td><span className="dot" style={{background:s.tenantColor}}/><strong>{s.name}</strong><div style={{fontFamily:"'DM Mono',monospace",fontSize:9,color:"var(--t3)"}}>{s.id?.slice(0,8)}…</div></td>
                          <td><TenantBadge name={s.tenant.split(" ")[0]} color={s.tenantColor}/></td>
                          <td className="r mono hi">{fmt(s.total,2)}</td>
                          <td className="r mono" style={{color:"var(--purple)"}}>{fmt(s.asTotal,2)}</td>
                          <td className="r mono">{P.allTotal?(s.total/P.allTotal*100).toFixed(1):0}%</td>
                          <td><Bar pct={P.subs[0]?s.total/P.subs[0].total*100:0} color={s.tenantColor}/></td>
                          <td className="r dim" style={{fontSize:11}}>{top?`${top[0].replace("Azure ","")} · ${fmt(top[1])}`:""}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
              <Pagination total={P.subs.length} page={getPage("subcards")} pageSize={PAGE_SIZE} onPage={p=>setPage("subcards",p)}/>
            </>
          )}

          {/* ═══ BREAKDOWN ═══ */}
          {view==="breakdown"&&(
            <div className="panel">
              <div className="phdr"><span className="pt">Full Cost Breakdown</span><span className="ps">{P.rows.length} records · {fmt(P.grandTotal)}</span></div>
              <table className="tbl">
                <thead><tr><th>Service</th><th>Tenant</th><th>Subscription</th><th>Resource Group</th><th>Type</th><th className="r">Cost</th></tr></thead>
                <tbody>
                  {P.rows.sort((a,b)=>sortDesc?b.cost-a.cost:a.cost-b.cost).slice((getPage('bd')-1)*PAGE_SIZE, getPage('bd')*PAGE_SIZE).map((r,i)=>(
                    <tr key={i}>
                      <td><span className="dot" style={{background:P.colorMap[r.service]||"#60a5fa"}}/>{r.service}</td>
                      <td><TenantBadge name={r.tenant.split(" ")[0]} color={r.tenantColor}/></td>
                      <td className="dim">{r.sub}</td>
                      <td><span className="rgt-sm">{r.rg||"—"}</span></td>
                      <td className="dim" style={{fontSize:11}}>{r.resourceType?.split("/").pop()||"—"}</td>
                      <td className="r mono hi">{fmt(r.cost,4)}</td>
                    </tr>
                  ))}
                  {P.rows.length>1000&&<tr><td colSpan={6} className="dim" style={{textAlign:"center",padding:12}}>Showing 1,000 of {P.rows.length} — use filters</td></tr>}
                </tbody>
              </table>
            </div>
          )}

          {/* ═══ INVOICES ═══ */}
          {view==="invoices"&&(
            <>
              {invoicesLoading&&(
                <div className="inv-loading">
                  <div className="sp-wrap" style={{width:40,height:40}}><div className="sp-ring"/><div className="sp-logo" style={{width:26,height:26,fontSize:13}}>A</div></div>
                  <div>Fetching invoices across all tenants…</div>
                </div>
              )}
              {!invoicesLoading&&invoices&&(
                <>
                  <div className="inv-toolbar">
                    <div className="inv-stats">
                      <div className="inv-stat"><span>Total invoices</span><strong>{invoices.length}</strong></div>
                      <div className="inv-stat"><span>Total billed</span><strong>{fmt(invoices.reduce((s,i)=>s+(i.totalAmount||0),0),2)}</strong></div>
                      <div className="inv-stat"><span>Tax</span><strong>{fmt(invoices.reduce((s,i)=>s+(i.taxAmount||0),0),2)}</strong></div>
                      <div className="inv-stat"><span>Amount due</span><strong style={{color:invoices.reduce((s,i)=>s+(i.amountDue||0),0)>0?"var(--red)":"var(--green)"}}>{fmt(invoices.reduce((s,i)=>s+(i.amountDue||0),0),2)}</strong></div>
                      <div className="inv-stat"><span>Paid</span><strong style={{color:"var(--green)"}}>{fmt(invoices.filter(i=>i.status==="Paid").reduce((s,i)=>s+(i.totalAmount||0),0),2)}</strong></div>
                    </div>
                    <div className="inv-filters">
                      <input className="si-in" placeholder="Search invoices…" value={invoiceSearch} onChange={e=>setInvoiceSearch(e.target.value)} style={{maxWidth:200}}/>
                      {["all","Paid","Due","PastDue","Void"].map(f=>(
                        <button key={f} className={`ps${invoiceFilter===f?" on":""}`} onClick={()=>{setInvoiceFilter(f);setPage('inv',1);}}>
                          {f==="all"?"All":f}
                        </button>
                      ))}
                      <button className="rb" onClick={loadInvoices} title="Refresh invoices">↻</button>
                    </div>
                  </div>

                  {invoices.length===0?(
                    <div className="panel">
                      <div className="empty-lg">
                        <div style={{fontSize:32}}>🧾</div>
                        <div style={{fontFamily:"'Syne',sans-serif",fontSize:16,fontWeight:700}}>No invoices found</div>
                        <div style={{fontSize:12,color:"var(--t3)",maxWidth:400,lineHeight:1.6}}>
                          Invoices require a billing account with the AzureReader app assigned the <strong>Billing Account Reader</strong> role, or the subscription must be on a Microsoft Customer Agreement.
                        </div>
                        <div style={{fontSize:11,color:"var(--t3)",fontFamily:"'DM Mono',monospace",marginTop:4}}>
                          Run in Cloud Shell:<br/>
                          az role assignment create --assignee 3977e66a --role "Billing Account Reader"
                        </div>
                      </div>
                    </div>
                  ):(
                    <div className="panel">
                      <div className="phdr">
                        <span className="pt">Invoices — All Tenants</span>
                        <span className="ps">{invoices.filter(i=>(invoiceFilter==="all"||i.status===invoiceFilter)&&(!invoiceSearch||i.name?.toLowerCase().includes(invoiceSearch.toLowerCase())||i.subName?.toLowerCase().includes(invoiceSearch.toLowerCase()))).length} invoices</span>
                      </div>
                      <table className="tbl">
                        <thead>
                          <tr>
                            <th>Invoice ID</th>
                            <th>Billing Profile</th>
                            <th>Invoice Date</th>
                            <th>Period</th>
                            <th>Status</th>
                            <th className="r">Subtotal</th>
                            <th className="r">Tax</th>
                            <th className="r">Total</th>
                            <th className="r">Amount Due</th>
                            <th className="r">Actions</th>
                          </tr>
                        </thead>
                        <tbody>
                          {invoices
                            .filter(i=>(invoiceFilter==="all"||i.status===invoiceFilter)&&(!invoiceSearch||i.name?.toLowerCase().includes(invoiceSearch.toLowerCase())||i.subName?.toLowerCase().includes(invoiceSearch.toLowerCase())||i.billingProfileName?.toLowerCase().includes(invoiceSearch.toLowerCase())))
                            .slice((getPage("inv")-1)*PAGE_SIZE, getPage("inv")*PAGE_SIZE).map((inv,i)=>{
                              const statusColor=inv.status==="Paid"?"var(--green)":inv.status==="PastDue"?"var(--red)":inv.status==="Due"?"var(--amber)":"var(--t3)";
                              const period=inv.periodStart&&inv.periodEnd?`${new Date(inv.periodStart+"T12:00:00Z").toLocaleDateString("en-US",{month:"short",year:"numeric"})} – ${new Date(inv.periodEnd+"T12:00:00Z").toLocaleDateString("en-US",{month:"short",year:"numeric"})}`:inv.invoiceDate?new Date(inv.invoiceDate+"T12:00:00Z").toLocaleDateString("en-US",{month:"short",year:"numeric"}):"—";
                              const due=inv.dueDate?new Date(inv.dueDate+"T12:00:00Z").toLocaleDateString("en-US",{month:"short",day:"numeric",year:"numeric"}):"—";
                              return(
                                <tr key={i} className={`inv-row${inv.status==="PastDue"?" inv-overdue":""}`} onClick={()=>openInvoice(inv)}>
                                  <td><span className="inv-id">{inv.name||inv.id}</span></td>
                                  <td className="dim" style={{fontSize:11}}>{inv.billingProfile||inv.subName||"—"}</td>
                                  <td className="dim" style={{fontSize:11}}>{inv.invoiceDate?new Date(inv.invoiceDate+"T12:00:00Z").toLocaleDateString("en-US",{month:"short",day:"numeric",year:"numeric"}):"—"}</td>
                                  <td className="dim" style={{fontSize:11}}>{period}</td>
                                  <td>
                                    <span className="inv-status" style={{color:statusColor,background:statusColor+"18",borderColor:statusColor+"33"}}>
                                      {inv.status==="Paid"?"✓ ":inv.status==="PastDue"?"⚠ ":""}{inv.status}
                                    </span>
                                  </td>
                                  <td className="r mono dim">{fmt(inv.subTotal||0,2)}</td>
                                  <td className="r mono dim">{fmt(inv.taxAmount||0,2)}</td>
                                  <td className="r mono hi">{fmt(inv.totalAmount||0,2)}</td>
                                  <td className="r mono" style={{color:inv.amountDue>0?"var(--red)":"var(--green)"}}>{fmt(inv.amountDue||0,2)}</td>
                                  <td className="r" style={{display:"flex",gap:5,justifyContent:"flex-end",alignItems:"center"}}>
                                    <button className="inv-dl-btn" onClick={e=>{e.stopPropagation();openInvoice(inv);}}>👁</button>
                                    <button className="inv-dl-btn" onClick={e=>{e.stopPropagation();exportInvoicePDF(inv);}} title="Export as PDF">↓ PDF</button>
                                  </td>
                                </tr>
                              );
                            })}
                        </tbody>
                      </table>
                      <Pagination
                        total={invoices.filter(i=>(invoiceFilter==="all"||i.status===invoiceFilter)&&(!invoiceSearch||i.name?.toLowerCase().includes(invoiceSearch.toLowerCase())||i.subName?.toLowerCase().includes(invoiceSearch.toLowerCase())||i.billingProfile?.toLowerCase().includes(invoiceSearch.toLowerCase()))).length}
                        page={getPage("inv")} pageSize={PAGE_SIZE} onPage={p=>setPage("inv",p)}/>
                    </div>
                  )}
                </>
              )}
              {!invoicesLoading&&!invoices&&(
                <div className="panel">
                  <div className="empty-lg">
                    <div style={{fontSize:32}}>🧾</div>
                    <div style={{fontFamily:"'Syne',sans-serif",fontSize:16,fontWeight:700}}>Load Invoices</div>
                    <button className="lc-btn" onClick={loadInvoices}>Fetch Invoices</button>
                  </div>
                </div>
              )}
            </>
          )}
        </div>
      </div>

      {/* ── Invoice Detail Modal ── */}
      {selectedInvoice&&(
        <div className="inv-modal-overlay" onClick={()=>setSelectedInvoice(null)}>
          <div className="inv-modal" onClick={e=>e.stopPropagation()}>
            <div className="inv-modal-hdr">
              <div>
                <div className="inv-modal-title">{selectedInvoice.label||selectedInvoice.name}</div>
                <div className="inv-modal-sub">{selectedInvoice.periodStart} — {selectedInvoice.periodEnd}</div>
              </div>
              <div style={{display:"flex",gap:8,alignItems:"center"}}>
                <button className="inv-dl-btn" style={{padding:"7px 14px",fontSize:13}} onClick={()=>exportInvoicePDF(selectedInvoice)}>↓ Download PDF</button>
                <button className="rb" onClick={()=>setSelectedInvoice(null)}>✕</button>
              </div>
            </div>
            <div className="inv-modal-body">
              {/* Summary */}
              <div className="inv-summary-cards">
                <div className="inv-sc blue"><div className="inv-sc-l">Total Amount</div><div className="inv-sc-v">{fmt(selectedInvoice.totalAmount||0,2)}</div><div className="inv-sc-s">{selectedInvoice.currency}</div></div>
                <div className="inv-sc amber"><div className="inv-sc-l">Subtotal</div><div className="inv-sc-v">{fmt(selectedInvoice.subTotal||0,2)}</div><div className="inv-sc-s">Before tax</div></div>
                <div className="inv-sc" style={{"--ac":"var(--t3)"}}><div className="inv-sc-l">Tax</div><div className="inv-sc-v">{fmt(selectedInvoice.taxAmount||0,2)}</div><div className="inv-sc-s">{selectedInvoice.currency}</div></div>
                <div className="inv-sc" style={{"--ac":selectedInvoice.amountDue>0?"var(--red)":"var(--green)"}}>
                  <div className="inv-sc-l">Amount Due</div>
                  <div className="inv-sc-v" style={{color:selectedInvoice.amountDue>0?"var(--red)":"var(--green)"}}>{fmt(selectedInvoice.amountDue||0,2)}</div>
                  <div className="inv-sc-s">{selectedInvoice.amountDue>0?"Outstanding":"Paid"}</div>
                </div>
              </div>

              {/* Invoice details */}
              <div className="inv-detail-grid">
                <div className="inv-detail-item"><span>Invoice ID</span><strong className="inv-id">{selectedInvoice.name}</strong></div>
                <div className="inv-detail-item"><span>Status</span><strong style={{color:selectedInvoice.status==="Paid"?"var(--green)":"var(--red)"}}>{selectedInvoice.status}</strong></div>
                <div className="inv-detail-item"><span>Invoice Date</span><strong>{selectedInvoice.invoiceDate?new Date(selectedInvoice.invoiceDate+"T12:00:00Z").toLocaleDateString("en-US",{month:"long",day:"numeric",year:"numeric"}):"—"}</strong></div>
                <div className="inv-detail-item"><span>Due Date</span><strong>{selectedInvoice.dueDate?new Date(selectedInvoice.dueDate+"T12:00:00Z").toLocaleDateString("en-US",{month:"long",day:"numeric",year:"numeric"}):"—"}</strong></div>
                <div className="inv-detail-item"><span>Billing Period</span><strong>{selectedInvoice.periodStart} – {selectedInvoice.periodEnd}</strong></div>
                <div className="inv-detail-item"><span>Billing Profile</span><strong>{selectedInvoice.billingProfile||"—"}</strong></div>
                {selectedInvoice.paymentMethod&&<div className="inv-detail-item"><span>Payment Method</span><strong>{selectedInvoice.paymentMethod}</strong></div>}
                {selectedInvoice.paymentDate&&<div className="inv-detail-item"><span>Payment Date</span><strong>{new Date(selectedInvoice.paymentDate+"T12:00:00Z").toLocaleDateString("en-US",{month:"long",day:"numeric",year:"numeric"})}</strong></div>}
              </div>

              {/* Transactions */}
              <div className="inv-section">
                <div className="inv-section-title" style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                  <span>Transactions</span>
                  {txnsLoading&&<span style={{fontSize:10,color:"var(--t3)"}}>Loading…</span>}
                  {!txnsLoading&&<span style={{fontSize:10,color:"var(--t3)"}}>{invoiceTxns.length} line items</span>}
                </div>
                {txnsLoading&&<div style={{padding:"20px",textAlign:"center",color:"var(--t3)",fontSize:12}}>Fetching transactions…</div>}
                {!txnsLoading&&invoiceTxns.length>0&&(
                  <table className="tbl">
                    <thead>
                      <tr>
                        <th>Date</th>
                        <th>Service Period</th>
                        <th>Product Family</th>
                        <th>Product / SKU</th>
                        <th>Invoice Section</th>
                        <th className="r">Charges</th>
                        <th className="r">Total</th>
                      </tr>
                    </thead>
                    <tbody>
                      {invoiceTxns.slice((getPage('txn')-1)*50, getPage('txn')*50).map((t,i)=>{
                        const p=t.properties||t;
                        const date=p.date||p.transactionDate||"";
                        const svcStart=p.servicePeriodStartDate||"";
                        const svcEnd=p.servicePeriodEndDate||"";
                        const period=svcStart&&svcEnd?`${svcStart.slice(0,10)} – ${svcEnd.slice(0,10)}`:"—";
                        const family=p.productFamily||p.productFamilyName||"—";
                        const sku=p.productDescription||p.productType||p.productName||"—";
                        const section=p.invoiceSectionDisplayName||p.invoiceSection||"—";
                        const charges=p.subTotal?.value??p.charges?.value??p.amount?.value??0;
                        const total=p.unitPrice?.value??p.subTotal?.value??charges;
                        return(
                          <tr key={i}>
                            <td className="dim" style={{fontSize:11,whiteSpace:"nowrap"}}>{date?new Date(date).toLocaleDateString("en-US",{month:"numeric",day:"numeric",year:"numeric"}):"—"}</td>
                            <td className="dim" style={{fontSize:10,whiteSpace:"nowrap"}}>{period}</td>
                            <td style={{fontSize:11}}>{family}</td>
                            <td style={{fontSize:11,maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={sku}>{sku}</td>
                            <td style={{fontSize:11}}><span className="rgt-sm">{section}</span></td>
                            <td className="r mono" style={{fontSize:11}}>{fmt(charges,2)}</td>
                            <td className="r mono hi" style={{fontSize:11}}>{fmt(total,2)}</td>
                          </tr>
                        );
                      })}
                      <tr style={{borderTop:"2px solid var(--b2)"}}>
                        <td colSpan={5}><strong>Total</strong></td>
                        <td className="r mono hi"><strong>{fmt(invoiceTxns.reduce((s,t)=>{const p=t.properties||t;return s+(p.subTotal?.value??p.charges?.value??0);},0),2)}</strong></td>
                        <td className="r mono hi"><strong>{fmt(selectedInvoice.totalAmount||0,2)}</strong></td>
                      </tr>
                    </tbody>
                  </table>
                  <Pagination total={invoiceTxns.length} page={getPage("txn")} pageSize={50} onPage={p=>setPage("txn",p)}/>
                )}
                {!txnsLoading&&invoiceTxns.length===0&&(
                  <div style={{padding:"16px",textAlign:"center",color:"var(--t3)",fontSize:12}}>
                    No transaction detail available — requires Billing Account Reader access on the specific billing profile.
                  </div>
                )}
              </div>

              {/* By Tenant */}
              {selectedInvoice.byTenant&&Object.keys(selectedInvoice.byTenant).length>0&&(
                <div className="inv-section">
                  <div className="inv-section-title">Cost by Tenant</div>
                  <table className="tbl">
                    <thead><tr><th>Tenant</th><th className="r">Amount</th><th className="r">% Share</th><th>Distribution</th></tr></thead>
                    <tbody>
                      {Object.entries(selectedInvoice.byTenant).sort((a,b)=>b[1]-a[1]).map(([t,c])=>{
                        const tc=TENANTS.find(x=>x.name===t);
                        return(
                          <tr key={t}>
                            <td><TenantBadge name={t} color={tc?.color||"#60a5fa"}/></td>
                            <td className="r mono hi">{fmt(c,2)}</td>
                            <td className="r mono dim">{selectedInvoice.amount?(c/selectedInvoice.amount*100).toFixed(1):0}%</td>
                            <td><Bar pct={selectedInvoice.amount?c/selectedInvoice.amount*100:0} color={tc?.color||"#60a5fa"}/></td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}

              {/* By Subscription */}
              {selectedInvoice.bySub&&Object.keys(selectedInvoice.bySub).length>0&&(
                <div className="inv-section">
                  <div className="inv-section-title">Cost by Subscription</div>
                  <table className="tbl">
                    <thead><tr><th>Subscription</th><th className="r">Amount</th><th className="r">% Share</th><th>Distribution</th></tr></thead>
                    <tbody>
                      {Object.entries(selectedInvoice.bySub).sort((a,b)=>b[1]-a[1]).map(([sub,cost])=>{
                        const subData=P?.subs?.find(s=>s.name===sub);
                        return(
                          <tr key={sub}>
                            <td><span className="dot" style={{background:subData?.tenantColor||"#60a5fa"}}/>{sub}</td>
                            <td className="r mono hi">{fmt(cost,2)}</td>
                            <td className="r mono dim">{selectedInvoice.amount?(cost/selectedInvoice.amount*100).toFixed(1):0}%</td>
                            <td><Bar pct={selectedInvoice.amount?cost/selectedInvoice.amount*100:0} color={subData?.tenantColor||"#60a5fa"}/></td>
                          </tr>
                        );
                      })}
                      <tr style={{borderTop:"2px solid var(--b2)"}}>
                        <td><strong>Total</strong></td>
                        <td className="r mono hi"><strong>{fmt(selectedInvoice.amount,2)}</strong></td>
                        <td className="r mono">100%</td>
                        <td/>
                      </tr>
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
