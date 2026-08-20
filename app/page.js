"use client";
import { useMemo, useState } from "react";
import * as XLSX from "xlsx";

const inr = new Intl.NumberFormat("en-IN", { style: "currency", currency: "INR", maximumFractionDigits: 0 });
const clean = (v) => String(v ?? "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
const n = (v) => typeof v === "number" ? v : Number(String(v ?? "").replace(/[,₹]/g, "")) || 0;
const show = (v) => inr.format(v || 0);
const col = (heads, choices) => choices.map(clean).map(x => heads.map(clean).indexOf(x)).find(x => x >= 0) ?? -1;

async function readFile(file) {
  const data = await file.arrayBuffer();
  const wb = XLSX.read(data, { type: "array", cellDates: true });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: "", raw: true });
  const at = rows.findIndex(r => r.map(clean).includes("vendorname") && r.map(clean).some(x => x === "billdate" || x === "date"));
  if (at < 0) throw new Error("I could not find the Zoho header row. Please use a Zoho purchase export containing Bill Date and Vendor Name.");
  const h = rows[at], i = {date:col(h,["Bill Date","Date"]),vendor:col(h,["Vendor Name","Vendor"]),bill:col(h,["Bill Number","Invoice Number"]),item:col(h,["Item Name","Item"]),qty:col(h,["Quantity","Qty"]),rate:col(h,["Rate","Item Rate"]),total:col(h,["Item Total","Line Item Total","Total"]),branch:col(h,["Branch Name","Branch"])};
  if (i.vendor < 0 || i.date < 0) throw new Error("Bill Date and Vendor Name are required.");
  const get = (row,key) => i[key] >= 0 ? row[i[key]] : "";
  const records = rows.slice(at+1).map((row,id) => { const qty=n(get(row,"qty")), rate=n(get(row,"rate")); return {id,date:String(get(row,"date")||""),vendor:String(get(row,"vendor")||"").trim(),bill:String(get(row,"bill")||"").trim(),item:String(get(row,"item")||"").trim(),qty,rate,total:n(get(row,"total")) || qty*rate,branch:String(get(row,"branch")||"").trim()};}).filter(r => r.vendor && clean(r.vendor) !== "vendorname" && (r.item || r.total));
  return { name:file.name, records };
}
function groups(rows,key) { return rows.reduce((m,r)=>{const k=key(r);m.set(k,[...(m.get(k)||[]),r]);return m;},new Map()); }
function rollup(rows,field) { return [...groups(rows,r=>clean(r[field])||"unassigned")].map(([key,list])=>({key,label:list[0][field]||"Unassigned",total:list.reduce((s,r)=>s+r.total,0)})); }
function analyse(current, previous, thresholds) {
  const {vendor=20,item=25,price=20}=thresholds||{};
  const vendorCut=vendor/100, itemCut=item/100, priceMult=1+price/100;
  const exact=groups(current.records,r=>[clean(r.date),clean(r.vendor),clean(r.bill),clean(r.item),r.qty,r.rate,clean(r.branch)].join("|"));
  const near=groups(current.records,r=>[clean(r.date),clean(r.vendor),clean(r.item),r.qty,r.rate,clean(r.branch)].join("|"));
  const duplicates=[];
  exact.forEach(rows=>{if(rows.length>1)duplicates.push({kind:"Confirmed duplicate",risk:"Critical",rows,total:rows.reduce((s,r)=>s+r.total,0)})});
  near.forEach(rows=>{const bills=new Set(rows.map(r=>clean(r.bill)).filter(Boolean));if(rows.length>1&&bills.size>1)duplicates.push({kind:"Same details, different bill no.",risk:"Review",rows,total:rows.reduce((s,r)=>s+r.total,0)})});
  const compare=(field,cutoff)=>{const prior=new Map(rollup(previous.records,field).map(x=>[x.key,x]));return rollup(current.records,field).map(x=>{const p=prior.get(x.key), old=p?.total||0, diff=x.total-old;return {...x,old,diff,pct:old?diff/old:null,status:p?"Changed":"New"};}).filter(x=>x.status==="New"||Math.abs(x.pct||0)>=cutoff).sort((a,b)=>Math.abs(b.diff)-Math.abs(a.diff));};
  const prices=[]; groups(current.records.filter(r=>r.item&&r.qty&&r.rate),r=>clean(r.item)).forEach(rows=>{const avg=rows.reduce((s,r)=>s+r.total,0)/rows.reduce((s,r)=>s+r.qty,0);rows.forEach(r=>{if(r.rate>avg*priceMult)prices.push({...r,avg,pct:(r.rate-avg)/avg})})});
  return {duplicates,vendors:compare("vendor",vendorCut),items:compare("item",itemCut),prices:prices.sort((a,b)=>b.pct-a.pct)};
}
function Upload({ title, file, onChange, help }) { return <label className="upload"><input type="file" accept=".xlsx,.xls,.csv" onChange={e=>e.target.files[0]&&onChange(e.target.files[0])}/><span>↑</span><strong>{title}</strong><small>{file?.name||help}</small><em>{file?"Replace file":"Choose Excel or CSV"}</em></label>; }
function Empty({children}) { return <p className="empty">{children}</p>; }
function Table({head,children}) { return <div className="table"><table><thead><tr>{head.map(x=><th key={x}>{x}</th>)}</tr></thead><tbody>{children}</tbody></table></div>; }
function ThresholdInput({label,value,onChange}) {
  return <label className="thresh-label">{label}<div className="thresh-wrap"><input className="thresh-input" type="number" min="1" max="100" value={value} onChange={e=>onChange(Math.max(1,Math.min(100,Number(e.target.value)||1))}/><span className="thresh-pct">%</span></div></label>;
}
export default function Home(){
  const [current,setCurrent]=useState(null),[previous,setPrevious]=useState(null),[error,setError]=useState(""),[tab,setTab]=useState("Overview");
  const [thresholds,setThresholds]=useState({vendor:20,item:25,price:20});
  const setT=(k,v)=>setThresholds(t=>({...t,[k]:v}));
  const result=useMemo(()=>current&&previous?analyse(current,previous,thresholds):null,[current,previous,thresholds]);
  const upload=async(file,setter)=>{try{setError("");setter(await readFile(file))}catch(e){setError(e.message)}};
  const spend=x=>x?.records.reduce((s,r)=>s+r.total,0)||0;
  const total=result&&result.duplicates.length+result.vendors.length+result.items.length+result.prices.length;
  return <main><header><div><p className="eyebrow">FINANCE CONTROL CENTER</p><h1>P&L Audit Desk</h1></div><p className="privacy"><i/> Local analysis - files stay on your device</p></header>
  {!result?<section className="landing"><p className="eyebrow">ZOHO BOOKS PURCHASE REVIEW</p><h2>Find what the spreadsheet misses.</h2><p className="lead">Compare two Zoho purchase exports to flag duplicate bills, vendor movements, item variation, and unusual prices.</p><div className="chips"><b>Duplicate bills</b><b>Vendor variation</b><b>Item variation</b><b>Price exceptions</b></div><div className="uploads"><Upload title="Current month" help="Upload the latest Zoho export" file={current} onChange={f=>upload(f,setCurrent)}/><Upload title="Previous month" help="Upload the month to compare" file={previous} onChange={f=>upload(f,setPrevious)}/></div>{error&&<p className="error">{error}</p>}<p className="note">Required: Bill Date and Vendor Name. Add Bill Number, Item Name, Quantity, Rate, Branch, and Item Total for the complete audit.</p></section>:
  <><section className="run"><div><strong>Analysis ready</strong><small>{current.name} compared with {previous.name}</small></div><button onClick={()=>{setCurrent(null);setPrevious(null);setTab("Overview")}}>New review</button></section><section className="thresholds"><p className="eyebrow">SENSITIVITY THRESHOLDS</p><div className="thresh-row"><ThresholdInput label="Vendor variation" value={thresholds.vendor} onChange={v=>setT("vendor",v)}/><ThresholdInput label="Item variation" value={thresholds.item} onChange={v=>setT("item",v)}/><ThresholdInput label="Price exception" value={thresholds.price} onChange={v=>setT("price",v)}/><button className="thresh-reset" onClick={()=>setThresholds({vendor:20,item:25,price:20})}>Reset to defaults</button></div></section><nav>{["Overview","Duplicate bills","Vendor variation","Purchase variation","Price exceptions"].map(x=><button className={tab===x?"active":""} onClick={()=>setTab(x)} key={x}>{x}</button>)}</nav>
  {tab==="Overview"&&<><section className="metrics"><Card label="Audit findings" value={total} warm/><Card label="Current-month spend" value={show(spend(current))}/><Card label="Spend movement" value={show(spend(current)-spend(previous))}/><Card label="Rows reviewed" value={current.records.length.toLocaleString("en-IN")}/></section><section className="panel"><div className="panelhead"><div><p className="eyebrow">PRIORITY QUEUE</p><h2>What to review first</h2></div><b className="badge">{result.duplicates.length} duplicate patterns</b></div>{result.duplicates.length||result.vendors.length||result.prices.length?<div className="queue">{[...result.duplicates.slice(0,3).map(x=>({title:x.kind,detail:`${x.rows[0].vendor} - ${x.rows.length} matching lines`,value:x.total,red:x.risk==="Critical"})),...result.vendors.slice(0,3).map(x=>({title:`${x.status} vendor`,detail:x.label,value:x.diff})),...result.prices.slice(0,2).map(x=>({title:"High item price",detail:`${x.item} - ${x.vendor}`,value:x.total}))].slice(0,7).map((x,i)=><div className="queueitem" key={i}><i className={x.red?"red":"amber"}/><div><strong>{x.title}</strong><small>{x.detail}</small></div><b>{show(x.value)}</b></div>)}</div>:<Empty>No material flags found.</Empty>}</section></>}
  {tab==="Duplicate bills"&&<section className="panel"><Panel title="Duplicate bill patterns" badge={`${result.duplicates.length} findings`}/><Table head={["Classification","Vendor","Item","Matching lines","Exposure"]}>{result.duplicates.length?result.duplicates.map((x,i)=><tr key={i}><td><b className={x.risk==="Critical"?"pill critical":"pill"}>{x.kind}</b></td><td>{x.rows[0].vendor}</td><td>{x.rows[0].item||"-"}</td><td>{x.rows.length}</td><td>{show(x.total)}</td></tr>):<tr><td colSpan="5"><Empty>No duplicate patterns found.</Empty></td></tr>}</Table></section>}
  {tab==="Vendor variation"&&<Changes title="Vendor variation" rows={result.vendors} field="Vendor" threshold={thresholds.vendor}/>}{tab==="Purchase variation"&&<Changes title="Purchase variation" rows={result.items} field="Item" threshold={thresholds.item}/>}{tab==="Price exceptions"&&<section className="panel"><Panel title="Items purchased above weighted average" badge={`${thresholds.price}%+ above average`}/><Table head={["Item","Vendor","Rate paid","Weighted average","Variance"]}>{result.prices.length?result.prices.map((x,i)=><tr key={i}><td>{x.item}</td><td>{x.vendor}</td><td>{show(x.rate)}</td><td>{show(x.avg)}</td><td className="bad">+{(x.pct*100).toFixed(0)}%</td></tr>):<tr><td colSpan="5"><Empty>No price exceptions found.</Empty></td></tr>}</Table></section>}</>}</main>;
}
function Card({label,value,warm}) { return <article className={warm?"card warm":"card"}><small>{label}</small><strong>{value}</strong></article> }
function Panel({title,badge}) { return <div className="panelhead"><div><p className="eyebrow">AUDIT REVIEW</p><h2>{title}</h2></div><b className="badge">{badge}</b></div> }
function Changes({title,rows,field,threshold}) { return <section className="panel"><Panel title={title} badge={`${threshold}%+ change or new entry`}/><Table head={[field,"Current month","Previous month","Change","Status"]}>{rows.length?rows.map((x,i)=><tr key={i}><td>{x.label}</td><td>{show(x.total)}</td><td>{show(x.old)}</td><td className={x.diff>=0?"bad":"good"}>{x.diff>=0?"+":""}{show(x.diff)}</td><td><b className="pill">{x.status}</b></td></tr>):<tr><td colSpan="5"><Empty>No material movements found.</Empty></td></tr>}</Table></section> }
