import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import { storage, subscribeToKey } from "./storage";

const MI = ({ name, size = 18, style = {} }) => (
  <span className="material-icons" style={{ fontSize: size, lineHeight: 1, display: "flex", alignItems: "center", userSelect: "none", ...style }}>{name}</span>
);

const CELL_W = 36;
const SQUAD_COLORS = ["#1565C0","#388E3C","#E65100","#4527A0","#AD1457","#0288D1"];
const DESIGNER_COLORS = ["#1565C0","#388E3C","#E65100","#4527A0","#AD1457","#0288D1","#C62828","#3949AB","#00796B","#BF360C","#8E24AA","#B71C1C","#0277BD","#558B2F","#7B1FA2","#E65100","#3949AB","#00897B","#FFB300","#7E57C2"];
const SIZE_POINTS = { XL:3, L:2, M:1.5, S:1, XS:0.5 };
const OVERLOAD_THRESHOLD = 6;
const MANAGERS = ["Annie","Ashot","Carson","Debbie","Rudy"];
const MANAGER_AS_DESIGNER = new Set(["Debbie"]);
const SUB_MANAGERS = {
  Carson: ["Christine","Ian","Michael"],
  Annie:  ["Kim"],
};
const MANAGER_DESIGNERS = {
  Ashot:     ["Arevik","Gayane","Inga","Laura","Lusine","Nor","Zara"],
  Annie:     ["Ken","Jozef","Hsin","Swaroop"],
  Kim:       ["Gohar","Heidi","Shawn"],
  Carson:    ["Brian","Doug","Armine","Aida","Kristen","Alexandra","Declan","Sabrina"],
  Christine: ["Lance"],
  Ian:       ["Kevin","Lejla","Ryan"],
  Michael:   ["Jennie"],
  Debbie:    ["Christy","Mary","Ying","Sofia"],
  Rudy:      ["Andrea","Mimi","Allison","Miranda","Yiyang","Aarti","Zoey","JJ"],
};
const DEF_DESIGNERS = MANAGER_DESIGNERS["Ashot"];
const DEF_SQUADS = ["Reporting","Intelligent Support Automation","Identity","Tenant Management","API Platform","Fleet Pro","Marketing Autopilot","Reputation","Document Template Engine","Org Model","Enterprise Hub","GUM"];
const DEF_PMS = ["Daniel Weiss","Khachik Tadevosyan","Jon Diamond","Lilit Ghukasyan","Vsevolod Morozov","Stephanie Dority","Krishna Kapila","Gamaliel Obinyan","Davit Arakelyan","John Billington","Derek Browers","Bella Arutyunyan","Austin Vogel","Lusine Martirosyan","Anna Arakelyan"];
const SIZES = ["XS","S","M","L","XL"];
const STATUS = { "Backlog":{color:"#90A4AE",bg:"#F5F5F5"}, "Not Started":{color:"#90A4AE",bg:"#F5F5F5"}, "In Progress":{color:"#1565C0",bg:"#E3F2FD"}, "Dev Support":{color:"#00796B",bg:"#E0F2F1"}, "Design QA":{color:"#BF360C",bg:"#FBE9E7"}, "In Review":{color:"#E65100",bg:"#FFF8E1"}, "On Hold":{color:"#C62828",bg:"#FFEBEE"}, "Done":{color:"#2E7D32",bg:"#E8F5E9"} };
const ALL_STATUSES = ["Backlog","Not Started","In Progress","Dev Support","Design QA","In Review","On Hold","Done"];
const PROJECT_TYPES = [
  { label:"North Star",                          color:"#7C3AED", bg:"#F3EEFF" },
  { label:"New Feature",                         color:"#1565C0", bg:"#E3F2FD" },
  { label:"Enhancement",                         color:"#00796B", bg:"#E0F2F1" },
  { label:"Discovery Insights",                  color:"#E65100", bg:"#FFF3E0" },
  { label:"Experience Evaluation",               color:"#AD1457", bg:"#FCE4EC" },
  { label:"Measurement / Performance Tracking",  color:"#0277BD", bg:"#E1F5FE" },
  { label:"Direction / Prioritization",          color:"#4527A0", bg:"#EDE7F6" },
  { label:"Framework Development",               color:"#2E7D32", bg:"#E8F5E9" },
  { label:"Infrastructure",                      color:"#546E7A", bg:"#ECEFF1" },
  { label:"Other",                               color:"#78909C", bg:"#F5F5F5" },
];
const LEAVE_TYPES = ["PTO","STO","Maternity Leave","Paternity Leave"];
const LEAVE_COLORS = {"PTO":"#607D8B","STO":"#E65100","Maternity Leave":"#880E4F","Paternity Leave":"#4527A0"};
const LEAVE_STATUSES = ["Planned","In Progress","Completed"];
const SQUAD_PM_MAP = {
  "Reporting":                ["Daniel Weiss","Khachik Tadevosyan","Jon Diamond","Lilit Ghukasyan"],
  "Identity":                 ["Vsevolod Morozov","Stephanie Dority"],
  "Tenant Management":        ["Gamaliel Obinyan"],
  "API Platform":             ["Krishna Kapila"],
  "Fleet Pro":                ["Davit Arakelyan"],
  "Marketing Autopilot":      ["John Billington"],
  "Reputation":               ["Derek Browers"],
  "Document Template Engine": ["Bella Arutyunyan"],
  "Org Model":                ["Austin Vogel"],
  "GUM":                      ["Lusine Martirosyan"],
  "Enterprise Hub":           ["Anna Arakelyan"],
};
const PM_SQUAD_MAP = {};
Object.entries(SQUAD_PM_MAP).forEach(([sq,pms])=>pms.forEach(pm=>{ PM_SQUAD_MAP[pm]=sq; }));

const ASHOT_SQUADS = [
  { id:"s1", name:"Reporting", colorIdx:0, designers:[
    { id:"d1", name:"Arevik", avatar:"AR", projects:[
      { id:"p9",  name:"Command Center - Financial Performance", size:"XL", startDate:"2026-03-10", endDate:"2026-04-07", squad:"Reporting", pms:["Jon Diamond"] },
      { id:"p15", name:"Embedded BI Experience Vision",          size:"XL", startDate:"2026-02-03", endDate:"2026-05-02", squad:"Reporting", pms:["Daniel Weiss","Khachik Tadevosyan"] },
      { id:"p61", name:"Reporting & Dashboards North Star Vision", size:"XL", startDate:"2026-03-30", endDate:"2026-06-15", squad:"Reporting", pms:["Khachik Tadevosyan"] },
    ]},
    { id:"d2", name:"Laura", avatar:"LA", projects:[
      { id:"p50", name:"Quadrant Movement",                size:"XL", startDate:"2026-01-01", endDate:"2026-03-01", squad:"Reporting", pms:["Lilit Ghukasyan"] },
      { id:"p51", name:"Benchmark TI BAU",                 size:"M",  startDate:"2026-02-01", endDate:"2026-04-30", squad:"Reporting", pms:["Lilit Ghukasyan"] },
      { id:"p52", name:"Front Funnel Metrics",             size:"M",  startDate:"2026-03-25", endDate:"2026-04-25", squad:"Reporting", pms:["Lilit Ghukasyan"] },
      { id:"p53", name:"Atlas Summaries",                  size:"M",  startDate:"2026-03-30", endDate:"2026-04-30", squad:"Reporting", pms:["Lilit Ghukasyan"] },
      { id:"p54", name:"Job Breakdown Details",            size:"M",  startDate:"2026-04-23", endDate:"2026-05-25", squad:"Reporting", pms:["Lilit Ghukasyan"] },
      { id:"p55", name:"Action Planning",                  size:"M",  startDate:"2026-04-30", endDate:"2026-05-30", squad:"Reporting", pms:["Lilit Ghukasyan"] },
      { id:"p56", name:"Balanced Scorecard",               size:"M",  startDate:"2026-05-30", endDate:"2026-06-30", squad:"Reporting", pms:["Lilit Ghukasyan"] },
      { id:"p16", name:"Business Unit Benchmarking",       size:"L",  startDate:"2026-03-10", endDate:"2026-03-24", squad:"Reporting", pms:["Lilit Ghukasyan"] },
      { id:"p17", name:"Customer Metrics",                 size:"L",  startDate:"2026-03-25", endDate:"2026-04-15", squad:"Reporting", pms:["Lilit Ghukasyan"] },
      { id:"p18", name:"Benchmarks PDF Updates",           size:"M",  startDate:"2026-04-16", endDate:"2026-04-23", squad:"Reporting", pms:["Lilit Ghukasyan"] },
      { id:"p19", name:"Ongoing Maintenance & UX Debts",  size:"S",  startDate:"2026-04-24", endDate:"2026-04-28", squad:"Reporting", pms:["Lilit Ghukasyan"] },
      { id:"p20", name:"Experiments (Atlas, DIY Metrics)", size:"S",  startDate:"2026-04-29", endDate:"2026-05-05", squad:"Reporting", pms:["Lilit Ghukasyan"] },
    ]},
  ]},
  { id:"s2", name:"Intelligent Support Automation", colorIdx:1, designers:[
    { id:"d3", name:"Zara", avatar:"ZA", projects:[
      { id:"leave_zara_mat", name:"Maternity Leave", type:"leave", leaveType:"Maternity Leave", leaveStatus:"In Progress", startDate:"2026-02-01", endDate:"2027-04-25", size:"XS", squad:"", pms:[], status:"" },
    ] },
  ]},
  { id:"s3", name:"Tenant Mgmt . Identity . API Platform", colorIdx:2, designers:[
    { id:"d4", name:"Inga", avatar:"IN", projects:[
      { id:"p29",  name:"Multi Session Management BAU",             size:"S", startDate:"2026-02-01", endDate:"2026-04-01", squad:"Identity",          pms:["Vsevolod Morozov"] },
      { id:"p24",  name:"Provisioning & Registry",                  size:"M", startDate:"2026-02-02", endDate:"2026-03-27", squad:"Tenant Management", pms:["Gamaliel Obinyan"] },
      { id:"p5",   name:"2-step Login",                             size:"L", startDate:"2026-02-16", endDate:"2026-03-26", squad:"Identity",          pms:["Stephanie Dority"] },
      { id:"p6",   name:"API Tunneling",                            size:"S", startDate:"2026-03-03", endDate:"2026-03-13", squad:"API Platform",      pms:["Krishna Kapila"] },
      { id:"p8",   name:"Dev Portal Sign Up Flow Redesign",         size:"M", startDate:"2026-03-16", endDate:"2026-03-30", squad:"API Platform",      pms:["Krishna Kapila"] },
      { id:"p22",  name:"MFA Admin Enforcement: Improvements",      size:"S", startDate:"2026-03-15", endDate:"2026-03-29", squad:"Identity",          pms:["Stephanie Dority"] },
      { id:"p28",  name:"Settings Page Granularity",                size:"M", startDate:"2026-04-10", endDate:"2026-04-15", squad:"Identity",          pms:["Vsevolod Morozov"], status:"Not Started" },
      { id:"p25",  name:"Context Switcher",                         size:"M", startDate:"2026-03-15", endDate:"2026-04-15", squad:"Identity",          pms:["Stephanie Dority"] },
      { id:"p23",  name:"SSO Phase 3 - Self-Service Administration", size:"M", startDate:"2026-04-10", endDate:"2026-05-01", squad:"Identity",         pms:["Stephanie Dority"] },
      { id:"p26",  name:"Account Linking & Dashboard",              size:"M", startDate:"2026-03-20", endDate:"2026-03-31", squad:"Identity",          pms:["Stephanie Dority"], status:"Done" },
      { id:"p16b", name:"Access Control & Auditability",            size:"S", startDate:"2026-03-20", endDate:"2026-04-05", squad:"Tenant Management", pms:["Gamaliel Obinyan"] },
      { id:"p27",  name:"User Profile Wizard",                      size:"M", startDate:"2026-05-01", endDate:"2026-06-10", squad:"Identity",          pms:["Stephanie Dority"] },
    ]},
  ]},
  { id:"s4", name:"Fleet Pro . Reputation . Marketing Autopilot", colorIdx:3, designers:[
    { id:"d5", name:"Lusine", avatar:"LU", projects:[
      { id:"p31", name:"Improve Autopilot Discovery BAU",    size:"S", startDate:"2026-02-01", endDate:"2026-04-30", squad:"Marketing Autopilot", pms:["John Billington"], status:"On Hold" },
      { id:"p21", name:"Fully Automated AI Review Response", size:"M", startDate:"2026-02-10", endDate:"2026-03-09", squad:"Reputation",          pms:["Derek Browers"] },
      { id:"p11", name:"Activity Timeline",                  size:"M", startDate:"2025-12-20", endDate:"2026-04-15", squad:"Fleet Pro",            pms:["Davit Arakelyan"], status:"On Hold" },
      { id:"p12", name:"Branding in Campaigns",              size:"M", startDate:"2026-04-10", endDate:"2026-05-10", squad:"Marketing Autopilot", pms:["John Billington"], status:"Not Started" },
      { id:"p10", name:"AEO Reporting",                      size:"M", startDate:"2026-03-11", endDate:"2026-04-05", squad:"Reputation",          pms:["Derek Browers"] },
      { id:"p14", name:"Image Swap Improvements",            size:"M", startDate:"2026-03-12", endDate:"2026-04-15", squad:"Marketing Autopilot", pms:["John Billington"] },
      { id:"p13", name:"Custom Campaigns",                   size:"S", startDate:"2026-04-10", endDate:"2026-05-10", squad:"Marketing Autopilot", pms:["John Billington"], status:"On Hold" },
      { id:"p30", name:"Featured Jobs Widget",               size:"M", startDate:"2026-04-20", endDate:"2026-05-20", squad:"Reputation",          pms:["Derek Browers"], status:"Not Started" },
      { id:"p33", name:"Atlas Summaries",                    size:"M", startDate:"2026-04-01", endDate:"2026-05-01", squad:"Fleet Pro",            pms:["Davit Arakelyan"] },
      { id:"p34", name:"Driver Roles",                       size:"M", startDate:"2026-04-01", endDate:"2026-05-01", squad:"Fleet Pro",            pms:["Davit Arakelyan"] },
      { id:"p32", name:"Safety Cam Remote Configurations",   size:"M", startDate:"2026-04-15", endDate:"2026-04-30", squad:"Fleet Pro",            pms:["Davit Arakelyan"] },
    ]},
  ]},
  { id:"s5", name:"Org Model . Doc Templates . Global User Mgmt", colorIdx:4, designers:[
    { id:"d6", name:"Gayane", avatar:"GA", projects:[
      { id:"p49", name:"Folder View",                   size:"L",  startDate:"2026-02-10", endDate:"2026-03-30", squad:"Org Model",                pms:["Austin Vogel"], status:"In Progress" },
      { id:"p45", name:"Org Selector Component",        size:"S",  startDate:"2026-03-20", endDate:"2026-04-30", squad:"Org Model",                pms:["Austin Vogel"] },
      { id:"p60", name:"Template Manager Northstar",    size:"XL", startDate:"2026-03-15", endDate:"2026-05-15", squad:"Document Template Engine", pms:["Bella Arutyunyan"] },
    ]},
  ]},
  { id:"s6", name:"Internal Billing . Enterprise Hub", colorIdx:5, designers:[
    { id:"d7", name:"Nor", avatar:"NO", projects:[
      { id:"p37", name:"User Creation",                     size:"M", startDate:"2026-03-20", endDate:"2026-04-30", squad:"GUM",            pms:["Lusine Martirosyan"] },
      { id:"p35", name:"Centralized Billing BAU",           size:"S", startDate:"2026-03-01", endDate:"2026-04-30", squad:"Enterprise Hub", pms:["Anna Arakelyan"] },
      { id:"p1",  name:"Org Studio North Star Ph.1",        size:"L", startDate:"2026-03-10", endDate:"2026-03-27", squad:"Enterprise Hub", pms:["Austin Vogel","David Matevosyan"] },
      { id:"p2",  name:"Titan Exchange Redesign",           size:"L", startDate:"2026-03-12", endDate:"2026-04-10", squad:"Enterprise Hub", pms:["Anna Arakelyan"] },
      { id:"p36", name:"Audit Trails Enhancements",         size:"M", startDate:"2026-03-20", endDate:"2026-05-05", squad:"Enterprise Hub", pms:["Anna Arakelyan"] },
      { id:"p38", name:"Role Creation",                     size:"M", startDate:"2026-04-30", endDate:"2026-05-30", squad:"GUM",            pms:["Lusine Martirosyan"] },
      { id:"p39", name:"Audit Trails and Roles Comparison", size:"M", startDate:"2026-05-30", endDate:"2026-06-30", squad:"GUM",            pms:["Lusine Martirosyan"] },
      { id:"p40", name:"Home Page",                         size:"M", startDate:"2026-06-30", endDate:"2026-07-30", squad:"GUM",            pms:["Lusine Martirosyan"] },
      { id:"p41", name:"Cleanup Tool",                      size:"M", startDate:"2026-07-30", endDate:"2026-08-30", squad:"GUM",            pms:["Lusine Martirosyan"] },
    ]},
  ]},
];

const INITIAL_STATE = {
  managers: {
    Ashot:  { squads: ASHOT_SQUADS },
    Carson:    { squads: [] },
    Christine: { squads: [] },
    Ian:       { squads: [] },
    Michael:   { squads: [] },
    Annie:  { squads: [] },
    Kim:    { squads: [] },
    Debbie: { squads: [] },
    Rudy:   { squads: [] },
  },
};

const mkDefaultSquads = mgr => {
  const designers = MANAGER_DESIGNERS[mgr] || [];
  if (!designers.length) return [];
  return [{
    id: `sq_default_${mgr}`,
    name: "Team",
    colorIdx: 0,
    designers: designers.map((name, i) => ({
      id: `d_default_${mgr}_${i}`,
      name,
      avatar: name.slice(0, 2).toUpperCase(),
      projects: [],
    })),
  }];
};

const parseD = iso => { const [y,m,d]=iso.split("-").map(Number); return new Date(y,m-1,d); };
const toIso = d => { const y=d.getFullYear(),m=String(d.getMonth()+1).padStart(2,"0"),dd=String(d.getDate()).padStart(2,"0"); return `${y}-${m}-${dd}`; };
const addW = (iso,n) => { const d=parseD(iso); d.setDate(d.getDate()+n*7); return toIso(d); };
const fmtS = iso => parseD(iso).toLocaleDateString("en-US",{month:"numeric",day:"numeric"});
const fmtM = iso => parseD(iso).toLocaleDateString("en-US",{month:"short",day:"numeric",year:"numeric"});

function qBounds() {
  const now=new Date(), q=Math.floor(now.getMonth()/3);
  return { start:new Date(now.getFullYear(),q*3,1), end:new Date(now.getFullYear(),q*3+6,0), label:`Q${q+1}-Q${q+2} ${now.getFullYear()}` };
}
function getWeeks() {
  const {start,end}=qBounds(); const weeks=[]; const cur=new Date(start);
  while(cur.getDay()!==1) cur.setDate(cur.getDate()+1);
  while(cur<=end){weeks.push(new Date(cur));cur.setDate(cur.getDate()+7);}
  return weeks;
}
const isActive = proj => { const t=new Date();t.setHours(0,0,0,0); return parseD(proj.startDate)<=t&&parseD(proj.endDate)>=t; };
const isLeaveItem = p => p.type==="leave"||(p.id&&p.id.indexOf("leave_")===0)||LEAVE_TYPES.indexOf(p.name)!==-1;
const autoStatus = proj => { if(proj.status&&STATUS[proj.status]) return proj.status; const t=new Date();t.setHours(0,0,0,0);const s=parseD(proj.startDate),e=parseD(proj.endDate); return e<t?"Done":s<=t?"In Progress":"Not Started"; };
const OVERLOAD_SKIP_STATUSES = new Set(["On Hold","Done","Backlog","Deferred","Complete"]);

function getFlags(designer) {
  const flags=[];
  if(!designer.projects.length) return flags;
  const today=new Date(); today.setHours(0,0,0,0);
  const active=designer.projects.filter(p=>parseD(p.startDate)<=today&&parseD(p.endDate)>=today&&!OVERLOAD_SKIP_STATUSES.has(autoStatus(p))&&!isLeaveItem(p));
  const score=active.reduce((s,p)=>s+(SIZE_POINTS[p.size]||0),0);
  if(active.length>=2&&score>=OVERLOAD_THRESHOLD) flags.push({week:today.toLocaleDateString("en-US",{month:"short",day:"numeric"}),detail:active.map(p=>`${p.name} (${p.size})`).join(", "),score,count:active.length});
  return flags;
}
const allOverloaded = squads => { const r=[]; squads.forEach(sq=>sq.designers.forEach(d=>{const f=getFlags(d);if(f.length)r.push({designer:d.name,squad:sq.name,flags:f,projects:d.projects});})); return r; };

// ---------- MultiSelect ----------
function MultiSelect({options, selected, onChange, placeholder}) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const ref = useRef(null);
  useEffect(() => {
    const h = e => { if(ref.current&&!ref.current.contains(e.target)){setOpen(false);setQuery("");} };
    const k = e => { if(e.key==="Escape"){setOpen(false);setQuery("");} };
    document.addEventListener("mousedown",h); document.addEventListener("keydown",k);
    return () => { document.removeEventListener("mousedown",h); document.removeEventListener("keydown",k); };
  }, []);
  const toggle = v => onChange(selected.includes(v)?selected.filter(x=>x!==v):[...selected,v]);
  const filtered = options.filter(o=>o.toLowerCase().includes(query.toLowerCase()));
  return (
    <div ref={ref} style={{position:"relative"}}>
      <div onClick={()=>setOpen(o=>!o)} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"0 10px",minHeight:40,borderRadius:4,border:`1px solid ${open?"#2563EB":"#ECEFF1"}`,background:"#FAFAFA",cursor:"pointer",fontFamily:"'Inter',sans-serif"}}>
        <div style={{display:"flex",flexWrap:"wrap",gap:6,flex:1}}>
          {selected.length===0
            ? <span style={{fontSize:14,color:"#90A4AE",lineHeight:"24px"}}>{placeholder}</span>
            : selected.map(s=>(
              <span key={s} style={{display:"inline-flex",alignItems:"center",gap:4,fontSize:12,fontWeight:500,padding:"3px 4px 3px 10px",borderRadius:8,background:"#E8F0FE",color:"#1A73E8",border:"1px solid #C5D8FD",boxShadow:"0 1px 2px rgba(0,0,0,0.06)",letterSpacing:"0.01em"}}>
                {s}
                <span onClick={e=>{e.stopPropagation();toggle(s);}} style={{cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",width:18,height:18,borderRadius:"50%",background:"transparent",transition:"background 0.1s"}} onMouseEnter={e=>e.currentTarget.style.background="#C5D8FD"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                  <MI name="close" size={16} style={{color:"#1A73E8"}} />
                </span>
              </span>
            ))
          }
        </div>
        <div style={{display:"flex",alignItems:"center",gap:4,flexShrink:0,marginLeft:6}}>
          {selected.length>0&&<span onClick={e=>{e.stopPropagation();onChange([]);}} style={{color:"#B0BEC5",cursor:"pointer",display:"flex"}} onMouseEnter={e=>e.currentTarget.style.color="#EF5350"} onMouseLeave={e=>e.currentTarget.style.color="#B0BEC5"}><MI name="close" size={16} /></span>}
          <span style={{color:"#90A4AE",transform:open?"rotate(180deg)":"none",transition:"transform 0.15s",display:"flex"}}><MI name="expand_more" size={18} /></span>
        </div>
      </div>
      {open&&<div style={{position:"absolute",top:"calc(100% + 4px)",left:0,right:0,background:"#fff",border:"1px solid #ECEFF1",borderRadius:8,boxShadow:"0 4px 20px rgba(0,0,0,0.12)",zIndex:9999}}>
        <div style={{padding:"8px 10px",borderBottom:"1px solid #F5F5F5"}}><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Search..." onClick={e=>e.stopPropagation()} autoFocus style={{width:"100%",border:"none",outline:"none",fontSize:12,color:"#0F172A",background:"transparent",fontFamily:"'Inter',sans-serif"}}/></div>
        <div style={{maxHeight:180,overflowY:"auto"}}>
          {filtered.length===0&&<div style={{padding:"8px 12px",fontSize:12,color:"#B0BEC5"}}>No results</div>}
          {filtered.map(o=>{const sel=selected.includes(o);return(
            <div key={o} onClick={()=>toggle(o)} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"8px 12px",fontSize:14,color:"#0F172A",cursor:"pointer",background:sel?"#EEF4FD":"transparent",fontFamily:"'Inter',sans-serif"}} onMouseEnter={e=>e.currentTarget.style.background=sel?"#E1EDFC":"#FAFAFA"} onMouseLeave={e=>e.currentTarget.style.background=sel?"#EEF4FD":"transparent"}>
              <span style={{fontWeight:sel?600:400,color:sel?"#1A73E8":"#0F172A"}}>{o}</span>
              {sel&&<span style={{display:"flex",color:"#1A73E8"}}><MI name="check" size={13} /></span>}
            </div>
          );})}
        </div>
      </div>}
    </div>
  );
}

// ---------- CustomSelect ----------
function CustomSelect({value, onChange, options, placeholder, clearable, style={}}) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [dropPos, setDropPos] = useState({top:0,left:0,width:0});
  const ref = useRef(null);
  const searchRef = useRef(null);
  useEffect(() => {
    const h = e => { if(ref.current&&!ref.current.contains(e.target)){setOpen(false);setQuery("");} };
    const k = e => { if(e.key==="Escape"){setOpen(false);setQuery("");} };
    document.addEventListener("mousedown",h); document.addEventListener("keydown",k);
    return () => { document.removeEventListener("mousedown",h); document.removeEventListener("keydown",k); };
  }, []);
  useEffect(() => { if(open) setTimeout(()=>searchRef.current?.focus(),50); }, [open]);
  const handleOpen = () => {
    if(ref.current) {
      const r = ref.current.getBoundingClientRect();
      setDropPos({top: r.bottom+4, left: r.left, width: r.width});
    }
    setOpen(o=>!o);
  };
  const filtered = options.filter(o=>o.isGroup||((o.label||o).toLowerCase().includes(query.toLowerCase())));
  const sel = options.find(o=>!o.isGroup&&(o.value||o)===value);
  const label = sel?(sel.triggerLabel||sel.label||sel):<span style={{color:"#90A4AE"}}>{placeholder}</span>;
  return (
    <div ref={ref} style={{position:"relative",...style}}>
      <div onClick={handleOpen} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"0 10px",minHeight:40,borderRadius:4,border:`1px solid ${open?"#2563EB":"#ECEFF1"}`,background:"#FAFAFA",cursor:"pointer",fontSize:14,color:"#0F172A",userSelect:"none",fontFamily:"'Inter',sans-serif"}}>
        <span>{label}</span>
        <span style={{color:"#90A4AE",transform:open?"rotate(180deg)":"none",transition:"transform 0.15s",display:"flex"}}><MI name="expand_more" size={18} /></span>
      </div>
      {open&&<div style={{position:"fixed",top:dropPos.top,left:dropPos.left,width:dropPos.width,background:"#fff",border:"1px solid #ECEFF1",borderRadius:8,boxShadow:"0 8px 24px rgba(0,0,0,0.12)",zIndex:99999}}>
        <div style={{padding:"8px 10px",borderBottom:"1px solid #F5F5F5"}}><input ref={searchRef} value={query} onChange={e=>setQuery(e.target.value)} placeholder="Search..." onClick={e=>e.stopPropagation()} style={{width:"100%",border:"none",outline:"none",fontSize:12,color:"#0F172A",background:"transparent",fontFamily:"'Inter',sans-serif"}}/></div>
        <div style={{maxHeight:180,overflowY:"auto"}}>
          {placeholder&&clearable!==false&&<div onClick={()=>{onChange("");setOpen(false);setQuery("");}} style={{padding:"8px 12px",fontSize:14,color:"#90A4AE",cursor:"pointer",fontFamily:"'Inter',sans-serif"}} onMouseEnter={e=>e.currentTarget.style.background="#FAFAFA"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>{placeholder}</div>}
          {filtered.length===0&&<div style={{padding:"8px 12px",fontSize:12,color:"#B0BEC5"}}>No results</div>}
          {filtered.map((o,i)=>{
            if(o.isGroup) return <div key={"g"+i} style={{padding:"6px 12px 2px",fontSize:11,fontWeight:700,color:"#90A4AE",textTransform:"uppercase",letterSpacing:"0.07em",userSelect:"none"}}>{o.label}</div>;
            const v=o.value||o;const l=o.label||o;return(
            <div key={v} onClick={()=>{onChange(v);setOpen(false);setQuery("");}} style={{padding:"8px 12px",fontSize:14,color:"#0F172A",cursor:"pointer",background:v===value?"#F5F5F5":"transparent",fontWeight:v===value?700:400,fontFamily:"'Inter',sans-serif"}} onMouseEnter={e=>e.currentTarget.style.background="#FAFAFA"} onMouseLeave={e=>e.currentTarget.style.background=v===value?"#F5F5F5":"transparent"}>{l}</div>
          );})}
        </div>
      </div>}
    </div>
  );
}

// ---------- Toast ----------
function Toast({toasts}) {
  return (
    <div style={{position:"fixed",bottom:24,left:24,zIndex:99999,display:"flex",flexDirection:"column",alignItems:"flex-start",gap:8,pointerEvents:"none"}}>
      {toasts.map(t=>{
        const isError = t.type==="remove"||t.type==="error";
        return (
          <div key={t.id} style={{display:"flex",alignItems:"center",gap:10,padding:"14px 20px",borderRadius:4,background:"#1C1B1F",color:"#F4EFF4",fontSize:14,fontWeight:500,fontFamily:"'Inter',sans-serif",boxShadow:"0 4px 12px rgba(0,0,0,0.3)",animation:"slideIn 0.2s ease",whiteSpace:"nowrap",letterSpacing:"0.01em"}}>
            <span style={{display:"flex",flexShrink:0,color:isError?"#F2B8B5":"#6DD58C"}}>
              <MI name={isError?"error_outline":"check_circle"} size={20} />
            </span>
            {t.msg}
          </div>
        );
      })}
    </div>
  );
}

// ---------- InfoTip ----------
function InfoTip({text}) {
  const [show,setShow]=useState(false);
  return (
    <span onMouseEnter={()=>setShow(true)} onMouseLeave={()=>setShow(false)} style={{color:"#B0BEC5",cursor:"help",display:"inline-flex",position:"relative",flexShrink:0}}>
      <MI name="info" size={14} />
      {show&&<div style={{position:"absolute",top:"calc(100% + 6px)",left:"50%",transform:"translateX(-50%)",width:240,background:"#1C1B1F",color:"#fff",fontSize:11,lineHeight:1.5,padding:"8px 12px",borderRadius:8,zIndex:100,pointerEvents:"none",boxShadow:"0 4px 12px rgba(0,0,0,0.2)",textTransform:"none",letterSpacing:"normal",fontWeight:400}}>{text}</div>}
    </span>
  );
}

// ---------- KpiCard ----------
function KpiCard({label,value,color,tip,onClick,sub}) {
  const [show,setShow]=useState(false);
  return (
    <div onClick={onClick} style={{background:"#fff",borderRadius:16,padding:"20px 24px",border:"1px solid #E8EAED",boxShadow:"0 1px 3px rgba(0,0,0,0.06)",position:"relative",cursor:onClick?"pointer":"default",transition:"box-shadow 0.2s,border-color 0.2s"}}
      onMouseEnter={e=>{if(onClick){e.currentTarget.style.boxShadow="0 4px 16px rgba(0,0,0,0.1)";e.currentTarget.style.borderColor="#C5D8FD";}}}
      onMouseLeave={e=>{if(onClick){e.currentTarget.style.boxShadow="0 1px 3px rgba(0,0,0,0.06)";e.currentTarget.style.borderColor="#E8EAED";}}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
        <div style={{fontSize:12,fontWeight:500,color:"#6B7280",letterSpacing:"0.01em"}}>{label}</div>
        {tip&&<div onMouseEnter={()=>setShow(true)} onMouseLeave={()=>setShow(false)} style={{color:"#C4C9D4",cursor:"default",flexShrink:0,display:"flex",alignItems:"center"}}><MI name="info" size={16}/></div>}
      </div>
      <div style={{fontSize:32,fontWeight:700,color,fontFamily:"'Inter',sans-serif",letterSpacing:"-0.02em",lineHeight:1}}>{value}</div>
      {sub&&<div style={{fontSize:12,color:"#9CA3AF",marginTop:8}}>{sub}</div>}
      {show&&tip&&<div style={{position:"absolute",top:"calc(100% + 6px)",left:0,width:220,background:"#1C1B1F",color:"#fff",fontSize:11,lineHeight:1.5,padding:"8px 12px",borderRadius:8,zIndex:100,pointerEvents:"none",boxShadow:"0 4px 12px rgba(0,0,0,0.2)"}}>{tip}</div>}
    </div>
  );
}

// ---------- Avatar ----------
const Avatar = ({initials,color,emoji}) => (
  <div style={{width:30,height:30,borderRadius:"50%",flexShrink:0,background:emoji?"#F5F5F5":color+"18",border:emoji?"1px solid #ECEFF1":`1px solid ${color}55`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:emoji?15:10,fontWeight:700,color,letterSpacing:"0.04em",fontFamily:emoji?"inherit":"monospace"}}>{emoji||initials}</div>
);

// ---------- Emoji keyword index ----------
const EMOJI_KEYWORDS = {
  "😀":"smile grin happy","😃":"smile big eyes happy","😄":"smile happy laugh","😁":"grin happy","😆":"laugh happy","😅":"sweat laugh","🤣":"rofl laugh","😂":"joy cry laugh","🙂":"smile","🙃":"upside smile","😉":"wink","😊":"blush smile happy","😇":"angel halo","🥰":"love hearts","😍":"heart eyes love","🤩":"star eyes wow","😘":"kiss","😚":"kiss","😋":"yum tongue","😛":"tongue","😜":"wink tongue","😝":"squint tongue","🤑":"money","😎":"cool sunglasses","🥳":"party celebrate","😕":"confused","😟":"worried","😢":"cry sad","😭":"sob cry sad","😱":"scream shocked","😤":"angry frustrated","😡":"angry red","😠":"angry","🤬":"cursing","😈":"devil evil","👿":"devil angry","💀":"skull death","🤡":"clown","👻":"ghost","❤️":"heart love red","🧡":"heart orange","💛":"heart yellow","💚":"heart green","💙":"heart blue","💜":"heart purple","🖤":"heart black","💔":"broken heart","💯":"100 perfect","✅":"check done yes","❌":"x no wrong","⚠️":"warning","🔴":"red circle","🟢":"green circle","🔵":"blue circle",
  "🐶":"dog puppy","🐱":"cat kitten","🐭":"mouse","🐹":"hamster","🐰":"rabbit bunny","🦊":"fox","🐻":"bear","🐼":"panda","🐨":"koala","🐯":"tiger","🦁":"lion","🐮":"cow","🐷":"pig","🐸":"frog","🐵":"monkey","🐔":"chicken","🐧":"penguin","🐦":"bird","🦆":"duck","🦅":"eagle","🦉":"owl","🦇":"bat","🐺":"wolf","🐴":"horse","🦄":"unicorn","🐝":"bee","🦋":"butterfly","🐢":"turtle","🐙":"octopus","🦈":"shark","🐳":"whale","🐬":"dolphin","🐘":"elephant","🦒":"giraffe","🐕":"dog","🐈":"cat",
  "🌵":"cactus","🎄":"christmas tree","🌲":"tree","🌳":"tree","🌴":"palm tree","🌱":"sprout plant","🌿":"herb plant","🍀":"clover lucky","🌷":"tulip flower","🌹":"rose flower","🌺":"hibiscus flower","🌸":"cherry blossom flower","🌼":"blossom flower","🌻":"sunflower","🌞":"sun happy","⭐":"star","🌟":"star glowing","✨":"sparkles","⚡":"lightning bolt","🔥":"fire hot","🌈":"rainbow","❄️":"snow cold ice","💧":"water drop","🌊":"wave ocean",
  "🍕":"pizza food","🍔":"burger food","🍟":"fries food","🌮":"taco food","🍣":"sushi food","🍜":"noodles food","🍩":"donut food","🍪":"cookie food","🎂":"cake birthday","🍰":"cake slice","☕":"coffee hot","🍵":"tea","🍺":"beer","🍷":"wine","🥂":"champagne celebrate",
  "💻":"laptop computer","📱":"phone mobile","⌨️":"keyboard","🖥️":"desktop monitor","📷":"camera photo","🔋":"battery","💡":"light idea","🔧":"wrench tool","⚙️":"gear settings","🔑":"key","🔒":"lock","🔓":"unlock","🔍":"search magnify","✏️":"pencil write","✂️":"scissors cut","📌":"pin","📎":"paperclip","📊":"chart graph","📈":"chart up","📁":"folder","📚":"books","📖":"book read","🎁":"gift present","🎉":"party celebrate","🎈":"balloon party","🏆":"trophy win","🥇":"gold medal win","🎯":"target dart","🎮":"game controller","🎵":"music note","🎨":"art palette","🔮":"crystal ball","💎":"diamond gem","💰":"money bag","💳":"card",
  "👶":"baby","👦":"boy","👧":"girl","🧑":"person","👱":"blonde","👨":"man","👩":"woman","🧔":"beard","👴":"old man","👵":"old woman","🏃":"running","💃":"dance","🕺":"dance man","👯":"dancers","🧗":"climbing","🏋️":"weightlifting","🚴":"cycling","🏊":"swimming","🧘":"yoga meditation","👍":"thumbs up good","👎":"thumbs down bad","👋":"wave hello","🤝":"handshake","👏":"clap applause","🙏":"pray thanks","💪":"muscle strong flex","🫶":"heart hands love",
  "🏁":"flag finish","🚩":"red flag","🌍":"earth globe","🌎":"earth globe","🌏":"earth globe asia","🏔️":"mountain","🌋":"volcano","🏠":"house home","🏢":"office building","🚀":"rocket","✈️":"airplane travel","🚗":"car","🚕":"taxi","🚌":"bus","🛒":"shopping cart",
};

// ---------- TitlePopover ----------
const TITLE_OPTIONS = ["Product Designer","Senior Product Designer","Staff Product Designer","Senior Staff Product Designer"];

function TitlePopover({name, title, squads, assignedSquads, emoji, onSave, onClose}) {
  const [selTitle, setSelTitle] = useState(title||"");
  const [selSquads, setSelSquads] = useState(assignedSquads||[]);
  const [selEmoji, setSelEmoji] = useState(emoji||"");
  const [showEmojiPicker, setShowEmojiPicker] = useState(false);
  const [emojiSearch, setEmojiSearch] = useState("");
  const [emojiCat, setEmojiCat] = useState("Smileys");
  const EMOJI_DATA = {"Smileys":["😀","😃","😄","😁","😆","😅","🤣","😂","🙂","🙃","😉","😊","😇","🥰","😍","🤩","😘","😗","😚","😙","🥲","😋","😛","😜","🤪","😝","🤑","🤗","🤭","🤫","🤔","🫡","🤐","🤨","😐","😑","😶","🫥","😏","😒","🙄","😬","🤥","😌","😔","😪","🤤","😴","😷","🤒","🤕","🤢","🤮","🥵","🥶","🥴","😵","🤯","🤠","🥳","🥸","😎","🤓","🧐","😕","🫤","😟","🙁","😮","😯","😲","😳","🥺","🥹","😦","😧","😨","😰","😥","😢","😭","😱","😖","😣","😞","😓","😩","😫","🥱","😤","😡","😠","🤬","😈","👿","💀","☠️","💩","🤡","👹","👺","👻","👽","👾","🤖"],"People":["👶","🧒","👦","👧","🧑","👱","👨","🧔","👩","🧓","👴","👵","💆","💇","🚶","🧍","🧎","🏃","💃","🕺","👯","🧖","🧗","🤸","⛹️","🏋️","🚴","🚵","🤼","🤽","🤾","🤺","🏇","⛷️","🏂","🏌️","🏄","🚣","🏊","🧘"],"Animals":["🐶","🐱","🐭","🐹","🐰","🦊","🐻","🐼","🐨","🐯","🦁","🐮","🐷","🐸","🐵","🙈","🙉","🙊","🐔","🐧","🐦","🦆","🦅","🦉","🦇","🐺","🐗","🐴","🦄","🐝","🐛","🦋","🐌","🐞","🐜","🐢","🐍","🦎","🐙","🦑","🦐","🦀","🐡","🐠","🐟","🐬","🐳","🦈","🐊","🐘","🦒","🐕","🐈"],"Nature":["🌵","🎄","🌲","🌳","🌴","🌱","🌿","☘️","🍀","🍃","🍂","🍁","🍄","🌾","💐","🌷","🌹","🥀","🌺","🌸","🌼","🌻","🌞","🌝","🌛","🌜","🌚","🌕","🌙","🌎","🌍","🌏","💫","⭐","🌟","✨","⚡","☄️","💥","🔥","🌪️","🌈","☀️","❄️","⛄","💧","💦","🌊","🏔️","🌋"],"Food":["🍇","🍈","🍉","🍊","🍋","🍌","🍍","🥭","🍎","🍏","🍐","🍑","🍒","🍓","🫐","🥝","🍅","🥑","🍆","🥕","🌽","🌶️","🥒","🥦","🍞","🥐","🧀","🍖","🍗","🍔","🍟","🍕","🌭","🥪","🌮","🌯","🍳","🍲","🥗","🍿","🍱","🍣","🍤","🍦","🍧","🍨","🍩","🍪","🎂","🍰","🧁","🍫","🍬","🍭","☕","🍵","🍺","🍻","🥂","🍷","🍸","🍹"],"Objects":["⌚","📱","💻","⌨️","🖥️","🖨️","🖱️","📷","📹","🎥","📞","☎️","📺","📻","🎙️","🔋","💡","🔦","🕯️","💸","💵","💰","💳","💎","⚖️","🔧","🔩","⚙️","🔫","💣","🔪","🛡️","🔮","📿","🧿","🔭","🔬","💊","💉","🧬","🔑","🗝️","🚪","🛋️","🛏️","🧸","🖼️","🛍️","🎁","🎈","🎀","🎊","🎉","✉️","📦","📜","📄","📊","📈","📋","📁","📰","📚","📖","🔖","🔗","📎","📌","📍","✂️","🖊️","✏️","🖌️","🔍","🔎","🔒","🔓"],"Symbols":["❤️","🧡","💛","💚","💙","💜","🖤","🤍","🤎","💔","❣️","💕","💞","💓","💗","💖","💘","💝","💯","💢","♨️","🚫","✅","❌","⭕","🛑","⚠️","♻️","💤","🏧","♿","🅿️","🆙","🆒","🆕","🆓","🔴","🟠","🟡","🟢","🔵","🟣","⚫","⚪","🟤","▶️","⏸️","⏹️","⏺️","⏭️","⏮️","⏩","⏪","🔼","🔽","➡️","⬅️","⬆️","⬇️","↗️","↘️","↙️","↖️","↕️","↔️","🔀","🔁","🔂","🔄","➕","➖","➗","✖️","💲","™️","©️","®️"],"Flags":["🏁","🚩","🎌","🏴","🏳️","🏳️‍🌈","🏴‍☠️","🇦🇲","🇺🇸","🇬🇧","🇫🇷","🇩🇪","🇮🇹","🇪🇸","🇯🇵","🇰🇷","🇨🇳","🇮🇳","🇧🇷","🇲🇽","🇨🇦","🇦🇺","🇷🇺","🇹🇷","🇺🇦","🇬🇪","🇵🇱","🇳🇱","🇨🇭","🇸🇪","🇳🇴","🇩🇰","🇫🇮","🇮🇪","🇦🇷","🇨🇴","🇵🇪"]};
  const EMOJI_CAT_ICONS = {"Smileys":"😀","People":"🧑","Animals":"🐱","Nature":"🌿","Food":"🍕","Objects":"💻","Symbols":"❤️","Flags":"🏳️"};
  const emojiCats = Object.keys(EMOJI_DATA);
  const filteredEmojis = emojiSearch.trim()
    ? Object.keys(EMOJI_DATA).flatMap(c=>EMOJI_DATA[c]).filter(e=>(EMOJI_KEYWORDS[e]||"").toLowerCase().includes(emojiSearch.toLowerCase())||e===emojiSearch)
    : (EMOJI_DATA[emojiCat]||[]);
  useEffect(()=>{
    const esc=e=>{if(e.key==="Escape"){if(showEmojiPicker)setShowEmojiPicker(false);else onClose();}};
    document.addEventListener("keydown",esc); return ()=>document.removeEventListener("keydown",esc);
  },[onClose,showEmojiPicker]);
  return (
    <>
      <div onClick={onClose} style={{position:"fixed",inset:0,zIndex:10000,background:"rgba(0,0,0,0.2)"}}/>
      <div style={{position:"fixed",zIndex:10001,top:"50%",left:"50%",transform:"translate(-50%,-50%)",background:"#fff",border:"1px solid #ECEFF1",borderRadius:28,boxShadow:"0 4px 24px rgba(0,0,0,0.12)",padding:20,width:340,maxHeight:"80vh",overflowY:"auto"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <div style={{fontSize:14,fontWeight:700,color:"#0F172A"}}>{name}</div>
          <button onClick={onClose} style={{background:"transparent",border:"none",cursor:"pointer",color:"#90A4AE",display:"flex",borderRadius:"50%",padding:4}}><MI name="close" size={16} /></button>
        </div>
        <label style={{fontSize:11,fontWeight:700,color:"#90A4AE",textTransform:"uppercase",letterSpacing:"0.05em",display:"block",marginBottom:6}}>AVATAR EMOJI</label>
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:14}}>
          <div onClick={()=>setShowEmojiPicker(!showEmojiPicker)} style={{width:36,height:36,borderRadius:"50%",background:"#F5F5F5",border:"1px solid #ECEFF1",display:"flex",alignItems:"center",justifyContent:"center",fontSize:selEmoji?20:12,cursor:"pointer"}} onMouseEnter={e=>e.currentTarget.style.borderColor="#2563EB"} onMouseLeave={e=>e.currentTarget.style.borderColor="#ECEFF1"}>
            {selEmoji||<span style={{color:"#90A4AE"}}>{name.slice(0,2).toUpperCase()}</span>}
          </div>
          {selEmoji&&<button onClick={()=>setSelEmoji("")} style={{fontSize:11,color:"#90A4AE",background:"none",border:"none",cursor:"pointer",fontFamily:"'Inter',sans-serif"}}>Remove</button>}
          {!selEmoji&&<span style={{fontSize:11,color:"#90A4AE"}}>Click to pick</span>}
        </div>
        {showEmojiPicker&&(
          <div style={{background:"#FAFAFA",borderRadius:8,border:"1px solid #ECEFF1",marginBottom:14,overflow:"hidden"}}>
            <div style={{padding:"8px 8px 4px"}}><input value={emojiSearch} onChange={e=>setEmojiSearch(e.target.value)} placeholder="Search emoji..." style={{width:"100%",padding:"6px 10px",borderRadius:4,border:"1px solid #ECEFF1",fontSize:12,outline:"none",background:"#fff",fontFamily:"'Inter',sans-serif",boxSizing:"border-box"}}/></div>
            {!emojiSearch.trim()&&<div style={{display:"flex",gap:1,padding:"2px 6px",overflowX:"auto"}}>{emojiCats.map(cat=><button key={cat} onClick={()=>setEmojiCat(cat)} title={cat} style={{padding:"3px 5px",borderRadius:4,border:"none",background:emojiCat===cat?"#ECEFF1":"transparent",fontSize:14,cursor:"pointer",flexShrink:0,lineHeight:1}}>{EMOJI_CAT_ICONS[cat]}</button>)}</div>}
            <div style={{display:"flex",flexWrap:"wrap",gap:2,padding:"6px 8px 8px",maxHeight:160,overflowY:"auto"}}>
              {filteredEmojis.length===0&&<div style={{fontSize:11,color:"#90A4AE",padding:8,width:"100%",textAlign:"center"}}>No emoji found</div>}
              {filteredEmojis.map((em,i)=><button key={em+i} onClick={()=>{setSelEmoji(em);setShowEmojiPicker(false);setEmojiSearch("");}} style={{width:30,height:30,borderRadius:5,border:"none",background:"transparent",fontSize:18,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}} onMouseEnter={e=>e.currentTarget.style.background="#ECEFF1"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>{em}</button>)}
            </div>
          </div>
        )}
        <label style={{fontSize:11,fontWeight:700,color:"#90A4AE",textTransform:"uppercase",letterSpacing:"0.05em",display:"block",marginBottom:8}}>TITLE</label>
        <div style={{display:"flex",flexDirection:"column",gap:4,marginBottom:2}}>
          {TITLE_OPTIONS.map(t=>{const on=selTitle===t;return(
            <button key={t} onClick={()=>setSelTitle(on?"":t)} style={{padding:"8px 12px",borderRadius:8,border:"1px solid"+(on?"#2563EB":"#ECEFF1"),background:on?"#2563EB":"#FAFAFA",color:on?"#fff":"#546E7A",fontSize:12,fontWeight:on?700:500,cursor:"pointer",textAlign:"left",fontFamily:"'Inter',sans-serif",transition:"all 0.12s"}}
              onMouseEnter={e=>{if(!on){e.currentTarget.style.borderColor="#90A4AE";e.currentTarget.style.background="#F5F5F5";}}}
              onMouseLeave={e=>{if(!on){e.currentTarget.style.borderColor="#ECEFF1";e.currentTarget.style.background="#FAFAFA";}}}
            >{t}</button>
          );})}
        </div>
        {squads&&squads.length>0&&(<>
          <label style={{fontSize:11,fontWeight:700,color:"#90A4AE",textTransform:"uppercase",letterSpacing:"0.05em",display:"block",marginBottom:6,marginTop:14}}>SQUADS</label>
          <div style={{display:"flex",flexWrap:"wrap",gap:4,maxHeight:120,overflowY:"auto"}}>
            {squads.map(s=>{const on=selSquads.includes(s);return<button key={s} onClick={()=>setSelSquads(p=>p.includes(s)?p.filter(x=>x!==s):[...p,s])} style={{padding:"4px 10px",borderRadius:8,fontSize:11,fontWeight:600,cursor:"pointer",border:"1px solid"+(on?"#2563EB":"#ECEFF1"),background:on?"#2563EB":"#FAFAFA",color:on?"#fff":"#607D8B",fontFamily:"'Inter',sans-serif",transition:"all 0.15s"}}>{s}</button>;})}
          </div>
        </>)}
        <div style={{display:"flex",gap:8,marginTop:16}}>
          <button onClick={onClose} style={{flex:1,padding:"10px 20px",borderRadius:20,border:"1px solid #B0BEC5",background:"transparent",color:"#49454F",fontSize:12,fontWeight:500,cursor:"pointer",fontFamily:"'Inter',sans-serif"}}>Cancel</button>
          <button onClick={()=>onSave(selTitle,selSquads,selEmoji)} style={{flex:1,padding:"10px 20px",borderRadius:20,border:"none",background:"#2563EB",color:"#fff",fontSize:12,fontWeight:500,cursor:"pointer",fontFamily:"'Inter',sans-serif",boxShadow:"0 1px 2px rgba(0,0,0,0.3)"}}>Save</button>
        </div>
      </div>
    </>
  );
}

// ---------- NoteField ----------
function NoteField({value,onChange,placeholder}) {
  const [v,setV]=useState(value);
  const [focused,setFocused]=useState(false);
  useEffect(()=>setV(value),[value]);
  return (
    <div>
      <textarea value={v} onChange={e=>setV(e.target.value)}
        onFocus={()=>setFocused(true)}
        onBlur={()=>{if(v!==value)onChange(v);setFocused(false);}}
        placeholder={placeholder} rows={2}
        style={{width:"100%",boxSizing:"border-box",padding:"7px 10px",borderRadius:8,border:`1px solid ${focused?"#2563EB":"#E5E7EB"}`,fontSize:12,color:"#0F172A",fontFamily:"'Inter',sans-serif",resize:"none",outline:"none",lineHeight:1.5,background:"#FAFAFA"}}/>
    </div>
  );
}

// ---------- BlockerField ----------
function BlockerField({value,onChange}) {
  const [v,setV]=useState(value);
  useEffect(()=>setV(value),[value]);
  return (
    <div style={{display:"flex",alignItems:"center",gap:6}}>
      <MI name="flag" size={15} style={{color:v?"#EF5350":"#D1D5DB",flexShrink:0,cursor:"pointer"}} onClick={()=>{const nv=v?"":v;setV(nv);onChange(nv);}}/>
      <input value={v} onChange={e=>setV(e.target.value)}
        onBlur={()=>{if(v!==value)onChange(v);}}
        placeholder="Flag a blocker..."
        style={{flex:1,padding:"6px 10px",borderRadius:8,border:`1px solid ${v?"#FCA5A5":"#E5E7EB"}`,fontSize:12,color:v?"#C53030":"#6B7280",fontFamily:"'Inter',sans-serif",outline:"none",background:v?"#FFF5F5":"#FAFAFA"}}/>
    </div>
  );
}

// ---------- UrlField ----------
function UrlField({icon,label,value,onChange}) {
  const [v,setV]=useState(value||"");
  const [focused,setFocused]=useState(false);
  useEffect(()=>setV(value||""),[value]);
  return (
    <div style={{flex:1,display:"flex",alignItems:"center",gap:0,border:`1px solid ${focused?"#2563EB":"#E5E7EB"}`,borderRadius:8,background:"#FAFAFA",overflow:"hidden",boxShadow:focused?"0 0 0 3px rgba(37,99,235,0.1)":"none",transition:"box-shadow 0.15s"}}>
      <MI name={icon} size={13} style={{color:"#90A4AE",marginLeft:8,flexShrink:0}}/>
      <input value={v} onChange={e=>setV(e.target.value)}
        onFocus={()=>setFocused(true)}
        onBlur={()=>{setFocused(false);if(v!==value)onChange(v);}}
        placeholder={label+" link..."}
        style={{flex:1,border:"none",background:"transparent",fontSize:11,color:"#0F172A",padding:"6px 6px",outline:"none",fontFamily:"'Inter',sans-serif",minWidth:0}}/>
      {v&&<a href={v} target="_blank" rel="noopener noreferrer" onClick={e=>e.stopPropagation()}
        style={{color:"#2563EB",display:"flex",padding:"4px 6px",flexShrink:0,textDecoration:"none"}} title={"Open "+label+" in new tab"}>
        <MI name="open_in_new" size={12}/>
      </a>}
    </div>
  );
}

// ---------- Popover ----------
function Popover({proj,color,onClose,onRemove,onEditProject,onOpenEdit}) {
  const ref=useRef(null);
  const [showCalc,setShowCalc]=useState(false);
  useEffect(()=>{
    const esc=e=>{if(e.key==="Escape"){if(showCalc)setShowCalc(false);else onClose();}};
    document.addEventListener("keydown",esc); return ()=>document.removeEventListener("keydown",esc);
  },[onClose,showCalc]);
  const st=autoStatus(proj),sc=STATUS[st];
  const isLeave=isLeaveItem(proj);
  function editAndClose(id,fields){if(onEditProject){onEditProject(id,fields);onClose();}}
  return (
    <>
    <div onClick={onClose} style={{position:"fixed",inset:0,zIndex:10000,background:"rgba(0,0,0,0.25)"}}/>
    <div ref={ref} style={{position:"fixed",zIndex:10001,top:"50%",left:"50%",transform:"translate(-50%,-50%)",background:"#fff",border:"1px solid #ECEFF1",borderRadius:28,boxShadow:"0 4px 24px rgba(0,0,0,0.12)",padding:18,minWidth:280,maxWidth:340}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
        <div style={{fontSize:14,fontWeight:700,color:"#0F172A",lineHeight:1.3,flex:1,paddingRight:8}}>{isLeave?(proj.leaveType||proj.name):proj.name}</div>
        <button onClick={onClose} style={{background:"transparent",border:"none",cursor:"pointer",color:"#90A4AE",display:"flex",borderRadius:"50%",padding:4,flexShrink:0}}><MI name="close" size={16} /></button>
      </div>
      {isLeave?(
        <>{[["TYPE",<div style={{display:"flex",gap:6}}>{LEAVE_TYPES.map(lt=>{const active=proj.leaveType===lt;const c=LEAVE_COLORS[lt];return<button key={lt} onClick={onEditProject?()=>onEditProject(proj.id,{leaveType:lt,name:lt}):undefined} style={{fontSize:11,fontWeight:800,padding:"2px 7px",borderRadius:4,background:active?c+"18":"#FAFAFA",color:active?c:"#90A4AE",border:"1px solid "+(active?c+"33":"#ECEFF1"),cursor:onEditProject?"pointer":"default"}}>{lt}</button>;})}</div>],["STATUS",<div style={{display:"flex",gap:6}}>{LEAVE_STATUSES.map(s=>{const active=(proj.leaveStatus||"Planned")===s;const sc=s==="Planned"?"#3949AB":s==="In Progress"?"#1565C0":"#2E7D32";return<button key={s} onClick={onEditProject?()=>onEditProject(proj.id,{leaveStatus:s}):undefined} style={{fontSize:11,fontWeight:800,padding:"2px 7px",borderRadius:4,background:active?sc+"18":"#FAFAFA",color:active?sc:"#90A4AE",border:"1px solid "+(active?sc+"33":"#ECEFF1"),cursor:onEditProject?"pointer":"default"}}>{s}</button>;})}</div>],["DATES",<span style={{fontSize:11,color:"#546E7A",fontFamily:"monospace"}}>{fmtM(proj.startDate)} → {fmtM(proj.endDate)}</span>]].map(([label,val])=><div key={label} style={{display:"flex",gap:6,alignItems:"flex-start",marginBottom:6}}><span style={{fontSize:11,fontWeight:700,color:"#90A4AE",width:44,flexShrink:0,paddingTop:2}}>{label}</span>{val}</div>)}</>
      ):(
        <>{[["SIZE",<div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}><div style={{display:"flex",gap:6}}>{SIZES.map(s=><button key={s} onClick={onEditProject?()=>onEditProject(proj.id,{size:s}):undefined} style={{fontSize:11,fontWeight:800,padding:"2px 7px",borderRadius:8,background:proj.size===s?color+"18":"#FAFAFA",color:proj.size===s?color:"#90A4AE",border:"1px solid "+(proj.size===s?color+"33":"#ECEFF1"),cursor:onEditProject?"pointer":"default",fontFamily:"monospace"}}>{s}</button>)}</div>{onEditProject&&<button onClick={()=>setShowCalc(true)} style={{fontSize:11,fontWeight:700,color:"#2563EB",background:"none",border:"none",cursor:"pointer",fontFamily:"'Inter',sans-serif",padding:0}}>{proj.calcScores?"Recalculate":"Calculate"}</button>}</div>],
          ...(proj.calcScores?[["CALC",<div style={{display:"flex",flexDirection:"column",gap:2}}>{SIZE_FACTORS.map(({key,label,options})=>{const idx=(proj.calcScores[key]||1)-1;return<span key={key} style={{fontSize:12,color:"#607D8B"}}><span style={{fontWeight:700,color:"#90A4AE"}}>{label.split(" ")[0]}:</span> {options[idx]||"--"}</span>;})}</div>]]:[]),
          ["STATUS",<div style={{display:"inline-flex",alignItems:"center",position:"relative"}}><select value={st} onChange={onEditProject?e=>onEditProject(proj.id,{status:e.target.value}):undefined} disabled={!onEditProject} style={{fontSize:11,fontWeight:700,padding:"3px 24px 3px 6px",borderRadius:6,border:"1px solid "+sc.color+"33",background:sc.bg,color:sc.color,cursor:onEditProject?"pointer":"default",fontFamily:"'Inter',sans-serif",appearance:"none",opacity:1}}>{ALL_STATUSES.map(s=><option key={s} value={s}>{s}</option>)}</select><MI name="expand_more" size={12} style={{position:"absolute",right:4,pointerEvents:"none",color:sc.color}} /></div>],
          ["TYPE",(()=>{const pt=PROJECT_TYPES.find(t=>t.label===proj.type);return onEditProject?(<div style={{display:"inline-flex",alignItems:"center",position:"relative"}}><select value={proj.type||""} onChange={e=>onEditProject(proj.id,{type:e.target.value})} style={{fontSize:11,fontWeight:600,padding:"3px 24px 3px 8px",borderRadius:20,border:`1px solid ${pt?pt.color+"33":"#ECEFF1"}`,background:pt?pt.bg:"#FAFAFA",color:pt?pt.color:"#90A4AE",cursor:"pointer",fontFamily:"'Inter',sans-serif",appearance:"none",opacity:1}}><option value="">No type</option>{PROJECT_TYPES.map(t=><option key={t.label} value={t.label}>{t.label}</option>)}</select><MI name="expand_more" size={12} style={{position:"absolute",right:6,pointerEvents:"none",color:pt?pt.color:"#90A4AE"}}/></div>):pt?(<span style={{fontSize:11,fontWeight:600,padding:"3px 10px",borderRadius:20,background:pt.bg,color:pt.color,border:`1px solid ${pt.color}33`}}>{pt.label}</span>):(<span style={{fontSize:12,color:"#90A4AE"}}>--</span>);})()],
          ["SQUAD",<span style={{fontSize:12,color:"#0F172A"}}>{proj.squad||"--"}</span>],["PM",<span style={{fontSize:12,color:"#0F172A",lineHeight:1.5}}>{proj.pms?.join(", ")||"--"}</span>],["DATES",<span style={{fontSize:11,color:"#546E7A",fontFamily:"monospace"}}>{fmtM(proj.startDate)} → {fmtM(proj.endDate)}</span>]].map(([label,val])=><div key={label} style={{display:"flex",gap:6,alignItems:"flex-start",marginBottom:6}}><span style={{fontSize:11,fontWeight:700,color:"#90A4AE",width:44,flexShrink:0,paddingTop:2}}>{label}</span>{val}</div>)}</>
      )}
      {/* Notes & Links section */}
      {(proj.figmaUrl||proj.jiraUrl||proj.note||proj.blocker||onEditProject)&&(
        <div style={{marginTop:10,paddingTop:10,borderTop:"1px solid #E0E0E0",display:"flex",flexDirection:"column",gap:8}}>
          {/* Figma + Jira links */}
          {onEditProject?(
            <div style={{display:"flex",gap:8}}>
              <UrlField icon="brush" label="Figma" value={proj.figmaUrl||""} onChange={v=>onEditProject(proj.id,{figmaUrl:v})}/>
              <UrlField icon="link" label="Jira" value={proj.jiraUrl||""} onChange={v=>onEditProject(proj.id,{jiraUrl:v})}/>
            </div>
          ):(proj.figmaUrl||proj.jiraUrl)&&(
            <div style={{display:"flex",gap:8}}>
              {proj.figmaUrl&&(
                <a href={proj.figmaUrl} target="_blank" rel="noopener noreferrer"
                  style={{display:"flex",alignItems:"center",gap:4,fontSize:12,fontWeight:500,color:"#2563EB",textDecoration:"none",padding:"4px 10px",borderRadius:20,border:"1px solid #C5D8FD",background:"#EEF4FD"}}>
                  <MI name="brush" size={13}/>Figma
                </a>
              )}
              {proj.jiraUrl&&(
                <a href={proj.jiraUrl} target="_blank" rel="noopener noreferrer"
                  style={{display:"flex",alignItems:"center",gap:4,fontSize:12,fontWeight:500,color:"#2563EB",textDecoration:"none",padding:"4px 10px",borderRadius:20,border:"1px solid #C5D8FD",background:"#EEF4FD"}}>
                  <MI name="link" size={13}/>Jira
                </a>
              )}
            </div>
          )}
          {/* Note */}
          {onEditProject?(
            <NoteField value={proj.note||""} onChange={v=>onEditProject(proj.id,{note:v})} placeholder="Add note or recent progress..."/>
          ):(proj.note&&(
            <div style={{fontSize:12,color:"#546E7A",lineHeight:1.5,background:"#F8FAFC",borderRadius:8,padding:"8px 10px"}}>{proj.note}</div>
          ))}
          {/* Blocker */}
          {onEditProject?(
            <BlockerField value={proj.blocker||""} onChange={v=>onEditProject(proj.id,{blocker:v})}/>
          ):(proj.blocker&&(
            <div style={{display:"flex",alignItems:"flex-start",gap:6,padding:"8px 10px",borderRadius:8,background:"#FFF5F5",border:"1px solid #FED7D7"}}>
              <MI name="flag" size={14} style={{color:"#EF5350",flexShrink:0,marginTop:1}}/>
              <span style={{fontSize:12,color:"#C53030",lineHeight:1.5}}>{proj.blocker}</span>
            </div>
          ))}
        </div>
      )}
      <div style={{marginTop:12,paddingTop:12,borderTop:"1px solid #E0E0E0",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        {onRemove
          ? <button onClick={()=>{onRemove(proj.id);onClose();}} style={{color:"#EF5350",background:"transparent",border:"none",cursor:"pointer",display:"flex",padding:6,borderRadius:"50%"}}><MI name="delete" size={16}/></button>
          : <div/>}
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          {onOpenEdit&&<button onClick={()=>{onOpenEdit(proj);onClose();}} style={{display:"flex",alignItems:"center",gap:5,height:36,padding:"0 16px",borderRadius:20,border:"none",background:"#E8EAF6",color:"#3949AB",fontSize:12,fontWeight:500,cursor:"pointer",fontFamily:"'Inter',sans-serif"}}><MI name="edit" size={13}/>Edit</button>}
          <button onClick={onClose} style={{display:"flex",alignItems:"center",gap:5,height:36,padding:"0 16px",borderRadius:20,border:"none",background:"#2563EB",color:"#fff",fontSize:12,fontWeight:500,cursor:"pointer",fontFamily:"'Inter',sans-serif",boxShadow:"0 1px 2px rgba(0,0,0,0.2)"}}><MI name="check" size={13}/>Confirm</button>
        </div>
      </div>
    </div>
    {showCalc&&(
      <div style={{position:"fixed",inset:0,zIndex:10002,display:"flex",alignItems:"center",justifyContent:"center",background:"rgba(0,0,0,0.35)"}} onClick={()=>setShowCalc(false)}>
        <div onClick={e=>e.stopPropagation()} style={{width:420}}>
          <div style={{background:"#fff",borderRadius:28,boxShadow:"0 20px 60px rgba(0,0,0,0.2)",overflow:"hidden"}}>
            <div style={{padding:"16px 20px",borderBottom:"1px solid #F5F5F5",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div><div style={{fontSize:14,fontWeight:700,color:"#0F172A",fontFamily:"'Inter',sans-serif"}}>Size Calculator</div><div style={{fontSize:11,color:"#90A4AE",marginTop:2}}>{proj.name}</div></div>
              <button onClick={()=>setShowCalc(false)} style={{background:"transparent",border:"none",cursor:"pointer",color:"#90A4AE",display:"flex",borderRadius:"50%",padding:4}}><MI name="close" size={16} /></button>
            </div>
            <div style={{padding:20}}><SizeCalculator bare initialScores={proj.calcScores} onResult={(s,sc)=>{onEditProject(proj.id,{size:s,calcScores:sc});setShowCalc(false);}} onClose={()=>setShowCalc(false)}/></div>
          </div>
        </div>
      </div>
    )}
    </>
  );
}

// ---------- Block ----------
function Block({proj,color,left,top,width,onUpdate,onRemove,onEditProject,onToast,isOpen,onOpen}) {
  const isLeave=isLeaveItem(proj);
  const barColor=isLeave?(LEAVE_COLORS[proj.leaveType]||"#90A4AE"):color;
  const [hover,setHover]=useState(false);
  const [mode,setMode]=useState(null);
  const [dw,setDw]=useState(0);
  const [cursorEdge,setCursorEdge]=useState(false);
  const drag=useRef(null);
  const EDGE=10; // px zone for resize handles at each edge
  const active=isActive(proj);
  const leaveCompleted=isLeave&&proj.leaveStatus==="Completed";
  let opacity=0.3;if(hover)opacity=0.75;if(active||isLeave)opacity=1;if(mode)opacity=0.85;if(leaveCompleted)opacity=0.5;
  const cL=(mode==="move"||mode==="resizeL")?left+dw*CELL_W:left;
  let cW=width;if(mode==="resizeL")cW=Math.max(CELL_W,width-dw*CELL_W);if(mode==="resizeR")cW=Math.max(CELL_W,width+dw*CELL_W);

  function getMode(e){
    const r=e.currentTarget.getBoundingClientRect();
    const x=e.clientX-r.left;
    if(x<=EDGE)return"resizeL";
    if(x>=r.width-EDGE)return"resizeR";
    return"move";
  }
  function startDrag(e){
    if(e.button!==0)return;
    if(!onUpdate){// read-only: treat as click to open popover
      onOpen(isOpen?null:{proj,color});return;
    }
    e.preventDefault();e.stopPropagation();
    const m=getMode(e);
    e.currentTarget.setPointerCapture(e.pointerId);
    drag.current={x:e.clientX,moved:false,mode:m};
    setMode(m);setDw(0);
  }
  function onMove(e){
    if(!onUpdate)return;
    if(!drag.current){
      // update cursor based on position
      const r=e.currentTarget.getBoundingClientRect();
      const x=e.clientX-r.left;
      setCursorEdge(x<=EDGE||x>=r.width-EDGE);
      return;
    }
    const dx=e.clientX-drag.current.x;
    if(Math.abs(dx)>4)drag.current.moved=true;
    setDw(Math.round(dx/CELL_W));
  }
  function onUp(e){
    if(!drag.current)return;
    e.currentTarget.releasePointerCapture(e.pointerId);
    const{moved,mode:m}=drag.current,w=dw;
    drag.current=null;setMode(null);setDw(0);
    if(moved&&w!==0&&onUpdate){
      if(m==="move"){onUpdate({...proj,startDate:addW(proj.startDate,w),endDate:addW(proj.endDate,w)});onToast&&onToast(`"${proj.name}" moved`);}
      if(m==="resizeL"){onUpdate({...proj,startDate:addW(proj.startDate,w)});onToast&&onToast(`"${proj.name}" start → ${fmtS(addW(proj.startDate,w))}`);}
      if(m==="resizeR"){onUpdate({...proj,endDate:addW(proj.endDate,w)});onToast&&onToast(`"${proj.name}" end → ${fmtS(addW(proj.endDate,w))}`);}
    } else if(!moved&&m==="move") onOpen(isOpen?null:{proj,color});
  }
  const tip=mode&&dw!==0?(mode==="move"?`${fmtS(addW(proj.startDate,dw))} → ${fmtS(addW(proj.endDate,dw))}`:mode==="resizeL"?`Start ${fmtS(addW(proj.startDate,dw))}`:`End ${fmtS(addW(proj.endDate,dw))}`):null;
  const cursor=!onUpdate?"pointer":mode?(mode==="move"?"grabbing":"ew-resize"):(cursorEdge?"ew-resize":"grab");
  return (
    <div style={{position:"absolute",left:cL+1,top,width:cW,height:24,zIndex:mode?9999:isOpen?999:1}}>
      <div
        onPointerDown={startDrag}
        onPointerMove={onMove}
        onPointerUp={onUp}
        onMouseEnter={()=>setHover(true)}
        onMouseLeave={()=>{setHover(false);setCursorEdge(false);}}
        style={{width:"100%",height:"100%",borderRadius:4,background:isLeave?`repeating-linear-gradient(135deg,${barColor}DD,${barColor}DD 6px,${barColor}99 6px,${barColor}99 10px)`:barColor+"DD",display:"flex",alignItems:"center",paddingLeft:10,paddingRight:10,overflow:"hidden",cursor,opacity,boxShadow:mode?`0 4px 16px ${barColor}66`:isOpen?`0 0 0 2px ${barColor}`:`0 1px 3px ${barColor}44`,transition:"opacity 0.2s,box-shadow 0.15s",userSelect:"none",position:"relative"}}
      >
        <div style={{position:"absolute",left:0,top:"50%",transform:"translateY(-50%)",width:3,height:12,background:"#ffffff33",borderRadius:2,marginLeft:4}}/>
        {!isLeave&&autoStatus(proj)==="Done"&&<span style={{fontSize:9,marginRight:4,color:"#fff",flexShrink:0}}><MI name="check" size={13} style={{color:"#fff"}} /></span>}
        {leaveCompleted&&<span style={{fontSize:9,marginRight:4,color:"#fff",flexShrink:0}}><MI name="check" size={13} style={{color:"#fff"}} /></span>}
        {proj.blocker&&<MI name="flag" size={10} style={{color:"#EF5350",flexShrink:0,opacity:0.9}}/>}
        <span style={{fontSize:11,color:"#fff",fontWeight:700,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",textShadow:isLeave?"0 1px 2px rgba(0,0,0,0.4)":"none",flex:1}}>{isLeave?(proj.leaveType||proj.name):proj.name}</span>
        {!isLeave&&<span style={{fontSize:9,color:"#ffffff99",marginLeft:4,flexShrink:0}}>{proj.size}</span>}
        <div style={{position:"absolute",right:0,top:"50%",transform:"translateY(-50%)",width:3,height:12,background:"#ffffff33",borderRadius:2,marginRight:4}}/>
      </div>
      {tip&&<div style={{position:"absolute",bottom:"calc(100% + 4px)",left:"50%",transform:"translateX(-50%)",background:"#1C1B1F",color:"#fff",fontSize:11,fontWeight:700,padding:"3px 8px",borderRadius:4,whiteSpace:"nowrap",pointerEvents:"none",zIndex:10000}}>{tip}</div>}
    </div>
  );
}

// ---------- SizeCalculator ----------
const SIZE_FACTORS = [
  {key:"screens",label:"Screens / Flows",options:["1-3 screens","4-8 screens","9+ screens"]},
  {key:"novelty",label:"Net New vs Iteration",options:["Iteration","Mix of new & existing","Greenfield"]},
  {key:"research",label:"Research Needed",options:["None","Lightweight","Generative / Evaluative"]},
  {key:"stakeholders",label:"Stakeholder Complexity",options:["1 squad","2-3 squads","Cross-org"]},
  {key:"system",label:"Design System Impact",options:["Existing patterns","Minor new patterns","Major new work"]},
];
function scoreToSize(score){if(score<=6)return "XS";if(score<=8)return "S";if(score<=11)return "M";if(score<=13)return "L";return "XL";}
function SizeCalculator({onResult,onClose:onCalcClose,bare,initialScores}) {
  const [scores,setScores]=useState(initialScores||{screens:0,novelty:0,research:0,stakeholders:0,system:0});
  const total=Object.values(scores).reduce((a,b)=>a+b,0);
  const allScored=Object.values(scores).every(v=>v>0);
  const result=allScored?scoreToSize(total):null;
  return (
    <div style={{border:bare?"none":"1px solid #ECEFF1",borderRadius:bare?0:10,overflow:"hidden",background:"#fff"}}>
      {!bare&&<div style={{padding:"12px 14px",background:"#FAFAFA",borderBottom:"1px solid #ECEFF1",display:"flex",justifyContent:"space-between",alignItems:"center"}}><div style={{display:"flex",alignItems:"center",gap:8}}><MI name="bar_chart" size={14} style={{color:"#546E7A"}} /><span style={{fontSize:12,fontWeight:700,color:"#0F172A",fontFamily:"'Inter',sans-serif"}}>Size Calculator</span></div><button onClick={onCalcClose} style={{background:"transparent",border:"none",cursor:"pointer",color:"#90A4AE",display:"flex",borderRadius:"50%",padding:4}}><MI name="close" size={16} /></button></div>}
      <div style={{padding:"12px 14px"}}>
        {SIZE_FACTORS.map(({key,label,options})=>(
          <div key={key} style={{marginBottom:10}}>
            <div style={{fontSize:11,fontWeight:700,color:"#90A4AE",letterSpacing:"0.05em",marginBottom:5}}>{label}</div>
            <div style={{display:"flex",gap:4}}>{options.map((opt,i)=>{const val=i+1;const sel=scores[key]===val;return<button key={i} onClick={()=>setScores(s=>({...s,[key]:val}))} style={{flex:1,padding:"6px 4px",borderRadius:6,border:"1px solid"+(sel?"#2563EB":"#ECEFF1"),background:sel?"#2563EB":"#FAFAFA",color:sel?"#fff":"#607D8B",fontSize:11,fontWeight:sel?700:500,cursor:"pointer",fontFamily:"'Inter',sans-serif",lineHeight:1.3,textAlign:"center"}}>{opt}</button>;})}</div>
          </div>
        ))}
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginTop:12,paddingTop:12,borderTop:"1px solid #F5F5F5"}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}>
            <span style={{fontSize:11,color:"#90A4AE"}}>Score: <strong style={{color:"#0F172A"}}>{allScored?total:"--"}</strong> / 15</span>
            {result&&<span style={{fontSize:14,fontWeight:700,padding:"3px 12px",borderRadius:6,background:"#2563EB",color:"#fff",fontFamily:"monospace"}}>{result}</span>}
          </div>
          {result&&<button onClick={()=>{onResult(result,scores);onCalcClose();}} style={{fontSize:12,fontWeight:700,color:"#fff",background:"#2563EB",border:"none",borderRadius:20,padding:"6px 14px",cursor:"pointer",fontFamily:"'Inter',sans-serif"}}>Use {result}</button>}
        </div>
      </div>
    </div>
  );
}

// ---------- Settings ----------
function Settings({settings,onSave,onClose,hideSquadsPMs=false,stateSquads=[]}) {
  const [designers,setDesigners]=useState([...(settings.designers||[])]);
  // Merge string names from teams with actual state squad names (union, deduped)
  const initSquads=()=>[...new Set([...(settings.squads||[]),...stateSquads.map(sq=>sq.name)])];
  const [squads,setSquads]=useState(initSquads);
  const [pms,setPms]=useState([...(settings.pms||[])]);
  const [removedSquadIds,setRemovedSquadIds]=useState([]);
  // squad→[pms] map; seed from saved squadPmMap or fall back to SQUAD_PM_MAP
  const initSqPm = ()=>{
    const base={...SQUAD_PM_MAP,...(settings.squadPmMap||{})};
    return base;
  };
  const [sqPmMap,setSqPmMap]=useState(initSqPm);
  const [tab,setTab]=useState("designers");
  const [val,setVal]=useState("");
  const cur=tab==="designers"?designers:(tab==="squads"?squads:pms);
  const setCur=tab==="designers"?setDesigners:(tab==="squads"?setSquads:setPms);
  function add(){const v=val.trim();if(!v||cur.includes(v))return;setCur([...cur,v]);setVal("");}
  function rem(item){
    setCur(cur.filter(i=>i!==item));
    if(tab==="squads"){
      // Also track the state squad id so we can delete from state on save
      const stateSquad=stateSquads.find(sq=>sq.name===item);
      if(stateSquad) setRemovedSquadIds(prev=>[...prev,stateSquad.id]);
    }
  }
  // Count designers in a state squad (for info display)
  function squadDesignerCount(name){return(stateSquads.find(sq=>sq.name===name)?.designers||[]).length;}
  // When PM assignment changes for a squad
  function setSquadPms(sq,selectedPms){setSqPmMap(m=>({...m,[sq]:selectedPms}));}
  // When squad assignment changes for a PM (update reverse mapping)
  function setPmSquad(pm,sq){
    setSqPmMap(m=>{
      const next={...m};
      // Remove pm from all squads first
      Object.keys(next).forEach(s=>{next[s]=(next[s]||[]).filter(p=>p!==pm);});
      // Add to selected squad
      if(sq){next[sq]=[...(next[sq]||[]).filter(p=>p!==pm),pm];}
      return next;
    });
  }
  // Get which squad a PM is currently assigned to
  function getPmSquad(pm){return Object.entries(sqPmMap).find(([,pmsArr])=>pmsArr.includes(pm))?.[0]||"";}
  const inp={flex:1,padding:"0 12px",height:40,boxSizing:"border-box",borderRadius:4,border:"1px solid #ECEFF1",fontSize:14,color:"#0F172A",outline:"none",background:"#FAFAFA",fontFamily:"'Inter',sans-serif"};
  function buildFinalMap(){
    // sqPmMap already has the truth; also ensure PM-level changes are reflected
    return sqPmMap;
  }
  return (
    <div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.4)",zIndex:10000,display:"flex",alignItems:"flex-start",justifyContent:"flex-end"}}>
      <div onClick={e=>e.stopPropagation()} style={{background:"#fff",width:440,height:"100vh",display:"flex",flexDirection:"column",boxShadow:"-2px 0 16px rgba(0,0,0,0.10)",borderRadius:"16px 0 0 16px"}}>
        <div style={{padding:"24px 24px 16px",borderBottom:"1px solid #F5F5F5",display:"flex",justifyContent:"space-between",flexShrink:0}}>
          <div><div style={{fontSize:18,fontWeight:700,color:"#0F172A",fontFamily:"'Inter',sans-serif",marginBottom:4}}>Settings</div><div style={{fontSize:12,color:"#90A4AE"}}>Manage your team</div></div>
          <button onClick={onClose} style={{background:"none",border:"none",fontSize:18,cursor:"pointer",color:"#90A4AE",display:"flex"}}><MI name="close" size={16} /></button>
        </div>
        {!hideSquadsPMs&&<div style={{display:"flex",borderBottom:"1px solid #F5F5F5",flexShrink:0}}>{["designers","squads","pms"].map(t=><button key={t} onClick={()=>{setTab(t);setVal("");}} style={{flex:1,padding:"12px 0",border:"none",background:"transparent",cursor:"pointer",fontSize:14,fontWeight:tab===t?700:500,color:tab===t?"#2563EB":"#6B7280",borderBottom:tab===t?"3px solid #2563EB":"2px solid transparent",marginBottom:-1,fontFamily:"'Inter',sans-serif"}}>{t==="pms"?"PMs":t==="designers"?"Designers":"Squads"}</button>)}</div>}
        <div style={{padding:24,flex:1,overflowY:"auto"}}>
          <div style={{display:"flex",gap:8,marginBottom:16}}>
            <input value={val} onChange={e=>setVal(e.target.value)} onKeyDown={e=>e.key==="Enter"&&add()} placeholder={tab==="designers"?"Designer name...":tab==="squads"?"Squad name...":"PM name..."} style={inp}/>
            <button onClick={add} style={{height:40,padding:"0 20px",borderRadius:20,border:"none",background:"#2563EB",color:"#fff",fontSize:14,fontWeight:500,cursor:"pointer",fontFamily:"'Inter',sans-serif",display:"flex",alignItems:"center",gap:6,boxShadow:"0 1px 2px rgba(0,0,0,0.3)"}}><MI name="add" size={14} /> Add</button>
          </div>
          {cur.length===0&&<div style={{fontSize:14,color:"#90A4AE",textAlign:"center",padding:"24px 0"}}>None yet</div>}
          {cur.map((item,i)=>(
            <div key={i} style={{padding:"10px 14px",background:"#FAFAFA",borderRadius:8,border:"1px solid #ECEFF1",marginBottom:6}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:(tab==="squads"||tab==="pms")?6:0}}>
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <span style={{fontSize:14,fontWeight:600,color:"#0F172A"}}>{item}</span>
                  {tab==="squads"&&squadDesignerCount(item)>0&&<span style={{fontSize:11,fontWeight:600,color:"#607D8B",background:"#ECEFF1",borderRadius:4,padding:"1px 6px"}}>{squadDesignerCount(item)} designer{squadDesignerCount(item)!==1?"s":""}</span>}
                </div>
                <button onClick={()=>rem(item)} style={{background:"none",border:"none",cursor:"pointer",color:"#B0BEC5",fontSize:16,lineHeight:1,display:"flex"}} onMouseEnter={e=>e.currentTarget.style.color="#EF5350"} onMouseLeave={e=>e.currentTarget.style.color="#B0BEC5"}><MI name="close" size={16} /></button>
              </div>
              {tab==="squads"&&(
                <div>
                  <div style={{fontSize:11,fontWeight:700,color:"#90A4AE",letterSpacing:"0.07em",marginBottom:4}}>ASSIGNED PMs</div>
                  <MultiSelect options={pms.length?pms:DEF_PMS} selected={sqPmMap[item]||[]} onChange={v=>setSquadPms(item,v)} placeholder="Assign PMs to this squad..."/>
                </div>
              )}
              {tab==="pms"&&(
                <div>
                  <div style={{fontSize:11,fontWeight:700,color:"#90A4AE",letterSpacing:"0.07em",marginBottom:4}}>ASSIGNED SQUAD</div>
                  <CustomSelect value={getPmSquad(item)} onChange={v=>setPmSquad(item,v)} options={squads.length?squads:DEF_SQUADS} placeholder="Assign to squad..."/>
                </div>
              )}
            </div>
          ))}
        </div>
        <div style={{padding:24,borderTop:"1px solid #F5F5F5",display:"flex",gap:10,flexShrink:0}}>
          <button onClick={onClose} style={{flex:1,height:40,padding:"0 20px",borderRadius:20,border:"none",background:"#F1F5FF",color:"#3B5BDB",fontSize:14,fontWeight:500,cursor:"pointer",fontFamily:"'Inter',sans-serif"}}>Cancel</button>
          <button onClick={()=>{onSave({designers,squads,pms,squadPmMap:buildFinalMap(),removedSquadIds});onClose();}} style={{flex:1,height:40,padding:"0 20px",borderRadius:20,border:"none",background:"#2563EB",color:"#fff",fontSize:14,fontWeight:500,cursor:"pointer",fontFamily:"'Inter',sans-serif",boxShadow:"0 1px 2px rgba(0,0,0,0.3)"}}>Save Changes</button>
        </div>
      </div>
    </div>
  );
}

// ---------- AddModal ----------
function AddModal({team,onAdd,onAddBatch,onClose,role,lockedDesigner,prefill}) {
  const theRole=role||"Designer";
  const prefillDesigner=lockedDesigner||(prefill&&prefill.designer)||"";
  const [mode,setMode]=useState(prefill?"manual":null);
  const [step,setStep]=useState(1);
  const [pasteVal,setPasteVal]=useState("");
  const [pasteErr,setPasteErr]=useState("");
  const [copied,setCopied]=useState(false);
  const [boardItems,setBoardItems]=useState([]);
  const [f,setF]=useState({
    id:         (prefill&&prefill.id)||"",
    name:       (prefill&&prefill.name)||"",
    squad:      (prefill&&prefill.squad)||"",
    designer:   prefillDesigner,
    pms:        (prefill&&prefill.pms)||[],
    startDate:  (prefill&&prefill.startDate)||"",
    endDate:    (prefill&&prefill.endDate)||"",
    size:       (prefill&&prefill.size)||"M",
    status:     (prefill&&prefill.status)||"",
    calcScores: (prefill&&prefill.calcScores)||null,
    figmaUrl:   (prefill&&prefill.figmaUrl)||"",
    jiraUrl:    (prefill&&prefill.jiraUrl)||"",
    note:       (prefill&&prefill.note)||"",
    blocker:    (prefill&&prefill.blocker)||"",
    type:       (prefill&&prefill.type)||"",
  });
  const [leaveF,setLeaveF]=useState({designer:prefillDesigner,leaveType:"PTO",leaveStatus:"Planned",startDate:(prefill&&prefill.startDate)||"",endDate:""});
  const [leaveErr,setLeaveErr]=useState("");
  const [err,setErr]=useState("");
  const [importUrl,setImportUrl]=useState("");
  const [showCalc,setShowCalc]=useState(false);
  const xlsxRef=useRef(null);
  const [xlsxLoading,setXlsxLoading]=useState(false);
  const [xlsxErr,setXlsxErr]=useState("");
  const inp={padding:"0 11px",height:40,borderRadius:4,border:"1px solid #ECEFF1",background:"#FAFAFA",fontSize:14,color:"#0F172A",fontFamily:"'Inter',sans-serif",width:"100%",outline:"none",boxSizing:"border-box"};
  const lbl={fontSize:11,fontWeight:700,color:"#90A4AE",letterSpacing:"0.08em",marginBottom:5,display:"block"};
  const stepBadge=(n,active)=>({display:"inline-flex",alignItems:"center",justifyContent:"center",width:22,height:22,borderRadius:"50%",fontSize:11,fontWeight:700,flexShrink:0,background:active?"#2563EB":"#ECEFF1",color:active?"#fff":"#90A4AE",fontFamily:"'Inter',sans-serif"});
  function setField(k,v){setF(p=>({...p,[k]:v}));}
  function toggleItem(idx){setBoardItems(b=>b.map((it,i)=>i===idx?{...it,selected:!it.selected}:it));}
  function updateItem(idx,k,v){setBoardItems(b=>b.map((it,i)=>i===idx?{...it,[k]:v}:it));}
  function toggleAll(){const allSel=boardItems.every(i=>i.selected);setBoardItems(b=>b.map(i=>({...i,selected:!allSel})));}
  function copyPrompt(text){try{navigator.clipboard.writeText(text).then(()=>{setCopied(true);setTimeout(()=>setCopied(false),2000);});}catch(e){}}
  const boardPrompt="Connect to Atlassian and search for all active (non-Done) Jira issues in my project board. For each issue return: key (ticket ID), name (summary), assignee (display name), status.\n\nReturn ONLY a valid JSON array in a code block, nothing else:\n```json\n[{\"key\":\"PROJ-123\", \"name\":\"Issue title\", \"assignee\":\"Person Name\", \"status\":\"In Progress\"}]\n```\n\nHere is my board link: ";
  const ticketPrompt="Connect to Atlassian and fetch this Jira ticket. Extract the ticket title, assignee name, and estimate the t-shirt size (XS/S/M/L/XL, default M).\n\nReturn ONLY a valid JSON object in a code block, nothing else:\n```json\n{\"projectName\":\"Ticket title\", \"designer\":\"Assignee Name\", \"size\":\"M\"}\n```\n\nHere is the ticket link: ";
  const sheetsPrompt="Read the Google Sheet at the link below. Extract projects with: name, designer, squad, pm, status, startDate (YYYY-MM-DD), endDate (YYYY-MM-DD), size (XS/S/M/L/XL).\n\nReturn ONLY a valid JSON array in a code block:\n```json\n[{\"name\":\"Project\", \"designer\":\"Name\", \"squad\":\"TM\", \"pm\":\"PM\", \"status\":\"In Progress\", \"startDate\":\"2026-03-01\", \"endDate\":\"2026-04-01\", \"size\":\"M\"}]\n```\n\nHere is the sheet link: ";
  function parseBoardJson(){
    setPasteErr("");
    try {
      let text=pasteVal.replace(/```json|```/g,"").trim();
      let items=null;
      const a1=text.indexOf("["),a2=text.lastIndexOf("]");
      if(a1!==-1&&a2>a1){try{const arr=JSON.parse(text.slice(a1,a2+1));if(Array.isArray(arr)&&arr.length)items=arr;}catch(e){}}
      if(!items||!items.length){setPasteErr("Could not find a JSON array.");return;}
      const sizeMap={"extra small":"XS","xs":"XS","small":"S","s":"S","medium":"M","m":"M","large":"L","l":"L","extra large":"XL","xl":"XL"};
      setBoardItems(items.map(i=>{
        const assignee=i.assignee||i.designer||"";
        const dm=team.designers.find(d=>assignee.toLowerCase().includes(d.toLowerCase()))||"";
        const pmRaw=i.pm||"";
        const pmMatch=(team.pms||[]).find(p=>pmRaw.toLowerCase().includes(p.split(" ")[0].toLowerCase()))||"";
        const rawSize=(i.size||i.effort||"M").toString().toLowerCase().trim();
        const size=sizeMap[rawSize]||(SIZES.includes((i.size||"").toUpperCase())?(i.size||"").toUpperCase():"M");
        const statusRaw=i.status||i.phase||"";
        const statusMatch=ALL_STATUSES.find(s=>s.toLowerCase()===statusRaw.toLowerCase())||statusRaw;
        return {key:i.key||"",name:i.name||i.projectName||i.project||"",assignee,status:statusMatch,designer:dm,squad:i.squad||"",startDate:i.startDate||"",endDate:i.endDate||"",size,pms:pmMatch?[pmMatch]:[],selected:false};
      }));
      setStep(3);
    }catch(e){setPasteErr("Invalid JSON: "+e.message);}
  }
  function parseTicketJson(){
    setPasteErr("");
    try{
      let text=pasteVal.replace(/```json|```/g,"").trim();
      const o1=text.indexOf("{"),o2=text.lastIndexOf("}");
      if(o1===-1||o2<=o1){setPasteErr("Could not find a JSON object.");return;}
      const parsed=JSON.parse(text.slice(o1,o2+1));
      const designer=team.designers.find(d=>(parsed.designer||parsed.assignee||"").toLowerCase().includes(d.toLowerCase()))||"";
      setF({id:"p_"+Date.now(),name:parsed.projectName||parsed.name||"",squad:parsed.squad||"",designer,pms:[],startDate:parsed.startDate||"",endDate:parsed.endDate||"",size:parsed.size||"M",status:parsed.status||""});
      setStep(3);
    }catch(e){setPasteErr("Invalid JSON: "+e.message);}
  }
  function submitBoard(){const sel=boardItems.filter(i=>i.selected);if(!sel.length)return;onAddBatch(sel);onClose();}
  function submitManual(){
    if(!f.name.trim()){setErr("Project name required");return;}
    if(!f.designer){setErr("Select a "+theRole.toLowerCase());return;}
    if(!f.startDate||!f.endDate){setErr("Both dates required");return;}
    onAdd({...f, id:f.id||"p_"+Date.now()});onClose();
  }
  function submitLeave(){
    if(!leaveF.designer){setLeaveErr("Select a "+theRole.toLowerCase());return;}
    if(!leaveF.startDate||!leaveF.endDate){setLeaveErr("Both dates required");return;}
    onAdd({id:"leave_"+Date.now(),name:leaveF.leaveType,type:"leave",leaveType:leaveF.leaveType,leaveStatus:leaveF.leaveStatus||"Planned",designer:leaveF.designer,squad:"",pms:[],startDate:leaveF.startDate,endDate:leaveF.endDate,size:"XS",status:""});
    onClose();
  }
  async function handleXlsxUpload(e){
    const file=e.target.files?.[0];if(!file)return;
    setXlsxLoading(true);setXlsxErr("");
    try{
      const buf=await file.arrayBuffer();
      const wb=XLSX.read(buf,{type:"array"});
      const dataSheets=wb.SheetNames.filter(s=>s!=="Timeline"&&s!=="Planning");
      let bestSheet=dataSheets.find(s=>/FY\d+H/i.test(s))||dataSheets.find(s=>/FY/i.test(s))||dataSheets[0];
      if(!bestSheet){setXlsxErr("No data sheet found.");setXlsxLoading(false);return;}
      const ws=wb.Sheets[bestSheet];
      const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:"",raw:true});
      let headerIdx=-1;
      for(let i=0;i<Math.min(8,rows.length);i++){const r=rows[i].map(c=>String(c||"").toLowerCase());if(r.some(c=>c.includes("project")||c==="start date"||c==="start"||c.includes("name"))){headerIdx=i;break;}}
      if(headerIdx===-1){setXlsxErr("Could not find header row.");setXlsxLoading(false);return;}
      const headers=rows[headerIdx].map(c=>String(c||"").toLowerCase().trim());
      // Exact match first, then partial contains match
      const colIdx=names=>{
        for(const n of names){const i=headers.indexOf(n);if(i>=0)return i;}
        for(const n of names){const i=headers.findIndex(h=>h.includes(n));if(i>=0)return i;}
        return -1;
      };
      const iStart=colIdx(["start date","est. design start","start"]);
      const iEnd=colIdx(["end date","est. design hand off","end"]);
      // "project name" before "name" so "Project Name, Size*" takes priority over "Name*"
      const iProject=colIdx(["project name","project","task"]);
      // "name*"/"name" as fallback for designer (e.g. "Name*" column)
      const iDesigner=colIdx(["designer","assignee","owner","name*","name"]);
      const iSquad=colIdx(["squad*","squad","team"]);
      const iPM=colIdx(["pm","product manager"]);
      const iPhase=colIdx(["phase","status"]);
      const iType=colIdx(["type*","type","category"]);
      const iEffort=colIdx(["effort","size","t-shirt"]);
      if(iProject===-1){setXlsxErr("Could not find a project name column (expected 'Project', 'Project Name', or 'Task').");setXlsxLoading(false);return;}
      // Parse dates: handles both Excel serial numbers (e.g. 46132) and string dates
      const parseDate=v=>{
        if(!v&&v!==0)return "";
        const n=Number(v);
        if(!isNaN(n)&&n>1000){
          // Excel serial: days since Dec 30, 1899 (accounts for Excel's leap-year bug)
          const d=new Date(Math.round((n-25569)*86400*1000));
          return d.toISOString().slice(0,10);
        }
        const d=new Date(String(v));
        if(!isNaN(d.getTime()))return d.toISOString().slice(0,10);
        return "";
      };
      const sizeMap={"xs":"XS","s":"S","m":"M","l":"L","xl":"XL","small":"S","medium":"M","large":"L"};
      const items=[];
      for(let i=headerIdx+1;i<rows.length;i++){
        const row=rows[i];
        const project=String(row[iProject]||"").trim();
        const startRaw=iStart>=0?row[iStart]:"";
        if(!project||!startRaw)continue;
        const startDate=parseDate(startRaw);
        const endDate=iEnd>=0?parseDate(row[iEnd]):"";
        const designerRaw=iDesigner>=0?String(row[iDesigner]||"").trim():"";
        const dm=team.designers.find(d=>designerRaw.toLowerCase().includes(d.toLowerCase()))||designerRaw||"";
        const squadRaw=iSquad>=0?String(row[iSquad]||"").trim():"";
        const pmRaw=iPM>=0?String(row[iPM]||"").trim():"";
        const pmMatch=(team.pms||[]).find(p=>pmRaw.toLowerCase().includes(p.split(" ")[0].toLowerCase()))||"";
        const pmValue=pmMatch||(pmRaw||"");
        const phaseRaw=iPhase>=0?String(row[iPhase]||"").trim():"";
        const statusMatch=ALL_STATUSES.find(s=>s.toLowerCase()===phaseRaw.toLowerCase())||"";
        const typeRaw=iType>=0?String(row[iType]||"").trim():"";
        const typeMatch=PROJECT_TYPES.find(t=>t.label.toLowerCase()===typeRaw.toLowerCase())?.label||
          PROJECT_TYPES.find(t=>typeRaw.toLowerCase().includes(t.label.toLowerCase().split(" ")[0].toLowerCase()))?.label||typeRaw||"";
        const effortRaw=iEffort>=0?String(row[iEffort]||"m").toLowerCase().trim():"m";
        const size=sizeMap[effortRaw]||(SIZES.includes(effortRaw.toUpperCase())?effortRaw.toUpperCase():"M");
        items.push({key:"",name:project,assignee:designerRaw,pmRaw,status:statusMatch,type:typeMatch,designer:dm,startDate,endDate,size,pms:pmValue?[pmValue]:[],selected:false,squad:squadRaw});
      }
      if(!items.length){setXlsxErr("No project rows found.");setXlsxLoading(false);return;}
      setBoardItems(items);setStep(3);
    }catch(e){setXlsxErr("Could not read file: "+e.message);}
    setXlsxLoading(false);
  }
  const selectedCount=boardItems.filter(i=>i.selected).length;
  const boardMissingCount=boardItems.filter(i=>i.selected&&(!i.designer||!i.startDate||!i.endDate)).length;
  const maxW=(mode==="board"||mode==="sheets"||mode==="xlsx")&&step===3?1000:500;
  const isEditing = !!(prefill&&prefill.id&&prefill.name);
  const modalTitle=mode==="xlsx"?(step===3?"Review & Import":"Import from Spreadsheet"):mode==="manual"?(isEditing?"Edit Project":(lockedDesigner||prefill)?"Add Project":"Add Manually"):mode==="leave"?"Add Leave":"Add Project";

  // Effective squad-PM map: team's saved map merged over the global default
  const effectiveSqPmMap = {...SQUAD_PM_MAP,...(team.squadPmMap||{})};
  const effectivePmSqMap = {};
  Object.entries(effectiveSqPmMap).forEach(([sq,pms])=>pms.forEach(pm=>{effectivePmSqMap[pm]=sq;}));

  function handleSquadChange(sq){
    setF(p=>({...p,squad:sq,pms:[]}));
  }
  function handlePmsChange(v){
    setF(p=>({...p,pms:v}));
    // Auto-fill squad if currently empty and exactly one PM selected
    if(!f.squad && v.length===1){
      const autoSq=effectivePmSqMap[v[0]];
      if(autoSq) setF(p=>({...p,pms:v,squad:autoSq}));
    }
  }

  const manualForm = (
    <div style={{display:"flex",flexDirection:"column",gap:14}}>
      <div><label style={lbl}>PROJECT NAME</label><input value={f.name} onChange={e=>setField("name",e.target.value)} placeholder="Project name" style={inp}/></div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        <div><label style={lbl}>{theRole.toUpperCase()}</label>
          {lockedDesigner?<div style={{padding:"0 12px",height:40,boxSizing:"border-box",display:"flex",alignItems:"center",borderRadius:8,border:"1px solid #ECEFF1",background:"#FAFAFA",fontSize:14,color:"#0F172A",fontWeight:600,fontFamily:"'Inter',sans-serif"}}>{lockedDesigner}</div>:<CustomSelect value={f.designer} onChange={v=>setField("designer",v)} options={team.designers} placeholder={"Select "+theRole+"..."}/>}
        </div>
        <div><label style={lbl}>SQUAD</label><CustomSelect value={f.squad} onChange={handleSquadChange} options={team.squads||[]} placeholder="Select squad..."/></div>
      </div>
      {!team.hidePM&&<div><label style={lbl}>PM {f.pms.length>0?"("+f.pms.length+" selected)":""}</label><MultiSelect options={f.squad&&effectiveSqPmMap[f.squad]?.length?effectiveSqPmMap[f.squad]:(team.pms||[])} selected={f.pms} onChange={handlePmsChange} placeholder="Select PMs..."/></div>}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
        <div><label style={lbl}>START DATE</label><div style={{position:"relative"}}><input type="date" value={f.startDate} onChange={e=>setField("startDate",e.target.value)} className="date-no-icon" style={inp}/><span style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",color:"#90A4AE",pointerEvents:"none",display:"flex"}}><MI name="calendar_today" size={16} /></span></div></div>
        <div><label style={lbl}>END DATE</label><div style={{position:"relative"}}><input type="date" value={f.endDate} onChange={e=>setField("endDate",e.target.value)} className="date-no-icon" style={inp}/><span style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",color:"#90A4AE",pointerEvents:"none",display:"flex"}}><MI name="calendar_today" size={16} /></span></div></div>
      </div>
      <div>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:5}}>
          <label style={{...lbl,marginBottom:0}}>SIZE</label>
          <button onClick={()=>setShowCalc(c=>!c)} style={{fontSize:11,fontWeight:700,color:showCalc?"#90A4AE":"#2563EB",background:"none",border:"none",cursor:"pointer",fontFamily:"'Inter',sans-serif"}}>{showCalc?"Close calculator":"Calculate"}</button>
        </div>
        <div style={{display:"flex",gap:8}}>{SIZES.map(s=><button key={s} onClick={()=>setField("size",s)} style={{flex:1,padding:"8px 4px",borderRadius:7,border:"1px solid"+(f.size===s?"#2563EB":"#ECEFF1"),background:f.size===s?"#2563EB":"#FAFAFA",color:f.size===s?"#fff":"#546E7A",fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"monospace"}}>{s}</button>)}</div>
        {showCalc&&<div style={{marginTop:10}}><SizeCalculator initialScores={f.calcScores} onResult={(s,sc)=>{setF(p=>({...p,size:s,calcScores:sc}));setShowCalc(false);}} onClose={()=>setShowCalc(false)}/></div>}
      </div>
      <div><label style={lbl}>STATUS</label><CustomSelect value={f.status} onChange={v=>setField("status",v)} options={ALL_STATUSES} placeholder="Select status..."/></div>
      <div>
        <label style={lbl}>TYPE <span style={{fontWeight:400,color:"#B0BEC5"}}>(optional)</span></label>
        <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
          {PROJECT_TYPES.map(t=>{const on=f.type===t.label;return<button key={t.label} type="button" onClick={()=>setField("type",on?"":t.label)} style={{padding:"4px 10px",borderRadius:20,fontSize:11,fontWeight:600,cursor:"pointer",border:`1px solid ${on?t.color+"66":"#ECEFF1"}`,background:on?t.bg:"#FAFAFA",color:on?t.color:"#607D8B",fontFamily:"'Inter',sans-serif",transition:"all 0.12s"}}>{t.label}</button>;})}
        </div>
      </div>
      {/* Links */}
      <div style={{display:"flex",gap:8}}>
        <div style={{flex:1}}>
          <label style={lbl}>FIGMA LINK <span style={{fontWeight:400,color:"#B0BEC5"}}>(optional)</span></label>
          <input value={f.figmaUrl} onChange={e=>setField("figmaUrl",e.target.value)} placeholder="https://figma.com/..." style={inp}/>
        </div>
        <div style={{flex:1}}>
          <label style={lbl}>JIRA LINK <span style={{fontWeight:400,color:"#B0BEC5"}}>(optional)</span></label>
          <input value={f.jiraUrl} onChange={e=>setField("jiraUrl",e.target.value)} placeholder="https://jira.com/..." style={inp}/>
        </div>
      </div>
      {/* Note */}
      <div>
        <label style={lbl}>NOTE <span style={{fontWeight:400,color:"#B0BEC5"}}>(optional)</span></label>
        <textarea value={f.note} onChange={e=>setField("note",e.target.value)} placeholder="Recent progress, comments..." rows={2} style={{...inp,height:"auto",padding:"8px 11px",resize:"vertical",lineHeight:1.5}}/>
      </div>
      {/* Blocker */}
      <div>
        <label style={lbl}>BLOCKER <span style={{fontWeight:400,color:"#B0BEC5"}}>(optional)</span></label>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <MI name="flag" size={16} style={{color:f.blocker?"#EF5350":"#D1D5DB",flexShrink:0}}/>
          <input value={f.blocker} onChange={e=>setField("blocker",e.target.value)} placeholder="Describe the blocker..." style={{...inp,borderColor:f.blocker?"#EF535033":"#ECEFF1"}}/>
        </div>
      </div>
      {err&&<div style={{fontSize:12,color:"#C62828",padding:"8px 12px",background:"#FFEBEE",borderRadius:6}}>{err}</div>}
      <div style={{display:"flex",gap:8}}>
        {!lockedDesigner&&!prefill&&<button onClick={()=>setMode(null)} style={{flex:1,height:40,padding:"0 20px",borderRadius:20,border:"none",background:"#F1F5FF",color:"#3B5BDB",fontSize:14,fontWeight:500,cursor:"pointer",fontFamily:"'Inter',sans-serif"}}>Back</button>}
        <button onClick={submitManual} style={{flex:2,height:40,padding:"0 24px",borderRadius:20,border:"none",background:"#2563EB",color:"#fff",fontSize:14,fontWeight:500,cursor:"pointer",fontFamily:"'Inter',sans-serif",boxShadow:"0 1px 2px rgba(0,0,0,0.3)"}}>{isEditing?"Save Changes":"Add Project"}</button>
      </div>
    </div>
  );

  return (
    <div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.4)",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center",padding:32}}>
      <div onClick={e=>e.stopPropagation()} style={{background:"#fff",borderRadius:28,width:"100%",maxWidth:maxW,maxHeight:"calc(100vh - 64px)",display:"flex",flexDirection:"column",boxShadow:"0 4px 24px rgba(0,0,0,0.12)"}}>
        <div style={{padding:"20px 24px 16px",borderBottom:"1px solid #F5F5F5",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0}}>
          <div style={{fontSize:16,fontWeight:700,color:"#0F172A",fontFamily:"'Inter',sans-serif"}}>{modalTitle}</div>
          <button onClick={onClose} style={{background:"transparent",border:"none",cursor:"pointer",color:"#90A4AE",display:"flex",borderRadius:"50%",padding:4}}><MI name="close" size={16} /></button>
        </div>
        <div style={{padding:24,overflowY:(mode==="board"||mode==="sheets"||mode==="xlsx")&&step===3?"hidden":"auto",flex:1,display:"flex",flexDirection:"column",minHeight:0}}>
          {mode===null&&(
            <div style={{display:"flex",flexDirection:"column",gap:10}}>
              <p style={{fontSize:14,color:"#607D8B",margin:"0 0 8px",fontFamily:"'Inter',sans-serif"}}>How would you like to add a project?</p>
              {[
                {id:"xlsx",icon:"table_view",title:"Import from Spreadsheet (.xlsx)",desc:"Upload an Excel file with project data"},
                {id:"manual",icon:"edit_note",title:"Add manually",desc:"Fill in all project details yourself"},
                {id:"leave",icon:"event_busy",title:"Add leave",desc:"PTO, STO, Maternity or Paternity Leave"},
              ].map(({id,icon,title,desc})=>(
                <button key={id} onClick={()=>{setMode(id);setStep(1);setPasteVal("");setPasteErr("");setBoardItems([]);setCopied(false);if(id==="manual")setF({id:"",name:"",squad:"",designer:prefillDesigner,pms:[],startDate:"",endDate:"",size:"M",status:""});if(id==="leave")setLeaveF({designer:prefillDesigner,leaveType:"PTO",leaveStatus:"Planned",startDate:"",endDate:""});}}
                  style={{display:"flex",alignItems:"center",gap:14,padding:"14px 16px",borderRadius:10,border:"1px solid #ECEFF1",background:"#FAFAFA",cursor:"pointer",textAlign:"left",width:"100%"}}
                  onMouseEnter={e=>{e.currentTarget.style.borderColor="#2563EB";e.currentTarget.style.background="#fff";}}
                  onMouseLeave={e=>{e.currentTarget.style.borderColor="#ECEFF1";e.currentTarget.style.background="#FAFAFA";}}>
                  <MI name={icon} size={22} style={{flexShrink:0,color:"#607D8B"}}/>
                  <div><div style={{fontSize:14,fontWeight:700,color:"#0F172A",fontFamily:"'Inter',sans-serif"}}>{title}</div><div style={{fontSize:11,color:"#90A4AE",marginTop:2}}>{desc}</div></div>
                </button>
              ))}
            </div>
          )}
          {mode==="manual"&&manualForm}
          {mode==="leave"&&(
            <div style={{display:"flex",flexDirection:"column",gap:14}}>
              <div><label style={{fontSize:11,fontWeight:700,color:"#90A4AE",textTransform:"uppercase",letterSpacing:"0.05em",display:"block",marginBottom:6}}>{theRole.toUpperCase()}</label>
                {lockedDesigner?<div style={{padding:"0 12px",height:40,boxSizing:"border-box",display:"flex",alignItems:"center",borderRadius:8,border:"1px solid #ECEFF1",background:"#FAFAFA",fontSize:14,color:"#0F172A",fontWeight:600,fontFamily:"'Inter',sans-serif"}}>{lockedDesigner}</div>:<CustomSelect value={leaveF.designer} onChange={v=>setLeaveF(p=>({...p,designer:v}))} options={team.designers} placeholder={"Select "+theRole+"..."}/>}
              </div>
              <div><label style={{fontSize:11,fontWeight:700,color:"#90A4AE",textTransform:"uppercase",letterSpacing:"0.05em",display:"block",marginBottom:6}}>LEAVE TYPE</label>
                <div style={{display:"flex",gap:6}}>{LEAVE_TYPES.map(lt=>{const active=leaveF.leaveType===lt;const c=LEAVE_COLORS[lt];return<button key={lt} onClick={()=>setLeaveF(p=>({...p,leaveType:lt}))} style={{flex:1,padding:"8px 4px",borderRadius:7,border:"1px solid"+(active?c:"#ECEFF1"),background:active?c+"18":"#FAFAFA",color:active?c:"#607D8B",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"'Inter',sans-serif"}}>{lt}</button>;})}</div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
                <div><label style={{fontSize:11,fontWeight:700,color:"#90A4AE",textTransform:"uppercase",letterSpacing:"0.05em",display:"block",marginBottom:6}}>START DATE</label><div style={{position:"relative"}}><input type="date" value={leaveF.startDate} onChange={e=>setLeaveF(p=>({...p,startDate:e.target.value}))} className="date-no-icon" style={{width:"100%",padding:"0 12px",height:40,boxSizing:"border-box",borderRadius:4,border:"1px solid #ECEFF1",fontSize:14,color:"#0F172A",outline:"none",background:"#FAFAFA",fontFamily:"'Inter',sans-serif"}}/><span style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",color:"#90A4AE",pointerEvents:"none",display:"flex"}}><MI name="calendar_today" size={16}/></span></div></div>
                <div><label style={{fontSize:11,fontWeight:700,color:"#90A4AE",textTransform:"uppercase",letterSpacing:"0.05em",display:"block",marginBottom:6}}>END DATE</label><div style={{position:"relative"}}><input type="date" value={leaveF.endDate} onChange={e=>setLeaveF(p=>({...p,endDate:e.target.value}))} className="date-no-icon" style={{width:"100%",padding:"0 12px",height:40,boxSizing:"border-box",borderRadius:4,border:"1px solid #ECEFF1",fontSize:14,color:"#0F172A",outline:"none",background:"#FAFAFA",fontFamily:"'Inter',sans-serif"}}/><span style={{position:"absolute",right:10,top:"50%",transform:"translateY(-50%)",color:"#90A4AE",pointerEvents:"none",display:"flex"}}><MI name="calendar_today" size={16}/></span></div></div>
              </div>
              <div><label style={{fontSize:11,fontWeight:700,color:"#90A4AE",textTransform:"uppercase",letterSpacing:"0.05em",display:"block",marginBottom:6}}>STATUS</label>
                <div style={{display:"flex",gap:6}}>{LEAVE_STATUSES.map(s=>{const active=leaveF.leaveStatus===s;const sc=s==="Planned"?"#3949AB":s==="In Progress"?"#1565C0":"#2E7D32";return<button key={s} onClick={()=>setLeaveF(p=>({...p,leaveStatus:s}))} style={{flex:1,padding:"8px 4px",borderRadius:7,border:"1px solid"+(active?sc:"#ECEFF1"),background:active?sc+"18":"#FAFAFA",color:active?sc:"#607D8B",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"'Inter',sans-serif"}}>{s}</button>;})}</div>
              </div>
              {leaveErr&&<div style={{fontSize:12,color:"#C62828",padding:"8px 12px",background:"#FFEBEE",borderRadius:6}}>{leaveErr}</div>}
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>setMode(null)} style={{flex:1,height:40,padding:"0 20px",borderRadius:20,border:"none",background:"#F1F5FF",color:"#3B5BDB",fontSize:14,fontWeight:500,cursor:"pointer",fontFamily:"'Inter',sans-serif"}}>Back</button>
                <button onClick={submitLeave} style={{flex:2,height:40,padding:"0 24px",borderRadius:20,border:"none",background:"#2563EB",color:"#fff",fontSize:14,fontWeight:500,cursor:"pointer",fontFamily:"'Inter',sans-serif",boxShadow:"0 1px 2px rgba(0,0,0,0.3)"}}>Add Leave</button>
              </div>
            </div>
          )}
          {/* Board/Sheets/XLSX step 1 — prompt copy */}
          {(mode==="board"||mode==="sheets")&&step===1&&(
            <div style={{display:"flex",flexDirection:"column",gap:16}}>
              <div style={{display:"flex",gap:12,alignItems:"flex-start"}}>
                <div style={stepBadge(1,true)}>1</div>
                <div style={{flex:1}}>
                  <div style={{fontSize:14,fontWeight:700,color:"#0F172A",marginBottom:6}}>Copy this prompt</div>
                  <input value={importUrl} onChange={e=>setImportUrl(e.target.value)} placeholder={mode==="board"?"https://company.atlassian.net/jira/boards/...":"https://docs.google.com/spreadsheets/d/..."} style={{...inp,marginBottom:10}}/>
                  <div style={{position:"relative",background:"#FAFAFA",border:"1px solid #ECEFF1",borderRadius:8,padding:"12px 14px",paddingRight:60}}>
                    <div style={{fontSize:11,color:"#546E7A",fontFamily:"monospace",lineHeight:1.6,whiteSpace:"pre-wrap",wordBreak:"break-word",maxHeight:120,overflowY:"auto"}}>{mode==="board"?boardPrompt:sheetsPrompt}{importUrl||"[paste your link here]"}</div>
                    <button onClick={()=>copyPrompt((mode==="board"?boardPrompt:sheetsPrompt)+(importUrl||"[paste your link here]"))} style={{position:"absolute",top:8,right:8,fontSize:11,fontWeight:700,color:copied?"#2E7D32":"#2563EB",background:copied?"#E8F5E9":"#EEF4FD",border:"1px solid "+(copied?"#A5D6A7":"#C5D8FD"),borderRadius:5,padding:"4px 10px",cursor:"pointer",fontFamily:"'Inter',sans-serif"}}>{copied?"Copied!":"Copy"}</button>
                  </div>
                </div>
              </div>
              <div style={{display:"flex",gap:8,marginTop:4}}>
                <button onClick={()=>setMode(null)} style={{flex:1,height:40,padding:"0 20px",borderRadius:20,border:"none",background:"#F1F5FF",color:"#3B5BDB",fontSize:14,fontWeight:500,cursor:"pointer",fontFamily:"'Inter',sans-serif"}}>Back</button>
                <button onClick={()=>setStep(2)} style={{flex:2,height:40,padding:"0 24px",borderRadius:20,border:"none",background:"#2563EB",color:"#fff",fontSize:14,fontWeight:500,cursor:"pointer",fontFamily:"'Inter',sans-serif",boxShadow:"0 1px 2px rgba(0,0,0,0.3)"}}>I copied the prompt →</button>
              </div>
            </div>
          )}
          {(mode==="board"||mode==="sheets")&&step===2&&(
            <div style={{display:"flex",flexDirection:"column",gap:16}}>
              <div style={{display:"flex",gap:12,alignItems:"flex-start"}}>
                <div style={stepBadge(2,true)}>2</div>
                <div style={{flex:1}}>
                  <div style={{fontSize:14,fontWeight:700,color:"#0F172A",marginBottom:6}}>Paste Claude JSON response</div>
                  <textarea value={pasteVal} onChange={e=>{setPasteVal(e.target.value);setPasteErr("");}} placeholder="Paste the JSON array here..." style={{...inp,minHeight:120,resize:"vertical",fontFamily:"monospace",fontSize:11}}/>
                  {pasteErr&&<div style={{fontSize:12,color:"#C62828",padding:"8px 12px",background:"#FFEBEE",borderRadius:6,marginTop:6}}>{pasteErr}</div>}
                </div>
              </div>
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>setStep(1)} style={{flex:1,height:40,padding:"0 20px",borderRadius:20,border:"none",background:"#F1F5FF",color:"#3B5BDB",fontSize:14,fontWeight:500,cursor:"pointer",fontFamily:"'Inter',sans-serif"}}>Back</button>
                <button onClick={parseBoardJson} disabled={!pasteVal.trim()} style={{flex:2,padding:11,borderRadius:20,border:"none",background:"#2563EB",color:"#fff",fontSize:14,fontWeight:700,cursor:pasteVal.trim()?"pointer":"not-allowed",fontFamily:"'Inter',sans-serif",opacity:pasteVal.trim()?1:0.5}}>Parse & Review →</button>
              </div>
            </div>
          )}
          {mode==="jira"&&step===1&&(
            <div style={{display:"flex",flexDirection:"column",gap:16}}>
              <div style={{display:"flex",gap:12,alignItems:"flex-start"}}>
                <div style={stepBadge(1,true)}>1</div>
                <div style={{flex:1}}>
                  <div style={{fontSize:14,fontWeight:700,color:"#0F172A",marginBottom:6}}>Copy this prompt</div>
                  <input value={importUrl} onChange={e=>setImportUrl(e.target.value)} placeholder="https://company.atlassian.net/browse/PROJ-123" style={{...inp,marginBottom:10}}/>
                  <div style={{position:"relative",background:"#FAFAFA",border:"1px solid #ECEFF1",borderRadius:8,padding:"12px 14px",paddingRight:60}}>
                    <div style={{fontSize:11,color:"#546E7A",fontFamily:"monospace",lineHeight:1.6,whiteSpace:"pre-wrap",wordBreak:"break-word",maxHeight:100,overflowY:"auto"}}>{ticketPrompt}{importUrl||"[paste your ticket link here]"}</div>
                    <button onClick={()=>copyPrompt(ticketPrompt+(importUrl||"[paste your ticket link here]"))} style={{position:"absolute",top:8,right:8,fontSize:11,fontWeight:700,color:copied?"#2E7D32":"#2563EB",background:copied?"#E8F5E9":"#EEF4FD",border:"1px solid "+(copied?"#A5D6A7":"#C5D8FD"),borderRadius:5,padding:"4px 10px",cursor:"pointer",fontFamily:"'Inter',sans-serif"}}>{copied?"Copied!":"Copy"}</button>
                  </div>
                </div>
              </div>
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>setMode(null)} style={{flex:1,height:40,padding:"0 20px",borderRadius:20,border:"none",background:"#F1F5FF",color:"#3B5BDB",fontSize:14,fontWeight:500,cursor:"pointer",fontFamily:"'Inter',sans-serif"}}>Back</button>
                <button onClick={()=>setStep(2)} style={{flex:2,height:40,padding:"0 24px",borderRadius:20,border:"none",background:"#2563EB",color:"#fff",fontSize:14,fontWeight:500,cursor:"pointer",fontFamily:"'Inter',sans-serif",boxShadow:"0 1px 2px rgba(0,0,0,0.3)"}}>I copied the prompt →</button>
              </div>
            </div>
          )}
          {mode==="jira"&&step===2&&(
            <div style={{display:"flex",flexDirection:"column",gap:16}}>
              <div style={{display:"flex",gap:12,alignItems:"flex-start"}}>
                <div style={stepBadge(2,true)}>2</div>
                <div style={{flex:1}}>
                  <div style={{fontSize:14,fontWeight:700,color:"#0F172A",marginBottom:6}}>Paste Claude JSON response</div>
                  <textarea value={pasteVal} onChange={e=>{setPasteVal(e.target.value);setPasteErr("");}} placeholder='{"projectName":"My Project","designer":"Name","size":"M"}' style={{...inp,minHeight:90,resize:"vertical",fontFamily:"monospace",fontSize:11}}/>
                  {pasteErr&&<div style={{fontSize:12,color:"#C62828",padding:"8px 12px",background:"#FFEBEE",borderRadius:6,marginTop:6}}>{pasteErr}</div>}
                </div>
              </div>
              <div style={{display:"flex",gap:8}}>
                <button onClick={()=>setStep(1)} style={{flex:1,height:40,padding:"0 20px",borderRadius:20,border:"none",background:"#F1F5FF",color:"#3B5BDB",fontSize:14,fontWeight:500,cursor:"pointer",fontFamily:"'Inter',sans-serif"}}>Back</button>
                <button onClick={parseTicketJson} disabled={!pasteVal.trim()} style={{flex:2,padding:11,borderRadius:20,border:"none",background:"#2563EB",color:"#fff",fontSize:14,fontWeight:700,cursor:pasteVal.trim()?"pointer":"not-allowed",fontFamily:"'Inter',sans-serif",opacity:pasteVal.trim()?1:0.5}}>Parse & Complete →</button>
              </div>
            </div>
          )}
          {mode==="jira"&&step===3&&<ManualForm/>}
          {mode==="xlsx"&&step===1&&(
            <div style={{display:"flex",flexDirection:"column",gap:16}}>
              <div style={{fontSize:12,color:"#607D8B",lineHeight:1.5}}>Supported column names:<br/><b>Project:</b> Project, Project Name, Task &nbsp;·&nbsp; <b>Designer:</b> Designer, Assignee, Owner, Name<br/><b>Dates:</b> Start Date, End Date &nbsp;·&nbsp; <b>Squad:</b> Squad &nbsp;·&nbsp; <b>Size:</b> Effort, Size &nbsp;·&nbsp; <b>PM:</b> PM</div>
              <div onClick={()=>xlsxRef.current?.click()} style={{border:"2px dashed #B0BEC5",borderRadius:10,padding:"32px 24px",textAlign:"center",cursor:"pointer",background:"#FAFAFA"}} onMouseEnter={e=>{e.currentTarget.style.borderColor="#2563EB";}} onMouseLeave={e=>{e.currentTarget.style.borderColor="#B0BEC5";}}>
                <div style={{fontSize:14,fontWeight:700,color:"#0F172A",fontFamily:"'Inter',sans-serif",marginBottom:4}}>{xlsxLoading?"Reading file...":"Upload .xlsx file"}</div>
                <div style={{fontSize:12,color:"#90A4AE"}}>Click to browse</div>
                <input ref={xlsxRef} type="file" accept=".xlsx,.xls" onChange={handleXlsxUpload} style={{display:"none"}}/>
              </div>
              {xlsxErr&&<div style={{fontSize:12,color:"#C62828",padding:"8px 12px",background:"#FFEBEE",borderRadius:6}}>{xlsxErr}</div>}
              <button onClick={()=>setMode(null)} style={{height:40,padding:"0 20px",borderRadius:20,border:"none",background:"#F1F5FF",color:"#3B5BDB",fontSize:14,fontWeight:500,cursor:"pointer",fontFamily:"'Inter',sans-serif"}}>Back</button>
            </div>
          )}
          {(mode==="board"||mode==="sheets"||mode==="xlsx")&&step===3&&(
            <div style={{display:"flex",flexDirection:"column",flex:1,minHeight:0,gap:8}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexShrink:0}}>
                <div style={{fontSize:12,color:"#607D8B"}}>{boardItems.length} items · {selectedCount} selected</div>
                <button onClick={toggleAll} style={{fontSize:11,fontWeight:600,color:"#2563EB",background:"none",border:"none",cursor:"pointer"}}>{boardItems.every(i=>i.selected)?"Deselect all":"Select all"}</button>
              </div>
              <div style={{flex:1,overflowY:"auto",overflowX:"auto",border:"1px solid #ECEFF1",borderRadius:8,minHeight:0}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,fontFamily:"'Inter',sans-serif"}}>
                  <thead><tr style={{background:"#FAFAFA",borderBottom:"1px solid #ECEFF1"}}>
                    <th style={{padding:"8px 10px",width:32}}></th>
                    <th style={{padding:"8px 10px",textAlign:"left",fontWeight:700,color:"#90A4AE",fontSize:11,minWidth:140}}>PROJECT</th>
                    <th style={{padding:"8px 10px",textAlign:"left",fontWeight:700,color:"#90A4AE",fontSize:11,width:110}}>DESIGNER</th>
                    <th style={{padding:"8px 10px",textAlign:"left",fontWeight:700,color:"#90A4AE",fontSize:11,width:110}}>SQUAD</th>
                    <th style={{padding:"8px 10px",textAlign:"left",fontWeight:700,color:"#90A4AE",fontSize:11,width:110}}>PM</th>
                    <th style={{padding:"8px 10px",textAlign:"left",fontWeight:700,color:"#90A4AE",fontSize:11,width:100}}>START</th>
                    <th style={{padding:"8px 10px",textAlign:"left",fontWeight:700,color:"#90A4AE",fontSize:11,width:100}}>END</th>
                    <th style={{padding:"8px 10px",textAlign:"left",fontWeight:700,color:"#90A4AE",fontSize:11,width:60}}>SIZE</th>
                  </tr></thead>
                  <tbody>{boardItems.map((it,idx)=>(
                    <tr key={idx} style={{borderBottom:"1px solid #F5F5F5",background:it.selected?"#F0F9FF":"#fff",opacity:it.selected?1:0.6}}>
                      <td style={{padding:"6px 10px",textAlign:"center"}}><input type="checkbox" checked={it.selected} onChange={()=>toggleItem(idx)} style={{cursor:"pointer"}}/></td>
                      <td style={{padding:"6px 10px"}}><div style={{fontSize:12,fontWeight:600,color:"#0F172A"}}>{it.name}</div></td>
                      <td style={{padding:"4px 6px"}}>
                        <select value={it.designer} onChange={e=>updateItem(idx,"designer",e.target.value)} style={{width:"100%",padding:"4px 6px",borderRadius:5,border:"1px solid "+(it.selected&&!it.designer?"#FFCDD2":"#ECEFF1"),fontSize:11,background:it.selected&&!it.designer?"#FFEBEE":"#FAFAFA",fontFamily:"'Inter',sans-serif"}}>
                          <option value="">Select...</option>
                          {[...(team.designers||[]),...(it.assignee&&!(team.designers||[]).some(d=>d.toLowerCase()===it.assignee.toLowerCase())?[it.assignee]:[])].map(d=><option key={d} value={d}>{d}</option>)}
                        </select>
                        {it.designer&&!(team.designers||[]).some(d=>d.toLowerCase()===it.designer.toLowerCase())&&<div style={{fontSize:9,color:"#E65100",marginTop:1}}>New designer — will be added</div>}
                      </td>
                      <td style={{padding:"4px 6px"}}><input value={it.squad||""} onChange={e=>updateItem(idx,"squad",e.target.value)} placeholder="Squad..." style={{width:"100%",padding:"4px 6px",borderRadius:5,border:"1px solid #ECEFF1",fontSize:11,background:"#FAFAFA",fontFamily:"'Inter',sans-serif",boxSizing:"border-box"}}/></td>
                      <td style={{padding:"4px 6px"}}>
                        <select value={it.pms[0]||""} onChange={e=>updateItem(idx,"pms",e.target.value?[e.target.value]:[])} style={{width:"100%",padding:"4px 6px",borderRadius:5,border:"1px solid #ECEFF1",fontSize:11,background:"#FAFAFA",fontFamily:"'Inter',sans-serif"}}>
                          <option value="">None</option>
                          {[...(team.pms||[]),...(it.pmRaw&&!(team.pms||[]).includes(it.pmRaw)&&!(team.pms||[]).some(p=>it.pmRaw.toLowerCase().includes(p.split(" ")[0].toLowerCase()))?[it.pmRaw]:[])].map(p=><option key={p} value={p}>{p}</option>)}
                        </select>
                        {it.pmRaw&&!(team.pms||[]).some(p=>it.pmRaw.toLowerCase().includes(p.split(" ")[0].toLowerCase()))&&it.pmRaw&&<div style={{fontSize:9,color:"#E65100",marginTop:1}}>New PM — will be added</div>}
                      </td>
                      <td style={{padding:"4px 6px"}}><input type="date" value={it.startDate} onChange={e=>updateItem(idx,"startDate",e.target.value)} className="date-no-icon" style={{width:"100%",padding:"4px 6px",borderRadius:5,border:"1px solid "+(it.selected&&!it.startDate?"#FFCDD2":"#ECEFF1"),fontSize:11,background:it.selected&&!it.startDate?"#FFEBEE":"#FAFAFA"}}/></td>
                      <td style={{padding:"4px 6px"}}><input type="date" value={it.endDate} onChange={e=>updateItem(idx,"endDate",e.target.value)} className="date-no-icon" style={{width:"100%",padding:"4px 6px",borderRadius:5,border:"1px solid "+(it.selected&&!it.endDate?"#FFCDD2":"#ECEFF1"),fontSize:11,background:it.selected&&!it.endDate?"#FFEBEE":"#FAFAFA"}}/></td>
                      <td style={{padding:"4px 6px"}}><select value={it.size} onChange={e=>updateItem(idx,"size",e.target.value)} style={{width:"100%",padding:"4px 6px",borderRadius:5,border:"1px solid #ECEFF1",fontSize:11,background:"#FAFAFA",fontFamily:"monospace"}}>{SIZES.map(s=><option key={s} value={s}>{s}</option>)}</select></td>
                    </tr>
                  ))}</tbody>
                </table>
              </div>
              {selectedCount>0&&boardMissingCount>0&&<div style={{fontSize:12,color:"#E65100",padding:"8px 12px",background:"#FFF8E1",borderRadius:6,border:"1px solid #FFE082"}}>{boardMissingCount} selected item{boardMissingCount>1?"s":""} missing designer or dates</div>}
              <div style={{display:"flex",gap:8,flexShrink:0,paddingTop:4,borderTop:"1px solid #F5F5F5"}}>
                <button onClick={()=>{setStep(mode==="xlsx"?1:2);setBoardItems([]);}} style={{flex:1,height:40,padding:"0 20px",borderRadius:20,border:"none",background:"#F1F5FF",color:"#3B5BDB",fontSize:14,fontWeight:500,cursor:"pointer",fontFamily:"'Inter',sans-serif"}}>Back</button>
                <button onClick={submitBoard} disabled={selectedCount===0||boardMissingCount>0} style={{flex:2,padding:11,borderRadius:20,border:"none",background:"#2563EB",color:"#fff",fontSize:14,fontWeight:700,cursor:selectedCount>0&&boardMissingCount===0?"pointer":"not-allowed",fontFamily:"'Inter',sans-serif",opacity:selectedCount===0||boardMissingCount>0?0.5:1}}>Add {selectedCount>0?selectedCount+" ":""}{selectedCount===1?"Project":"Projects"}</button>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ---------- OverloadPanel ----------
function OverloadPanel({overloaded,onClose}) {
  return (
    <div onClick={onClose} style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.4)",zIndex:10000,display:"flex",alignItems:"flex-start",justifyContent:"flex-end"}}>
      <div onClick={e=>e.stopPropagation()} style={{background:"#fff",width:440,height:"100vh",overflowY:"auto",boxShadow:"-2px 0 16px rgba(0,0,0,0.10)",display:"flex",flexDirection:"column",borderRadius:"16px 0 0 16px"}}>
        <div style={{padding:"24px 24px 16px",borderBottom:"1px solid #F5F5F5",display:"flex",justifyContent:"space-between",flexShrink:0}}>
          <div><div style={{fontSize:18,fontWeight:700,color:"#0F172A",fontFamily:"'Inter',sans-serif",marginBottom:4}}>Capacity Alert</div><div style={{fontSize:12,color:"#90A4AE"}}>{overloaded.length} designer{overloaded.length>1?"s":""} overloaded</div></div>
          <button onClick={onClose} style={{background:"none",border:"none",fontSize:18,cursor:"pointer",color:"#90A4AE",display:"flex"}}><MI name="close" size={16} /></button>
        </div>
        <div style={{padding:24,flex:1,display:"flex",flexDirection:"column",gap:16}}>
          {overloaded.map((o,i)=>{
            const today=new Date();today.setHours(0,0,0,0);
            const active=o.projects.filter(p=>parseD(p.startDate)<=today&&parseD(p.endDate)>=today&&!isLeaveItem(p));
            const score=active.reduce((s,p)=>s+(SIZE_POINTS[p.size]||0),0);
            return(
              <div key={i} style={{border:"1px solid #FFCDD2",borderRadius:10,overflow:"hidden",background:"#FFFAFA"}}>
                <div style={{padding:"12px 16px",background:"#FFEBEE",borderBottom:"1px solid #FFCDD2",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                  <div style={{fontSize:14,fontWeight:700,color:"#C62828"}}>{o.designer}</div>
                  <span style={{fontSize:11,fontWeight:700,color:"#C62828",background:"#fff",border:"1px solid #FFCDD2",borderRadius:6,padding:"2px 8px"}}>{score} pts · {active.length} active</span>
                </div>
                <div style={{padding:"12px 16px",display:"flex",flexDirection:"column",gap:6}}>
                  {active.map((p,j)=>(
                    <div key={j} style={{display:"flex",alignItems:"center",gap:8,fontSize:12,color:"#546E7A"}}>
                      <span style={{fontFamily:"monospace",fontWeight:700,fontSize:11,color:"#607D8B",background:"#F5F5F5",padding:"1px 5px",borderRadius:6}}>{p.size}</span>
                      <span>{p.name}</span>
                      <span style={{color:"#B0BEC5",marginLeft:"auto",whiteSpace:"nowrap"}}>{fmtS(p.startDate)} – {fmtS(p.endDate)}</span>
                    </div>
                  ))}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

// ---------- Timeline ----------
function Timeline({squads,onUpdate,onRemove,onEditProject,onOpenEdit,onToast,filter,onFilter,team,hideFilter,onUpdateTitle,onClickEmpty}) {
  const weeks=getWeeks();
  const {start:qS,end:qE}=qBounds();
  const [openPopover,setOpenPopover]=useState(null);
  const [titleEdit,setTitleEdit]=useState(null);
  return (
    <div>
      {!hideFilter&&<div style={{marginBottom:16}}>
        <div style={{display:"flex",flexWrap:"wrap",gap:6,alignItems:"center"}}>
          {[...team.designers].sort().map(name=>{const on=filter.includes(name);return(
            <div key={name} onClick={()=>onFilter(p=>p.includes(name)?p.filter(d=>d!==name):[...p,name])} style={{display:"flex",alignItems:"center",padding:"0 12px",height:32,borderRadius:8,fontSize:12,fontWeight:500,cursor:"pointer",userSelect:"none",background:on?"#EEF4FD":"transparent",color:on?"#2563EB":"#6B7280",border:`1px solid ${on?"#2563EB":"#D1D5DB"}`,transition:"all 0.15s",lineHeight:1}}>{name}</div>
          );})}
          {filter.length>0&&<div onClick={()=>onFilter([])} style={{display:"flex",alignItems:"center",padding:"0 12px",height:32,borderRadius:8,fontSize:12,fontWeight:500,cursor:"pointer",userSelect:"none",color:"#9CA3AF",border:"1px solid #E5E7EB",background:"transparent",transition:"all 0.15s"}}>Clear</div>}
        </div>
      </div>}
      <div style={{overflowX:"auto"}}>
        <div style={{minWidth:weeks.length*CELL_W+180,position:"relative"}}>
          {squads.length>0&&(()=>{
            const today=new Date();today.setHours(0,0,0,0);
            const lo=(today-weeks[0])/(24*3600*1000)/7*CELL_W;
            if(lo<0||lo>weeks.length*CELL_W)return null;
            return <div style={{position:"absolute",left:lo+160,top:0,width:2,bottom:0,background:"#EF5350",zIndex:100,opacity:0.7,pointerEvents:"none"}}/>;
          })()}
          <div style={{display:"flex",marginLeft:160,marginBottom:4}}>
            {weeks.map((w,i)=><div key={i} style={{width:CELL_W,fontSize:9,color:i%2===0?"#90A4AE":"transparent",textAlign:"center",fontFamily:"monospace",flexShrink:0}}>{i%2===0?w.toLocaleDateString("en-US",{month:"short",day:"numeric"}):""}</div>)}
          </div>
          {squads.map((sq,sqIdx)=>{
            const vis=sq.designers.filter(d=>team.designers.includes(d.name)&&(!filter.length||filter.includes(d.name)));
            if(!vis.length)return null;
            return (
              <div key={sq.id} style={{marginBottom:16}}>
                <div style={{height:2,background:SQUAD_COLORS[sq.colorIdx%SQUAD_COLORS.length]+"22",marginBottom:8,borderRadius:2}}/>
                {vis.map(designer=>{
                  const dIdx=team.designers.indexOf(designer.name);
                  const color=DESIGNER_COLORS[(dIdx>=0?dIdx:sqIdx)%DESIGNER_COLORS.length];
                  const flags=getFlags(designer),flagSet=new Set(flags.map(f=>f.week));
                  const isOver=flags.length>0,isEmpty=!designer.projects.length;
                  const sorted=[...designer.projects].sort((a,b)=>parseD(a.startDate)-parseD(b.startDate));
                  const lanes=[],laneEnds=[];
                  sorted.forEach(p=>{const ps=parseD(p.startDate),pe=parseD(p.endDate);let lane=laneEnds.findIndex(e=>ps>=e);if(lane===-1){lane=laneEnds.length;laneEnds.push(pe);}else laneEnds[lane]=pe;lanes.push({p,lane});});
                  const rowH=Math.max(1,laneEnds.length)*30+4;
                  return (
                    <div key={designer.id} style={{marginBottom:isOver||isEmpty?10:6}}>
                      {isOver&&<div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4,padding:"6px 12px",borderRadius:6,background:"#FFEBEE",border:"1px solid #FFCDD2"}}><MI name="flag" size={15} style={{color:"#C62828"}}/><span style={{fontSize:12,fontWeight:700,color:"#C62828"}}>{designer.name} is overloaded</span></div>}
                      {isEmpty&&<div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4,padding:"6px 12px",borderRadius:6,background:"#E8F5E9",border:"1px solid #A5D6A7"}}><MI name="info" size={15} style={{color:"#2E7D32"}}/><span style={{fontSize:12,fontWeight:700,color:"#2E7D32"}}>{designer.name} has no active projects</span></div>}
                      <div style={{display:"flex",alignItems:"flex-start"}}>
                        <div style={{width:180,display:"flex",alignItems:"flex-start",gap:8,flexShrink:0,paddingRight:8,paddingTop:6,overflow:"hidden"}}>
                          <div onClick={e=>{e.stopPropagation();if(onUpdateTitle)setTitleEdit(titleEdit===designer.name?null:designer.name);}} style={{cursor:onUpdateTitle?"pointer":"default",marginTop:1}}>
                            <Avatar initials={designer.avatar} color={color} emoji={(team.emojis||{})[designer.name]}/>
                          </div>
                          <div style={{minWidth:0,overflow:"hidden",flex:1}}>
                            <div style={{fontSize:12,fontWeight:700,color:isOver?"#C62828":"#0F172A",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{designer.name}</div>
                            {(team.titles||{})[designer.name]
                              ? <div style={{marginTop:2}}><span onClick={e=>{e.stopPropagation();if(onUpdateTitle)setTitleEdit(titleEdit===designer.name?null:designer.name);}} style={{fontSize:11,fontWeight:600,color:"#388E3C",background:"#C8E6C9",padding:"1px 5px",borderRadius:3,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",display:"inline-block",maxWidth:"100%",cursor:onUpdateTitle?"pointer":"default"}}>{(team.titles||{})[designer.name]}</span></div>
                              : onUpdateTitle?<div style={{marginTop:2}}><span onClick={e=>{e.stopPropagation();setTitleEdit(titleEdit===designer.name?null:designer.name);}} style={{fontSize:11,fontWeight:500,color:"#B0BEC5",cursor:"pointer",display:"inline-block"}}>+ add title</span></div>:null
                            }
                            {(()=>{
                              const dsMap=team.designerSquads||{};
                              const hasCustom=Object.prototype.hasOwnProperty.call(dsMap,designer.name);
                              const custom=dsMap[designer.name];
                              let chips;
                              if(hasCustom){
                                chips=custom||[]; // explicit — even empty means "show nothing"
                              } else {
                                const projSquads=[...new Set(designer.projects.map(p=>p.squad).filter(Boolean))];
                                chips=projSquads.length?projSquads:sq.name.split(" . ").map(s=>s.trim()).filter(Boolean);
                              }
                              const removeChip=(chip)=>{
                                if(!onUpdateTitle)return;
                                const newSqs=chips.filter(c=>c!==chip);
                                onUpdateTitle(designer.name,(team.titles||{})[designer.name]||"",newSqs,(team.emojis||{})[designer.name]||"");
                              };
                              return(<div style={{display:"flex",flexWrap:"wrap",gap:2,marginTop:2}}>{chips.map(s=>(
                                <span key={s} style={{display:"inline-flex",alignItems:"center",gap:2,fontSize:11,fontWeight:600,color:"#3949AB",background:"#E8EAF6",padding:"1px 4px 1px 5px",borderRadius:3,maxWidth:130}}>
                                  <span onClick={e=>{e.stopPropagation();if(onUpdateTitle)setTitleEdit(designer.name);}} style={{cursor:onUpdateTitle?"pointer":"default",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",maxWidth:100}}>{s}</span>
                                  {onUpdateTitle&&<span onClick={e=>{e.stopPropagation();removeChip(s);}} style={{cursor:"pointer",color:"#9FA8DA",lineHeight:1,flexShrink:0,display:"flex",alignItems:"center"}} onMouseEnter={e=>e.currentTarget.style.color="#3949AB"} onMouseLeave={e=>e.currentTarget.style.color="#9FA8DA"}><MI name="close" size={10}/></span>}
                                </span>
                              ))}</div>);
                            })()}
                          </div>
                        </div>
                        <div style={{display:"flex",position:"relative",height:rowH}}>
                          {weeks.map((w,wi)=>{
                            const isoDate=w.getFullYear()+"-"+String(w.getMonth()+1).padStart(2,"0")+"-"+String(w.getDate()).padStart(2,"0");
                            const wl=w.toLocaleDateString("en-US",{month:"short",day:"numeric"});
                            return <div key={wi} onClick={()=>{if(onClickEmpty)onClickEmpty(designer.name,isoDate);}} style={{width:CELL_W,height:rowH,flexShrink:0,background:flagSet.has(wl)?"#FFEBEE":wi%2===0?"#FAFAFA":"#F5F5F5",borderRight:"1px solid #ECEFF1",cursor:"pointer"}}/>;
                          })}
                          {lanes.map(({p,lane})=>{
                            const ps=new Date(Math.max(parseD(p.startDate),qS));
                            const pe=new Date(Math.min(parseD(p.endDate),qE));
                            if(ps>qE||pe<qS)return null;
                            const lo=Math.max(0,(ps-weeks[0])/(7*24*3600*1000))*CELL_W;
                            const ww=Math.max(1,(pe-ps)/(7*24*3600*1000));
                            return <Block key={p.id} proj={p} color={color} left={lo} top={lane*30+4} width={ww*CELL_W-2} onUpdate={onUpdate} onRemove={onRemove} onEditProject={onEditProject} onToast={onToast} isOpen={openPopover?.proj?.id===p.id} onOpen={setOpenPopover}/>;
                          })}
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            );
          })}
          {squads.length===0&&<div style={{textAlign:"center",padding:"60px 0",color:"#90A4AE",fontSize:14}}>No designers yet — use 'Manage Team' to add your team</div>}
        </div>
      </div>
      {openPopover&&(()=>{
        const latest=squads.flatMap(sq=>sq.designers.flatMap(d=>d.projects)).find(p=>p.id===openPopover.proj.id);
        return latest?<Popover proj={latest} color={openPopover.color} onClose={()=>setOpenPopover(null)} onRemove={onRemove?id=>{onRemove(id);setOpenPopover(null);}:undefined} onEditProject={onEditProject} onOpenEdit={onOpenEdit}/>:null;
      })()}
      {titleEdit&&<TitlePopover name={titleEdit} title={(team.titles||{})[titleEdit]||""} squads={team.squads||[]} assignedSquads={(team.designerSquads||{})[titleEdit]||[]} emoji={(team.emojis||{})[titleEdit]||""} onSave={(t,sq,em)=>{onUpdateTitle&&onUpdateTitle(titleEdit,t,sq,em);setTitleEdit(null);}} onClose={()=>setTitleEdit(null)}/>}
    </div>
  );
}

// ---------- AnalyticsPage ----------
function AnalyticsPage({state,teams,setDrawer,setManager,setPage}) {
  const [tab,setTab]=useState("summary");
  const today=new Date();today.setHours(0,0,0,0);
  const allManagerKeys=[...new Set(MANAGERS.flatMap(m=>{const sbs=SUB_MANAGERS[m]||[];return sbs.length?[m,...sbs]:[m];}))];
  const allSquadsA=allManagerKeys.flatMap(k=>state&&state.managers[k]?state.managers[k].squads||[]:[]);
  const stateDesignerMap={};
  allSquadsA.forEach(sq=>sq.designers.forEach(d=>{stateDesignerMap[d.name]=d;}));
  // Merge names from teams settings AND from actual state squads so newly added designers always appear
  const allDesignerNames=[...new Set(allManagerKeys.flatMap(k=>{
    const fromTeams=teams[k]?teams[k].designers||[]:MANAGER_DESIGNERS[k]||[];
    const fromState=(state&&state.managers[k]?state.managers[k].squads||[]:[]).flatMap(sq=>sq.designers.map(d=>d.name));
    return [...fromTeams,...fromState];
  }))].filter(n=>!MANAGER_AS_DESIGNER.has(n));
  const allDesigners=allDesignerNames.map(name=>stateDesignerMap[name]||{name,projects:[]});
  const allProjects=allDesigners.flatMap(d=>d.projects).filter(p=>!isLeaveItem(p));
  const activeProjects=allProjects.filter(p=>parseD(p.startDate)<=today&&parseD(p.endDate)>=today);
  const upcomingProjects=allProjects.filter(p=>{const s=parseD(p.startDate);return s>today&&s<=new Date(today.getTime()+28*24*3600*1000);});
  const endingSoon=allProjects.filter(p=>{const e=parseD(p.endDate);return e>=today&&e<=new Date(today.getTime()+28*24*3600*1000);});
  const availableDesigners=allDesigners.filter(d=>!d.projects.some(p=>parseD(p.startDate)<=today&&parseD(p.endDate)>=today));
  const overloadedAll=allDesigners.filter(d=>getFlags(d).length>0);
  const sizeCount={XS:0,S:0,M:0,L:0,XL:0};
  allProjects.forEach(p=>{if(sizeCount[p.size]!==undefined)sizeCount[p.size]++;});
  const qb=qBounds();
  const qWeeks=Math.round((qb.end-qb.start)/(7*24*3600*1000));
  const totalDesignerWeeks=allDesigners.length*qWeeks;
  const occupiedWeeks=allDesigners.reduce((acc,d)=>{let w=0;d.projects.forEach(p=>{const s=Math.max(parseD(p.startDate),qb.start);const e=Math.min(parseD(p.endDate),qb.end);if(e>s)w+=Math.round((e-s)/(7*24*3600*1000));});return acc+w;},0);
  const utilPct=totalDesignerWeeks>0?Math.round(occupiedWeeks/totalDesignerWeeks*100):0;
  const byManager=MANAGERS.map(m=>{
    const sbs=SUB_MANAGERS[m]||[];const mKeys=sbs.length?[m,...sbs]:[m];
    const mNames=[...new Set(mKeys.flatMap(k=>{
      const fromTeams=teams[k]?teams[k].designers||[]:MANAGER_DESIGNERS[k]||[];
      const fromState=(state&&state.managers[k]?state.managers[k].squads||[]:[]).flatMap(sq=>sq.designers.map(d=>d.name));
      return [...fromTeams,...fromState];
    }))].filter(n=>!MANAGER_AS_DESIGNER.has(n));
    const mSquads=mKeys.flatMap(k=>state&&state.managers[k]?state.managers[k].squads||[]:[]);
    const mMap={};mSquads.forEach(sq=>sq.designers.forEach(d=>{mMap[d.name]=d;}));
    const mDesigners=mNames.map(name=>mMap[name]||{name,projects:[]});
    const mActive=mDesigners.filter(d=>d.projects.some(p=>parseD(p.startDate)<=today&&parseD(p.endDate)>=today));
    const mOver=mDesigners.filter(d=>getFlags(d).length>0);
    return {name:m,total:mDesigners.length,active:mActive.length,overloaded:mOver.length,projects:mDesigners.flatMap(d=>d.projects).length};
  });
  const flags=[];
  overloadedAll.forEach(d=>{if(getFlags(d).length)flags.push({type:"overload",text:d.name+" is overloaded",severity:"red"});});
  availableDesigners.forEach(d=>flags.push({type:"available",text:d.name+" has no active projects",severity:"green"}));

  function mkDG(list){return allManagerKeys.map(k=>{
    const fromTeams=teams[k]?teams[k].designers||[]:MANAGER_DESIGNERS[k]||[];
    const fromState=(state&&state.managers[k]?state.managers[k].squads||[]:[]).flatMap(sq=>sq.designers.map(d=>d.name));
    const ds=[...new Set([...fromTeams,...fromState])].filter(n=>list.some(d=>d.name===n));
    return {manager:k,designers:ds.map(n=>{const d=stateDesignerMap[n]||{name:n,projects:[]};return {name:n,projects:d.projects.length,overloaded:getFlags(d).length>0};})};
  }).filter(g=>g.designers.length>0);}
  function mkPG(list){return allManagerKeys.map(k=>{const ds=(teams[k]?teams[k].designers||[]:MANAGER_DESIGNERS[k]||[]).map(n=>stateDesignerMap[n]).filter(Boolean);return {manager:k,projects:ds.flatMap(d=>d.projects.map(p=>({...p,designer:d.name}))).filter(p=>list.some(lp=>lp.id===p.id))};}).filter(g=>g.projects.length>0);}

  const blockedProjects = allProjects.filter(p => p.blocker && p.blocker.trim());
  const projectDesignerMap = {};
  allDesigners.forEach(d => d.projects.forEach(p => { projectDesignerMap[p.id] = d.name; }));
  const utilColor=utilPct>=80?"#C62828":utilPct>=60?"#0F172A":"#E65100";

  return (
    <div style={{padding:"24px 40px",maxWidth:1200,margin:"0 auto"}}>
      <div style={{fontSize:16,fontWeight:700,color:"#0F172A",fontFamily:"'Inter',sans-serif",letterSpacing:"-0.01em",marginBottom:16}}>Analytics</div>
      <div style={{display:"flex",borderBottom:"1px solid #E0E0E0",marginBottom:28}}>
        {[{id:"summary",label:"Executive Summary"},{id:"capacity",label:"Capacity"}].map(t=>(
          <button key={t.id} onClick={()=>setTab(t.id)} style={{position:"relative",height:48,padding:"0 24px",border:"none",background:"transparent",color:tab===t.id?"#2563EB":"#607D8B",fontSize:14,fontWeight:tab===t.id?600:500,cursor:"pointer",fontFamily:"'Inter',sans-serif",letterSpacing:"0.01em",transition:"color 0.15s"}}>
            {t.label}
            {tab===t.id&&<span style={{position:"absolute",bottom:0,left:0,right:0,height:3,background:"#2563EB",borderRadius:"3px 3px 0 0"}}/>}
          </button>
        ))}
      </div>
      {tab==="summary"&&(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(6,1fr)",gap:16,marginBottom:24}}>
            <KpiCard label="Active Projects" value={activeProjects.length} color="#1565C0" tip="Projects currently in progress today." onClick={()=>setDrawer({title:"Active Projects",type:"projects",dateKey:"endDate",groups:mkPG(activeProjects)})}/>
            <KpiCard label="Upcoming (28d)" value={upcomingProjects.length} color="#0288D1" tip="Projects starting in the next 28 days." onClick={()=>setDrawer({title:"Upcoming Projects",type:"projects",dateKey:"startDate",groups:mkPG(upcomingProjects)})}/>
            <KpiCard label="Ending Soon (28d)" value={endingSoon.length} color="#E65100" tip="Projects ending in the next 28 days." onClick={()=>setDrawer({title:"Ending Soon",type:"projects",dateKey:"endDate",groups:mkPG(endingSoon)})}/>
            <KpiCard label="Overloaded" value={overloadedAll.length} color={overloadedAll.length>0?"#C62828":"#2E7D32"} tip="Designers with capacity score ≥ 6." onClick={()=>setDrawer({title:"Overloaded Designers",type:"designers",groups:mkDG(overloadedAll)})}/>
            <KpiCard label="Available" value={availableDesigners.length} color={availableDesigners.length>0?"#388E3C":"#90A4AE"} tip="Designers with no active projects." onClick={()=>setDrawer({title:"Available Now",type:"designers",groups:mkDG(availableDesigners)})}/>
            <KpiCard label="Flags" value={blockedProjects.length} color={blockedProjects.length>0?"#EF5350":"#2E7D32"} tip="Projects with an active flag/blocker raised by a designer." onClick={blockedProjects.length>0?()=>setDrawer({title:"Flagged Projects",type:"flags",dateKey:"endDate",groups:mkPG(blockedProjects)}):undefined}/>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(2,1fr)",gap:16,marginBottom:24}}>
            <KpiCard label="Total Designers" value={allDesigners.length} color="#0F172A" tip="All designers across all managers." sub={`${availableDesigners.length} available · ${overloadedAll.length} overloaded`}/>
            <KpiCard label="Utilization" value={`${utilPct}%`} color={utilColor} tip="Percentage of designer-weeks occupied by projects this quarter." sub={`${occupiedWeeks} of ${totalDesignerWeeks} designer-weeks`}/>
          </div>
          <div style={{background:"#fff",borderRadius:16,border:"1px solid #E8EAED",boxShadow:"0 1px 3px rgba(0,0,0,0.06)",overflow:"hidden",marginBottom:24}}>
            <div style={{padding:"16px 20px",borderBottom:"1px solid #F5F5F5"}}><div style={{fontSize:14,fontWeight:600,color:"#0F172A"}}>Manager Scorecard</div></div>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,fontFamily:"'Inter',sans-serif"}}>
              <thead><tr style={{background:"#FAFAFA",borderBottom:"1px solid #ECEFF1"}}>{["MANAGER","DESIGNERS","PROJECTS","OVERLOADED"].map(h=><th key={h} style={{padding:"8px 16px",textAlign:h==="MANAGER"?"left":"center",fontWeight:700,color:"#90A4AE",fontSize:11}}>{h}</th>)}</tr></thead>
              <tbody>{[...byManager].sort((a,b)=>a.name.localeCompare(b.name)).map(m=>(
                <tr key={m.name} style={{borderBottom:"1px solid #F5F5F5",cursor:"pointer"}} onClick={()=>{setManager(m.name);setPage("timeline");}} onMouseEnter={e=>e.currentTarget.style.background="#FAFAFA"} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                  <td style={{padding:"10px 16px",fontWeight:700,color:"#0F172A"}}>{m.name}</td>
                  <td style={{padding:"10px 12px",textAlign:"center",color:"#546E7A"}}>{m.total}</td>
                  <td style={{padding:"10px 12px",textAlign:"center",color:"#546E7A"}}>{m.projects}</td>
                  <td style={{padding:"10px 12px",textAlign:"center"}}>{m.overloaded>0?<span style={{fontWeight:700,color:"#C62828",background:"#FFEBEE",padding:"2px 8px",borderRadius:6}}>{m.overloaded}</span>:<span style={{color:"#B0BEC5"}}>0</span>}</td>
                </tr>
              ))}</tbody>
            </table>
          </div>
          {flags.length>0&&<div style={{background:"#fff",borderRadius:16,border:"1px solid #E8EAED",boxShadow:"0 1px 3px rgba(0,0,0,0.06)",padding:"16px 20px"}}><div style={{fontSize:14,fontWeight:600,color:"#0F172A",marginBottom:12}}>Attention Needed</div><div style={{display:"flex",flexDirection:"column",gap:6}}>{flags.slice(0,8).map((fl,i)=><div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 10px",borderRadius:6,background:fl.severity==="red"?"#FFEBEE":"#E8F5E9"}}><div style={{width:6,height:6,borderRadius:"50%",background:fl.severity==="red"?"#C62828":"#2E7D32",flexShrink:0}}/><div style={{fontSize:12,color:"#0F172A"}}>{fl.text}</div></div>)}</div></div>}
        </div>
      )}
      {tab==="capacity"&&(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:16,marginBottom:32}}>
            <KpiCard label="Total Designers" value={allDesigners.length} color="#1565C0" tip="All designers across all managers." onClick={()=>setDrawer({title:"Total Designers",type:"designers",groups:mkDG(allDesigners)})}/>
            <KpiCard label="Active Projects" value={activeProjects.length} color="#388E3C" tip="Projects in progress today." onClick={()=>setDrawer({title:"Active Projects",type:"projects",dateKey:"endDate",groups:mkPG(activeProjects)})}/>
            <KpiCard label="Available Now" value={availableDesigners.length} color={availableDesigners.length>0?"#388E3C":"#90A4AE"} tip="Designers with no active projects." onClick={()=>setDrawer({title:"Available Now",type:"designers",groups:mkDG(availableDesigners)})}/>
            <KpiCard label="Overloaded" value={overloadedAll.length} color={overloadedAll.length>0?"#C62828":"#2E7D32"} tip="Designers with 2+ active projects scoring ≥ 6." onClick={()=>setDrawer({title:"Overloaded Designers",type:"designers",groups:mkDG(overloadedAll)})}/>
          </div>
          <div style={{background:"#fff",borderRadius:16,padding:"20px 24px",border:"1px solid #E8EAED",boxShadow:"0 1px 3px rgba(0,0,0,0.06)",marginBottom:32}}>
            <div style={{fontSize:14,fontWeight:600,color:"#0F172A",marginBottom:12}}>Team capacity overview</div>
            {(()=>{
              const occDs=allDesigners.filter(d=>d.projects.some(p=>parseD(p.startDate)<=today&&parseD(p.endDate)>=today)&&getFlags(d).length===0);
              const segs=[{key:"over",label:"Overloaded",list:overloadedAll,count:overloadedAll.length,color:"#EF5350",title:"Overloaded"},{key:"occ",label:"Occupied",list:occDs,count:occDs.length,color:"#3B82F6",title:"Occupied"},{key:"avail",label:"Available",list:availableDesigners,count:availableDesigners.length,color:"#22C55E",title:"Available"}];
              return(<>
                <div style={{display:"flex",height:20,borderRadius:6,overflow:"hidden",marginBottom:12}}>{segs.map(s=>s.count>0?<div key={s.key} onClick={()=>setDrawer({title:s.title,type:"designers",groups:mkDG(s.list)})} style={{flex:s.count,background:s.color,cursor:"pointer"}} onMouseEnter={e=>e.currentTarget.style.opacity="0.8"} onMouseLeave={e=>e.currentTarget.style.opacity="1"}/>:null)}</div>
                <div style={{display:"flex",gap:20}}>{segs.map(s=><div key={s.key} onClick={()=>setDrawer({title:s.title,type:"designers",groups:mkDG(s.list)})} style={{display:"flex",alignItems:"center",gap:6,cursor:"pointer"}}><div style={{width:10,height:10,borderRadius:2,background:s.color}}/><span style={{fontSize:12,color:"#607D8B"}}>{s.label}</span><span style={{fontSize:12,fontWeight:700,color:"#0F172A"}}>{s.count}</span></div>)}</div>
              </>);
            })()}
          </div>
          {blockedProjects.length>0&&(
            <div style={{background:"#fff",borderRadius:16,padding:"20px 24px",border:"1px solid #E8EAED",boxShadow:"0 1px 3px rgba(0,0,0,0.06)",marginBottom:32}}>
              <div style={{fontSize:14,fontWeight:600,color:"#0F172A",marginBottom:16,display:"flex",alignItems:"center",gap:6}}>
                <MI name="flag" size={16} style={{color:"#EF5350"}}/>
                Flagged Projects <span style={{fontSize:12,fontWeight:500,color:"#9CA3AF",marginLeft:4}}>{blockedProjects.length} project{blockedProjects.length!==1?"s":""}</span>
              </div>
              {blockedProjects.map(p=>(
                <div key={p.id} style={{display:"flex",alignItems:"flex-start",gap:10,padding:"10px 0",borderBottom:"1px solid #F5F5F5"}}>
                  <MI name="flag" size={14} style={{color:"#EF5350",flexShrink:0,marginTop:2}}/>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:2}}>
                      <span style={{fontSize:13,fontWeight:600,color:"#0F172A"}}>{p.name}</span>
                      {projectDesignerMap[p.id]&&<span style={{fontSize:11,fontWeight:500,color:"#6B7280",background:"#F3F4F6",borderRadius:20,padding:"1px 8px"}}>{projectDesignerMap[p.id]}</span>}
                    </div>
                    <div style={{fontSize:12,color:"#EF5350",lineHeight:1.4}}>{p.blocker}</div>
                  </div>
                  <span style={{fontSize:11,color:"#9CA3AF",flexShrink:0,fontFamily:"monospace"}}>{fmtS(p.endDate)}</span>
                </div>
              ))}
            </div>
          )}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:32}}>
            {[{title:"Starting in next 4 weeks",projects:upcomingProjects,dateKey:"startDate"},{title:"Finishing in next 4 weeks",projects:endingSoon,dateKey:"endDate"}].map(item=>(
              <div key={item.title} style={{background:"#fff",borderRadius:16,padding:"20px 24px",border:"1px solid #E8EAED",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
                  <div style={{fontSize:14,fontWeight:700,color:"#0F172A"}}>{item.title} <span style={{fontSize:11,fontWeight:500,color:"#90A4AE"}}>{item.projects.length} projects</span></div>
                  {item.projects.length>0&&<button onClick={()=>setDrawer({title:item.title,type:"projects",dateKey:item.dateKey,groups:mkPG(item.projects)})} style={{fontSize:11,fontWeight:700,color:"#2563EB",background:"none",border:"none",cursor:"pointer"}}>View all</button>}
                </div>
                {item.projects.length===0&&<div style={{fontSize:12,color:"#B0BEC5"}}>None</div>}
                {item.projects.slice(0,5).map(p=><div key={p.id} style={{display:"flex",justifyContent:"space-between",padding:"6px 0",borderBottom:"1px solid #F5F5F5"}}><span style={{fontSize:12,color:"#0F172A",fontWeight:600}}>{p.name.length>28?p.name.slice(0,28)+"...":p.name}</span><span style={{fontSize:11,color:"#90A4AE",fontFamily:"monospace"}}>{fmtS(p[item.dateKey])}</span></div>)}
                {item.projects.length>5&&<div style={{fontSize:11,color:"#90A4AE",marginTop:8}}>+{item.projects.length-5} more</div>}
              </div>
            ))}
          </div>
          <div style={{background:"#fff",borderRadius:16,padding:"20px 24px",border:"1px solid #E8EAED",boxShadow:"0 1px 3px rgba(0,0,0,0.06)",marginBottom:32}}>
            <div style={{fontSize:14,fontWeight:600,color:"#0F172A",marginBottom:16}}>Project size breakdown <span style={{fontSize:11,fontWeight:500,color:"#90A4AE"}}>{allProjects.length} total</span></div>
            <div style={{display:"flex",gap:16,alignItems:"flex-end",height:80}}>
              {Object.entries(sizeCount).map(([size,count])=>{
                const max=Math.max(...Object.values(sizeCount),1);
                const h=Math.max(4,Math.round(count/max*64));
                const colors={XS:"#B0BEC5",S:"#93C5FD",M:"#60A5FA",L:"#3B82F6",XL:"#1565C0"};
                const sizeProjects=allProjects.filter(p=>p.size===size);
                return(
                  <div key={size} onClick={count>0?()=>setDrawer({title:size+" Projects",type:"projects",dateKey:"endDate",groups:mkPG(sizeProjects)}):undefined}
                    style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",gap:4,cursor:count>0?"pointer":"default",borderRadius:6,padding:"4px 2px"}}
                    onMouseEnter={e=>{if(count>0)e.currentTarget.style.background="#F5F5F5";}} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
                    <div style={{fontSize:11,fontWeight:700,color:count>0?"#0F172A":"#B0BEC5"}}>{count}</div>
                    <div style={{width:"100%",height:h,background:colors[size],borderRadius:4,opacity:count>0?1:0.4}}/>
                    <div style={{fontSize:11,fontWeight:700,color:"#90A4AE",fontFamily:"monospace"}}>{size}</div>
                  </div>
                );
              })}
            </div>
          </div>
          <div style={{background:"#fff",borderRadius:16,padding:"20px 24px",border:"1px solid #E8EAED",boxShadow:"0 1px 3px rgba(0,0,0,0.06)"}}>
            <div style={{fontSize:14,fontWeight:600,color:"#0F172A",marginBottom:16}}>By manager</div>
            <div style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr 1fr",gap:8,marginBottom:8}}>{["Manager","Designers","Active","Overloaded","Projects"].map(h=><div key={h} style={{fontSize:11,fontWeight:700,color:"#90A4AE",textTransform:"uppercase",letterSpacing:"0.04em"}}>{h}</div>)}</div>
            {[...byManager].sort((a,b)=>a.name.localeCompare(b.name)).map(row=>(
              <div key={row.name} style={{display:"grid",gridTemplateColumns:"2fr 1fr 1fr 1fr 1fr",gap:8,padding:"10px 0",borderTop:"1px solid #F5F5F5",alignItems:"center"}}>
                <div onClick={()=>{setManager(row.name);setPage("timeline");}} style={{fontSize:14,fontWeight:700,color:"#2563EB",cursor:"pointer"}}>{row.name}</div>
                <div style={{fontSize:14,color:"#546E7A"}}>{row.total}</div>
                <div style={{fontSize:14,color:"#546E7A"}}>{row.active}</div>
                <div style={{fontSize:14,color:row.overloaded>0?"#C62828":"#2E7D32",fontWeight:row.overloaded>0?700:400}}>{row.overloaded}</div>
                <div style={{fontSize:14,color:"#546E7A"}}>{row.projects}</div>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// ---------- MultiSelectPill (MD3 Filter Chip + Menu) ----------
function MultiSelectPill({ label, options, selected, onChange }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  useEffect(() => {
    if (!open) return;
    const h = e => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", h);
    return () => document.removeEventListener("mousedown", h);
  }, [open]);
  const allSelected = !selected || selected.length === options.length;
  const count = allSelected ? options.length : (selected ? selected.length : 0);
  const isActive = !allSelected;
  const toggle = opt => {
    if (allSelected) {
      onChange(options.filter(o => o !== opt));
    } else {
      const next = selected.includes(opt) ? selected.filter(o => o !== opt) : [...selected, opt];
      onChange(next.length === options.length ? null : (next.length === 0 ? [] : next));
    }
  };
  const toggleAll = () => { onChange(allSelected ? [] : null); };
  return (
    <div ref={ref} style={{ position:"relative", flexShrink:0 }}>
      {/* MD3 Filter Chip: 32dp height, 8dp corner radius, outline border when unselected */}
      <button onClick={() => setOpen(o => !o)}
        style={{
          display:"flex", alignItems:"center", gap:4,
          height:32, padding: isActive ? "0 12px 0 8px" : "0 12px",
          borderRadius:8,
          border: isActive ? "none" : "1px solid #79747E",
          background: isActive ? "#EEF4FD" : "transparent",
          color: isActive ? "#2563EB" : "#49454F",
          fontSize:14, fontWeight:500, cursor:"pointer",
          fontFamily:"'Inter',sans-serif", transition:"background 0.15s, border 0.15s",
        }}>
        {/* Leading check — MD3 shows it only in selected state */}
        {isActive && <MI name="check" size={18} style={{ display:"flex", flexShrink:0 }} />}
        {label}
        {isActive && <span style={{ fontSize:13, fontWeight:600 }}>({count})</span>}
        {/* Trailing dropdown arrow */}
        <MI name={open ? "arrow_drop_up" : "arrow_drop_down"} size={18} style={{ display:"flex", flexShrink:0 }} />
      </button>

      {open && (
        /* MD3 Menu: 4dp corner radius, elevation-2 shadow, 8dp padding top/bottom */
        <div style={{
          position:"absolute", top:"calc(100% + 4px)", left:0,
          background:"#FFFBFE",
          borderRadius:4,
          boxShadow:"0 1px 2px rgba(0,0,0,0.30), 0 2px 6px 2px rgba(0,0,0,0.15)",
          zIndex:9999, minWidth:220, maxHeight:320, overflowY:"auto",
          padding:"8px 0",
        }}>
          {/* Select All — separated by a divider */}
          <div onClick={toggleAll}
            onMouseEnter={e => e.currentTarget.style.background = "rgba(37,99,235,0.08)"}
            onMouseLeave={e => e.currentTarget.style.background = "transparent"}
            style={{ display:"flex", alignItems:"center", gap:12, padding:"0 16px", height:48, cursor:"pointer", fontFamily:"'Inter',sans-serif" }}>
            {/* MD3 checkbox: 18dp, 2dp corner radius */}
            <div style={{ width:18, height:18, borderRadius:2, border:`2px solid ${allSelected?"#2563EB":"#49454F"}`, background:allSelected?"#2563EB":"transparent", display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0, transition:"all 0.1s" }}>
              {allSelected && <MI name="check" size={12} style={{ color:"#fff", display:"flex" }} />}
            </div>
            <span style={{ fontSize:14, fontWeight:600, color:"#1C1B1F" }}>Select All</span>
          </div>
          {/* MD3 divider */}
          <div style={{ height:1, background:"#CAC4D0", margin:"4px 0" }} />
          {options.map(opt => {
            const checked = allSelected || (selected && selected.includes(opt));
            return (
              <div key={opt} onClick={() => toggle(opt)}
                onMouseEnter={e => e.currentTarget.style.background = "rgba(37,99,235,0.08)"}
                onMouseLeave={e => e.currentTarget.style.background = "transparent"}
                style={{ display:"flex", alignItems:"center", gap:12, padding:"0 16px", height:48, cursor:"pointer", fontFamily:"'Inter',sans-serif" }}>
                <div style={{ width:18, height:18, borderRadius:2, border:`2px solid ${checked?"#2563EB":"#49454F"}`, background:checked?"#2563EB":"transparent", display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0, transition:"all 0.1s" }}>
                  {checked && <MI name="check" size={12} style={{ color:"#fff", display:"flex" }} />}
                </div>
                <span style={{ fontSize:14, color:"#1C1B1F" }}>{opt}</span>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ---------- Drawer ----------
function Drawer({drawer,setDrawer,drawerSearch,setDrawerSearch,drawerSort,setDrawerSort,drawerManagers,setDrawerManagers}) {
  if(!drawer)return null;
  let groups=drawer.groups?(drawerManagers.length>0?drawer.groups.filter(g=>drawerManagers.includes(g.manager)):drawer.groups):null;
  if(groups&&drawerSearch)groups=groups.map(g=>({...g,designers:g.designers?.filter(d=>d.name.toLowerCase().includes(drawerSearch.toLowerCase())),projects:g.projects?.filter(p=>p.name.toLowerCase().includes(drawerSearch.toLowerCase()))})).filter(g=>(g.designers||g.projects||[]).length>0);
  if(groups&&drawerSort)groups=groups.map(g=>({...g,designers:g.designers?[...g.designers].sort((a,b)=>a.name.localeCompare(b.name)):undefined,projects:g.projects?[...g.projects].sort((a,b)=>a.name.localeCompare(b.name)):undefined})).sort((a,b)=>a.manager.localeCompare(b.manager));
  const isDesigners=drawer.type==="designers";
  const isFlags=drawer.type==="flags";
  const totalCount=groups?groups.reduce((a,g)=>a+(isDesigners?(g.designers||[]).length:(g.projects||[]).length),0):0;
  return (
    <div style={{position:"fixed",inset:0,zIndex:10000,display:"flex"}}>
      <div style={{flex:1,background:"rgba(0,0,0,0.3)"}} onClick={()=>setDrawer(null)}/>
      <div style={{width:520,background:"#fff",height:"100%",overflowY:"hidden",boxShadow:"-2px 0 16px rgba(0,0,0,0.10)",display:"flex",flexDirection:"column"}}>
        <div style={{padding:"24px 24px 16px",borderBottom:"1px solid #F5F5F5",display:"flex",justifyContent:"space-between",alignItems:"center",flexShrink:0}}>
          <div><div style={{fontSize:16,fontWeight:700,color:"#0F172A",fontFamily:"'Inter',sans-serif"}}>{drawer.title}</div><div style={{fontSize:12,color:"#90A4AE",marginTop:2}}>{totalCount} {isDesigners?"designer"+(totalCount!==1?"s":""):"project"+(totalCount!==1?"s":"")}</div></div>
          <button onClick={()=>setDrawer(null)} style={{background:"transparent",border:"none",cursor:"pointer",color:"#90A4AE",display:"flex",borderRadius:"50%",padding:4}}><MI name="close" size={16} /></button>
        </div>
        <div style={{padding:"10px 16px",borderBottom:"1px solid #F5F5F5",display:"flex",gap:8,flexShrink:0,flexWrap:"wrap"}}>
          <div style={{flex:1,display:"flex",alignItems:"center",gap:8,padding:"7px 10px",border:"1px solid #ECEFF1",borderRadius:4,background:"#FAFAFA",minWidth:140}}>
            <MI name="search" size={16} style={{color:"#90A4AE"}} />
            <input value={drawerSearch} onChange={e=>setDrawerSearch(e.target.value)} placeholder="Search..." style={{border:"none",outline:"none",fontSize:12,color:"#0F172A",background:"transparent",width:"100%",fontFamily:"'Inter',sans-serif"}}/>
            {drawerSearch&&<span onClick={()=>setDrawerSearch("")} style={{cursor:"pointer",color:"#B0BEC5",display:"flex"}}><MI name="close" size={16} /></span>}
          </div>
          <button onClick={()=>setDrawerSort(s=>!s)} style={{display:"flex",alignItems:"center",gap:4,padding:"7px 12px",borderRadius:100,border:"1px solid"+(drawerSort?"#2563EB":"#ECEFF1"),background:drawerSort?"#2563EB":"#FAFAFA",color:drawerSort?"#fff":"#607D8B",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"'Inter',sans-serif"}}>A→Z</button>
          {groups&&<div style={{display:"flex",flexWrap:"wrap",gap:4,width:"100%"}}>
            {drawer.groups.map(g=><button key={g.manager} onClick={()=>setDrawerManagers(m=>m.includes(g.manager)?m.filter(x=>x!==g.manager):[...m,g.manager])} style={{padding:"3px 10px",borderRadius:100,fontSize:11,fontWeight:600,cursor:"pointer",background:drawerManagers.includes(g.manager)?"#2563EB":"transparent",color:drawerManagers.includes(g.manager)?"#fff":"#607D8B",border:drawerManagers.includes(g.manager)?"none":"1px solid #D1D5DB",fontFamily:"'Inter',sans-serif"}}>{g.manager}</button>)}
            {drawerManagers.length>0&&<button onClick={()=>setDrawerManagers([])} style={{padding:"3px 10px",borderRadius:100,fontSize:11,fontWeight:600,cursor:"pointer",background:"transparent",color:"#90A4AE",border:"1px solid #ECEFF1",fontFamily:"'Inter',sans-serif"}}>Clear</button>}
          </div>}
        </div>
        <div style={{padding:16,flex:1,overflowY:"auto"}}>
          {groups?groups.map(({manager,designers,projects})=>(
            <div key={manager} style={{marginBottom:20}}>
              <div style={{fontSize:11,fontWeight:500,color:"#90A4AE",textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:8,paddingBottom:6,borderBottom:"1px solid #F5F5F5"}}>{manager}</div>
              {isDesigners?(designers||[]).map((d,i)=>(
                <div key={i} style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"10px 12px",borderRadius:8,border:"1px solid #ECEFF1",marginBottom:6,background:"#FAFAFA"}}>
                  <div style={{display:"flex",alignItems:"center",gap:10}}>
                    <div style={{width:30,height:30,borderRadius:"50%",background:"#F5F5F5",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,color:"#546E7A",flexShrink:0}}>{d.name.slice(0,2).toUpperCase()}</div>
                    <div><div style={{fontSize:14,fontWeight:700,color:"#0F172A"}}>{d.name}</div><div style={{fontSize:11,color:"#90A4AE"}}>{d.projects} project{d.projects!==1?"s":""}</div></div>
                  </div>
                  {d.overloaded&&<span style={{fontSize:11,fontWeight:700,color:"#C62828",background:"#FFEBEE",border:"1px solid #FFCDD2",borderRadius:6,padding:"2px 8px"}}>Overloaded</span>}
                </div>
              )):(projects||[]).map(p=>(
                <div key={p.id} style={{padding:"12px 16px",borderRadius:8,border:`1px solid ${p.blocker?"#FED7D7":"#ECEFF1"}`,marginBottom:6,background:p.blocker?"#FFF5F5":"#FAFAFA"}}>
                  <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:6,flexWrap:"wrap"}}>
                    {p.blocker&&<MI name="flag" size={13} style={{color:"#EF5350",flexShrink:0}}/>}
                    <div style={{fontSize:14,fontWeight:700,color:"#0F172A"}}>{p.name}</div>
                    {p.designer&&<span style={{fontSize:11,fontWeight:500,color:"#6B7280",background:"#F3F4F6",borderRadius:20,padding:"1px 8px",flexShrink:0}}>{p.designer}</span>}
                    {p.type&&(()=>{const pt=PROJECT_TYPES.find(t=>t.label===p.type);return pt?<span style={{fontSize:10,fontWeight:600,padding:"1px 8px",borderRadius:20,background:pt.bg,color:pt.color,border:`1px solid ${pt.color}33`,flexShrink:0}}>{pt.label}</span>:null;})()}
                  </div>
                  {p.blocker&&(
                    <div style={{fontSize:12,color:"#C53030",lineHeight:1.5,marginBottom:8,padding:"6px 10px",borderRadius:6,background:"rgba(254,215,215,0.3)",border:"1px solid #FED7D7"}}>{p.blocker}</div>
                  )}
                  <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
                    <span title="End date" style={{fontSize:11,color:"#607D8B",display:"flex",alignItems:"center",gap:3,cursor:"default"}}><MI name="event" size={13}/>{fmtM(p.endDate)}</span>
                    {p.squad&&<span title="Squad" style={{fontSize:11,color:"#607D8B",display:"flex",alignItems:"center",gap:3,cursor:"default"}}><MI name="workspaces" size={13}/>{p.squad}</span>}
                    <span title="Size" style={{fontSize:11,fontWeight:700,padding:"1px 7px",borderRadius:3,background:"#F5F5F5",color:"#546E7A",fontFamily:"monospace",cursor:"default"}}>{p.size}</span>
                    {p.pms?.length>0&&<span title={`PM: ${p.pms.join(", ")}`} style={{fontSize:11,color:"#607D8B",display:"flex",alignItems:"center",gap:3,cursor:"default"}}><MI name="person" size={13}/>{p.pms.join(", ")}</span>}
                    {p.figmaUrl&&<a href={p.figmaUrl} target="_blank" rel="noopener noreferrer" title="Open in Figma" style={{display:"flex",alignItems:"center",color:"#2563EB",textDecoration:"none",flexShrink:0}}><MI name="brush" size={15}/></a>}
                    {p.jiraUrl&&<a href={p.jiraUrl} target="_blank" rel="noopener noreferrer" title="Open in Jira" style={{display:"flex",alignItems:"center",color:"#2563EB",textDecoration:"none",flexShrink:0}}><MI name="link" size={15}/></a>}
                  </div>
                </div>
              ))}
            </div>
          )):null}
          {totalCount===0&&<div style={{textAlign:"center",padding:"40px 0",color:"#B0BEC5",fontSize:14}}>No results</div>}
        </div>
      </div>
    </div>
  );
}


// ---------- INIT_TEAMS ----------
const INIT_TEAMS = {
  Ashot:     { designers: MANAGER_DESIGNERS.Ashot,     squads: DEF_SQUADS, pms: DEF_PMS },
  Annie:     { designers: MANAGER_DESIGNERS.Annie,     squads: [], pms: [] },
  Kim:       { designers: MANAGER_DESIGNERS.Kim,       squads: [], pms: [] },
  Carson:    { designers: MANAGER_DESIGNERS.Carson,    squads: [], pms: [] },
  Christine: { designers: MANAGER_DESIGNERS.Christine, squads: [], pms: [] },
  Ian:       { designers: MANAGER_DESIGNERS.Ian,       squads: [], pms: [] },
  Michael:   { designers: MANAGER_DESIGNERS.Michael,   squads: [], pms: [] },
  Debbie:    { designers: MANAGER_DESIGNERS.Debbie,    squads: [], pms: [] },
  Rudy:      { designers: MANAGER_DESIGNERS.Rudy,      squads: [], pms: [] },
};

// ---------- Role map ----------
function buildUserMap() {
  const m = {};
  ["Ashot"].forEach(n => (m[n] = "vp"));
  MANAGERS.forEach(n => { if (!m[n]) m[n] = "manager"; });
  Object.values(SUB_MANAGERS).flat().forEach(n => { if (!m[n]) m[n] = "manager"; });
  return m;
}
const USER_MAP = buildUserMap();

// ---------- LoginScreen ----------
function LoginScreen({ onLogin }) {
  const [role, setRole] = useState("");
  const [name, setName] = useState("");
  const managers = MANAGERS.flatMap(m => [
    m,
    ...(SUB_MANAGERS[m] || []).map(sub => ({ value: sub, label: `↳ ${sub}`, triggerLabel: sub })),
  ]);
  const allManagerKeys = MANAGERS.flatMap(m => [m, ...(SUB_MANAGERS[m] || [])]);
  const designerOptions = allManagerKeys.flatMap(mgr => {
    const designers = MANAGER_DESIGNERS[mgr] || [];
    if (!designers.length) return [];
    return [{ isGroup: true, label: mgr }, ...designers];
  });
  const options = role === "Manager" ? managers : role === "Designer" ? designerOptions : [];
  const canEnter = role === "VP" || (role && name);
  const submit = () => { if (canEnter) onLogin(role === "VP" ? "VP" : name, role); };
  return (
    <div style={{ minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center", background:"linear-gradient(135deg,#FAFAFA 0%,#EEF2F7 100%)", fontFamily:"'Inter',sans-serif" }}>
      <style>{"@keyframes fadeIn{from{opacity:0;transform:translateY(-6px)}to{opacity:1;transform:translateY(0)}}"}</style>
      <div style={{ background:"#fff", borderRadius:28, padding:"48px 40px", boxShadow:"0 4px 24px rgba(0,0,0,0.10)", width:400, maxWidth:"90vw" }}>
        <div style={{ textAlign:"center", marginBottom:32 }}>
          <img src="/XDlogo.svg" alt="XD Management Hub" style={{ width:88, height:88, objectFit:"contain", margin:"0 auto 16px", display:"block" }} />
          <div style={{ fontSize:24, fontWeight:800, color:"#0F172A", marginBottom:6, letterSpacing:"-0.01em" }}>XD Management Hub</div>
          <div style={{ fontSize:14, color:"#90A4AE" }}>ServiceTitan Design Team</div>
        </div>
        <div style={{ fontSize:11, fontWeight:700, color:"#90A4AE", textTransform:"uppercase", letterSpacing:"0.08em", marginBottom:10 }}>I am a</div>
        <div style={{ display:"flex", gap:10, marginBottom:24 }}>
          {["VP","Manager","Designer"].map(r => (
            <button key={r} onClick={() => { setRole(r); setName(""); }} style={{ flex:1, padding:"12px 0", borderRadius:12, border:"none", background:role===r?"#2563EB":"#F5F5F5", color:role===r?"#fff":"#607D8B", fontSize:14, fontWeight:500, cursor:"pointer", fontFamily:"'Inter',sans-serif", transition:"all 0.15s", boxShadow:role===r?"0 1px 2px rgba(0,0,0,0.3)":"none" }}>
              {r}
            </button>
          ))}
        </div>
        {role === "VP" && (
          <div style={{ animation:"fadeIn 0.2s ease" }}>
            <button onClick={submit} style={{ width:"100%", padding:"12px 0", borderRadius:20, border:"none", background:"#2563EB", color:"#fff", fontSize:14, fontWeight:500, cursor:"pointer", fontFamily:"'Inter',sans-serif", transition:"all 0.15s", boxShadow:"0 1px 2px rgba(0,0,0,0.3)", display:"flex", alignItems:"center", justifyContent:"center", gap:4 }}>
              Enter Hub<MI name="arrow_forward" size={16} style={{marginLeft:4}} />
            </button>
          </div>
        )}
        {(role === "Manager" || role === "Designer") && (
          <div style={{ animation:"fadeIn 0.2s ease" }}>
            <div style={{ marginBottom:16 }}>
              <CustomSelect value={name} onChange={setName} options={options} placeholder={"Select your name..."} clearable={false} style={{ width:"100%" }} />
            </div>
            <button onClick={submit} disabled={!name} style={{ width:"100%", padding:"12px 0", borderRadius:20, border:"none", background:name?"#2563EB":"#ECEFF1", color:name?"#fff":"#90A4AE", fontSize:14, fontWeight:500, cursor:name?"pointer":"not-allowed", fontFamily:"'Inter',sans-serif", transition:"all 0.15s", boxShadow:name?"0 1px 2px rgba(0,0,0,0.3)":"none", display:"flex", alignItems:"center", justifyContent:"center", gap:4 }}>
              Enter Hub<MI name="arrow_forward" size={16} style={{marginLeft:4}} />
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

// ---------- App ----------
export default function App() {
  const [appReady, setAppReady] = useState(false);
  const [user, setUser]         = useState(null);
  const [loginRole, setLoginRole] = useState(null); // "Manager" | "Designer"
  const [state, setState]       = useState(INITIAL_STATE);
  const [teams, setTeams]       = useState(INIT_TEAMS);
  const [page, setPage]         = useState("timeline");
  const [manager, setManager]   = useState("Ashot");
  const [vpFilter, setVpFilter] = useState("all");
  const [vpMgrFilter, setVpMgrFilter] = useState(null);  // null = all selected
  const [vpSqFilter, setVpSqFilter]   = useState(null);
  const [vpDsnFilter, setVpDsnFilter] = useState(null);
  const [filter, setFilter]     = useState([]);
  const [toasts, setToasts]     = useState([]);
  const [drawer, setDrawer]     = useState(null);
  const [drawerSearch, setDrawerSearch]   = useState("");
  const [drawerSort, setDrawerSort]       = useState(false);
  const [drawerManagers, setDrawerManagers] = useState([]);
  const [showAdd, setShowAdd]   = useState(false);
  const [managerDropOpen, setManagerDropOpen] = useState(false);
  const managerDropRef = useRef(null);
  const [addPrefill, setAddPrefill] = useState(null);
  const [editProj, setEditProj] = useState(null);
  const [showSettings, setShowSettings] = useState(false);
  const [showOverload, setShowOverload] = useState(false);
  const lastLocalChange = useRef(0);

  // Reset VP sub-filters when switching between "all" and a specific manager
  useEffect(() => {
    setVpMgrFilter(null);
    setVpSqFilter(null);
    setVpDsnFilter(null);
  }, [vpFilter]);

  // Close manager dropdown on outside click
  useEffect(() => {
    if (!managerDropOpen) return;
    const h = e => { if (managerDropRef.current && !managerDropRef.current.contains(e.target)) setManagerDropOpen(false); };
    document.addEventListener("mousedown", h);
    return () => document.removeEventListener("mousedown", h);
  }, [managerDropOpen]);

  // ---- Load on mount ----
  useEffect(() => {
    (async () => {
      const ur = await storage.get("xdh-current-user");
      if (ur) setUser(ur.value);
      const rr = await storage.get("xdh-current-role");
      if (rr) {
        setLoginRole(rr.value);
        if (rr.value === "Manager" && ur) setManager(ur.value);
      }
      const sr = await storage.get("xdh-state-v7", true);
      if (sr) try { setState(JSON.parse(sr.value)); } catch (e) {}
      const tr = await storage.get("xdh-teams-v34", true);
      if (tr) try { setTeams(JSON.parse(tr.value)); } catch (e) {}
      setAppReady(true);
    })();
  }, []);

  // ---- Real-time Firebase listeners ----
  useEffect(() => {
    const unsubState = subscribeToKey("xdh-state-v7", val => {
      if (Date.now() - lastLocalChange.current < 5000) return;
      try {
        const f = JSON.parse(val);
        setState(p => JSON.stringify(p) !== JSON.stringify(f) ? f : p);
      } catch (e) {}
    });
    const unsubTeams = subscribeToKey("xdh-teams-v34", val => {
      if (Date.now() - lastLocalChange.current < 5000) return;
      try {
        const f = JSON.parse(val);
        setTeams(p => JSON.stringify(p) !== JSON.stringify(f) ? f : p);
      } catch (e) {}
    });
    return () => { unsubState(); unsubTeams(); };
  }, []);

  // ---- Helpers ----
  const saveState = useCallback(async ns => {
    lastLocalChange.current = Date.now();
    setState(ns);
    await storage.set("xdh-state-v7", JSON.stringify(ns), true);
  }, []);

  const saveTeams = useCallback(async nt => {
    setTeams(nt);
    await storage.set("xdh-teams-v34", JSON.stringify(nt), true);
  }, []);

  const addToast = (msg, type = "success") => {
    const id = Date.now();
    setToasts(t => [...t, { id, msg, type }]);
    setTimeout(() => setToasts(t => t.filter(x => x.id !== id)), 3500);
  };

  const handleLogin = async (name, role) => {
    await storage.set("xdh-current-user", name);
    await storage.set("xdh-current-role", role);
    setUser(name);
    setLoginRole(role);
    if (role === "Manager") setManager(name);
  };

  const handleLogout = async () => {
    await storage.delete("xdh-current-user");
    await storage.delete("xdh-current-role");
    setUser(null);
    setLoginRole(null);
  };

  // ---- Effective manager (used by all mutations) ----
  const isDesignerRole = loginRole === "Designer";
  const isVP = loginRole === "VP";
  const designerManager = isDesignerRole
    ? Object.keys(state.managers).find(k =>
        state.managers[k]?.squads?.some(sq => sq.designers?.some(d => d.name === user))
      ) || "Ashot"
    : null;
  const effectiveManager = isDesignerRole ? designerManager : manager;

  // ---- Project mutations ----
  const getTeam = m => {
    const t = teams[m] || {};
    const init = INIT_TEAMS[m] || {};
    return {
      ...init,
      ...t,
      designers: (t.designers && t.designers.length) ? t.designers : (init.designers || MANAGER_DESIGNERS[m] || []),
      squads:    (t.squads    && t.squads.length)    ? t.squads    : (init.squads    || []),
      pms:       (t.pms       && t.pms.length)       ? t.pms       : (init.pms       || []),
    };
  };

  const ensureDesignerInSquads = (squads, designerName, teamSettings) => {
    for (const sq of squads) {
      if (sq.designers.find(d => d.name === designerName)) return;
    }
    const dsq = (teamSettings?.designerSquads || {})[designerName] || [];
    let targetSq = squads.find(sq => dsq.includes(sq.name));
    if (!targetSq && squads.length > 0) targetSq = squads[0];
    if (targetSq) {
      targetSq.designers.push({ id:"d_"+Date.now()+"_"+designerName, name:designerName, avatar:designerName.slice(0,2).toUpperCase(), projects:[] });
      return;
    }
    // No squads exist — use a single neutral "Unassigned" bucket instead of per-designer squads
    let unassigned = squads.find(sq => sq.name === "Unassigned");
    if (!unassigned) {
      unassigned = { id:"sq_unassigned_"+Date.now(), name:"Unassigned", colorIdx:0, designers:[] };
      squads.push(unassigned);
    }
    unassigned.designers.push({ id:"d_"+Date.now()+"_"+designerName, name:designerName, avatar:designerName.slice(0,2).toUpperCase(), projects:[] });
  };

  const handleAdd = useCallback(async proj => {
    const dn = proj.designer || "";
    const ns = { ...state, managers: { ...state.managers } };
    const squads = JSON.parse(JSON.stringify(ns.managers[effectiveManager]?.squads || []));
    ensureDesignerInSquads(squads, dn, getTeam(manager));
    for (const sq of squads) {
      const d = sq.designers.find(x => x.name === dn);
      if (d) { const { designer: _d, ...rest } = proj; d.projects = [...d.projects, rest]; break; }
    }
    ns.managers[effectiveManager] = { squads };
    await saveState(ns);
    addToast("Project added");
  }, [state, manager, teams, saveState]);

  const handleAddBatch = useCallback(async items => {
    const ns = { ...state, managers: { ...state.managers } };
    const squads = JSON.parse(JSON.stringify(ns.managers[effectiveManager]?.squads || []));
    const teamSettings = getTeam(manager);
    items.forEach(item => {
      const dn = item.designer || ""; if (!dn) return;
      ensureDesignerInSquads(squads, dn, teamSettings);
      for (const sq of squads) {
        const d = sq.designers.find(x => x.name === dn);
        if (d) { d.projects = [...d.projects, { id:"p_"+Date.now()+"_"+Math.random().toString(36).slice(2), name:item.name, size:item.size||"M", startDate:item.startDate, endDate:item.endDate, squad:item.squad||"", pms:item.pms||[], status:item.status||"", type:item.type||"" }]; break; }
      }
    });
    ns.managers[effectiveManager] = { squads };
    await saveState(ns);
    // Auto-add any new PMs found in imported items
    const existingPms = teamSettings.pms || [];
    const newPms = [];
    items.forEach(item => {
      (item.pms||[]).forEach(pm => {
        if (pm && !existingPms.includes(pm) && !newPms.includes(pm)) newPms.push(pm);
      });
    });
    // Auto-add any new designers found in imported items
    const existingDesigners = teamSettings.designers || [];
    const newDesigners = [];
    items.forEach(item => {
      if (item.designer && !existingDesigners.includes(item.designer) && !newDesigners.includes(item.designer)) newDesigners.push(item.designer);
    });
    const base = teams[manager] || INIT_TEAMS[manager] || {};
    const updatedTeam = {
      ...base,
      ...(newPms.length ? { pms: [...existingPms, ...newPms] } : {}),
      ...(newDesigners.length ? { designers: [...existingDesigners, ...newDesigners] } : {}),
    };
    if (newPms.length || newDesigners.length) {
      await saveTeams({ ...teams, [manager]: updatedTeam });
    }
    addToast(`${items.length} project${items.length !== 1 ? "s" : ""} added${newDesigners.length ? ` · ${newDesigners.length} new designer${newDesigners.length !== 1 ? "s" : ""} added` : ""}${newPms.length ? ` · ${newPms.length} new PM${newPms.length !== 1 ? "s" : ""} added` : ""}`);
  }, [state, manager, teams, saveState, saveTeams]);

  // For VP: find which manager key owns a given project id
  const findManagerForProject = (projectId) => {
    return allManagerKeys.find(k =>
      state.managers[k]?.squads?.some(sq => sq.designers.some(d => d.projects.some(p => p.id === projectId)))
    ) || effectiveManager;
  };

  const handleUpdate = useCallback(async proj => {
    const ns = { ...state, managers: { ...state.managers } };
    const targetMgr = isVP ? findManagerForProject(proj.id) : effectiveManager;
    const squads = JSON.parse(JSON.stringify(ns.managers[targetMgr]?.squads || []));
    squads.forEach(sq => sq.designers.forEach(d => {
      const i = d.projects.findIndex(p => p.id === proj.id);
      if (i !== -1) d.projects[i] = proj;
    }));
    ns.managers[targetMgr] = { squads };
    await saveState(ns);
  }, [state, manager, isVP, saveState]);

  const handleRemove = useCallback(async id => {
    const ns = { ...state, managers: { ...state.managers } };
    const targetMgr = isVP ? findManagerForProject(id) : effectiveManager;
    const squads = JSON.parse(JSON.stringify(ns.managers[targetMgr]?.squads || []));
    squads.forEach(sq => sq.designers.forEach(d => { d.projects = d.projects.filter(p => p.id !== id); }));
    ns.managers[targetMgr] = { squads };
    await saveState(ns);
    addToast("Project removed", "info");
  }, [state, manager, isVP, saveState]);

  const handleInlineEdit = useCallback(async (id, fields) => {
    const ns = { ...state, managers: { ...state.managers } };
    const targetMgr = isVP ? findManagerForProject(id) : effectiveManager;
    const squads = JSON.parse(JSON.stringify(ns.managers[targetMgr]?.squads || []));
    squads.forEach(sq => sq.designers.forEach(d => {
      const i = d.projects.findIndex(p => p.id === id);
      if (i !== -1) d.projects[i] = { ...d.projects[i], ...fields };
    }));
    ns.managers[targetMgr] = { squads };
    await saveState(ns);
  }, [state, manager, isVP, saveState]);

  const handleEditSubmit = useCallback(async proj => {
    await handleUpdate(proj);
    setEditProj(null);
    addToast("Project updated");
  }, [handleUpdate]);

  const handleUpdateTitle = (designerName, title, designerSqs, emoji) => {
    const key = effectiveManager;
    const base = teams[key] || INIT_TEAMS[key] || {};
    const nt = {
      ...teams,
      [key]: {
        ...base,
        titles:        { ...(base.titles||{}),        [designerName]: title },
        designerSquads:{ ...(base.designerSquads||{}), [designerName]: designerSqs },
        emojis:        { ...(base.emojis||{}),         [designerName]: emoji },
      }
    };
    setTeams(nt);
    storage.set("xdh-teams-v34", JSON.stringify(nt), true).catch(() => {});
    addToast("Designer details updated");
  };

  const handleSaveSettings = useCallback(async settings => {
    const { removedSquadIds = [], ...teamSettings } = settings;
    const base = teams[manager] || INIT_TEAMS[manager] || {};
    const nt = { ...teams, [manager]: { ...base, ...teamSettings } };
    await saveTeams(nt);

    const ns = { ...state, managers: { ...state.managers } };
    const rawSquads = JSON.parse(JSON.stringify(ns.managers[effectiveManager]?.squads || []));

    // Add any new designers that aren't already in a squad
    const existingNames = new Set(rawSquads.flatMap(sq => sq.designers.map(d => d.name)));
    const newDesigners = (teamSettings.designers || []).filter(n => !existingNames.has(n));
    newDesigners.forEach(name => ensureDesignerInSquads(rawSquads, name, teamSettings));

    // Remove designers that were deleted from the team list
    const activeDesignerSet = new Set(teamSettings.designers || []);
    const removedDesigners = [...existingNames].filter(n => !activeDesignerSet.has(n));
    if (removedDesigners.length) {
      rawSquads.forEach(sq => {
        sq.designers = sq.designers.filter(d => activeDesignerSet.has(d.name));
      });
    }

    // Remove deleted squads, moving orphaned designers to the first remaining squad
    if (removedSquadIds.length) {
      const toRemove = rawSquads.filter(sq => removedSquadIds.includes(sq.id));
      const remaining = rawSquads.filter(sq => !removedSquadIds.includes(sq.id));
      const orphans = toRemove.flatMap(sq => sq.designers);
      if (orphans.length) {
        if (remaining.length > 0) {
          remaining[0].designers = [...remaining[0].designers, ...orphans];
        } else {
          remaining.push({ id:"sq_unassigned_"+Date.now(), name:"Unassigned", colorIdx:0, designers:orphans });
        }
      }
      ns.managers[effectiveManager] = { squads: remaining };
    } else if (newDesigners.length || removedDesigners.length) {
      ns.managers[effectiveManager] = { squads: rawSquads };
    }

    if (removedSquadIds.length || newDesigners.length || removedDesigners.length) await saveState(ns);
    addToast("Settings saved");
  }, [teams, manager, saveTeams, state, effectiveManager, saveState]);

  // ---- Derived ----
  const isDesigner = isDesignerRole;
  const isManager = !isDesigner;
  // True when the logged-in manager is viewing their own team (or a sub-manager's team they own).
  // Ashot is admin and can edit everything.
  const isOwnTeam = user === "Ashot" || !isManager || manager === user || (SUB_MANAGERS[user] || []).includes(manager);

  const currentTeam = getTeam(effectiveManager);
  const resolvedSquads = m => {
    const sq = state.managers[m]?.squads || [];
    const base = sq.length ? JSON.parse(JSON.stringify(sq)) : mkDefaultSquads(m);
    // Defensively ensure every designer in the team settings appears in a squad
    const teamDesigners = getTeam(m).designers || [];
    const inSquads = new Set(base.flatMap(s => s.designers.map(d => d.name)));
    teamDesigners.forEach(name => {
      if (!inSquads.has(name)) ensureDesignerInSquads(base, name, getTeam(m));
    });
    return base;
  };
  const currentSquads = resolvedSquads(effectiveManager);

  // VP: aggregate squads from all managers (or filter to one)
  const allManagerKeys = [...new Set(MANAGERS.flatMap(m => { const sbs=SUB_MANAGERS[m]||[]; return sbs.length?[m,...sbs]:[m]; }))];

  // Annotated squads: each squad tagged with its owning manager key
  const vpSquadsAnnotated = isVP && vpFilter === "all"
    ? allManagerKeys.flatMap(k => resolvedSquads(k).map(sq => ({ ...sq, _manager: k })))
    : null;

  // Filter option lists for the VP multiselect bar
  const vpAllManagers = allManagerKeys;
  const vpAllSquads   = vpSquadsAnnotated ? [...new Set(vpSquadsAnnotated.map(sq => sq.name))] : [];
  const vpAllDesigners= vpSquadsAnnotated ? [...new Set(vpSquadsAnnotated.flatMap(sq => sq.designers.map(d => d.name)))] : [];

  const vpSquads = isVP
    ? (vpFilter === "all"
        ? (vpSquadsAnnotated
            .filter(sq => !vpMgrFilter || vpMgrFilter.length === 0 || vpMgrFilter.includes(sq._manager))
            .filter(sq => !vpSqFilter  || vpSqFilter.length  === 0 || vpSqFilter.includes(sq.name))
            .map(sq => ({
              ...sq,
              designers: sq.designers.filter(d => !vpDsnFilter || vpDsnFilter.length === 0 || vpDsnFilter.includes(d.name))
            }))
            .filter(sq => sq.designers.length > 0))
        : resolvedSquads(vpFilter))
    : null;
  const displaySquads = vpSquads ?? currentSquads;

  // Team used by Timeline (drives filter chips + designer visibility)
  const allTeamDesigners = [...new Set(allManagerKeys.flatMap(k => getTeam(k).designers || []))];
  const displayTeam = isVP
    ? (vpFilter === "all" ? { designers: allTeamDesigners } : getTeam(vpFilter))
    : currentTeam;

  const overloaded = allOverloaded(displaySquads);
  const managerOptions = MANAGERS.flatMap(m => [
    m,
    ...(SUB_MANAGERS[m] || []).map(sub => ({ value: sub, label: `↳ ${sub}`, triggerLabel: sub })),
  ]);

  if (!user) return <LoginScreen onLogin={handleLogin} />;

  const navItems = !isDesigner
    ? [
        { id:"timeline",  label:"Timeline",  icon:<MI name="calendar_today" size={18} /> },
        { id:"analytics", label:"Analytics", icon:<MI name="bar_chart" size={18} /> },
      ]
    : [];

  if (!appReady) return (
    <div style={{ height:"100vh", background:"#FAFAFA", fontFamily:"'Inter',sans-serif", display:"flex", alignItems:"center", justifyContent:"center", flexDirection:"column", gap:16 }}>
      <img src="/XDlogo.svg" alt="" style={{ width:48, height:48, objectFit:"contain", opacity:0.8 }} />
      <div style={{ width:32, height:32, border:"3px solid #E8EAED", borderTopColor:"#2563EB", borderRadius:"50%", animation:"spin 0.8s linear infinite" }} />
      <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
    </div>
  );

  return (
    <div style={{ height:"100vh", background:"#FAFAFA", fontFamily:"'Inter',sans-serif", display:"flex", overflow:"hidden" }}>

        {/* Left Sidebar */}
        <div style={{ width:240, background:"#fff", borderRight:"1px solid #ECEFF1", display:"flex", flexDirection:"column", flexShrink:0 }}>

          {/* Logo */}
          <div style={{ padding:"20px 20px 0", display:"flex", alignItems:"center", gap:12, flexShrink:0 }}>
            <img src="/XDlogo.svg" alt="" style={{ width:40, height:40, objectFit:"contain", flexShrink:0 }} />
            <span style={{ fontSize:16, fontWeight:600, color:"#0F172A", letterSpacing:"-0.01em", lineHeight:1.25 }}>XD Management</span>
          </div>

          {/* Nav items */}
          <div style={{ padding:"0 12px 16px", flex:1, overflowY:"auto", marginTop:24 }}>
            <div style={{ fontSize:11, fontWeight:700, color:"#B0BEC5", letterSpacing:"0.1em", textTransform:"uppercase", padding:"0 14px", marginBottom:8 }}>Menu</div>
            {navItems.map(({ id, label, icon }) => {
              const active = id === "settings" ? showSettings : page === id;
              const handleClick = id === "settings" ? () => setShowSettings(true) : () => setPage(id);
              return (
                <button key={id} onClick={handleClick}
                  onMouseEnter={e=>{ if(!active) e.currentTarget.style.background="#F5F5F5"; }}
                  onMouseLeave={e=>{ if(!active) e.currentTarget.style.background="transparent"; }}
                  style={{ width:"100%", display:"flex", alignItems:"center", gap:12, height:48, padding:"0 16px", borderRadius:24, border:"none", background:active?"#EEF4FD":"transparent", color:active?"#2563EB":"#546E7A", fontSize:14, fontWeight:active?600:500, cursor:"pointer", fontFamily:"'Inter',sans-serif", marginBottom:2, textAlign:"left", transition:"background 0.15s", boxSizing:"border-box" }}>
                  <span style={{ color:active?"#2563EB":"#78909C", display:"flex", flexShrink:0, transition:"color 0.15s" }}>{icon}</span>
                  {label}
                </button>
              );
            })}
          </div>

          {/* Bottom: user */}
          <div style={{ padding:"12px 12px 16px", borderTop:"1px solid #ECEFF1" }}>
            <div style={{ display:"flex", alignItems:"center", gap:10, padding:"10px 12px", borderRadius:12, background:"#F5F5F5" }}>
              <div style={{ width:36, height:36, borderRadius:"50%", background:"#2563EB", color:"#fff", display:"flex", alignItems:"center", justifyContent:"center", fontSize:12, fontWeight:700, flexShrink:0, fontFamily:"'Inter',sans-serif", letterSpacing:"0.03em" }}>
                {user.slice(0,2).toUpperCase()}
              </div>
              <div style={{ minWidth:0, flex:1, overflow:"hidden" }}>
                <div style={{ fontSize:11, fontWeight:500, color:"#90A4AE", letterSpacing:"0.1em", textTransform:"uppercase", marginBottom:2 }}>Logged in as</div>
                <div style={{ fontSize:12, fontWeight:700, color:"#0F172A", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{user}</div>
              </div>
              <button onClick={handleLogout} title="Sign out"
                onMouseEnter={e=>e.currentTarget.style.color="#C62828"}
                onMouseLeave={e=>e.currentTarget.style.color="#90A4AE"}
                style={{ background:"transparent", border:"none", cursor:"pointer", color:"#90A4AE", padding:6, display:"flex", flexShrink:0, borderRadius:"50%", transition:"color 0.15s" }}>
                <MI name="logout" size={15} />
              </button>
            </div>
          </div>
        </div>

        {/* Main content */}
        <div style={{ flex:1, display:"flex", flexDirection:"column", minWidth:0, overflow:"hidden" }}>

        {/* Content */}
        <div style={{ padding:page==="analytics"?0:24, flex:1, overflowY:"auto", overflowX:"hidden" }}>
          {page === "timeline" && <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:24 }}>
            <div ref={managerDropRef} style={{ position:"relative" }}>
              <div onClick={() => !isDesigner && setManagerDropOpen(o => !o)}
                style={{ display:"flex", alignItems:"center", gap:6, cursor: isDesigner ? "default" : "pointer", userSelect:"none" }}>
                <span style={{ fontSize:16, fontWeight:700, color:"#0F172A", fontFamily:"'Inter',sans-serif", letterSpacing:"-0.01em" }}>
                  {isVP ? (vpFilter === "all" ? "XD Team Timeline" : `${vpFilter}'s Team Timeline`) : `${manager}'s Team Timeline`}
                </span>
                {!isDesigner && <MI name="expand_more" size={20} style={{ color:"#90A4AE", transform: managerDropOpen ? "rotate(180deg)" : "none", transition:"transform 0.15s", display:"flex" }} />}
              </div>
              {managerDropOpen && (
                <div style={{ position:"absolute", top:"calc(100% + 8px)", left:0, background:"#fff", border:"1px solid #ECEFF1", borderRadius:12, boxShadow:"0 8px 24px rgba(0,0,0,0.12)", zIndex:9999, minWidth:220, overflow:"hidden" }}>
                  {(isVP ? [{ value:"all", label:"All Teams" }, ...managerOptions] : managerOptions).map(opt => {
                    const val = typeof opt === "string" ? opt : (opt.value || opt);
                    const lbl = typeof opt === "string" ? opt : (opt.label || opt);
                    const isSub = typeof opt !== "string" && opt.triggerLabel;
                    const current = isVP ? vpFilter : manager;
                    const isActive = val === current;
                    return (
                      <div key={val} onClick={() => { isVP ? setVpFilter(val) : setManager(val); setFilter([]); setManagerDropOpen(false); }}
                        onMouseEnter={e => e.currentTarget.style.background = isActive ? "#E1EDFC" : "#FAFAFA"}
                        onMouseLeave={e => e.currentTarget.style.background = isActive ? "#EEF4FD" : "transparent"}
                        style={{ display:"flex", alignItems:"center", justifyContent:"space-between", padding: isSub ? "8px 16px 8px 28px" : "10px 16px", fontSize: isSub ? 12 : 13, cursor:"pointer", color: isActive ? "#2563EB" : isSub ? "#607D8B" : "#0F172A", fontWeight: isActive ? 600 : 400, background: isActive ? "#EEF4FD" : "transparent" }}>
                        {lbl}
                        {isActive && <MI name="check" size={14} style={{ color:"#2563EB", display:"flex" }} />}
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
            <div style={{ flex:1 }} />
            {!isDesigner && isOwnTeam && overloaded.length > 0 && (
              <button onClick={() => setShowOverload(true)} style={{ display:"flex", alignItems:"center", gap:6, height:40, padding:"0 16px", borderRadius:20, border:"none", background:"#FFEBEE", color:"#C62828", fontSize:12, fontWeight:500, cursor:"pointer", flexShrink:0 }}>
                <MI name="flag" size={14}/>{overloaded.length} designer{overloaded.length !== 1 ? "s" : ""} overloaded
              </button>
            )}
            {isManager && !isVP && isOwnTeam && (
              <button onClick={() => setShowSettings(true)} style={{ display:"flex", alignItems:"center", gap:6, height:40, padding:"0 18px", borderRadius:20, border:"none", background:"#F1F5FF", color:"#3B5BDB", fontSize:14, fontWeight:500, cursor:"pointer", flexShrink:0, fontFamily:"'Inter',sans-serif" }}>
                Manage Team
              </button>
            )}
            {!isVP && isOwnTeam && (
              <button onClick={() => { setAddPrefill(isDesigner ? { designer: user } : null); setShowAdd(true); }} style={{ display:"flex", alignItems:"center", gap:6, height:40, padding:"0 20px", borderRadius:20, border:"none", background:"#2563EB", color:"#fff", fontSize:14, fontWeight:500, cursor:"pointer", flexShrink:0, boxShadow:"0 1px 2px rgba(0,0,0,0.3)" }}>
                <MI name="add" size={13} />
                Add Project
              </button>
            )}
          </div>}
          {page === "timeline" && isVP && vpFilter === "all" && vpAllManagers.length > 0 && (
            <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:16, flexWrap:"wrap" }}>
              <span style={{ fontSize:12, fontWeight:600, color:"#90A4AE", letterSpacing:"0.04em", textTransform:"uppercase", marginRight:4 }}>Filter</span>
              <MultiSelectPill
                label="Managers"
                options={vpAllManagers}
                selected={vpMgrFilter}
                onChange={setVpMgrFilter}
              />
              <MultiSelectPill
                label="Squads"
                options={vpAllSquads}
                selected={vpSqFilter}
                onChange={setVpSqFilter}
              />
              <MultiSelectPill
                label="Designers"
                options={vpAllDesigners}
                selected={vpDsnFilter}
                onChange={setVpDsnFilter}
              />
              {(vpMgrFilter || vpSqFilter || vpDsnFilter) && (
                /* MD3 Text Button */
                <button
                  onClick={() => { setVpMgrFilter(null); setVpSqFilter(null); setVpDsnFilter(null); }}
                  onMouseEnter={e => e.currentTarget.style.background = "rgba(37,99,235,0.08)"}
                  onMouseLeave={e => e.currentTarget.style.background = "transparent"}
                  style={{ height:32, padding:"0 12px", borderRadius:8, border:"none", background:"transparent", color:"#2563EB", fontSize:14, fontWeight:500, cursor:"pointer", fontFamily:"'Inter',sans-serif", transition:"background 0.15s" }}>
                  Clear all
                </button>
              )}
            </div>
          )}
          {page === "timeline" && (
            <Timeline
              squads={displaySquads}
              onUpdate={isVP || !isOwnTeam ? undefined : handleUpdate}
              onRemove={isVP || !isOwnTeam ? undefined : handleRemove}
              onEditProject={isVP || !isOwnTeam ? undefined : handleInlineEdit}
              onOpenEdit={isVP || !isOwnTeam ? undefined : proj => setEditProj(proj)}
              onToast={addToast}
              filter={isDesigner ? [user] : filter}
              onFilter={isDesigner ? () => {} : setFilter}
              team={displayTeam}
              hideFilter={isDesigner || (isVP && vpFilter === "all")}
              onUpdateTitle={isManager && !isVP && isOwnTeam ? handleUpdateTitle : undefined}
              onClickEmpty={isVP || !isOwnTeam ? undefined : (dn, iso) => {
                setAddPrefill({ designer: isDesigner ? user : dn, startDate: iso });
                setShowAdd(true);
              }}
            />
          )}
          {page === "analytics" && (
            <AnalyticsPage
              state={state}
              teams={teams}
              setDrawer={d => { setDrawer(d); setDrawerSearch(""); setDrawerSort(false); setDrawerManagers([]); }}
              setManager={m => { setManager(m); setPage("timeline"); }}
              setPage={setPage}
            />
          )}
        </div>
        </div>{/* end Main content */}

      {/* Modals */}
      {showAdd && (
        <AddModal team={currentTeam} onAdd={handleAdd} onAddBatch={handleAddBatch} onClose={() => { setShowAdd(false); setAddPrefill(null); }} prefill={addPrefill} lockedDesigner={isDesigner ? user : undefined} />
      )}
      {editProj && (
        <AddModal team={currentTeam} onAdd={handleEditSubmit} onAddBatch={() => {}} onClose={() => setEditProj(null)} prefill={editProj} />
      )}
      {showSettings && (
        <Settings settings={currentTeam} stateSquads={currentSquads} onSave={handleSaveSettings} onClose={() => setShowSettings(false)} />
      )}
      {showOverload && (
        <OverloadPanel overloaded={overloaded} onClose={() => setShowOverload(false)} />
      )}

      <Drawer drawer={drawer} setDrawer={setDrawer} drawerSearch={drawerSearch} setDrawerSearch={setDrawerSearch} drawerSort={drawerSort} setDrawerSort={setDrawerSort} drawerManagers={drawerManagers} setDrawerManagers={setDrawerManagers} />
      <Toast toasts={toasts} />
    </div>
  );
}
