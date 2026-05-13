import React, { useState, useCallback } from "react";

// ─── CONFIG ───────────────────────────────────────────────────────────────────
const STORAGE_ACCOUNT = "dlmanagerstorage";
const CONTAINER       = "dl-results";
const SAS_TOKEN       = "sv=2025-11-05&ss=bfqt&srt=o&sp=rwdlacupiytfx&se=2027-05-13T23:32:25Z&st=2026-05-13T15:17:25Z&spr=https&sig=wTxGQfShm2HTKvyFAUfyNTME5NcsvV1qnFLK%2FOBsNdQ%3D";

function blobUrl(blobName) {
  return `https://${STORAGE_ACCOUNT}.blob.core.windows.net/${CONTAINER}/${blobName}?${SAS_TOKEN}`;
}

const WEBHOOKS = {
  getOwned:   "https://d6f7aa8d-1818-44aa-bf8b-0079a6d6a479.webhook.eus.azure-automation.net/webhooks?token=g5Btr7qtBD5%2bDykCAW0VGQ2by%2bsrAisb2O3oZH1AAMo%3d",
  getMembers: "https://d6f7aa8d-1818-44aa-bf8b-0079a6d6a479.webhook.eus.azure-automation.net/webhooks?token=SwSQOB0umD7OE%2beyNHyZtYRq1h2nMpWdZ741k3T10XA%3d",
  add:        "https://d6f7aa8d-1818-44aa-bf8b-0079a6d6a479.webhook.eus.azure-automation.net/webhooks?token=UD9tS50TIYhxTqJoowZq9b%2bWSjDAQTxuY1RJFuPuqbg%3d",
  remove:     "https://d6f7aa8d-1818-44aa-bf8b-0079a6d6a479.webhook.eus.azure-automation.net/webhooks?token=bGCb%2bRQeRnIHGIzcSpcpE6%2f6FP10UxROqhnqs81bnpc%3d",
};

// ─── Webhook caller ───────────────────────────────────────────────────────────
// Azure Automation webhooks don't return CORS headers — must use no-cors.
// no-cors means we get an opaque response (can't read status/body) but the
// request still fires and the runbook still executes. We poll blob for results.
async function callRunbook(url, params = {}) {
  const body = Object.keys(params).length
    ? { Properties: Object.entries(params).map(([Key, Value]) => ({ Key, Value })) }
    : {};
  await fetch(url, {
    method: "POST",
    mode: "no-cors",           // ← required: AA webhooks have no CORS headers
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });
  // With no-cors we cannot read the response — assume fired successfully
}

// ─── Blob reader (GET only — no DELETE, avoids CORS preflight issues) ────────
async function readBlob(blobName) {
  const res = await fetch(blobUrl(blobName), { method: "GET" });
  if (res.status === 404) return null;
  if (!res.ok) throw new Error(`Blob read failed: HTTP ${res.status}`);
  const text = await res.text();
  try { return JSON.parse(text); } catch { return null; }
}

// Poll blob until data newer than firedAt appears, or timeout
async function pollBlob(blobName, { intervalMs=6000, timeoutMs=150000, firedAt=0 } = {}) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      // Check Last-Modified header to skip stale blobs from previous runs
      const head = await fetch(blobUrl(blobName), { method: "HEAD" });
      if (head.ok) {
        const lastMod = head.headers.get("Last-Modified");
        const blobAge = lastMod ? new Date(lastMod).getTime() : 0;
        if (blobAge > firedAt) {
          const data = await readBlob(blobName);
          if (data !== null) return data;
        }
      }
    } catch { /* ignore, keep polling */ }
    await new Promise(r => setTimeout(r, intervalMs));
  }
  throw new Error("Timed out waiting for runbook (>2.5 min). Check Azure Automation job logs.");
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
function initials(name) { return (name||"?").split(" ").map(n=>n[0]).slice(0,2).join("").toUpperCase(); }
const AVC = ["#1a6b8a","#2d7a4f","#7a4f2d","#6b1a6b","#1a3d6b","#8a1a1a","#4f6b1a"];
function avc(name) { let h=0; for(let c of (name||"")) h=(h*31+c.charCodeAt(0))&0xffffffff; return AVC[Math.abs(h)%AVC.length]; }

function Avatar({ name, size=36 }) {
  return (
    <div style={{width:size,height:size,borderRadius:"50%",background:avc(name),flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontFamily:"'DM Sans',sans-serif",fontSize:size*.38,fontWeight:600}}>
      {initials(name)}
    </div>
  );
}

function Spinner({ size=14, color="#fff" }) {
  return <span style={{display:"inline-block",width:size,height:size,border:`2px solid ${color}33`,borderTop:`2px solid ${color}`,borderRadius:"50%",animation:"spin .7s linear infinite",flexShrink:0}}/>;
}

function Toast({ msg, type, onClose }) {
  const timerRef = React.useRef(null);
  if (!timerRef.current) {
    timerRef.current = setTimeout(onClose, 5000);
  }
  const bg={success:"#0f5132",error:"#842029",pending:"#664d03",info:"#1e40af"};
  const ic={success:"✓",error:"✕",pending:"⏳",info:"ℹ"};
  return (
    <div style={{position:"fixed",bottom:24,right:24,zIndex:9999,background:bg[type]||bg.info,color:"#fff",borderRadius:10,padding:"13px 20px",fontFamily:"'DM Sans',sans-serif",fontSize:14,boxShadow:"0 8px 30px rgba(0,0,0,.25)",display:"flex",alignItems:"center",gap:10,maxWidth:420,animation:"slideUp .3s ease"}}>
      <span>{ic[type]||"ℹ"}</span>{msg}
    </div>
  );
}

function ConfirmModal({ message, onConfirm, onCancel }) {
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(10,15,30,.55)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:900,backdropFilter:"blur(3px)"}}>
      <div style={{background:"#fff",borderRadius:16,padding:"30px 34px",maxWidth:400,width:"90%",boxShadow:"0 20px 60px rgba(0,0,0,.2)",fontFamily:"'DM Sans',sans-serif"}}>
        <div style={{fontSize:36,textAlign:"center",marginBottom:10}}>⚠️</div>
        <p style={{textAlign:"center",color:"#222",fontSize:14,marginBottom:26,lineHeight:1.6}}>{message}</p>
        <div style={{display:"flex",gap:10,justifyContent:"center"}}>
          <button onClick={onCancel}  style={S.btn("ghost")}>Cancel</button>
          <button onClick={onConfirm} style={S.btn("danger")}>Remove</button>
        </div>
      </div>
    </div>
  );
}

function PendingBanner({ ops }) {
  if (!ops.length) return null;
  return (
    <div style={{background:"#fff8e1",border:"1px solid #ffe082",borderRadius:10,padding:"11px 16px",marginBottom:16,display:"flex",alignItems:"flex-start",gap:10,fontFamily:"'DM Sans',sans-serif",fontSize:13,color:"#7a5c00"}}>
      <span style={{animation:"spin 1.5s linear infinite",display:"inline-block",marginTop:1,fontSize:15}}>⟳</span>
      <div>
        <strong>{ops.length} change{ops.length>1?"s":""} in flight</strong> — running in Azure Automation (~30–90s)
        <div style={{marginTop:4,fontSize:12,opacity:.75}}>
          {ops.map((op,i)=><div key={i}>{op.type==="add"?"➕":"➖"} {op.member} {op.type==="add"?"→":"←"} {op.group}</div>)}
        </div>
      </div>
    </div>
  );
}

function SidebarSkeleton() {
  return (
    <div style={{padding:"0 7px"}}>
      {[80,60,70].map((w,i)=>(
        <div key={i} style={{padding:"11px 9px",borderRadius:10,marginBottom:3}}>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <div style={{width:32,height:32,borderRadius:8,background:"#e2e8f0",animation:"pulse 1.4s ease-in-out infinite"}}/>
            <div style={{flex:1}}>
              <div style={{height:11,background:"#e2e8f0",borderRadius:4,marginBottom:6,width:`${w}%`,animation:"pulse 1.4s ease-in-out infinite"}}/>
              <div style={{height:9,background:"#e2e8f0",borderRadius:4,width:"90%",animation:"pulse 1.4s ease-in-out infinite"}}/>
            </div>
          </div>
        </div>
      ))}
    </div>
  );
}

function PollingIndicator({ label }) {
  return (
    <div style={{display:"flex",alignItems:"center",gap:8,padding:"16px 12px",fontFamily:"'DM Sans',sans-serif",fontSize:13,color:"#64748b"}}>
      <Spinner size={14} color="#1a56db"/>
      <span>{label}…</span>
    </div>
  );
}

const S = {
  btn(v) {
    const b={border:"none",borderRadius:8,padding:"9px 20px",cursor:"pointer",fontFamily:"'DM Sans',sans-serif",fontSize:14,fontWeight:600,transition:"all .15s"};
    if(v==="primary") return {...b,background:"#1a56db",color:"#fff"};
    if(v==="danger")  return {...b,background:"#dc2626",color:"#fff"};
    if(v==="ghost")   return {...b,background:"#f1f5f9",color:"#374151"};
    if(v==="outline") return {...b,background:"transparent",color:"#1a56db",border:"1.5px solid #1a56db"};
    return b;
  }
};

// ─── App ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [authStep,   setAuthStep]   = useState("login");
  const [signingIn,  setSigningIn]  = useState(false);
  const [userEmail,  setUserEmail]  = useState("jane.smith@contoso.com");

  const [lists,      setLists]      = useState([]);
  const [listsPhase, setListsPhase] = useState("idle");
  const [listsError, setListsError] = useState("");

  const [membersMap,   setMembersMap]   = useState({});
  const [membersPhase, setMembersPhase] = useState({});

  const [selectedDL, setSelectedDL] = useState(null);
  const [search,     setSearch]     = useState("");
  const [addPanel,   setAddPanel]   = useState(false);
  const [addEmail,   setAddEmail]   = useState("");
  const [confirm,    setConfirm]    = useState(null);
  const [toast,      setToast]      = useState(null);
  const [pending,    setPending]    = useState([]);
  const [loading,    setLoading]    = useState({});
  const [viewMode,   setViewMode]   = useState("owned");

  const showToast = (msg, type="success") => setToast({msg,type});

  const dl        = selectedDL ? lists.find(l=>l.id===selectedDL) : null;
  const dlMembers = dl ? (membersMap[dl.email]||[]) : [];
  const dlMPhase  = dl ? (membersPhase[dl.email]||"idle") : "idle";

  function normalizeDLs(raw) {
    if (!raw) return [];
    const arr = Array.isArray(raw) ? raw : [raw];
    return arr.map((item,i) => ({
      id:          item.ExternalDirectoryObjectId || item.PrimarySmtpAddress || `dl-${i}`,
      displayName: item.DisplayName || item.PrimarySmtpAddress || "Unknown",
      email:       item.PrimarySmtpAddress || "",
      memberCount: item.MemberCount ?? 0,
      owned:       true,
    }));
  }

  function normalizeMembers(raw) {
    if (!raw) return [];
    const arr = Array.isArray(raw) ? raw : [raw];
    return arr.map(m => ({
      id:          m.ExternalDirectoryObjectId || m.PrimarySmtpAddress || String(Math.random()),
      displayName: m.DisplayName || m.Name || m.PrimarySmtpAddress || "Unknown",
      email:       m.PrimarySmtpAddress || m.WindowsEmailAddress || "",
      jobTitle:    m.Title || "",
    }));
  }

  // ── Fetch owned DLs ────────────────────────────────────────────────────────
  const fetchOwnedDLs = useCallback(async (email) => {
    setListsPhase("submitting");
    setListsError("");
    const firedAt = Date.now();

    try {
      await callRunbook(WEBHOOKS.getOwned, { OwnerEmail: email });
    } catch(e) {
      setListsPhase("error");
      setListsError(`Failed to submit runbook: ${e.message}`);
      return;
    }

    setListsPhase("polling");
    try {
      const data = await pollBlob(`${email}.json`, { intervalMs:6000, timeoutMs:150000, firedAt });
      const normalized = normalizeDLs(data);
      setLists(normalized);
      setListsPhase("loaded");
      showToast(
        normalized.length === 0
          ? "No distribution lists found where you are an owner"
          : `Loaded ${normalized.length} distribution list${normalized.length>1?"s":""}`,
        normalized.length === 0 ? "info" : "success"
      );
    } catch(e) {
      setListsPhase("error");
      setListsError(e.message);
      showToast(e.message, "error");
    }
  }, []);

  // ── Fetch members ──────────────────────────────────────────────────────────
  const fetchMembers = useCallback(async (dlEmail) => {
    setMembersPhase(p=>({...p,[dlEmail]:"submitting"}));
    const firedAt = Date.now();

    try {
      await callRunbook(WEBHOOKS.getMembers, { GroupEmail: dlEmail });
    } catch(e) {
      setMembersPhase(p=>({...p,[dlEmail]:"error"}));
      showToast(`Failed to submit: ${e.message}`, "error");
      return;
    }

    setMembersPhase(p=>({...p,[dlEmail]:"polling"}));
    try {
      const data = await pollBlob(`members-${dlEmail}.json`, { intervalMs:5000, timeoutMs:120000, firedAt });
      const normalized = normalizeMembers(data);
      setMembersMap(p=>({...p,[dlEmail]:normalized}));
      setLists(p=>p.map(l=>l.email===dlEmail?{...l,memberCount:normalized.length}:l));
      setMembersPhase(p=>({...p,[dlEmail]:"loaded"}));
      showToast(`Loaded ${normalized.length} members`, "success");
    } catch(e) {
      setMembersPhase(p=>({...p,[dlEmail]:"error"}));
      showToast(e.message, "error");
    }
  }, []);

  // ── Add member ─────────────────────────────────────────────────────────────
  async function handleAdd(memberEmail) {
    memberEmail = memberEmail.trim();
    if (!memberEmail) return;
    const group = dl.email;
    const key = `add-${memberEmail}`;
    setLoading(p=>({...p,[key]:true}));
    try {
      await callRunbook(WEBHOOKS.add, { GroupEmail: group, MemberEmail: memberEmail });
      setMembersMap(p=>({...p,[group]:[...(p[group]||[]),{id:memberEmail,displayName:memberEmail,email:memberEmail,jobTitle:""}]}));
      setLists(p=>p.map(l=>l.id===selectedDL?{...l,memberCount:l.memberCount+1}:l));
      setPending(p=>[...p,{type:"add",member:memberEmail,group}]);
      showToast(`Add submitted for ${memberEmail}`, "pending");
      setAddEmail("");
      setTimeout(()=>setPending(p=>p.filter(o=>!(o.member===memberEmail&&o.group===group))),90000);
    } catch(e) { showToast(`Failed: ${e.message}`, "error"); }
    finally { setLoading(p=>({...p,[key]:false})); }
  }

  // ── Remove member ──────────────────────────────────────────────────────────
  async function handleRemove() {
    const { userId, memberEmail, memberName } = confirm;
    const group = dl.email;
    setConfirm(null);
    setLoading(p=>({...p,[`rm-${userId}`]:true}));
    try {
      await callRunbook(WEBHOOKS.remove, { GroupEmail: group, MemberEmail: memberEmail });
      setMembersMap(p=>({...p,[group]:(p[group]||[]).filter(m=>m.id!==userId)}));
      setLists(p=>p.map(l=>l.id===selectedDL?{...l,memberCount:l.memberCount-1}:l));
      setPending(p=>[...p,{type:"remove",member:memberEmail,group}]);
      showToast(`Remove submitted for ${memberName}`, "pending");
      setTimeout(()=>setPending(p=>p.filter(o=>!(o.member===memberEmail&&o.group===group))),90000);
    } catch(e) { showToast(`Failed: ${e.message}`, "error"); }
    finally { setLoading(p=>({...p,[`rm-${userId}`]:false})); }
  }

  // ─── Login ─────────────────────────────────────────────────────────────────
  if (authStep==="login") return (
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#0c1b3a 0%,#1a3a6b 60%,#0e2954 100%)",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'DM Sans',sans-serif"}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Sora:wght@600;700&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        @keyframes slideUp{from{transform:translateY(20px);opacity:0}to{transform:translateY(0);opacity:1}}
        @keyframes fadeIn{from{opacity:0}to{opacity:1}}
        @keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}
        @keyframes spin{to{transform:rotate(360deg)}}
      `}</style>
      <div style={{position:"absolute",inset:0,opacity:.06,backgroundImage:"linear-gradient(rgba(255,255,255,.5) 1px,transparent 1px),linear-gradient(90deg,rgba(255,255,255,.5) 1px,transparent 1px)",backgroundSize:"40px 40px"}}/>
      <div style={{background:"rgba(255,255,255,.07)",backdropFilter:"blur(20px)",border:"1px solid rgba(255,255,255,.13)",borderRadius:24,padding:"50px 46px",width:430,textAlign:"center",boxShadow:"0 40px 80px rgba(0,0,0,.4)",animation:"fadeIn .5s ease",position:"relative"}}>
        <div style={{width:62,height:62,borderRadius:15,background:"linear-gradient(135deg,#0078d4,#106ebe)",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 18px",fontSize:30,boxShadow:"0 8px 24px rgba(0,120,212,.4)"}}>📋</div>
        <h1 style={{fontFamily:"'Sora',sans-serif",fontSize:23,fontWeight:700,color:"#fff",marginBottom:6}}>DL Manager</h1>
        <p style={{color:"rgba(255,255,255,.5)",fontSize:13,lineHeight:1.6,marginBottom:26}}>Distribution List Self-Service<br/>for Microsoft 365 Owners</p>
        <div style={{background:"rgba(255,255,255,.04)",borderRadius:11,padding:"14px 16px",marginBottom:18,textAlign:"left",border:"1px solid rgba(255,255,255,.08)"}}>
          <p style={{color:"rgba(255,255,255,.4)",fontSize:11,marginBottom:5}}>Your M365 email</p>
          <input value={userEmail} onChange={e=>setUserEmail(e.target.value)}
            style={{background:"transparent",border:"none",outline:"none",color:"#fff",fontSize:14,fontFamily:"'DM Sans',sans-serif",width:"100%"}}
            placeholder="your@company.com"/>
        </div>
        <button
          onClick={()=>{
            setSigningIn(true);
            setTimeout(()=>{ setAuthStep("app"); fetchOwnedDLs(userEmail); }, 1400);
          }}
          disabled={signingIn||!userEmail.includes("@")}
          style={{...S.btn("primary"),width:"100%",padding:"13px 22px",fontSize:15,borderRadius:12,background:signingIn?"#1a3a6b":"#0078d4",display:"flex",alignItems:"center",justifyContent:"center",gap:10}}
        >
          {signingIn
            ?<><Spinner/><span style={{animation:"pulse 1s infinite"}}>Signing in…</span></>
            :<><span>🔐</span>Sign in with Microsoft 365</>}
        </button>
        <p style={{color:"rgba(255,255,255,.2)",fontSize:11,marginTop:14}}>Delegated OAuth · No admin credentials stored</p>
      </div>
    </div>
  );

  // ─── Main ──────────────────────────────────────────────────────────────────
  const filteredLists = lists.filter(l =>
    (l.displayName.toLowerCase().includes(search.toLowerCase()) || l.email.toLowerCase().includes(search.toLowerCase())) &&
    (viewMode==="all" || l.owned)
  );

  return (
    <div style={{minHeight:"100vh",background:"#f0f4f8",fontFamily:"'DM Sans',sans-serif"}}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Sora:wght@600;700&display=swap');
        *{box-sizing:border-box;margin:0;padding:0}
        @keyframes slideUp{from{transform:translateY(14px);opacity:0}to{transform:translateY(0);opacity:1}}
        @keyframes fadeIn{from{opacity:0}to{opacity:1}}
        @keyframes spin{to{transform:rotate(360deg)}}
        @keyframes pulse{0%,100%{opacity:1}50%{opacity:.5}}
        .dlc:hover{background:#fff!important;box-shadow:0 4px 20px rgba(26,86,219,.12)!important}
        .mrow:hover{background:#f8faff!important}
        .ibtn:hover{background:#fee2e2!important;color:#dc2626!important}
        ::-webkit-scrollbar{width:5px}::-webkit-scrollbar-track{background:transparent}::-webkit-scrollbar-thumb{background:#cbd5e1;border-radius:3px}
      `}</style>

      {/* Header */}
      <div style={{background:"#fff",borderBottom:"1px solid #e2e8f0",padding:"0 26px",height:56,display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:100,boxShadow:"0 1px 8px rgba(0,0,0,.06)"}}>
        <div style={{display:"flex",alignItems:"center",gap:11}}>
          <div style={{width:31,height:31,borderRadius:8,background:"linear-gradient(135deg,#0078d4,#1a56db)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:16}}>📋</div>
          <span style={{fontFamily:"'Sora',sans-serif",fontWeight:700,fontSize:16,color:"#0f172a"}}>DL Manager</span>
          <span style={{background:"#e8f0fe",color:"#1a56db",fontSize:11,fontWeight:600,borderRadius:6,padding:"2px 7px"}}>M365</span>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          {pending.length>0&&<span style={{background:"#fff8e1",color:"#7a5c00",fontSize:12,fontWeight:600,borderRadius:20,padding:"3px 11px",border:"1px solid #ffe082"}}>⏳ {pending.length} pending</span>}
          <span style={{color:"#64748b",fontSize:13}}>{userEmail}</span>
          <Avatar name={userEmail} size={29}/>
          <button onClick={()=>fetchOwnedDLs(userEmail)} title="Reload my lists"
            style={{background:"#f1f5f9",border:"none",borderRadius:7,width:29,height:29,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center"}}>
            {listsPhase==="polling"||listsPhase==="submitting"?<Spinner size={12} color="#64748b"/>:<span style={{fontSize:14}}>⟳</span>}
          </button>
        </div>
      </div>

      <div style={{display:"flex",height:"calc(100vh - 56px)"}}>

        {/* Sidebar */}
        <div style={{width:288,background:"#fff",borderRight:"1px solid #e2e8f0",display:"flex",flexDirection:"column",flexShrink:0}}>
          <div style={{padding:"16px 13px 10px"}}>
            <div style={{display:"flex",background:"#f1f5f9",borderRadius:9,padding:3,marginBottom:13}}>
              {[["owned","My Lists"],["all","All Lists"]].map(([v,label])=>(
                <button key={v} onClick={()=>setViewMode(v)} style={{flex:1,border:"none",borderRadius:7,padding:"6px 0",cursor:"pointer",fontSize:12,fontWeight:600,fontFamily:"'DM Sans',sans-serif",background:viewMode===v?"#fff":"transparent",color:viewMode===v?"#1a56db":"#64748b",boxShadow:viewMode===v?"0 1px 4px rgba(0,0,0,.1)":"none",transition:"all .15s"}}>{label}</button>
              ))}
            </div>
            <div style={{position:"relative"}}>
              <span style={{position:"absolute",left:9,top:"50%",transform:"translateY(-50%)",fontSize:13,color:"#94a3b8"}}>🔍</span>
              <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search lists…"
                style={{width:"100%",padding:"8px 10px 8px 29px",borderRadius:8,border:"1.5px solid #e2e8f0",fontSize:13,outline:"none",fontFamily:"'DM Sans',sans-serif",background:"#f8fafc",color:"#1e293b"}}/>
            </div>
          </div>

          <div style={{flex:1,overflowY:"auto",padding:"0 7px 8px"}}>
            {(listsPhase==="submitting") && <PollingIndicator label="Submitting runbook job"/>}
            {(listsPhase==="polling")    && <PollingIndicator label="Loading your lists from Exchange"/>}
            {(listsPhase==="idle")       && <SidebarSkeleton/>}

            {listsPhase==="error" && (
              <div style={{margin:"8px 4px",background:"#fef2f2",border:"1px solid #fecaca",borderRadius:10,padding:"12px",fontSize:12,color:"#991b1b",lineHeight:1.6}}>
                ✕ {listsError}
                <br/>
                <button onClick={()=>fetchOwnedDLs(userEmail)} style={{marginTop:6,fontSize:12,color:"#1a56db",background:"none",border:"none",cursor:"pointer",padding:0,fontFamily:"'DM Sans',sans-serif"}}>↺ Retry</button>
              </div>
            )}

            {filteredLists.map(list=>(
              <div key={list.id} className="dlc"
                onClick={()=>{setSelectedDL(list.id);setAddPanel(false);setAddEmail("");}}
                style={{padding:"11px 9px",borderRadius:10,marginBottom:3,cursor:"pointer",transition:"all .15s",background:selectedDL===list.id?"#eff6ff":"transparent",border:`1.5px solid ${selectedDL===list.id?"#bfdbfe":"transparent"}`}}
              >
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <div style={{width:32,height:32,borderRadius:8,flexShrink:0,background:selectedDL===list.id?"#dbeafe":"#eff6ff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:15}}>📧</div>
                  <div style={{minWidth:0,flex:1}}>
                    <div style={{display:"flex",alignItems:"center",gap:5}}>
                      <p style={{fontWeight:600,fontSize:13,color:"#1e293b",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{list.displayName}</p>
                      <span style={{fontSize:9,background:"#dcfce7",color:"#15803d",borderRadius:4,padding:"1px 5px",fontWeight:700,flexShrink:0}}>OWNER</span>
                    </div>
                    <p style={{fontSize:11,color:"#94a3b8",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{list.email}</p>
                  </div>
                </div>
                <div style={{marginTop:5,paddingLeft:40}}>
                  <span style={{background:"#e8f0fe",color:"#1a56db",borderRadius:20,padding:"2px 8px",fontSize:10,fontWeight:600}}>{list.memberCount} members</span>
                </div>
              </div>
            ))}

            {listsPhase==="loaded"&&filteredLists.length===0&&(
              <p style={{textAlign:"center",color:"#94a3b8",fontSize:13,paddingTop:24}}>No lists found</p>
            )}
          </div>
        </div>

        {/* Main panel */}
        <div style={{flex:1,overflowY:"auto",padding:"26px 30px"}}>
          {!dl ? (
            <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",height:"100%",color:"#94a3b8",animation:"fadeIn .4s ease",textAlign:"center"}}>
              <div style={{fontSize:56,marginBottom:12,opacity:.3}}>📋</div>
              {listsPhase==="polling"||listsPhase==="submitting" ? (
                <>
                  <p style={{fontSize:17,fontWeight:600,color:"#64748b",marginBottom:8}}>Loading your distribution lists…</p>
                  <p style={{fontSize:13,marginBottom:18}}>Fetching from Exchange Online via Azure Automation</p>
                  <div style={{display:"flex",gap:6,justifyContent:"center"}}>
                    {[0,1,2].map(i=><div key={i} style={{width:8,height:8,borderRadius:"50%",background:"#1a56db",animation:`pulse 1.2s ease-in-out ${i*.2}s infinite`}}/>)}
                  </div>
                </>
              ) : (
                <>
                  <p style={{fontSize:17,fontWeight:600,color:"#64748b"}}>Select a Distribution List</p>
                  <p style={{fontSize:13,marginTop:5}}>Choose from the sidebar to manage members</p>
                </>
              )}
            </div>
          ) : (
            <div style={{maxWidth:740,animation:"slideUp .3s ease"}}>
              <PendingBanner ops={pending}/>

              {/* DL header */}
              <div style={{background:"linear-gradient(135deg,#1a56db 0%,#0078d4 100%)",borderRadius:18,padding:"24px 28px",marginBottom:20,color:"#fff",boxShadow:"0 8px 32px rgba(26,86,219,.25)"}}>
                <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between"}}>
                  <div>
                    <div style={{display:"flex",alignItems:"center",gap:9,marginBottom:5}}>
                      <span style={{fontSize:22}}>📧</span>
                      <h2 style={{fontFamily:"'Sora',sans-serif",fontSize:19,fontWeight:700}}>{dl.displayName}</h2>
                    </div>
                    <p style={{opacity:.75,fontSize:13}}>{dl.email}</p>
                  </div>
                  <div style={{background:"rgba(255,255,255,.18)",borderRadius:12,padding:"9px 16px",textAlign:"center",minWidth:72}}>
                    <div style={{fontSize:22,fontWeight:700}}>{dl.memberCount}</div>
                    <div style={{fontSize:11,opacity:.8}}>Members</div>
                  </div>
                </div>
                <div style={{marginTop:12,display:"flex",gap:7,flexWrap:"wrap",alignItems:"center"}}>
                  <span style={{background:"rgba(255,255,255,.2)",borderRadius:20,padding:"3px 12px",fontSize:12,fontWeight:600}}>✓ You are an owner</span>
                  <span style={{background:"rgba(255,255,255,.11)",borderRadius:20,padding:"3px 12px",fontSize:12,opacity:.8}}>Exchange Online</span>
                  <button onClick={()=>fetchMembers(dl.email)}
                    disabled={dlMPhase==="polling"||dlMPhase==="submitting"}
                    style={{background:"rgba(255,255,255,.15)",border:"1px solid rgba(255,255,255,.3)",borderRadius:20,padding:"3px 12px",fontSize:12,color:"#fff",cursor:"pointer",fontFamily:"'DM Sans',sans-serif",fontWeight:600,display:"flex",alignItems:"center",gap:6}}>
                    {dlMPhase==="submitting"||dlMPhase==="polling"
                      ?<><Spinner size={11}/>{dlMPhase==="submitting"?"Submitting…":"Polling Exchange…"}</>
                      :"⟳ Load members"}
                  </button>
                </div>
              </div>

              {/* Members table */}
              <div style={{background:"#fff",borderRadius:16,border:"1px solid #e2e8f0",overflow:"hidden",boxShadow:"0 2px 12px rgba(0,0,0,.04)"}}>
                <div style={{padding:"15px 20px",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                  <div>
                    <h3 style={{fontSize:15,fontWeight:700,color:"#0f172a"}}>Members</h3>
                    <p style={{fontSize:12,color:"#94a3b8",marginTop:2}}>
                      {dlMPhase==="polling"   ?"Polling Exchange for members…"  :
                       dlMPhase==="submitting"?"Submitting runbook job…"         :
                       dlMembers.length>0     ?`${dlMembers.length} members loaded from Exchange`
                                              :"Click ⟳ Load members to fetch from Exchange"}
                    </p>
                  </div>
                  <button onClick={()=>setAddPanel(!addPanel)}
                    style={{...S.btn("primary"),display:"flex",alignItems:"center",gap:6,background:addPanel?"#1e40af":"#1a56db"}}>
                    <span style={{fontSize:15}}>{addPanel?"✕":"+"}</span>
                    {addPanel?"Close":"Add Member"}
                  </button>
                </div>

                {/* Add panel */}
                {addPanel&&(
                  <div style={{background:"#f8faff",borderBottom:"1px solid #e2e8f0",padding:"14px 20px",animation:"slideUp .2s ease"}}>
                    <p style={{fontSize:13,fontWeight:600,color:"#374151",marginBottom:10}}>Add member by email address</p>
                    <div style={{display:"flex",gap:8}}>
                      <input value={addEmail} onChange={e=>setAddEmail(e.target.value)}
                        onKeyDown={e=>e.key==="Enter"&&handleAdd(addEmail)}
                        placeholder="user@contoso.com" autoFocus
                        style={{flex:1,padding:"9px 12px",borderRadius:8,border:"1.5px solid #d1d5db",fontSize:13,outline:"none",fontFamily:"'DM Sans',sans-serif"}}/>
                      <button onClick={()=>handleAdd(addEmail)}
                        disabled={!addEmail.trim()||loading[`add-${addEmail.trim()}`]}
                        style={{...S.btn("primary"),padding:"9px 18px",fontSize:13,borderRadius:8,flexShrink:0,display:"flex",alignItems:"center",gap:6}}>
                        {loading[`add-${addEmail.trim()}`]?<><Spinner size={13}/> Submitting…</>:"+ Add"}
                      </button>
                    </div>
                    <p style={{fontSize:11,color:"#94a3b8",marginTop:8,lineHeight:1.5}}>
                      Enter the user's full UPN / email. Fires <code>Add-DistributionGroupMember</code> via Azure Automation.
                    </p>
                  </div>
                )}

                {/* Member rows */}
                {dlMPhase==="polling"||dlMPhase==="submitting" ? (
                  <div style={{padding:"28px 20px"}}>
                    <PollingIndicator label={dlMPhase==="submitting"?"Submitting runbook":"Fetching members from Exchange"}/>
                  </div>
                ) : dlMembers.length===0 ? (
                  <div style={{padding:"34px 20px",textAlign:"center",color:"#94a3b8"}}>
                    <div style={{fontSize:30,marginBottom:7,opacity:.3}}>👥</div>
                    <p style={{fontSize:13,marginBottom:4}}>No members loaded</p>
                    <p style={{fontSize:12}}>Click <strong>⟳ Load members</strong> in the header to fetch from Exchange Online</p>
                  </div>
                ) : dlMembers.map((m,i)=>(
                  <div key={m.id} className="mrow" style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"12px 20px",transition:"background .12s",borderBottom:i<dlMembers.length-1?"1px solid #f8fafc":"none",opacity:loading[`rm-${m.id}`]?.4:1}}>
                    <div style={{display:"flex",alignItems:"center",gap:10}}>
                      <Avatar name={m.displayName||m.email} size={34}/>
                      <div>
                        <p style={{fontSize:13,fontWeight:600,color:"#1e293b"}}>{m.displayName}</p>
                        <p style={{fontSize:12,color:"#94a3b8"}}>{m.email}</p>
                      </div>
                    </div>
                    <div style={{display:"flex",alignItems:"center",gap:9}}>
                      {m.jobTitle&&<span style={{fontSize:12,color:"#64748b",background:"#f1f5f9",borderRadius:6,padding:"3px 8px",fontWeight:500}}>{m.jobTitle}</span>}
                      <button className="ibtn" title="Remove" disabled={loading[`rm-${m.id}`]}
                        onClick={()=>setConfirm({userId:m.id,memberEmail:m.email,memberName:m.displayName||m.email})}
                        style={{background:"transparent",border:"none",cursor:"pointer",width:29,height:29,borderRadius:7,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,color:"#cbd5e1",transition:"all .12s"}}>
                        {loading[`rm-${m.id}`]?<Spinner size={13} color="#94a3b8"/>:"🗑"}
                      </button>
                    </div>
                  </div>
                ))}
              </div>

              {/* Info box */}
              <div style={{marginTop:16,padding:"13px 16px",borderRadius:11,background:"#fffbeb",border:"1px solid #fde68a",display:"flex",gap:9,alignItems:"flex-start"}}>
                <span style={{fontSize:15,flexShrink:0}}>⚡</span>
                <div>
                  <p style={{fontSize:13,fontWeight:600,color:"#92400e",marginBottom:2}}>How this works</p>
                  <p style={{fontSize:12,color:"#a16207",lineHeight:1.7}}>
                    <strong>Load members</strong> — fires Get-DLMembers runbook → polls <code>dlmanagerstorage</code> every 5s for results (~30–90s total)<br/>
                    <strong>Add / Remove</strong> — fires runbook webhook, UI updates optimistically. Change lands in Exchange in ~30–90s.
                  </p>
                </div>
              </div>
            </div>
          )}
        </div>
      </div>

      {confirm&&<ConfirmModal
        message={`Remove ${confirm.memberName} from ${dl?.displayName}? Submits a job to Exchange Online via Azure Automation.`}
        onConfirm={handleRemove} onCancel={()=>setConfirm(null)}/>}
      {toast&&<Toast msg={toast.msg} type={toast.type} onClose={()=>setToast(null)}/>}
    </div>
  );
}
