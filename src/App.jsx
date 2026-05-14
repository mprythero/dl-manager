import React, { useState, useEffect, useCallback } from "react";
import { msalInstance, graphScopes, graphRequest, graphRequestAsApp, getToken, CLIENT_ID } from "./auth";

// ─── Font Awesome icon helper ─────────────────────────────────────────────────
function Fa({ icon, style={} }) {
  return <i className={`fa-solid fa-${icon}`} style={style}/>;
}
function Far({ icon, style={} }) {
  return <i className={`fa-regular fa-${icon}`} style={style}/>;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
function initials(name) {
  return (name || "?").split(" ").map(n => n[0]).slice(0, 2).join("").toUpperCase();
}
const AVC = ["#1a6b8a","#2d7a4f","#7a4f2d","#6b1a6b","#1a3d6b","#8a1a1a","#4f6b1a","#1a5c8a","#5c1a8a"];
function avc(s) { let h=0; for(let c of (s||"")) h=(h*31+c.charCodeAt(0))&0xffffffff; return AVC[Math.abs(h)%AVC.length]; }

function Avatar({ name, size=36, isGuest=false }) {
  return (
    <div style={{position:"relative",flexShrink:0}}>
      <div style={{width:size,height:size,borderRadius:"50%",background:avc(name),display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontFamily:"'DM Sans',sans-serif",fontSize:size*.36,fontWeight:600}}>
        {initials(name)}
      </div>
      {isGuest && (
        <div style={{position:"absolute",bottom:-1,right:-1,width:15,height:15,borderRadius:"50%",background:"#f59e0b",border:"2px solid #fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:7,color:"#fff",fontWeight:700}}>
          <Fa icon="user-tie" style={{fontSize:6}}/>
        </div>
      )}
    </div>
  );
}

function Spinner({ size=14, color="#fff" }) {
  return <span style={{display:"inline-block",width:size,height:size,border:`2px solid ${color}33`,borderTop:`2px solid ${color}`,borderRadius:"50%",animation:"spin .7s linear infinite",flexShrink:0}}/>;
}

function Toast({ msg, type, onClose }) {
  const ref = React.useRef();
  React.useEffect(() => { ref.current = setTimeout(onClose, 5000); return () => clearTimeout(ref.current); }, []);
  const cfg = {
    success: { bg:"#0f5132", icon:"circle-check" },
    error:   { bg:"#842029", icon:"circle-xmark" },
    pending: { bg:"#664d03", icon:"clock" },
    info:    { bg:"#1e40af", icon:"circle-info" },
  };
  const c = cfg[type] || cfg.info;
  return (
    <div style={{position:"fixed",bottom:24,right:24,zIndex:9999,background:c.bg,color:"#fff",borderRadius:10,padding:"13px 18px",fontFamily:"'DM Sans',sans-serif",fontSize:14,boxShadow:"0 8px 30px rgba(0,0,0,.25)",display:"flex",alignItems:"center",gap:10,maxWidth:400,animation:"slideUp .3s ease"}}>
      <Fa icon={c.icon} style={{fontSize:16,flexShrink:0}}/>{msg}
    </div>
  );
}

function Modal({ title, onClose, children }) {
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(10,15,30,.55)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:900,backdropFilter:"blur(3px)"}}>
      <div style={{background:"#fff",borderRadius:16,padding:"28px 30px",maxWidth:460,width:"90%",boxShadow:"0 20px 60px rgba(0,0,0,.2)",fontFamily:"'DM Sans',sans-serif",maxHeight:"85vh",overflowY:"auto"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:20}}>
          <h3 style={{fontSize:16,fontWeight:700,color:"#0f172a",margin:0}}>{title}</h3>
          <button onClick={onClose} style={{background:"#f1f5f9",border:"none",cursor:"pointer",width:28,height:28,borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",color:"#64748b"}}>
            <Fa icon="xmark"/>
          </button>
        </div>
        {children}
      </div>
    </div>
  );
}

function ConfirmModal({ message, onConfirm, onCancel }) {
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(10,15,30,.55)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:950,backdropFilter:"blur(3px)"}}>
      <div style={{background:"#fff",borderRadius:16,padding:"28px 32px",maxWidth:390,width:"90%",boxShadow:"0 20px 60px rgba(0,0,0,.2)",fontFamily:"'DM Sans',sans-serif"}}>
        <div style={{width:52,height:52,borderRadius:14,background:"#fef2f2",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 14px"}}>
          <Fa icon="triangle-exclamation" style={{fontSize:22,color:"#dc2626"}}/>
        </div>
        <p style={{textAlign:"center",color:"#374151",fontSize:14,marginBottom:24,lineHeight:1.6}}>{message}</p>
        <div style={{display:"flex",gap:10,justifyContent:"center"}}>
          <button onClick={onCancel}  style={S.btn("ghost")}>Cancel</button>
          <button onClick={onConfirm} style={S.btn("danger")}><Fa icon="trash"/> Remove</button>
        </div>
      </div>
    </div>
  );
}

const S = {
  btn(v, extra={}) {
    const b={border:"none",borderRadius:8,padding:"9px 18px",cursor:"pointer",fontFamily:"'DM Sans',sans-serif",fontSize:13,fontWeight:600,transition:"all .15s",display:"flex",alignItems:"center",gap:7,...extra};
    if(v==="primary") return {...b,background:"#1a56db",color:"#fff"};
    if(v==="danger")  return {...b,background:"#dc2626",color:"#fff"};
    if(v==="ghost")   return {...b,background:"#f1f5f9",color:"#374151"};
    if(v==="outline") return {...b,background:"transparent",color:"#1a56db",border:"1.5px solid #1a56db"};
    if(v==="amber")   return {...b,background:"#f59e0b",color:"#fff"};
    return b;
  }
};

// ─── App ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [authState,    setAuthState]    = useState("idle");
  const [account,      setAccount]      = useState(null);
  const [groups,       setGroups]       = useState([]);
  const [groupsLoad,   setGroupsLoad]   = useState(false);
  const [selectedG,    setSelectedG]    = useState(null);
  const [members,      setMembers]      = useState([]);
  const [membersLoad,  setMembersLoad]  = useState(false);
  const [search,       setSearch]       = useState("");
  const [memberSearch, setMemberSearch] = useState("");
  const [toast,        setToast]        = useState(null);
  const [modal,        setModal]        = useState(null);
  const [confirmData,  setConfirmData]  = useState(null);
  const [userSearch,   setUserSearch]   = useState("");
  const [userResults,  setUserResults]  = useState([]);
  const [userSearching,setUserSearching]= useState(false);
  const [extEmail,     setExtEmail]     = useState("");
  const [extName,      setExtName]      = useState("");
  const [extMsg,       setExtMsg]       = useState("");
  const [sendInvite,   setSendInvite]   = useState(false);
  const [editDesc,     setEditDesc]     = useState("");
  const [savingDesc,   setSavingDesc]   = useState(false);
  // Guest name editing
  const [editingGuest, setEditingGuest] = useState(null); // { id, displayName }
  const [editGuestName,setEditGuestName]= useState("");
  const [savingGuest,  setSavingGuest]  = useState(false);
  const [submitting,   setSubmitting]   = useState(false);
  const [loading,      setLoading]      = useState({});

  const showToast = (msg, type="success") => setToast({msg,type});
  const notConfigured = CLIENT_ID === "PASTE_YOUR_APP_CLIENT_ID_HERE";

  useEffect(() => {
    if (notConfigured) return;
    msalInstance.initialize().then(() => {
      msalInstance.handleRedirectPromise().then(result => {
        if (result) setAccount(result.account);
        else {
          const accounts = msalInstance.getAllAccounts();
          if (accounts.length) setAccount(accounts[0]);
        }
        setAuthState("ready");
      }).catch(() => setAuthState("error"));
    });
  }, []);

  async function signIn() {
    setAuthState("loading");
    try {
      const result = await msalInstance.loginPopup({ scopes: graphScopes });
      setAccount(result.account);
      setAuthState("ready");
    } catch { setAuthState("ready"); showToast("Sign in cancelled", "error"); }
  }

  function signOut() {
    msalInstance.logoutPopup();
    setAccount(null); setGroups([]); setSelectedG(null); setMembers([]);
  }

  const loadGroups = useCallback(async () => {
    setGroupsLoad(true);
    try {
      const data = await graphRequest("GET",
        "/me/ownedObjects/microsoft.graph.group?$filter=groupTypes/any(c:c+eq+'Unified')&$select=id,displayName,mail,description,visibility&$top=50"
      );
      setGroups(data?.value || []);
    } catch {
      try {
        const data = await graphRequest("GET", "/me/ownedObjects?$select=id,displayName,mail,groupTypes,description");
        setGroups((data?.value || []).filter(g => g["@odata.type"] === "#microsoft.graph.group"));
      } catch(e) { showToast(`Failed to load groups: ${e.message}`, "error"); }
    } finally { setGroupsLoad(false); }
  }, []);

  useEffect(() => { if (account) loadGroups(); }, [account]);

  const loadMembers = useCallback(async (groupId) => {
    setMembersLoad(true);
    try {
      const data = await graphRequest("GET",
        `/groups/${groupId}/members?$select=id,displayName,mail,userPrincipalName,userType,jobTitle,externalUserState`
      );
      setMembers(data?.value || []);
    } catch(e) { showToast(`Failed to load members: ${e.message}`, "error"); }
    finally { setMembersLoad(false); }
  }, []);

  function selectGroup(g) {
    setSelectedG(g); setMembers([]); setMemberSearch(""); setModal(null);
    loadMembers(g.id);
  }

  // ── User search ─────────────────────────────────────────────────────────────
  useEffect(() => {
    if (!userSearch || userSearch.length < 2) { setUserResults([]); return; }
    const t = setTimeout(async () => {
      setUserSearching(true);
      try {
        const token = await getToken();
        const res = await fetch(
          `https://graph.microsoft.com/v1.0/users?$search="displayName:${encodeURIComponent(userSearch)}"&$select=id,displayName,mail,userPrincipalName,jobTitle&$top=15`,
          { headers: { Authorization: `Bearer ${token}`, ConsistencyLevel: "eventual" } }
        );
        if (!res.ok) {
          const res2 = await fetch(
            `https://graph.microsoft.com/v1.0/users?$filter=startswith(displayName,'${encodeURIComponent(userSearch)}')&$select=id,displayName,mail,userPrincipalName,jobTitle&$top=15`,
            { headers: { Authorization: `Bearer ${token}` } }
          );
          const d2 = await res2.json();
          const memberIds = new Set(members.map(m => m.id));
          setUserResults((Array.isArray(d2?.value) ? d2.value : []).filter(u => !memberIds.has(u.id)));
          return;
        }
        const data = await res.json();
        const memberIds = new Set(members.map(m => m.id));
        setUserResults((Array.isArray(data?.value) ? data.value : []).filter(u => !memberIds.has(u.id)));
      } catch { setUserResults([]); }
      finally { setUserSearching(false); }
    }, 400);
    return () => clearTimeout(t);
  }, [userSearch, members]);

  // ── Add internal ────────────────────────────────────────────────────────────
  async function addInternalMember(user) {
    setLoading(p=>({...p,[`add-${user.id}`]:true}));
    try {
      await graphRequest("POST", `/groups/${selectedG.id}/members/$ref`, {
        "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${user.id}`
      });
      setMembers(p=>[...p, {...user, userType:"Member"}]);
      showToast(`${user.displayName} added`);
      setUserSearch(""); setUserResults([]);
    } catch(e) { showToast(`Failed: ${e.message}`, "error"); }
    finally { setLoading(p=>({...p,[`add-${user.id}`]:false})); }
  }

  // ── Invite external ─────────────────────────────────────────────────────────
  async function inviteExternal() {
    if (!extEmail.trim()) return;
    setSubmitting(true);
    try {
      const invite = await graphRequest("POST", "/invitations", {
        invitedUserEmailAddress: extEmail.trim(),
        invitedUserDisplayName:  extName.trim() || extEmail.trim(),
        inviteRedirectUrl:       "https://myapps.microsoft.com",
        sendInvitationMessage:   sendInvite,
        invitedUserMessageInfo:  sendInvite && extMsg.trim() ? { customizedMessageBody: extMsg.trim() } : undefined,
      });
      const guestId = invite.invitedUser.id;
      await graphRequest("POST", `/groups/${selectedG.id}/members/$ref`, {
        "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${guestId}`
      });
      setMembers(p=>[...p, {
        id: guestId, displayName: extName.trim() || extEmail.trim(),
        mail: extEmail.trim(), userType:"Guest", jobTitle:"", externalUserState:"PendingAcceptance",
      }]);
      showToast(sendInvite ? `Invitation sent to ${extEmail.trim()}` : `${extEmail.trim()} added silently`, "success");
      setExtEmail(""); setExtName(""); setExtMsg(""); setSendInvite(false);
      setModal(null);
    } catch(e) { showToast(`Failed: ${e.message}`, "error"); }
    finally { setSubmitting(false); }
  }

  // ── Remove member ───────────────────────────────────────────────────────────
  async function removeMember() {
    const { id, displayName } = confirmData;
    setConfirmData(null); setModal(null);
    setLoading(p=>({...p,[`rm-${id}`]:true}));
    try {
      await graphRequest("DELETE", `/groups/${selectedG.id}/members/${id}/$ref`);
      setMembers(p=>p.filter(m=>m.id!==id));
      showToast(`${displayName} removed`);
    } catch(e) { showToast(`Failed: ${e.message}`, "error"); }
    finally { setLoading(p=>({...p,[`rm-${id}`]:false})); }
  }

  // ── Save description ────────────────────────────────────────────────────────
  async function saveDescription() {
    setSavingDesc(true);
    try {
      await graphRequest("PATCH", `/groups/${selectedG.id}`, { description: editDesc });
      setGroups(p=>p.map(g=>g.id===selectedG.id?{...g,description:editDesc}:g));
      setSelectedG(p=>({...p,description:editDesc}));
      showToast("Description updated");
      setModal(null);
    } catch(e) { showToast(`Failed: ${e.message}`, "error"); }
    finally { setSavingDesc(false); }
  }

  // ── Save guest display name ──────────────────────────────────────────────────
  async function saveGuestName() {
    if (!editGuestName.trim() || !editingGuest) return;
    setSavingGuest(true);
    try {
      await graphRequestAsApp("PATCH", `/users/${editingGuest.id}`, { displayName: editGuestName.trim() });
      setMembers(p=>p.map(m=>m.id===editingGuest.id?{...m,displayName:editGuestName.trim()}:m));
      showToast("Guest name updated");
      setEditingGuest(null); setEditGuestName("");
      setModal(null);
    } catch(e) { showToast(`Failed: ${e.message}`, "error"); }
    finally { setSavingGuest(false); }
  }

  const filteredGroups  = groups.filter(g =>
    g.displayName?.toLowerCase().includes(search.toLowerCase()) ||
    g.mail?.toLowerCase().includes(search.toLowerCase())
  );
  const filteredMembers = members.filter(m =>
    m.displayName?.toLowerCase().includes(memberSearch.toLowerCase()) ||
    m.mail?.toLowerCase().includes(memberSearch.toLowerCase())
  );
  const guestCount    = members.filter(m=>m.userType==="Guest").length;
  const internalCount = members.filter(m=>m.userType!=="Guest").length;

  // ─── Not configured ──────────────────────────────────────────────────────────
  if (notConfigured) return (
    <div style={{minHeight:"100vh",background:"#f0f4f8",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'DM Sans',sans-serif",padding:24}}>
      <style>{css}</style>
      <div style={{background:"#fff",borderRadius:20,padding:"40px 44px",maxWidth:520,boxShadow:"0 4px 40px rgba(0,0,0,.1)"}}>
        <div style={{width:52,height:52,borderRadius:14,background:"#f1f5f9",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 16px"}}>
          <Fa icon="gear" style={{fontSize:22,color:"#64748b"}}/>
        </div>
        <h2 style={{fontFamily:"'Sora',sans-serif",fontSize:20,fontWeight:700,color:"#0f172a",textAlign:"center",marginBottom:8}}>Setup Required</h2>
        <p style={{color:"#64748b",fontSize:14,lineHeight:1.7,marginBottom:24,textAlign:"center"}}>
          Open <code style={{background:"#f1f5f9",padding:"2px 6px",borderRadius:4}}>src/auth.js</code> and replace the two placeholder values.
        </p>
        <div style={{background:"#f8fafc",borderRadius:12,padding:"18px 20px",border:"1px solid #e2e8f0",fontSize:13,lineHeight:2}}>
          <div><strong>CLIENT_ID</strong> — App Registration → Application (client) ID</div>
          <div><strong>TENANT_ID</strong> — App Registration → Directory (tenant) ID</div>
        </div>
      </div>
    </div>
  );

  // ─── Login ────────────────────────────────────────────────────────────────────
  if (!account) return (
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#0c1b3a 0%,#1a3a6b 60%,#0e2954 100%)",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'DM Sans',sans-serif"}}>
      <style>{css}</style>
      <div style={{position:"absolute",inset:0,opacity:.05,backgroundImage:"linear-gradient(rgba(255,255,255,.5) 1px,transparent 1px),linear-gradient(90deg,rgba(255,255,255,.5) 1px,transparent 1px)",backgroundSize:"40px 40px"}}/>
      <div style={{background:"rgba(255,255,255,.08)",backdropFilter:"blur(20px)",border:"1px solid rgba(255,255,255,.14)",borderRadius:24,padding:"52px 48px",width:420,textAlign:"center",boxShadow:"0 40px 80px rgba(0,0,0,.4)",position:"relative",animation:"fadeIn .5s ease"}}>
        <div style={{width:64,height:64,borderRadius:16,background:"linear-gradient(135deg,#0078d4,#106ebe)",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 20px",boxShadow:"0 8px 24px rgba(0,120,212,.4)"}}>
          <Fa icon="users" style={{fontSize:26,color:"#fff"}}/>
        </div>
        <h1 style={{fontFamily:"'Sora',sans-serif",fontSize:24,fontWeight:700,color:"#fff",marginBottom:8}}>Group Manager</h1>
        <p style={{color:"rgba(255,255,255,.5)",fontSize:13,lineHeight:1.7,marginBottom:32}}>
          Manage your Microsoft 365 Groups<br/>Add members & invite external guests
        </p>
        <button onClick={signIn} disabled={authState==="loading"}
          style={{...S.btn("primary"),width:"100%",justifyContent:"center",padding:"14px 22px",fontSize:15,borderRadius:12,background:authState==="loading"?"#1a3a6b":"#0078d4"}}>
          {authState==="loading"
            ?<><Spinner/><span style={{animation:"pulse 1s infinite"}}>Signing in…</span></>
            :<><Fa icon="lock" style={{fontSize:14}}/> Sign in with Microsoft 365</>}
        </button>
        <p style={{color:"rgba(255,255,255,.2)",fontSize:11,marginTop:16}}>Delegated OAuth — only your groups, only your permissions</p>
      </div>
    </div>
  );

  // ─── Main app ─────────────────────────────────────────────────────────────────
  return (
    <div style={{minHeight:"100vh",background:"#f0f4f8",fontFamily:"'DM Sans',sans-serif"}}>
      <style>{css}</style>

      {/* Header */}
      <div style={{background:"#fff",borderBottom:"1px solid #e2e8f0",padding:"0 24px",height:56,display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:100,boxShadow:"0 1px 8px rgba(0,0,0,.06)"}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:30,height:30,borderRadius:8,background:"linear-gradient(135deg,#0078d4,#1a56db)",display:"flex",alignItems:"center",justifyContent:"center"}}>
            <Fa icon="users" style={{fontSize:13,color:"#fff"}}/>
          </div>
          <span style={{fontFamily:"'Sora',sans-serif",fontWeight:700,fontSize:15,color:"#0f172a"}}>Group Manager</span>
          <span style={{background:"#e8f0fe",color:"#1a56db",fontSize:10,fontWeight:700,borderRadius:5,padding:"2px 7px"}}>M365</span>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <span style={{color:"#64748b",fontSize:13}}>{account.name||account.username}</span>
          <Avatar name={account.name||account.username} size={28}/>
          <button onClick={signOut} style={{...S.btn("ghost"),padding:"5px 12px",fontSize:12}}>
            <Fa icon="right-from-bracket" style={{fontSize:12}}/> Sign out
          </button>
        </div>
      </div>

      <div style={{display:"flex",height:"calc(100vh - 56px)"}}>

        {/* Sidebar */}
        <div style={{width:280,background:"#fff",borderRight:"1px solid #e2e8f0",display:"flex",flexDirection:"column",flexShrink:0}}>
          <div style={{padding:"16px 12px 10px"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
              <p style={{fontSize:11,fontWeight:700,color:"#94a3b8",letterSpacing:".08em",textTransform:"uppercase"}}>My Groups</p>
              <button onClick={loadGroups} disabled={groupsLoad} title="Refresh"
                style={{background:"none",border:"none",cursor:"pointer",color:"#94a3b8",padding:4,borderRadius:5}}>
                {groupsLoad?<Spinner size={12} color="#94a3b8"/>:<Fa icon="rotate" style={{fontSize:13}}/>}
              </button>
            </div>
            <div style={{position:"relative"}}>
              <Fa icon="magnifying-glass" style={{position:"absolute",left:9,top:"50%",transform:"translateY(-50%)",fontSize:12,color:"#94a3b8"}}/>
              <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search groups…"
                style={{width:"100%",padding:"8px 10px 8px 28px",borderRadius:8,border:"1.5px solid #e2e8f0",fontSize:13,outline:"none",fontFamily:"'DM Sans',sans-serif",background:"#f8fafc",color:"#1e293b"}}/>
            </div>
          </div>

          <div style={{flex:1,overflowY:"auto",padding:"0 7px 12px"}}>
            {groupsLoad && (
              <div style={{padding:"20px 12px",display:"flex",alignItems:"center",gap:8,color:"#94a3b8",fontSize:13}}>
                <Spinner size={13} color="#94a3b8"/> Loading groups…
              </div>
            )}
            {!groupsLoad && filteredGroups.length===0 && (
              <p style={{textAlign:"center",color:"#94a3b8",fontSize:13,paddingTop:24}}>No groups found</p>
            )}
            {filteredGroups.map(g=>(
              <div key={g.id} className="dlc"
                onClick={()=>selectGroup(g)}
                style={{padding:"10px 9px",borderRadius:10,marginBottom:3,cursor:"pointer",transition:"all .15s",background:selectedG?.id===g.id?"#eff6ff":"transparent",border:`1.5px solid ${selectedG?.id===g.id?"#bfdbfe":"transparent"}`}}>
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <div style={{width:32,height:32,borderRadius:8,flexShrink:0,background:selectedG?.id===g.id?"#dbeafe":"#eff6ff",display:"flex",alignItems:"center",justifyContent:"center"}}>
                    <Fa icon="users" style={{fontSize:13,color:selectedG?.id===g.id?"#1a56db":"#64a4d8"}}/>
                  </div>
                  <div style={{minWidth:0}}>
                    <p style={{fontWeight:600,fontSize:13,color:"#1e293b",lineHeight:1.35,margin:0}}>{g.displayName}</p>
                    <p style={{fontSize:11,color:"#94a3b8",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis",margin:0}}>{g.mail||"No email"}</p>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* Main panel */}
        <div style={{flex:1,overflowY:"auto",padding:"24px 28px"}}>
          {!selectedG ? (
            <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",height:"100%",color:"#94a3b8",animation:"fadeIn .4s ease",textAlign:"center"}}>
              <div style={{width:72,height:72,borderRadius:20,background:"#e2e8f0",display:"flex",alignItems:"center",justifyContent:"center",marginBottom:16,opacity:.6}}>
                <Fa icon="users" style={{fontSize:30,color:"#94a3b8"}}/>
              </div>
              <p style={{fontSize:17,fontWeight:600,color:"#64748b"}}>Select a Group</p>
              <p style={{fontSize:13,marginTop:5}}>Choose a group from the sidebar to manage members</p>
            </div>
          ) : (
            <div style={{maxWidth:760,animation:"slideUp .3s ease"}}>

              {/* Group header */}
              <div style={{background:"linear-gradient(135deg,#1a56db 0%,#0078d4 100%)",borderRadius:18,padding:"22px 26px",marginBottom:20,color:"#fff",boxShadow:"0 8px 32px rgba(26,86,219,.25)"}}>
                <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between"}}>
                  <div style={{minWidth:0}}>
                    <div style={{display:"flex",alignItems:"center",gap:9,marginBottom:4}}>
                      <Fa icon="users" style={{fontSize:18,opacity:.9}}/>
                      <h2 style={{fontFamily:"'Sora',sans-serif",fontSize:18,fontWeight:700,margin:0}}>{selectedG.displayName}</h2>
                    </div>
                    <p style={{opacity:.75,fontSize:12,margin:0}}>{selectedG.mail||"No email address"}</p>
                    <div style={{display:"flex",alignItems:"center",gap:8,marginTop:6}}>
                      {selectedG.description
                        ? <p style={{opacity:.6,fontSize:12,margin:0}}>{selectedG.description}</p>
                        : <p style={{opacity:.4,fontSize:12,margin:0,fontStyle:"italic"}}>No description</p>
                      }
                      <button onClick={()=>{setEditDesc(selectedG.description||"");setModal("editDesc");}}
                        style={{background:"rgba(255,255,255,.15)",border:"1px solid rgba(255,255,255,.25)",borderRadius:6,padding:"2px 8px",fontSize:11,color:"#fff",cursor:"pointer",fontFamily:"'DM Sans',sans-serif",display:"flex",alignItems:"center",gap:5,flexShrink:0}}>
                        <Fa icon="pen" style={{fontSize:9}}/> Edit
                      </button>
                    </div>
                  </div>
                  <div style={{display:"flex",gap:8,flexShrink:0,marginLeft:12}}>
                    <div style={{background:"rgba(255,255,255,.18)",borderRadius:10,padding:"8px 14px",textAlign:"center"}}>
                      <div style={{fontSize:18,fontWeight:700}}>{internalCount}</div>
                      <div style={{fontSize:10,opacity:.8}}>Internal</div>
                    </div>
                    <div style={{background:"rgba(245,158,11,.3)",borderRadius:10,padding:"8px 14px",textAlign:"center",border:"1px solid rgba(245,158,11,.4)"}}>
                      <div style={{fontSize:18,fontWeight:700}}>{guestCount}</div>
                      <div style={{fontSize:10,opacity:.8}}>Guests</div>
                    </div>
                  </div>
                </div>
                <div style={{marginTop:12,display:"flex",gap:7,flexWrap:"wrap"}}>
                  <span style={{background:"rgba(255,255,255,.2)",borderRadius:20,padding:"3px 11px",fontSize:11,fontWeight:600,display:"flex",alignItems:"center",gap:5}}>
                    <Fa icon="circle-check" style={{fontSize:10}}/> You are an owner
                  </span>
                  <span style={{background:"rgba(255,255,255,.12)",borderRadius:20,padding:"3px 11px",fontSize:11,opacity:.8,display:"flex",alignItems:"center",gap:5}}>
                    <Fa icon="microsoft" style={{fontSize:10}}/> Microsoft 365 Group
                  </span>
                </div>
              </div>

              {/* Action buttons */}
              <div style={{display:"flex",gap:10,marginBottom:20,flexWrap:"wrap"}}>
                <button onClick={()=>{setModal("addInternal");setUserSearch("");setUserResults([]);}}
                  style={S.btn("primary")}>
                  <Fa icon="user-plus" style={{fontSize:12}}/> Add Internal Member
                </button>
                <button onClick={()=>{setModal("addExternal");setExtEmail("");setExtName("");setExtMsg("");setSendInvite(false);}}
                  style={S.btn("amber")}>
                  <Fa icon="envelope" style={{fontSize:12}}/> Invite External Guest
                </button>
                <button onClick={()=>loadMembers(selectedG.id)} disabled={membersLoad}
                  style={S.btn("ghost")}>
                  {membersLoad?<Spinner size={12} color="#64748b"/>:<Fa icon="rotate" style={{fontSize:12}}/>}
                  Refresh
                </button>
              </div>

              {/* Members table */}
              <div style={{background:"#fff",borderRadius:16,border:"1px solid #e2e8f0",overflow:"hidden",boxShadow:"0 2px 12px rgba(0,0,0,.04)"}}>
                <div style={{padding:"14px 18px",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",justifyContent:"space-between",gap:12}}>
                  <h3 style={{fontSize:14,fontWeight:700,color:"#0f172a",margin:0}}>
                    Members <span style={{color:"#94a3b8",fontWeight:400,fontSize:13}}>({members.length})</span>
                  </h3>
                  <div style={{position:"relative",flex:1,maxWidth:260}}>
                    <Fa icon="magnifying-glass" style={{position:"absolute",left:8,top:"50%",transform:"translateY(-50%)",fontSize:11,color:"#94a3b8"}}/>
                    <input value={memberSearch} onChange={e=>setMemberSearch(e.target.value)} placeholder="Filter members…"
                      style={{width:"100%",padding:"7px 10px 7px 24px",borderRadius:7,border:"1.5px solid #e2e8f0",fontSize:12,outline:"none",fontFamily:"'DM Sans',sans-serif"}}/>
                  </div>
                </div>

                {membersLoad ? (
                  <div style={{padding:"30px 18px",display:"flex",alignItems:"center",gap:8,color:"#94a3b8",fontSize:13}}>
                    <Spinner size={14} color="#94a3b8"/> Loading members…
                  </div>
                ) : filteredMembers.length===0 ? (
                  <div style={{padding:"32px 18px",textAlign:"center",color:"#94a3b8"}}>
                    <Fa icon="users" style={{fontSize:28,opacity:.3,display:"block",marginBottom:8}}/>
                    <p style={{fontSize:13}}>{memberSearch?"No members match your search":"No members yet"}</p>
                  </div>
                ) : (
                  <>
                    {guestCount>0 && !memberSearch && (
                      <div style={{padding:"8px 18px",background:"#fffbeb",borderBottom:"1px solid #fde68a",display:"flex",alignItems:"center",gap:7,fontSize:12,color:"#92400e"}}>
                        <Fa icon="user-tie" style={{fontSize:11,color:"#f59e0b"}}/>
                        {guestCount} external guest{guestCount>1?"s":""} — pending acceptance shown below. Click <Fa icon="pen" style={{fontSize:10}}/> to edit their display name.
                      </div>
                    )}
                    {filteredMembers.map((m,i)=>{
                      const isGuest   = m.userType==="Guest";
                      const isPending = m.externalUserState==="PendingAcceptance";
                      return (
                        <div key={m.id} className="mrow"
                          style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"11px 18px",transition:"background .12s",borderBottom:i<filteredMembers.length-1?"1px solid #f8fafc":"none",opacity:loading[`rm-${m.id}`]?.4:1,background:isGuest?"#fffdf5":"#fff"}}>
                          <div style={{display:"flex",alignItems:"center",gap:10,minWidth:0}}>
                            <Avatar name={m.displayName||m.mail} size={34} isGuest={isGuest}/>
                            <div style={{minWidth:0}}>
                              <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
                                <p style={{fontSize:13,fontWeight:600,color:"#1e293b",margin:0}}>{m.displayName||m.mail}</p>
                                {isGuest   && <span style={{fontSize:9,background:"#fef3c7",color:"#92400e",borderRadius:4,padding:"1px 5px",fontWeight:700,display:"flex",alignItems:"center",gap:3}}><Fa icon="user-tie" style={{fontSize:7}}/>GUEST</span>}
                                {isPending && <span style={{fontSize:9,background:"#fce7f3",color:"#9d174d",borderRadius:4,padding:"1px 5px",fontWeight:700,display:"flex",alignItems:"center",gap:3}}><Fa icon="clock" style={{fontSize:7}}/>PENDING</span>}
                              </div>
                              <p style={{fontSize:11,color:"#94a3b8",margin:0}}>{m.mail||m.userPrincipalName}</p>
                            </div>
                          </div>
                          <div style={{display:"flex",alignItems:"center",gap:6,flexShrink:0}}>
                            {m.jobTitle && <span style={{fontSize:11,color:"#64748b",background:"#f1f5f9",borderRadius:5,padding:"2px 7px"}}>{m.jobTitle}</span>}
                            {/* Edit name — guests only */}
                            {isGuest && (
                              <button title="Edit display name" className="ebtn"
                                onClick={()=>{setEditingGuest({id:m.id,displayName:m.displayName||m.mail});setEditGuestName(m.displayName||"");setModal("editGuest");}}
                                style={{background:"transparent",border:"none",cursor:"pointer",width:28,height:28,borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",color:"#94a3b8",transition:"all .12s"}}>
                                <Fa icon="pen" style={{fontSize:11}}/>
                              </button>
                            )}
                            <button title="Remove" className="ibtn"
                              disabled={loading[`rm-${m.id}`]}
                              onClick={()=>{setConfirmData({id:m.id,displayName:m.displayName||m.mail,isGuest});setModal("confirm");}}
                              style={{background:"transparent",border:"none",cursor:"pointer",width:28,height:28,borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",color:"#cbd5e1",transition:"all .12s"}}>
                              {loading[`rm-${m.id}`]?<Spinner size={12} color="#94a3b8"/>:<Fa icon="trash" style={{fontSize:11}}/>}
                            </button>
                          </div>
                        </div>
                      );
                    })}
                  </>
                )}
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Add Internal Modal */}
      {modal==="addInternal" && (
        <Modal title={<><Fa icon="user-plus" style={{marginRight:8,color:"#1a56db"}}/>Add Internal Member</>} onClose={()=>setModal(null)}>
          <p style={{fontSize:13,color:"#64748b",marginBottom:14}}>Search for users in your organization.</p>
          <div style={{position:"relative",marginBottom:12}}>
            <Fa icon="magnifying-glass" style={{position:"absolute",left:9,top:"50%",transform:"translateY(-50%)",fontSize:12,color:"#94a3b8"}}/>
            <input value={userSearch} onChange={e=>setUserSearch(e.target.value)}
              placeholder="Search by name or email…" autoFocus
              style={{width:"100%",padding:"9px 10px 9px 28px",borderRadius:8,border:"1.5px solid #d1d5db",fontSize:13,outline:"none",fontFamily:"'DM Sans',sans-serif",boxSizing:"border-box"}}/>
          </div>
          {userSearching && <div style={{display:"flex",alignItems:"center",gap:7,color:"#94a3b8",fontSize:13,padding:"8px 4px"}}><Spinner size={12} color="#94a3b8"/> Searching…</div>}
          {(Array.isArray(userResults)?userResults:[]).map(u=>(
            <div key={u.id} className="arow" style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"9px 8px",borderRadius:8,transition:"background .12s"}}>
              <div style={{display:"flex",alignItems:"center",gap:9}}>
                <Avatar name={u.displayName} size={32}/>
                <div>
                  <p style={{fontSize:13,fontWeight:600,color:"#1e293b",margin:0}}>{u.displayName}</p>
                  <p style={{fontSize:11,color:"#94a3b8",margin:0}}>{u.mail}{u.jobTitle?` · ${u.jobTitle}`:""}</p>
                </div>
              </div>
              <button onClick={()=>addInternalMember(u)} disabled={loading[`add-${u.id}`]}
                style={{...S.btn("outline"),padding:"5px 12px",fontSize:12,borderRadius:6}}>
                {loading[`add-${u.id}`]?<><Spinner size={11} color="#1a56db"/> Adding…</>:<><Fa icon="plus" style={{fontSize:10}}/> Add</>}
              </button>
            </div>
          ))}
          {userSearch.length>=2 && !userSearching && Array.isArray(userResults) && userResults.length===0 && (
            <p style={{fontSize:13,color:"#94a3b8",textAlign:"center",padding:"12px 0"}}>No users found</p>
          )}
        </Modal>
      )}

      {/* Invite External Modal */}
      {modal==="addExternal" && (
        <Modal title={<><Fa icon="envelope" style={{marginRight:8,color:"#f59e0b"}}/>Invite External Guest</>} onClose={()=>setModal(null)}>
          <p style={{fontSize:13,color:"#64748b",marginBottom:18,lineHeight:1.6}}>
            Add an external user to this group. They'll receive emails immediately regardless of whether you send an invitation.
          </p>
          <div style={{marginBottom:12}}>
            <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:5}}>
              <Fa icon="at" style={{marginRight:5,color:"#64748b"}}/>Email address *
            </label>
            <input value={extEmail} onChange={e=>setExtEmail(e.target.value)}
              placeholder="external@company.com" type="email" autoFocus
              style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"1.5px solid #d1d5db",fontSize:13,outline:"none",fontFamily:"'DM Sans',sans-serif",boxSizing:"border-box"}}/>
          </div>
          <div style={{marginBottom:12}}>
            <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:5}}>
              <Fa icon="id-card" style={{marginRight:5,color:"#64748b"}}/>Display name <span style={{fontWeight:400,color:"#94a3b8"}}>(optional)</span>
            </label>
            <input value={extName} onChange={e=>setExtName(e.target.value)} placeholder="Jane Smith"
              style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"1.5px solid #d1d5db",fontSize:13,outline:"none",fontFamily:"'DM Sans',sans-serif",boxSizing:"border-box"}}/>
          </div>

          {/* Send invite toggle */}
          <div onClick={()=>setSendInvite(p=>!p)}
            style={{display:"flex",alignItems:"flex-start",gap:10,padding:"12px 14px",borderRadius:9,border:`1.5px solid ${sendInvite?"#f59e0b":"#e2e8f0"}`,background:sendInvite?"#fffbeb":"#f8fafc",marginBottom:sendInvite?12:18,cursor:"pointer",transition:"all .15s"}}>
            <div style={{width:18,height:18,borderRadius:4,border:`2px solid ${sendInvite?"#f59e0b":"#cbd5e1"}`,background:sendInvite?"#f59e0b":"#fff",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,marginTop:1,transition:"all .15s"}}>
              {sendInvite && <Fa icon="check" style={{fontSize:9,color:"#fff"}}/>}
            </div>
            <div>
              <p style={{fontSize:13,fontWeight:600,color:sendInvite?"#92400e":"#374151",margin:0,marginBottom:2}}>
                <Fa icon="envelope" style={{marginRight:5,fontSize:11}}/> Send invitation email
              </p>
              <p style={{fontSize:11,color:sendInvite?"#a16207":"#94a3b8",margin:0,lineHeight:1.5}}>
                {sendInvite?"Microsoft will email the guest. They appear as Pending until accepted.":"Guest added silently — no email sent. Still receives group emails immediately."}
              </p>
            </div>
          </div>

          {sendInvite && (
            <div style={{marginBottom:18}}>
              <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:5}}>
                <Fa icon="message" style={{marginRight:5,color:"#64748b"}}/>Personal message <span style={{fontWeight:400,color:"#94a3b8"}}>(optional)</span>
              </label>
              <textarea value={extMsg} onChange={e=>setExtMsg(e.target.value)}
                placeholder="Hi! I'm inviting you to our Microsoft 365 group…" rows={3}
                style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"1.5px solid #d1d5db",fontSize:13,outline:"none",fontFamily:"'DM Sans',sans-serif",boxSizing:"border-box",resize:"vertical"}}/>
            </div>
          )}

          <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}>
            <button onClick={()=>setModal(null)} style={S.btn("ghost")}>Cancel</button>
            <button onClick={inviteExternal} disabled={!extEmail.trim()||submitting}
              style={{...S.btn("amber"),opacity:!extEmail.trim()||submitting?.6:1}}>
              {submitting?<><Spinner size={13}/> Adding…</>
                :sendInvite?<><Fa icon="envelope" style={{fontSize:11}}/> Send Invitation</>
                :<><Fa icon="user-plus" style={{fontSize:11}}/> Add Silently</>}
            </button>
          </div>
        </Modal>
      )}

      {/* Edit Guest Name Modal */}
      {modal==="editGuest" && editingGuest && (
        <Modal title={<><Fa icon="pen" style={{marginRight:8,color:"#1a56db"}}/>Edit Guest Display Name</>} onClose={()=>{setModal(null);setEditingGuest(null);}}>
          <p style={{fontSize:13,color:"#64748b",marginBottom:6,lineHeight:1.6}}>
            Editing display name for:
          </p>
          <div style={{display:"flex",alignItems:"center",gap:9,padding:"10px 12px",background:"#f8fafc",borderRadius:8,marginBottom:16,border:"1px solid #e2e8f0"}}>
            <Avatar name={editingGuest.displayName} size={32} isGuest/>
            <div>
              <p style={{fontSize:13,fontWeight:600,color:"#1e293b",margin:0}}>{editingGuest.displayName}</p>
            </div>
          </div>
          <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:6}}>
            <Fa icon="id-card" style={{marginRight:5,color:"#64748b"}}/>New display name
          </label>
          <input value={editGuestName} onChange={e=>setEditGuestName(e.target.value)}
            onKeyDown={e=>e.key==="Enter"&&saveGuestName()}
            placeholder="Full name…" autoFocus
            style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"1.5px solid #d1d5db",fontSize:13,outline:"none",fontFamily:"'DM Sans',sans-serif",boxSizing:"border-box",marginBottom:8}}/>
          <p style={{fontSize:11,color:"#94a3b8",marginBottom:18,lineHeight:1.5}}>
            <Fa icon="circle-info" style={{marginRight:4}}/>This updates the guest's display name in your Entra directory, visible across all Microsoft 365 apps.
          </p>
          <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}>
            <button onClick={()=>{setModal(null);setEditingGuest(null);}} style={S.btn("ghost")}>Cancel</button>
            <button onClick={saveGuestName} disabled={!editGuestName.trim()||savingGuest}
              style={{...S.btn("primary"),opacity:!editGuestName.trim()||savingGuest?.6:1}}>
              {savingGuest?<><Spinner size={13}/> Saving…</>:<><Fa icon="floppy-disk" style={{fontSize:11}}/> Save name</>}
            </button>
          </div>
        </Modal>
      )}

      {/* Edit Description Modal */}
      {modal==="editDesc" && (
        <Modal title={<><Fa icon="pen" style={{marginRight:8,color:"#1a56db"}}/>Edit Group Description</>} onClose={()=>setModal(null)}>
          <p style={{fontSize:13,color:"#64748b",marginBottom:14,lineHeight:1.6}}>
            This description appears in the group header and in Microsoft 365 directory listings.
          </p>
          <textarea value={editDesc} onChange={e=>setEditDesc(e.target.value)}
            placeholder="Enter a description…" rows={4} autoFocus
            style={{width:"100%",padding:"10px 12px",borderRadius:8,border:"1.5px solid #d1d5db",fontSize:13,outline:"none",fontFamily:"'DM Sans',sans-serif",boxSizing:"border-box",resize:"vertical",lineHeight:1.6}}/>
          <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:16}}>
            <button onClick={()=>setModal(null)} style={S.btn("ghost")}>Cancel</button>
            <button onClick={saveDescription} disabled={savingDesc}
              style={{...S.btn("primary"),opacity:savingDesc?.6:1}}>
              {savingDesc?<><Spinner size={13}/> Saving…</>:<><Fa icon="floppy-disk" style={{fontSize:11}}/> Save description</>}
            </button>
          </div>
        </Modal>
      )}

      {/* Confirm remove */}
      {modal==="confirm" && confirmData && (
        <ConfirmModal
          message={`Remove ${confirmData.displayName} from ${selectedG?.displayName}?${confirmData.isGuest?" This removes their guest access to this group.":""}`}
          onConfirm={removeMember}
          onCancel={()=>{setModal(null);setConfirmData(null);}}
        />
      )}

      {toast && <Toast msg={toast.msg} type={toast.type} onClose={()=>setToast(null)}/>}
    </div>
  );
}

const css = `
  @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Sora:wght@600;700&display=swap');
  @import url('https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.1/css/all.min.css');
  *{box-sizing:border-box;margin:0;padding:0}
  @keyframes slideUp{from{transform:translateY(14px);opacity:0}to{transform:translateY(0);opacity:1}}
  @keyframes fadeIn{from{opacity:0}to{opacity:1}}
  @keyframes spin{to{transform:rotate(360deg)}}
  @keyframes pulse{0%,100%{opacity:1}50%{opacity:.5}}
  .dlc:hover{background:#f8faff!important;box-shadow:0 2px 12px rgba(26,86,219,.08)!important}
  .mrow:hover{background:#f8faff!important}
  .arow:hover{background:#f0f4ff!important}
  .ibtn:hover{background:#fee2e2!important;color:#dc2626!important}
  .ebtn:hover{background:#eff6ff!important;color:#1a56db!important}
  ::-webkit-scrollbar{width:5px}::-webkit-scrollbar-track{background:transparent}::-webkit-scrollbar-thumb{background:#cbd5e1;border-radius:3px}
`;
