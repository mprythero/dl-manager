import React, { useState, useEffect, useCallback } from "react";
import { msalInstance, graphScopes, graphRequest, CLIENT_ID } from "./auth";

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
        <div style={{position:"absolute",bottom:-1,right:-1,width:14,height:14,borderRadius:"50%",background:"#f59e0b",border:"2px solid #fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:7,color:"#fff",fontWeight:700}}>G</div>
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
  const bg={success:"#0f5132",error:"#842029",pending:"#664d03",info:"#1e40af"};
  const ic={success:"✓",error:"✕",pending:"⏳",info:"ℹ"};
  return (
    <div style={{position:"fixed",bottom:24,right:24,zIndex:9999,background:bg[type]||bg.info,color:"#fff",borderRadius:10,padding:"12px 18px",fontFamily:"'DM Sans',sans-serif",fontSize:14,boxShadow:"0 8px 30px rgba(0,0,0,.25)",display:"flex",alignItems:"center",gap:10,maxWidth:400,animation:"slideUp .3s ease"}}>
      <span>{ic[type]||"ℹ"}</span>{msg}
    </div>
  );
}

function Modal({ title, onClose, children }) {
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(10,15,30,.55)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:900,backdropFilter:"blur(3px)"}}>
      <div style={{background:"#fff",borderRadius:16,padding:"28px 30px",maxWidth:460,width:"90%",boxShadow:"0 20px 60px rgba(0,0,0,.2)",fontFamily:"'DM Sans',sans-serif",maxHeight:"85vh",overflowY:"auto"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:20}}>
          <h3 style={{fontSize:16,fontWeight:700,color:"#0f172a",margin:0}}>{title}</h3>
          <button onClick={onClose} style={{background:"none",border:"none",cursor:"pointer",fontSize:18,color:"#94a3b8",padding:4}}>✕</button>
        </div>
        {children}
      </div>
    </div>
  );
}

function ConfirmModal({ message, onConfirm, onCancel, confirmLabel="Remove", danger=true }) {
  return (
    <div style={{position:"fixed",inset:0,background:"rgba(10,15,30,.55)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:950,backdropFilter:"blur(3px)"}}>
      <div style={{background:"#fff",borderRadius:16,padding:"28px 32px",maxWidth:390,width:"90%",boxShadow:"0 20px 60px rgba(0,0,0,.2)",fontFamily:"'DM Sans',sans-serif"}}>
        <div style={{fontSize:34,textAlign:"center",marginBottom:10}}>⚠️</div>
        <p style={{textAlign:"center",color:"#374151",fontSize:14,marginBottom:24,lineHeight:1.6}}>{message}</p>
        <div style={{display:"flex",gap:10,justifyContent:"center"}}>
          <button onClick={onCancel}  style={S.btn("ghost")}>Cancel</button>
          <button onClick={onConfirm} style={S.btn(danger?"danger":"primary")}>{confirmLabel}</button>
        </div>
      </div>
    </div>
  );
}

const S = {
  btn(v, extra={}) {
    const b={border:"none",borderRadius:8,padding:"9px 18px",cursor:"pointer",fontFamily:"'DM Sans',sans-serif",fontSize:13,fontWeight:600,transition:"all .15s",display:"flex",alignItems:"center",gap:6,...extra};
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
  const [authState,   setAuthState]   = useState("idle"); // idle|loading|ready|error
  const [account,     setAccount]     = useState(null);
  const [groups,      setGroups]      = useState([]);
  const [groupsLoad,  setGroupsLoad]  = useState(false);
  const [selectedG,   setSelectedG]   = useState(null);
  const [members,     setMembers]     = useState([]);
  const [membersLoad, setMembersLoad] = useState(false);
  const [search,      setSearch]      = useState("");
  const [memberSearch,setMemberSearch]= useState("");
  const [toast,       setToast]       = useState(null);
  const [modal,       setModal]       = useState(null); // null | "addInternal" | "addExternal" | "confirm"
  const [confirmData, setConfirmData] = useState(null);
  const [userSearch,  setUserSearch]  = useState("");
  const [userResults, setUserResults] = useState([]);
  const [userSearching,setUserSearching]=useState(false);
  const [extEmail,    setExtEmail]    = useState("");
  const [extName,     setExtName]     = useState("");
  const [extMsg,      setExtMsg]      = useState("");
  const [sendInvite,  setSendInvite]  = useState(false);
  const [editDesc,    setEditDesc]    = useState("");
  const [savingDesc,  setSavingDesc]  = useState(false);
  const [submitting,  setSubmitting]  = useState(false);
  const [loading,     setLoading]     = useState({});

  const showToast = (msg, type="success") => setToast({msg,type});

  const notConfigured = CLIENT_ID === "PASTE_YOUR_APP_CLIENT_ID_HERE";

  // ── Init MSAL ───────────────────────────────────────────────────────────────
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
      }).catch(e => { setAuthState("error"); console.error(e); });
    });
  }, []);

  // ── Sign in ─────────────────────────────────────────────────────────────────
  async function signIn() {
    setAuthState("loading");
    try {
      const result = await msalInstance.loginPopup({ scopes: graphScopes });
      setAccount(result.account);
      setAuthState("ready");
    } catch(e) {
      setAuthState("ready");
      showToast("Sign in cancelled or failed", "error");
    }
  }

  // ── Sign out ────────────────────────────────────────────────────────────────
  function signOut() {
    msalInstance.logoutPopup();
    setAccount(null);
    setGroups([]);
    setSelectedG(null);
    setMembers([]);
  }

  // ── Load groups owned by user ───────────────────────────────────────────────
  const loadGroups = useCallback(async () => {
    setGroupsLoad(true);
    try {
      const data = await graphRequest("GET",
        "/me/ownedObjects/microsoft.graph.group?$filter=groupTypes/any(c:c+eq+'Unified')&$select=id,displayName,mail,description,membershipRule,visibility&$top=50"
      );
      setGroups(data?.value || []);
    } catch(e) {
      // Fall back — get all owned objects and filter
      try {
        const data = await graphRequest("GET", "/me/ownedObjects?$select=id,displayName,mail,groupTypes,description");
        const filtered = (data?.value || []).filter(g => g["@odata.type"] === "#microsoft.graph.group");
        setGroups(filtered);
      } catch(e2) {
        showToast(`Failed to load groups: ${e2.message}`, "error");
      }
    } finally {
      setGroupsLoad(false);
    }
  }, []);

  useEffect(() => { if (account) loadGroups(); }, [account]);

  // ── Load members of a group ─────────────────────────────────────────────────
  const loadMembers = useCallback(async (groupId) => {
    setMembersLoad(true);
    try {
      const data = await graphRequest("GET",
        `/groups/${groupId}/members?$select=id,displayName,mail,userPrincipalName,userType,jobTitle,externalUserState`
      );
      setMembers(data?.value || []);
    } catch(e) {
      showToast(`Failed to load members: ${e.message}`, "error");
    } finally {
      setMembersLoad(false);
    }
  }, []);

  function selectGroup(g) {
    setSelectedG(g);
    setMembers([]);
    setMemberSearch("");
    setModal(null);
    loadMembers(g.id);
  }

  // ── Search internal users ───────────────────────────────────────────────────
  useEffect(() => {
    if (!userSearch || userSearch.length < 2) { setUserResults([]); return; }
    const t = setTimeout(async () => {
      setUserSearching(true);
      try {
        const data = await graphRequest("GET",
          `/users?$filter=startswith(displayName,'${encodeURIComponent(userSearch)}') or startswith(mail,'${encodeURIComponent(userSearch)}')&$select=id,displayName,mail,userPrincipalName,jobTitle&$top=10`
        );
        // Filter out already-members
        const memberIds = new Set(members.map(m => m.id));
        setUserResults((data?.value || []).filter(u => !memberIds.has(u.id)));
      } catch(e) { setUserResults([]); }
      finally { setUserSearching(false); }
    }, 400);
    return () => clearTimeout(t);
  }, [userSearch, members]);

  // ── Add internal member ─────────────────────────────────────────────────────
  async function addInternalMember(user) {
    setLoading(p=>({...p,[`add-${user.id}`]:true}));
    try {
      await graphRequest("POST", `/groups/${selectedG.id}/members/$ref`, {
        "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${user.id}`
      });
      setMembers(p=>[...p, {...user, userType:"Member"}]);
      showToast(`${user.displayName} added`, "success");
      setUserSearch("");
      setUserResults([]);
    } catch(e) { showToast(`Failed: ${e.message}`, "error"); }
    finally { setLoading(p=>({...p,[`add-${user.id}`]:false})); }
  }

  // ── Invite external user ────────────────────────────────────────────────────
  async function inviteExternal() {
    if (!extEmail.trim()) return;
    setSubmitting(true);
    try {
      // Step 1: Send invitation
      const invite = await graphRequest("POST", "/invitations", {
        invitedUserEmailAddress: extEmail.trim(),
        invitedUserDisplayName:  extName.trim() || extEmail.trim(),
        inviteRedirectUrl:       "https://myapps.microsoft.com",
        sendInvitationMessage:   sendInvite,
        invitedUserMessageInfo:  sendInvite && extMsg.trim() ? {
          customizedMessageBody: extMsg.trim()
        } : undefined,
      });

      // Step 2: Add the new guest to the group
      const guestId = invite.invitedUser.id;
      await graphRequest("POST", `/groups/${selectedG.id}/members/$ref`, {
        "@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${guestId}`
      });

      setMembers(p=>[...p, {
        id: guestId,
        displayName: extName.trim() || extEmail.trim(),
        mail: extEmail.trim(),
        userType: "Guest",
        jobTitle: "",
        externalUserState: "PendingAcceptance",
      }]);

      showToast(sendInvite ? `Invitation sent to ${extEmail.trim()}` : `${extEmail.trim()} added silently (no email sent)`, "success");
      setExtEmail(""); setExtName(""); setExtMsg(""); setSendInvite(false);
      setModal(null);
    } catch(e) { showToast(`Failed: ${e.message}`, "error"); }
    finally { setSubmitting(false); }
  }

  // ── Remove member ───────────────────────────────────────────────────────────
  async function removeMember() {
    const { id, displayName } = confirmData;
    setConfirmData(null);
    setModal(null);
    setLoading(p=>({...p,[`rm-${id}`]:true}));
    try {
      await graphRequest("DELETE", `/groups/${selectedG.id}/members/${id}/$ref`);
      setMembers(p=>p.filter(m=>m.id!==id));
      showToast(`${displayName} removed`, "success");
    } catch(e) { showToast(`Failed: ${e.message}`, "error"); }
    finally { setLoading(p=>({...p,[`rm-${id}`]:false})); }
  }

  // ── Save description ────────────────────────────────────────────────────────
  async function saveDescription() {
    setSavingDesc(true);
    try {
      await graphRequest("PATCH", `/groups/${selectedG.id}`, { description: editDesc });
      setGroups(p => p.map(g => g.id === selectedG.id ? {...g, description: editDesc} : g));
      setSelectedG(p => ({...p, description: editDesc}));
      showToast("Description updated", "success");
      setModal(null);
    } catch(e) { showToast(`Failed: ${e.message}`, "error"); }
    finally { setSavingDesc(false); }
  }

  // ── Filtered lists ──────────────────────────────────────────────────────────
  const filteredGroups  = groups.filter(g =>
    g.displayName?.toLowerCase().includes(search.toLowerCase()) ||
    g.mail?.toLowerCase().includes(search.toLowerCase())
  );
  const filteredMembers = members.filter(m =>
    m.displayName?.toLowerCase().includes(memberSearch.toLowerCase()) ||
    m.mail?.toLowerCase().includes(memberSearch.toLowerCase())
  );
  const guestCount    = members.filter(m => m.userType === "Guest").length;
  const internalCount = members.filter(m => m.userType !== "Guest").length;

  // ─── Not configured ──────────────────────────────────────────────────────────
  if (notConfigured) return (
    <div style={{minHeight:"100vh",background:"#f0f4f8",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'DM Sans',sans-serif",padding:24}}>
      <style>{css}</style>
      <div style={{background:"#fff",borderRadius:20,padding:"40px 44px",maxWidth:520,boxShadow:"0 4px 40px rgba(0,0,0,.1)"}}>
        <div style={{fontSize:48,marginBottom:16,textAlign:"center"}}>⚙️</div>
        <h2 style={{fontFamily:"'Sora',sans-serif",fontSize:20,fontWeight:700,color:"#0f172a",textAlign:"center",marginBottom:8}}>Setup Required</h2>
        <p style={{color:"#64748b",fontSize:14,lineHeight:1.7,marginBottom:24,textAlign:"center"}}>
          Open <code style={{background:"#f1f5f9",padding:"2px 6px",borderRadius:4}}>src/auth.js</code> and replace the two placeholder values with your Azure App Registration details.
        </p>
        <div style={{background:"#f8fafc",borderRadius:12,padding:"18px 20px",border:"1px solid #e2e8f0",fontSize:13,lineHeight:2}}>
          <div><strong>CLIENT_ID</strong> — App Registration → Overview → Application (client) ID</div>
          <div><strong>TENANT_ID</strong> — App Registration → Overview → Directory (tenant) ID</div>
        </div>
        <div style={{marginTop:20,background:"#fffbeb",borderRadius:12,padding:"16px 18px",border:"1px solid #fde68a",fontSize:12,color:"#92400e",lineHeight:1.8}}>
          <strong>App Registration setup:</strong><br/>
          1. Entra ID → App registrations → New registration<br/>
          2. Redirect URI → Single-page application → your Static Web App URL<br/>
          3. API permissions → Add → Microsoft Graph → Delegated:<br/>
          &nbsp;&nbsp;• <code>Group.ReadWrite.All</code><br/>
          &nbsp;&nbsp;• <code>User.ReadBasic.All</code><br/>
          &nbsp;&nbsp;• <code>Directory.AccessAsUser.All</code><br/>
          4. Grant admin consent
        </div>
      </div>
    </div>
  );

  // ─── Login screen ────────────────────────────────────────────────────────────
  if (!account) return (
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#0c1b3a 0%,#1a3a6b 60%,#0e2954 100%)",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"'DM Sans',sans-serif"}}>
      <style>{css}</style>
      <div style={{position:"absolute",inset:0,opacity:.05,backgroundImage:"linear-gradient(rgba(255,255,255,.5) 1px,transparent 1px),linear-gradient(90deg,rgba(255,255,255,.5) 1px,transparent 1px)",backgroundSize:"40px 40px"}}/>
      <div style={{background:"rgba(255,255,255,.08)",backdropFilter:"blur(20px)",border:"1px solid rgba(255,255,255,.14)",borderRadius:24,padding:"52px 48px",width:420,textAlign:"center",boxShadow:"0 40px 80px rgba(0,0,0,.4)",position:"relative",animation:"fadeIn .5s ease"}}>
        <div style={{width:64,height:64,borderRadius:16,background:"linear-gradient(135deg,#0078d4,#106ebe)",display:"flex",alignItems:"center",justifyContent:"center",margin:"0 auto 20px",fontSize:32,boxShadow:"0 8px 24px rgba(0,120,212,.4)"}}>👥</div>
        <h1 style={{fontFamily:"'Sora',sans-serif",fontSize:24,fontWeight:700,color:"#fff",marginBottom:8}}>Group Manager</h1>
        <p style={{color:"rgba(255,255,255,.5)",fontSize:13,lineHeight:1.7,marginBottom:32}}>
          Manage your Microsoft 365 Groups<br/>
          Add internal members & invite external guests
        </p>
        <button onClick={signIn} disabled={authState==="loading"}
          style={{...S.btn("primary"),width:"100%",justifyContent:"center",padding:"14px 22px",fontSize:15,borderRadius:12,background:authState==="loading"?"#1a3a6b":"#0078d4"}}>
          {authState==="loading"?<><Spinner/><span style={{animation:"pulse 1s infinite"}}>Signing in…</span></>:<><span>🔐</span>Sign in with Microsoft 365</>}
        </button>
        <p style={{color:"rgba(255,255,255,.2)",fontSize:11,marginTop:16}}>Uses delegated OAuth — only your groups, only your permissions</p>
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
          <div style={{width:30,height:30,borderRadius:8,background:"linear-gradient(135deg,#0078d4,#1a56db)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:15}}>👥</div>
          <span style={{fontFamily:"'Sora',sans-serif",fontWeight:700,fontSize:15,color:"#0f172a"}}>Group Manager</span>
          <span style={{background:"#e8f0fe",color:"#1a56db",fontSize:10,fontWeight:700,borderRadius:5,padding:"2px 7px"}}>M365</span>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <span style={{color:"#64748b",fontSize:13}}>{account.name || account.username}</span>
          <Avatar name={account.name || account.username} size={28}/>
          <button onClick={signOut} style={{...S.btn("ghost"),padding:"5px 12px",fontSize:12}}>Sign out</button>
        </div>
      </div>

      <div style={{display:"flex",height:"calc(100vh - 56px)"}}>

        {/* Sidebar */}
        <div style={{width:280,background:"#fff",borderRight:"1px solid #e2e8f0",display:"flex",flexDirection:"column",flexShrink:0}}>
          <div style={{padding:"16px 12px 10px"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
              <p style={{fontSize:11,fontWeight:700,color:"#94a3b8",letterSpacing:".08em",textTransform:"uppercase"}}>My Groups</p>
              <button onClick={loadGroups} disabled={groupsLoad} title="Refresh"
                style={{background:"none",border:"none",cursor:"pointer",fontSize:14,color:"#94a3b8",padding:2}}>
                {groupsLoad?<Spinner size={12} color="#94a3b8"/>:"⟳"}
              </button>
            </div>
            <div style={{position:"relative"}}>
              <span style={{position:"absolute",left:9,top:"50%",transform:"translateY(-50%)",fontSize:13,color:"#94a3b8"}}>🔍</span>
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
            {!groupsLoad && filteredGroups.length === 0 && (
              <p style={{textAlign:"center",color:"#94a3b8",fontSize:13,paddingTop:24}}>No groups found</p>
            )}
            {filteredGroups.map(g => (
              <div key={g.id} className="dlc"
                onClick={() => selectGroup(g)}
                style={{padding:"10px 9px",borderRadius:10,marginBottom:3,cursor:"pointer",transition:"all .15s",background:selectedG?.id===g.id?"#eff6ff":"transparent",border:`1.5px solid ${selectedG?.id===g.id?"#bfdbfe":"transparent"}`}}>
                <div style={{display:"flex",alignItems:"center",gap:8}}>
                  <div style={{width:32,height:32,borderRadius:8,flexShrink:0,background:selectedG?.id===g.id?"#dbeafe":"#eff6ff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:14}}>👥</div>
                  <div style={{minWidth:0}}>
                    <p style={{fontWeight:600,fontSize:13,color:"#1e293b",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{g.displayName}</p>
                    <p style={{fontSize:11,color:"#94a3b8",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{g.mail || "No email"}</p>
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
              <div style={{fontSize:52,marginBottom:12,opacity:.3}}>👥</div>
              <p style={{fontSize:17,fontWeight:600,color:"#64748b"}}>Select a Group</p>
              <p style={{fontSize:13,marginTop:5}}>Choose a group from the sidebar to manage members</p>
            </div>
          ) : (
            <div style={{maxWidth:760,animation:"slideUp .3s ease"}}>

              {/* Group header */}
              <div style={{background:"linear-gradient(135deg,#1a56db 0%,#0078d4 100%)",borderRadius:18,padding:"22px 26px",marginBottom:20,color:"#fff",boxShadow:"0 8px 32px rgba(26,86,219,.25)"}}>
                <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between"}}>
                  <div>
                    <div style={{display:"flex",alignItems:"center",gap:9,marginBottom:4}}>
                      <span style={{fontSize:20}}>👥</span>
                      <h2 style={{fontFamily:"'Sora',sans-serif",fontSize:18,fontWeight:700}}>{selectedG.displayName}</h2>
                    </div>
                    <p style={{opacity:.75,fontSize:12}}>{selectedG.mail || "No email address"}</p>
                    <div style={{display:"flex",alignItems:"center",gap:8,marginTop:4}}>
                    {selectedG.description && <p style={{opacity:.6,fontSize:12,margin:0}}>{selectedG.description}</p>}
                    <button onClick={()=>{setEditDesc(selectedG.description||"");setModal("editDesc");}}
                      style={{background:"rgba(255,255,255,.15)",border:"1px solid rgba(255,255,255,.25)",borderRadius:6,padding:"2px 8px",fontSize:11,color:"#fff",cursor:"pointer",fontFamily:"'DM Sans',sans-serif",flexShrink:0}}>
                      ✏ Edit description
                    </button>
                  </div>
                  </div>
                  <div style={{display:"flex",gap:8,flexShrink:0}}>
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
                  <span style={{background:"rgba(255,255,255,.2)",borderRadius:20,padding:"3px 11px",fontSize:11,fontWeight:600}}>✓ You are an owner</span>
                  <span style={{background:"rgba(255,255,255,.12)",borderRadius:20,padding:"3px 11px",fontSize:11,opacity:.8}}>Microsoft 365 Group</span>
                  <span style={{background:"rgba(255,255,255,.12)",borderRadius:20,padding:"3px 11px",fontSize:11,opacity:.8}}>Graph API</span>
                </div>
              </div>

              {/* Action buttons */}
              <div style={{display:"flex",gap:10,marginBottom:20,flexWrap:"wrap"}}>
                <button onClick={()=>{setModal("addInternal");setUserSearch("");setUserResults("");}}
                  style={S.btn("primary")}>
                  <span>+</span> Add Internal Member
                </button>
                <button onClick={()=>{setModal("addExternal");setExtEmail("");setExtName("");setExtMsg("");}}
                  style={S.btn("amber")}>
                  <span>✉</span> Invite External Guest
                </button>
                <button onClick={()=>loadMembers(selectedG.id)} disabled={membersLoad}
                  style={S.btn("ghost")}>
                  {membersLoad?<Spinner size={12} color="#64748b"/>:<span>⟳</span>} Refresh
                </button>
              </div>

              {/* Members table */}
              <div style={{background:"#fff",borderRadius:16,border:"1px solid #e2e8f0",overflow:"hidden",boxShadow:"0 2px 12px rgba(0,0,0,.04)"}}>
                <div style={{padding:"14px 18px",borderBottom:"1px solid #f1f5f9",display:"flex",alignItems:"center",justifyContent:"space-between",gap:12}}>
                  <h3 style={{fontSize:14,fontWeight:700,color:"#0f172a",margin:0}}>
                    Members <span style={{color:"#94a3b8",fontWeight:400,fontSize:13}}>({members.length})</span>
                  </h3>
                  <div style={{position:"relative",flex:1,maxWidth:260}}>
                    <span style={{position:"absolute",left:8,top:"50%",transform:"translateY(-50%)",fontSize:12,color:"#94a3b8"}}>🔍</span>
                    <input value={memberSearch} onChange={e=>setMemberSearch(e.target.value)} placeholder="Filter members…"
                      style={{width:"100%",padding:"7px 10px 7px 26px",borderRadius:7,border:"1.5px solid #e2e8f0",fontSize:12,outline:"none",fontFamily:"'DM Sans',sans-serif"}}/>
                  </div>
                </div>

                {membersLoad ? (
                  <div style={{padding:"30px 18px",display:"flex",alignItems:"center",gap:8,color:"#94a3b8",fontSize:13}}>
                    <Spinner size={14} color="#94a3b8"/> Loading members from Microsoft Graph…
                  </div>
                ) : filteredMembers.length === 0 ? (
                  <div style={{padding:"32px 18px",textAlign:"center",color:"#94a3b8"}}>
                    <div style={{fontSize:28,marginBottom:6,opacity:.3}}>👥</div>
                    <p style={{fontSize:13}}>{memberSearch ? "No members match your search" : "No members yet"}</p>
                  </div>
                ) : (
                  <>
                    {/* Guest banner if any */}
                    {guestCount > 0 && !memberSearch && (
                      <div style={{padding:"8px 18px",background:"#fffbeb",borderBottom:"1px solid #fde68a",display:"flex",alignItems:"center",gap:7,fontSize:12,color:"#92400e"}}>
                        <span style={{background:"#f59e0b",color:"#fff",borderRadius:"50%",width:16,height:16,display:"inline-flex",alignItems:"center",justifyContent:"center",fontSize:9,fontWeight:700,flexShrink:0}}>G</span>
                        {guestCount} external guest{guestCount>1?"s":""} — invitation emails sent via Microsoft, pending acceptance shown below
                      </div>
                    )}
                    {filteredMembers.map((m, i) => {
                      const isGuest = m.userType === "Guest";
                      const isPending = m.externalUserState === "PendingAcceptance";
                      return (
                        <div key={m.id} className="mrow"
                          style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"11px 18px",transition:"background .12s",borderBottom:i<filteredMembers.length-1?"1px solid #f8fafc":"none",opacity:loading[`rm-${m.id}`]?.4:1,background:isGuest?"#fffdf5":"#fff"}}>
                          <div style={{display:"flex",alignItems:"center",gap:10}}>
                            <Avatar name={m.displayName||m.mail} size={34} isGuest={isGuest}/>
                            <div>
                              <div style={{display:"flex",alignItems:"center",gap:6}}>
                                <p style={{fontSize:13,fontWeight:600,color:"#1e293b",margin:0}}>{m.displayName||m.mail}</p>
                                {isGuest && <span style={{fontSize:9,background:"#fef3c7",color:"#92400e",borderRadius:4,padding:"1px 5px",fontWeight:700}}>GUEST</span>}
                                {isPending && <span style={{fontSize:9,background:"#fce7f3",color:"#9d174d",borderRadius:4,padding:"1px 5px",fontWeight:700}}>PENDING</span>}
                              </div>
                              <p style={{fontSize:11,color:"#94a3b8",margin:0}}>{m.mail||m.userPrincipalName}</p>
                            </div>
                          </div>
                          <div style={{display:"flex",alignItems:"center",gap:8}}>
                            {m.jobTitle && <span style={{fontSize:11,color:"#64748b",background:"#f1f5f9",borderRadius:5,padding:"2px 7px"}}>{m.jobTitle}</span>}
                            <button className="ibtn" title="Remove"
                              disabled={loading[`rm-${m.id}`]}
                              onClick={()=>{setConfirmData({id:m.id,displayName:m.displayName||m.mail,isGuest});setModal("confirm");}}
                              style={{background:"transparent",border:"none",cursor:"pointer",width:28,height:28,borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,color:"#cbd5e1",transition:"all .12s"}}>
                              {loading[`rm-${m.id}`]?<Spinner size={12} color="#94a3b8"/>:"🗑"}
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

      {/* Add Internal Member Modal */}
      {modal==="addInternal" && (
        <Modal title="Add Internal Member" onClose={()=>setModal(null)}>
          <p style={{fontSize:13,color:"#64748b",marginBottom:14}}>Search for users in your organization.</p>
          <div style={{position:"relative",marginBottom:12}}>
            <span style={{position:"absolute",left:9,top:"50%",transform:"translateY(-50%)",fontSize:13,color:"#94a3b8"}}>🔍</span>
            <input value={userSearch} onChange={e=>setUserSearch(e.target.value)}
              placeholder="Search by name or email…" autoFocus
              style={{width:"100%",padding:"9px 10px 9px 30px",borderRadius:8,border:"1.5px solid #d1d5db",fontSize:13,outline:"none",fontFamily:"'DM Sans',sans-serif",boxSizing:"border-box"}}/>
          </div>
          {userSearching && <div style={{display:"flex",alignItems:"center",gap:7,color:"#94a3b8",fontSize:13,padding:"8px 4px"}}><Spinner size={12} color="#94a3b8"/> Searching…</div>}
          {userResults.map(u=>(
            <div key={u.id} className="arow" style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"9px 8px",borderRadius:8,transition:"background .12s"}}>
              <div style={{display:"flex",alignItems:"center",gap:9}}>
                <Avatar name={u.displayName} size={32}/>
                <div>
                  <p style={{fontSize:13,fontWeight:600,color:"#1e293b",margin:0}}>{u.displayName}</p>
                  <p style={{fontSize:11,color:"#94a3b8",margin:0}}>{u.mail} {u.jobTitle?`· ${u.jobTitle}`:""}</p>
                </div>
              </div>
              <button onClick={()=>addInternalMember(u)} disabled={loading[`add-${u.id}`]}
                style={{...S.btn("outline"),padding:"5px 12px",fontSize:12,borderRadius:6}}>
                {loading[`add-${u.id}`]?<><Spinner size={11} color="#1a56db"/> Adding…</>:"+ Add"}
              </button>
            </div>
          ))}
          {userSearch.length >= 2 && !userSearching && userResults.length === 0 && (
            <p style={{fontSize:13,color:"#94a3b8",textAlign:"center",padding:"12px 0"}}>No users found</p>
          )}
        </Modal>
      )}

      {/* Invite External Guest Modal */}
      {modal==="addExternal" && (
        <Modal title="✉ Invite External Guest" onClose={()=>setModal(null)}>
          <p style={{fontSize:13,color:"#64748b",marginBottom:18,lineHeight:1.6}}>
            Send a Microsoft invitation email to an external user. They'll receive an email to accept and join the group. Perfect for vendors, contractors, or partners.
          </p>
          <div style={{marginBottom:12}}>
            <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:5}}>Email address *</label>
            <input value={extEmail} onChange={e=>setExtEmail(e.target.value)}
              placeholder="external@company.com" type="email" autoFocus
              style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"1.5px solid #d1d5db",fontSize:13,outline:"none",fontFamily:"'DM Sans',sans-serif",boxSizing:"border-box"}}/>
          </div>
          <div style={{marginBottom:12}}>
            <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:5}}>Display name <span style={{fontWeight:400,color:"#94a3b8"}}>(optional)</span></label>
            <input value={extName} onChange={e=>setExtName(e.target.value)}
              placeholder="Jane Smith"
              style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"1.5px solid #d1d5db",fontSize:13,outline:"none",fontFamily:"'DM Sans',sans-serif",boxSizing:"border-box"}}/>
          </div>
          <div style={{marginBottom:20}}>
            <label style={{fontSize:12,fontWeight:600,color:"#374151",display:"block",marginBottom:5}}>Personal message <span style={{fontWeight:400,color:"#94a3b8"}}>(optional)</span></label>
            <textarea value={extMsg} onChange={e=>setExtMsg(e.target.value)}
              placeholder="Hi! I'm inviting you to our Microsoft 365 group…"
              rows={3}
              style={{width:"100%",padding:"9px 12px",borderRadius:8,border:"1.5px solid #d1d5db",fontSize:13,outline:"none",fontFamily:"'DM Sans',sans-serif",boxSizing:"border-box",resize:"vertical"}}/>
          </div>
          {/* Send invitation toggle */}
          <div onClick={()=>setSendInvite(p=>!p)}
            style={{display:"flex",alignItems:"flex-start",gap:10,padding:"12px 14px",borderRadius:9,border:`1.5px solid ${sendInvite?"#f59e0b":"#e2e8f0"}`,background:sendInvite?"#fffbeb":"#f8fafc",marginBottom:16,cursor:"pointer",transition:"all .15s"}}>
            <div style={{width:18,height:18,borderRadius:4,border:`2px solid ${sendInvite?"#f59e0b":"#cbd5e1"}`,background:sendInvite?"#f59e0b":"#fff",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,marginTop:1,transition:"all .15s"}}>
              {sendInvite && <span style={{color:"#fff",fontSize:11,fontWeight:700}}>✓</span>}
            </div>
            <div>
              <p style={{fontSize:13,fontWeight:600,color:sendInvite?"#92400e":"#374151",margin:0,marginBottom:2}}>Send invitation email</p>
              <p style={{fontSize:11,color:sendInvite?"#a16207":"#94a3b8",margin:0,lineHeight:1.5}}>
                {sendInvite
                  ? "Microsoft will email the guest to accept. They appear as Pending until accepted."
                  : "Guest added silently — no email sent. They can still receive group emails immediately."}
              </p>
            </div>
          </div>
          <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}>
            <button onClick={()=>setModal(null)} style={S.btn("ghost")}>Cancel</button>
            <button onClick={inviteExternal} disabled={!extEmail.trim()||submitting}
              style={{...S.btn("amber"),opacity:!extEmail.trim()||submitting?.6:1}}>
              {submitting?<><Spinner size={13}/> Adding…</>:sendInvite?<><span>✉</span> Send Invitation</>:<><span>+</span> Add Silently</>}
            </button>
          </div>
        </Modal>
      )}

      {/* Edit Description Modal */}
      {modal==="editDesc" && (
        <Modal title="Edit Group Description" onClose={()=>setModal(null)}>
          <p style={{fontSize:13,color:"#64748b",marginBottom:14,lineHeight:1.6}}>
            This description appears in the group header and in Microsoft 365 directory listings.
          </p>
          <textarea value={editDesc} onChange={e=>setEditDesc(e.target.value)}
            placeholder="Enter a description for this group…"
            rows={4}
            autoFocus
            style={{width:"100%",padding:"10px 12px",borderRadius:8,border:"1.5px solid #d1d5db",fontSize:13,outline:"none",fontFamily:"'DM Sans',sans-serif",boxSizing:"border-box",resize:"vertical",lineHeight:1.6}}/>
          <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:16}}>
            <button onClick={()=>setModal(null)} style={S.btn("ghost")}>Cancel</button>
            <button onClick={saveDescription} disabled={savingDesc}
              style={{...S.btn("primary"),opacity:savingDesc?.6:1}}>
              {savingDesc?<><Spinner size={13}/> Saving…</>:"Save description"}
            </button>
          </div>
        </Modal>
      )}

      {/* Confirm remove modal */}
      {modal==="confirm" && confirmData && (
        <ConfirmModal
          message={`Remove ${confirmData.displayName} from ${selectedG?.displayName}?${confirmData.isGuest?" This will remove their guest access to this group.":""}`}
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
  *{box-sizing:border-box;margin:0;padding:0}
  @keyframes slideUp{from{transform:translateY(14px);opacity:0}to{transform:translateY(0);opacity:1}}
  @keyframes fadeIn{from{opacity:0}to{opacity:1}}
  @keyframes spin{to{transform:rotate(360deg)}}
  @keyframes pulse{0%,100%{opacity:1}50%{opacity:.5}}
  .dlc:hover{background:#f8faff!important;box-shadow:0 2px 12px rgba(26,86,219,.08)!important}
  .mrow:hover{background:#f8faff!important}
  .arow:hover{background:#f0f4ff!important}
  .ibtn:hover{background:#fee2e2!important;color:#dc2626!important}
  ::-webkit-scrollbar{width:5px}::-webkit-scrollbar-track{background:transparent}::-webkit-scrollbar-thumb{background:#cbd5e1;border-radius:3px}
`;
