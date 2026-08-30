/* =========================================================================
   M365 CONFIGURATION  —  FILL THESE IN, then nothing else is needed below.
   ========================================================================= */
const M365_CONFIG = {
    // From Azure portal > Microsoft Entra ID > App registrations > [your app] > Overview
    tenantId:      "e35c0059-b98e-4fd5-abab-a72d3307c532",        // 7J Portal App > Overview > Directory (tenant) ID
    clientId:      "fe8072d4-f842-4d5e-9320-4948a0c0fd49",        // 7J Portal App > Overview > Application (client) ID
    // Production redirect URI is selected below from the current hosted origin.
    // For Netlify this becomes https://tims7jbdmportal.netlify.app/app.html.
    // For the existing SharePoint deployment it remains the SharePoint page.
    redirectUri:   "https://7jfinance.sharepoint.com/sites/BDM/SiteAssets/bdm-portal.html",
    // Your SharePoint site details (used to locate the site in Graph)
    sharePointHost:"7jfinance.sharepoint.com",
    sitePath:      "/sites/BDM",
    // Microsoft Graph permission scopes. These must be granted (delegated) in the Azure app.
    scopes:        ["User.Read", "Sites.ReadWrite.All", "Calendars.ReadWrite"],
    // The display names of the Microsoft Lists that store the data.
    lists: {
        brokers: "BDM Brokers",
        users:   "BDM Users",
        audit:   "BDM Audit Log",
        auditAccount: "BDM Audit Log",
        auditDialer: "BDM Audit - Dialler Activity",
        deals:   "BDM Deals",
        globalSettings: "BDM Global Settings",
        callGuide: "BDM Call Guide",
        performanceReviews: "BDM Performance Reviews",
        kpiSnapshots: "BDM KPI Snapshots",
        backupManifest: "BDM System Backup Manifest"
    },
    // People who are admins. Access is decided by Microsoft 365 sign-in (only your org
    // can sign in) — this list just marks who among them is an admin. Add more emails
    // to grant more admins, e.g. ["naz@7jfinance.com", "jane@7jfinance.com"].
    admins: ["naz@7jfinance.com"]
};

// Is this email an admin? Decided by the admins list in M365_CONFIG (not by a
// manually-edited role field), so only the people you list here ever get admin access.
function isAdmin(email){
    if(!email) return false;
    const e = email.toLowerCase();
    return (M365_CONFIG.admins || []).some(a => a.toLowerCase() === e);
}

// Local development must never use the production Microsoft 365 configuration.
// The packaged portal contains the real production tenant/client IDs, so simply
// checking whether those values are populated is unsafe: opening the portal from
// file:// or localhost would otherwise initialise MSAL against the SharePoint
// redirect URI and can leave the local app stuck/crashing at startup.
function isLocalRuntime(){
    const protocol = String(location.protocol || '').toLowerCase();
    const host = String(location.hostname || '').toLowerCase();
    return protocol === 'file:' || host === 'localhost' || host === '127.0.0.1' || host === '::1';
}

function getProductionRedirectUri(){
    const host = String(location.hostname || '').toLowerCase();
    if(host === 'tims7jbdmportal.netlify.app') return 'https://tims7jbdmportal.netlify.app/app.html';
    if(host === 'timdell1990-lgtm.github.io') return 'https://timdell1990-lgtm.github.io/7j-finance-bdm-portal/app.html';
    // Custom domain (portal.7jfinance.com) once DNS/GitHub Pages custom domain is live.
    if(host === 'portal.7jfinance.com') return 'https://portal.7jfinance.com/app.html';
    return M365_CONFIG.redirectUri;
}

function m365Configured(){
    if(isLocalRuntime()) return false;
    return M365_CONFIG.tenantId !== "YOUR-TENANT-ID-GUID" &&
           M365_CONFIG.clientId !== "YOUR-CLIENT-ID-GUID";
}

/* ----------------------- Microsoft Graph HTTP helpers ------------------- */
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
let _msalInstance = null;
let _siteId = null;
let _cloudReady = false;

function initMsal(){
    if(!m365Configured()) return null;
    // If the MSAL library didn't load (CDN blocked by SharePoint CSP / network),
    // give a clear, actionable error instead of failing later with a cryptic one.
    if(typeof msal === "undefined" || !msal || !msal.PublicClientApplication){
        throw new Error("The Microsoft sign-in library (msal-browser) failed to load. This is usually because SharePoint blocked the external CDN. Open the setup guide's 'If the sign-in library won't load' note and load msal-browser.min.js from the same folder as the portal.");
    }
    _msalInstance = new msal.PublicClientApplication({
        auth: {
            clientId: M365_CONFIG.clientId,
            authority: ((localStorage.getItem('adt7j_auth_mode_v1')||'tenant')==='multitenant' ? 'https://login.microsoftonline.com/organizations' : 'https://login.microsoftonline.com/' + M365_CONFIG.tenantId),
            redirectUri: getProductionRedirectUri()
        },
        cache: { cacheLocation: "localStorage", storeAuthStateInCookie: true }
    });
    return _msalInstance;
}
let _msalInitStarted = false;
// Ensure MSAL is created and initialised (v3 requires initialize() before use).
async function ensureMsal(){
    if(!_msalInstance){ initMsal(); }              // may throw if the library didn't load
    if(!_msalInitStarted){
        _msalInitStarted = true;
        await _msalInstance.initialize();
    }
    return _msalInstance;
}

async function getGraphToken(){
    if(!_msalInstance) throw new Error("M365 not configured");
    const acc = _msalInstance.getActiveAccount() || _msalInstance.getAllAccounts()[0];
    if(acc) _msalInstance.setActiveAccount(acc);
    const req = { scopes: M365_CONFIG.scopes, account: _msalInstance.getActiveAccount() };
    try{
        const r = await _msalInstance.acquireTokenSilent(req);
        return r.accessToken;
    }catch(e){
        const r = await _msalInstance.acquireTokenPopup(req);
        return r.accessToken;
    }
}

async function graphGet(url){
    const tok = await getGraphToken();
    const r = await fetch(url, { headers: { Authorization: "Bearer " + tok } });
    if(!r.ok) throw new Error("Graph GET failed ("+r.status+"): "+url);
    return r.status===204 ? null : r.json();
}
async function graphPost(url, body){
    const tok = await getGraphToken();
    const r = await fetch(url, { method:"POST", headers:{ Authorization:"Bearer "+tok, "Content-Type":"application/json" }, body: JSON.stringify(body||{}) });
    if(r.status!==201 && r.status!==200 && r.status!==204) throw new Error("Graph POST failed ("+r.status+"): "+url);
    return r.status===204 ? null : r.json();
}
async function graphPatch(url, body, etag){
    const tok=await getGraphToken();
    const headers={Authorization:"Bearer "+tok,"Content-Type":"application/json"};
    if(etag) headers["If-Match"]=etag;
    const r=await fetch(url,{method:"PATCH",headers,body:JSON.stringify(body||{})});
    if(!r.ok && r.status!==204){ const e=new Error("Graph PATCH failed ("+r.status+"): "+url); e.status=r.status; throw e; }
}
async function graphDelete(url, etag){
    const tok = await getGraphToken();
    const headers={ Authorization:"Bearer "+tok };
    if(etag) headers["If-Match"]=etag;
    const r = await fetch(url, { method:"DELETE", headers });
    if(!r.ok && r.status!==204){ const e=new Error("Graph DELETE failed ("+r.status+"): "+url); e.status=r.status; throw e; }
}

/* ----------------------- Calendar / Teams integration ------------------ */
function localDateISO(d){
    const x=new Date(d); return x.getFullYear()+"-"+String(x.getMonth()+1).padStart(2,"0")+"-"+String(x.getDate()).padStart(2,"0");
}
function calendarRange(){
    const y=calendarCursor.getFullYear(), m=calendarCursor.getMonth();
    const start=new Date(y,m,1); start.setDate(start.getDate()-start.getDay());
    const end=new Date(y,m+1,0); end.setDate(end.getDate()+(6-end.getDay())); end.setHours(23,59,59,999);
    return {start,end};
}
function getLocalCalendarEvents(){ try{return JSON.parse(localStorage.getItem(CALENDAR_STORAGE_KEY))||[];}catch(e){return [];} }
function saveLocalCalendarEvents(items){localStorage.setItem(CALENDAR_STORAGE_KEY,JSON.stringify(items||[]));}
function seedDemoCalendarEvents(){
    if(getLocalCalendarEvents().length) return;
    const now=new Date(); const y=now.getFullYear(),m=now.getMonth();
    const b1=(DEFAULT_BROKERS||[])[0], b2=(DEFAULT_BROKERS||[])[1];
    const mk=(day,hour,broker,title,teams)=>({id:"demo-"+day+"-"+hour,subject:title,start:new Date(y,m,day,hour,0).toISOString(),end:new Date(y,m,day,hour+1,0).toISOString(),brokerId:broker?.Id||"",brokerName:broker?.Title||"",company:broker?.Company||"",notes:"Demo calendar appointment",isTeams:teams,webLink:""});
    saveLocalCalendarEvents([mk(3,10,b1,"Broker follow-up",true),mk(9,14,b2,"Bridging case review",true),mk(17,11,b1,"Pipeline review",false),mk(24,15,b2,"Deal discussion",true)]);
}
async function loadCalendarEvents(){
    const r=calendarRange();
    if(isDemoMode() || !m365Configured()){ calendarEvents=getLocalCalendarEvents(); renderCallbackCount(); return; }
    const url=`${GRAPH_BASE}/me/calendarView?startDateTime=${encodeURIComponent(r.start.toISOString())}&endDateTime=${encodeURIComponent(r.end.toISOString())}&$top=200&$select=id,subject,start,end,isOnlineMeeting,onlineMeeting,webLink,bodyPreview`;
    const data=await graphGet(url);
    const localLinks=getLocalCalendarEvents();
    calendarEvents=(data.value||[]).map(e=>{const link=localLinks.find(x=>String(x.id)===String(e.id))||{};return {id:e.id,subject:e.subject||"(no subject)",start:e.start?.dateTime,end:e.end?.dateTime,brokerId:link.brokerId||"",brokerName:link.brokerName||"",company:link.company||"",notes:e.bodyPreview||link.notes||"",isTeams:!!e.isOnlineMeeting,webLink:e.onlineMeeting?.joinUrl||e.webLink||link.webLink||"",isCallback:!!link.isCallback||/^callback\b/i.test(e.subject||"")};});
}
function eventBrokerName(e){
    if(e.brokerName) return e.company ? e.brokerName+" · "+e.company : e.brokerName;
    return "";
}
function setCalendarView(mode){
    calendarViewMode=mode;
    ['Month','Week','Day'].forEach(x=>{const b=document.getElementById('calView'+x); if(b)b.classList.toggle('active',mode===x.toLowerCase());});
    renderCalendar();
}
function calendarWeekStart(date){
    const d=new Date(date); d.setHours(0,0,0,0); d.setDate(d.getDate()-d.getDay()); return d;
}
function calendarEventHtml(e){
    const time=new Date(e.start).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"});
    const broker=eventBrokerName(e); const tip=e.webLink?" · Teams meeting":"";
    return `<div class="calendar-event ${e.isTeams?'teams':''}" draggable="true" data-event-id="${escapeHtml(String(e.id))}" title="${escapeHtml((e.subject||'')+' '+broker+tip)}" ondragstart="calendarDragStart(event,'${String(e.id).replace(/'/g,"\\'")}')" onclick="event.stopPropagation();openExistingCalendarEvent('${String(e.id).replace(/'/g,"\\'")}')">${time} ${escapeHtml(e.subject||"")}${broker?' · '+escapeHtml(e.brokerName):''}</div>`;
}
function attachCalendarDragHandlers(){
    document.querySelectorAll('.calendar-cell[data-date]').forEach(cell=>{
        cell.addEventListener('dragover',e=>{e.preventDefault();cell.classList.add('drag-over');});
        cell.addEventListener('dragleave',()=>cell.classList.remove('drag-over'));
        cell.addEventListener('drop',e=>{
            e.preventDefault();cell.classList.remove('drag-over');
            const id=window._calendarDraggedId; if(!id)return;
            rescheduleCalendarEvent(id,cell.dataset.date);
        });
    });
}
function renderCalendar(){
    const title=document.getElementById('calendarTitle'); const grid=document.getElementById('calendarGrid'); if(!grid)return;
    if(calendarViewMode==='month'){
        const r=calendarRange(); if(title) title.textContent=calendarCursor.toLocaleDateString("en-GB",{month:"long",year:"numeric"});
        const d=new Date(r.start); let html="";
        for(let i=0;i<42;i++){
            const iso=localDateISO(d), other=d.getMonth()!==calendarCursor.getMonth(), today=iso===localDateISO(new Date());
            const evs=calendarEvents.filter(e=>localDateISO(new Date(e.start))===iso).sort((a,b)=>new Date(a.start)-new Date(b.start));
            html+=`<div class="calendar-cell ${other?'other-month':''} ${today?'today':''}" data-date="${iso}" onclick="openCalendarDay('${iso}')"><div class="calendar-day-number">${d.getDate()}</div>${evs.map(calendarEventHtml).join("")}</div>`;
            d.setDate(d.getDate()+1);
        }
        grid.className='calendar-grid'; grid.innerHTML=html;
    } else if(calendarViewMode==='week'){
        const start=calendarWeekStart(calendarCursor); if(title) title.textContent=`Week of ${start.toLocaleDateString("en-GB",{day:"numeric",month:"short",year:"numeric"})}`;
        let html="";
        for(let i=0;i<7;i++){const d=new Date(start);d.setDate(start.getDate()+i);const iso=localDateISO(d),today=iso===localDateISO(new Date());const evs=calendarEvents.filter(e=>localDateISO(new Date(e.start))===iso).sort((a,b)=>new Date(a.start)-new Date(b.start));html+=`<div class="calendar-cell ${today?'today':''}" data-date="${iso}" onclick="openCalendarDay('${iso}')"><div class="calendar-day-number">${d.toLocaleDateString("en-GB",{weekday:"short",day:"numeric"})}</div>${evs.map(calendarEventHtml).join("")}</div>`;}
        grid.className='calendar-grid'; grid.style.gridTemplateColumns='repeat(7,minmax(140px,1fr))'; grid.innerHTML=html;
    } else {
        const iso=localDateISO(calendarCursor); if(title) title.textContent=calendarCursor.toLocaleDateString("en-GB",{weekday:"long",day:"numeric",month:"long",year:"numeric"});
        let html='<div class="calendar-day-view">';
        for(let h=7;h<=19;h++){const evs=calendarEvents.filter(e=>localDateISO(new Date(e.start))===iso&&new Date(e.start).getHours()===h);html+=`<div class="calendar-day-hour">${String(h).padStart(2,'0')}:00</div><div class="calendar-day-slot" data-date="${iso}" onclick="openCalendarDay('${iso}')">${evs.map(e=>`<div class="calendar-day-event" draggable="true" ondragstart="calendarDragStart(event,'${String(e.id).replace(/'/g,"\\'")}')" onclick="event.stopPropagation();openExistingCalendarEvent('${String(e.id).replace(/'/g,"\\'")}')">${new Date(e.start).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"})} · ${escapeHtml(e.subject||"")}</div>`).join("")}</div>`;}
        html+='</div>'; grid.className=''; grid.style.display='block'; grid.innerHTML=html;
    }
    if(calendarViewMode!=='day') grid.style.display='grid';
    const st=document.getElementById("calendarSyncStatus"); if(st) st.textContent=isDemoMode()?"Demo calendar · local only":"Microsoft 365 · Teams sync";
    const acc=document.getElementById("calendarAccount"); if(acc) acc.textContent=currentUser?(currentUser.Title+" · "+currentUser.Email):"Demo user";
    renderCalendarBrokerSearch(); attachCalendarDragHandlers();
}
function calendarDragStart(e,id){window._calendarDraggedId=String(id);e.dataTransfer.effectAllowed='move';}
async function rescheduleCalendarEvent(id,newDate){
    const ev=calendarEvents.find(x=>String(x.id)===String(id)); if(!ev)return;
    const old=new Date(ev.start); const time=String(old.getHours()).padStart(2,'0')+':'+String(old.getMinutes()).padStart(2,'0');
    calendarEditingId=String(id);
    document.getElementById("calSubject").value=ev.subject||"";
    document.getElementById("calDate").value=newDate; document.getElementById("calTime").value=time;
    document.getElementById("calNotes").value=ev.notes||"";
    document.getElementById("calDuration").value=String(Math.max(30,Math.round((new Date(ev.end)-old)/60000)||60));
    populateCalendarBrokerSelect(ev.brokerId||""); document.getElementById("calTeams").checked=!!ev.isTeams&&!isDemoMode();
    await saveCalendarEvent();
}
async function refreshCalendar(){
    try{await loadCalendarEvents();renderCalendar();}catch(e){showCloudError(e);}
}
function calendarMove(delta){
    if(calendarViewMode==='day') calendarCursor.setDate(calendarCursor.getDate()+delta);
    else if(calendarViewMode==='week') calendarCursor.setDate(calendarCursor.getDate()+delta*7);
    else calendarCursor=new Date(calendarCursor.getFullYear(),calendarCursor.getMonth()+delta,1);
    refreshCalendar();
}
function calendarToday(){calendarCursor=new Date();refreshCalendar();}
function openCalendarDay(iso){
    openCalendarEventModal("");
    const d=document.getElementById("calDate"); if(d) d.value=iso;
}
function openCalendarEventModal(brokerId=""){
    calendarEditingId=null;
    document.getElementById("calendarEventModalTitle").textContent=calendarCallbackContext ? "Schedule callback" : "New calendar event";
    document.getElementById("calSubject").value=calendarCallbackContext?.subject||"";
    document.getElementById("calNotes").value=calendarCallbackContext?.notes||"";
    document.getElementById("calTeams").checked=!isDemoMode();
    const n=new Date(); n.setMinutes(0,0,0);
    n.setHours(Math.max(8,Math.min(17,n.getHours()+1)));
    document.getElementById("calDate").value=localDateISO(n);
    document.getElementById("calTime").value=String(n.getHours()).padStart(2,"0")+":00";
    populateCalendarBrokerSelect(calendarCallbackContext?.brokerId||brokerId);
    document.getElementById("calendarEventModal").classList.add("show");
}
function closeCalendarEventModal(){document.getElementById("calendarEventModal").classList.remove("show"); if(calendarCallbackContext && !calendarEditingId) calendarCallbackContext=null;}
function populateCalendarBrokerSelect(selected=""){
    const el=document.getElementById("calBrokerSelect"); if(!el)return; const list=getStoredBrokers(); el.innerHTML='<option value="">No broker linked</option>'+list.map(b=>`<option value="${escapeHtml(String(b.Id))}">${escapeHtml(b.Title||'')} — ${escapeHtml(b.Company||'')}</option>`).join(""); el.value=selected||"";
}
function renderCalendarBrokerSearch(){
    const q=(document.getElementById("calendarBrokerSearch")?.value||"").toLowerCase().trim(); const el=document.getElementById("calendarBrokerResults"); if(!el)return;
    const list=getStoredBrokers().filter(b=>!q||[b.Title,b.Company,b.Email,b.Phone,b.City].some(x=>String(x||"").toLowerCase().includes(q))).slice(0,12);
    el.innerHTML=list.map(b=>`<div class="calendar-broker-result" onclick="openCalendarEventModal('${String(b.Id)}')"><strong>${escapeHtml(b.Title||'')}</strong><br><span>${escapeHtml(b.Company||'')}</span></div>`).join("");
}
async function saveCalendarEvent(){
    const subject=document.getElementById("calSubject").value.trim()||"Broker follow-up"; const date=document.getElementById("calDate").value; const time=document.getElementById("calTime").value||"09:00"; const mins=Number(document.getElementById("calDuration").value)||60; const brokerId=document.getElementById("calBrokerSelect").value; const broker=getStoredBrokers().find(b=>String(b.Id)===String(brokerId));
    if(!date){alert("Choose a date.");return;}
    const start=new Date(`${date}T${time}:00`); const end=new Date(start.getTime()+mins*60000); const notes=document.getElementById("calNotes").value.trim(); const teams=document.getElementById("calTeams").checked && !isDemoMode();
    try{
        if(isDemoMode() || !m365Configured()){
            const items=getLocalCalendarEvents();
            const payload={id:calendarEditingId||"local-"+Date.now(),subject,start:start.toISOString(),end:end.toISOString(),brokerId:broker?.Id||"",brokerName:broker?.Title||"",company:broker?.Company||"",notes,isTeams:false,webLink:"",isCallback:!!calendarCallbackContext,createdBy:currentUser?.Email||currentUser?.Title||""};
            const idx=items.findIndex(x=>String(x.id)===String(calendarEditingId));
            if(idx>=0) items[idx]=payload; else items.push(payload);
            saveLocalCalendarEvents(items);
        }else{
            const body={subject,start:{dateTime:date+"T"+time+":00",timeZone:"GMT Standard Time"},end:{dateTime:date+"T"+new Date(start.getTime()+mins*60000).toTimeString().slice(0,5)+":00",timeZone:"GMT Standard Time"},body:{contentType:"HTML",content:escapeHtml(notes)},isOnlineMeeting:teams};
            if(teams) body.onlineMeetingProvider="teamsForBusiness";
            let created;
            if(calendarEditingId){
                await graphPatch(`${GRAPH_BASE}/me/events/${encodeURIComponent(calendarEditingId)}`,body);
                created={id:calendarEditingId,subject,start:{dateTime:start.toISOString()},end:{dateTime:end.toISOString()},isOnlineMeeting:teams,onlineMeeting:{joinUrl:""}};
            } else {
                created=await graphPost(`${GRAPH_BASE}/me/events`,body);
            }
            const links=getLocalCalendarEvents();
            const payload={id:created.id,subject:created.subject||subject,start:created.start?.dateTime||start.toISOString(),end:created.end?.dateTime||end.toISOString(),brokerId:broker?.Id||"",brokerName:broker?.Title||"",company:broker?.Company||"",notes,isTeams:!!created.isOnlineMeeting||teams,webLink:created.onlineMeeting?.joinUrl||created.webLink||"",isCallback:!!calendarCallbackContext,createdBy:currentUser?.Email||currentUser?.Title||""};
            const idx=links.findIndex(x=>String(x.id)===String(created.id)); if(idx>=0) links[idx]=payload; else links.push(payload); saveLocalCalendarEvents(links);
        }
        closeCalendarEventModal(); calendarEditingId=null; calendarCallbackContext=null; await refreshCalendar();
    }catch(e){alert("Could not save calendar event: "+(e.message||e));}
}
function openExistingCalendarEvent(id){
    const e=calendarEvents.find(x=>String(x.id)===String(id)); if(!e)return;
    calendarEditingId=String(id);
    document.getElementById("calendarEventModalTitle").textContent="Edit calendar event";
    document.getElementById("calSubject").value=e.subject||"";
    const dt=new Date(e.start);
    document.getElementById("calDate").value=localDateISO(dt);
    document.getElementById("calTime").value=String(dt.getHours()).padStart(2,"0")+":"+String(dt.getMinutes()).padStart(2,"0");
    document.getElementById("calNotes").value=e.notes||"";
    const mins=Math.max(30,Math.round((new Date(e.end)-new Date(e.start))/60000)||60);
    document.getElementById("calDuration").value=[30,60,90,120].includes(mins)?String(mins):"60";
    populateCalendarBrokerSelect(e.brokerId||"");
    document.getElementById("calTeams").checked=!!e.isTeams && !isDemoMode();
    document.getElementById("calendarEventModal").classList.add("show");
}

async function getSiteId(){
    if(_siteId) return _siteId;
    const host = M365_CONFIG.sharePointHost;
    const path = M365_CONFIG.sitePath;
    const data = await graphGet(`${GRAPH_BASE}/sites/${host}:${path}`);
    _siteId = data.id;
    return _siteId;
}

// SharePoint list IDs are not reliably usable from their display name in the URL,
// so resolve each list's stable id once by display name and cache it.
const _listIdCache = {};

/* ===================== CALL GUIDE / SCRIPT CONTENT ===================== */
const CALL_GUIDE_STORAGE_KEY = "adt7j_call_guide_v1";
const DEFAULT_CALL_GUIDE_SECTIONS = [{"SectionKey":"mindset","Title":"1. Execution & Mindset Hacks (For Tim)","Content":"<p>Practical strategies to stay calm, avoid rushing, and stay resilient during call blocks:</p>\n<ul>\n<li><strong>The \"3-Second Rule\" for Pacing:</strong> When you feel yourself stumbling, it's usually because you're rushing to get all the words out. Practice saying your company name (<em>\"7J Finance\"</em>) with a deliberate half-second pause right before and after it.</li>\n<li><strong>Keep a Physical Cheat Sheet:</strong> Print out or keep this guide open right at eye level next to your monitor. Having it visually pinned removes the cognitive load of having to remember what to say next under pressure.</li>\n<li><strong>Track Your \"No's\":</strong> On a cold-calling block, keep a tally of total attempts/rejections instead of fixating on wins. Shifting the game to \"let's see how many conversations I can run today\" reduces performance anxiety and keeps your voice relaxed.</li>\n</ul>","SortOrder":1,"Active":true},{"SectionKey":"fullflow","Title":"2. Complete Loose Call Flow (Start to Finish)","Content":"<p>Use this connected sequence as your overarching road map when you want to read a continuous script from top to bottom:</p>\n<div class=\"script-box\">\n                    \"Hi [Name], it's Tim calling from 7J Finance. I know I'm catching you completely out of the blue—did I happen to catch you at an okay time for a quick chat?<br/><br/>\n                    [Wait for response]<br/><br/>\n                    Reason for my call is, we specialize specifically in secured loans and bridging finance for brokers who run into clients needing fast or complex funding solutions. I'm just reaching out to see what you're currently doing with those kinds of enquiries when they cross your desk?\"\n                </div>","SortOrder":2,"Active":true},{"SectionKey":"gatekeeper","Title":"3. Passing the Gatekeeper","Content":"<p>Do not pitch the receptionist. Sound casual, calm, and familiar like an internal peer.</p>\n<div class=\"script-box\">\n                    \"Hi, it's Tim here. Can you put me through to [Target Name] please?\"\n                </div>\n<div class=\"tip-box\">\n<strong>If asked what it's regarding:</strong> \"It's just regarding some bridging and secured loan enquiries. Is he at his desk?\"\n                </div>","SortOrder":3,"Active":true},{"SectionKey":"opener","Title":"4. The Smooth Opener (First 10 Seconds)","Content":"<p>Breathe, slow down your pace, and state your identity cleanly without stuttering.</p>\n<div class=\"script-box\">\n                    \"Hi [Name], it's Tim calling from 7J Finance. I know I'm catching you completely out of the blue—did I happen to catch you at an okay time for a quick chat?\"\n                </div>\n<div class=\"tip-box\">\n<strong>Key Habit:</strong> Pause right after asking if it's an okay time and wait for them to answer.\n                </div>","SortOrder":4,"Active":true},{"SectionKey":"value","Title":"5. Core Value Proposition","Content":"<p>Deliver your specialty cleanly and with absolute clarity:</p>\n<div class=\"script-box\">\n                    \"The reason for my call earlier is, we specialize specifically in secured loans and bridging finance for brokers who run into clients needing fast or complex funding solutions.\"\n                </div>","SortOrder":5,"Active":true},{"SectionKey":"question","Title":"6. The Engagement Question","Content":"<p>Transition smoothly into opening up the conversation:</p>\n<div class=\"script-box\">\n                    \"I'm just reaching out to see what you're currently doing with those kinds of enquiries when they cross your desk?\"\n                </div>","SortOrder":6,"Active":true},{"SectionKey":"handling","Title":"7. Common Objections & Objection Pivots","Content":"<p>Quick fallback lines if they throw an objection or try to brush you off:</p>\n<ul>\n<li><strong>\"We already have a lender / network we use:\"</strong><br/><em>\"Makes sense. Are you getting the turnaround times and communication you need from them, or do cases ever get stuck?\"</em></li>\n<li><strong>\"We don't really do much bridging / Only standard mortgages:\"</strong><br/><em>\"Fair enough. So when a regular mortgage client suddenly mentions they need short-term funds for a knockdown-rebuild or a quick purchase, do you guys have an outlet for that, or do you have to turn them away?\"</em></li>\n<li><strong>\"Just send me an email:\"</strong><br/><em>\"Will do, I'll send over my direct contact details. But just so I send something relevant, what type of cases cross your desk most often—residential bridging or commercial?\"</em></li>\n</ul>","SortOrder":7,"Active":true},{"SectionKey":"postcall","Title":"8. Post-Call & Email Follow-Up Template","Content":"<p>When someone tells you to email them, or if you had a brief introductory chat, send this template within 15 minutes while you're still fresh in their mind:</p>\n<div class=\"script-box\" style=\"font-style: normal; background: #ffffff; border-left-color: var(--primary); color: var(--text);\">\n<strong>Subject:</strong> Quick intro / 7J Finance - Bridging &amp; Secured Loans<br/><br/>\n                    Hi [Name],<br/><br/>\n                    Great speaking with you briefly earlier today (or: \"Apologies I missed you on the phone earlier!\").<br/><br/>\n                    As mentioned, I’m Tim, the BDM with <strong>7J Finance</strong>. We work closely with brokers to handle complex or fast-moving secured loan and bridging scenarios—especially when timelines are tight.<br/><br/>\n                    I know you have your standard panel, but whenever a tricky or time-sensitive case pops up that doesn't fit the box, keep us in mind as a reliable backup option.<br/><br/>\n                    Let me know if it makes sense to have a quick 5-minute chat next week, or feel free to drop a scenario over if you have one on your desk right now.<br/><br/>\n                    Best regards,<br/>\n<strong>Tim</strong><br/>\n                    BDM | 7J Finance\n                </div>","SortOrder":8,"Active":true},{"SectionKey":"openers","Title":"9. Opener Variations (First 30 Seconds)","Content":"<p>Have two or three openers ready so you don't sound robotic on back-to-back calls. Pick the one that fits the situation.</p>\n<div class=\"tip-box\"><strong>Soft intro (default):</strong> <em>\"Hi [Name], it's Tim here from 7J Finance. I know I'm catching you completely out of the blue — did I catch you at an okay time for a quick chat?\"</em></div>\n<div class=\"tip-box\"><strong>Direct intro (busy brokers):</strong> <em>\"Hi [Name], Tim from 7J Finance — I'll be brief. We specialise in bridging and secured loans for brokers. Is it worth a two-minute chat to see if we'd be a useful backup outlet for you?\"</em></div>\n<div class=\"tip-box\"><strong>Referral-style intro (warm lead):</strong> <em>\"Hi [Name], Tim from 7J Finance — [Referrer] suggested I give you a ring. We handle bridging and secured loan cases that need moving quickly. Have you got a couple of minutes?\"</em></div>","SortOrder":9,"Active":true},{"SectionKey":"discovery","Title":"10. Discovery & Qualification Questions","Content":"<p>Once they're talking, diagnose the case quickly. Don't interrogate — weave these in naturally.</p>\n<ul>\n<li><strong>Loan type:</strong> \"What sort of case is it — residential bridging, commercial, or development?\"</li>\n<li><strong>Amount:</strong> \"Roughly what sort of figure are we looking at?\"</li>\n<li><strong>Security/property:</strong> \"What's the security — residential, commercial, or a mix? And is it a first or second charge?\"</li>\n<li><strong>Timeline:</strong> \"How quickly does this need to complete — days or weeks?\"</li>\n<li><strong>Exit route:</strong> \"What's the exit — a refinance, a sale, or something else?\"</li>\n<li><strong>Regulated/unregulated:</strong> \"Is the borrower an individual (regulated) or a company (unregulated)?\"</li>\n<li><strong>Adverse credit:</strong> \"Any credit blips we should know about up front?\"</li>\n<li><strong>Current lender issue:</strong> \"Why isn't their usual lender fitting this one?\"</li>\n</ul>","SortOrder":10,"Active":true},{"SectionKey":"products","Title":"11. Bridging & Secured Loan Cheat-Sheet","Content":"<p>Generic reminders of what you're offering. Keep it compliance-safe — say \"we can look at\" or \"we specialise in finding routes for\", never \"we can definitely do X\". Confirm exact criteria, rates, and timelines on a case-by-case basis with the desk.</p>\n<ul>\n<li><strong>Bridging finance:</strong> short-term, fast-completion funding secured against property — used for purchases, chain breaks, auction buys, or refurb before a refinance.</li>\n<li><strong>Secured loans (second charge):</strong> longer-term borrowing behind an existing mortgage — often for capital raising, debt consolidation, or funding where a remortgage isn't ideal.</li>\n<li><strong>Development finance:</strong> for ground-up builds or heavy refurb, usually rolled-up interest and staged drawdowns.</li>\n<li><strong>Typical use cases:</strong> speed, complex/non-standard cases, or where mainstream lenders can't move quickly enough.</li>\n</ul>\n<div class=\"tip-box\"><strong>Don't promise:</strong> guaranteed approvals, exact rates, or fixed completion dates. <em>\"We can usually move quickly — I'll get the case in front of the right person and come back to you with realistic timelines.\"</em></div>","SortOrder":11,"Active":true},{"SectionKey":"checklist","Title":"12. Qualification Checklist","Content":"<p>Have these answers before you end the call — they let the desk assess whether it's a live case.</p>\n<ul>\n<li>☐ Loan type &amp; purpose</li>\n<li>☐ Approx. loan amount</li>\n<li>☐ Property/security type &amp; value</li>\n<li>☐ First or second charge</li>\n<li>☐ Required term</li>\n<li>☐ Exit route</li>\n<li>☐ Urgency / target completion date</li>\n<li>☐ Borrower type (individual vs company)</li>\n<li>☐ Any adverse credit</li>\n<li>☐ Documents available (valuation, ID, proof of income)</li>\n</ul>","SortOrder":12,"Active":true},{"SectionKey":"voicemail","Title":"13. Voicemail Script (Under 20 Seconds)","Content":"<p>If you hit voicemail, leave a short, specific message — never a rambling one.</p>\n<div class=\"script-box\">\n                    \"Hi [Name], it's Tim from 7J Finance. I specialise in bridging and secured loans for brokers. I'll try you again later today, or feel free to call me back on [your number]. Thanks, [Name].\"\n                </div>","SortOrder":13,"Active":true},{"SectionKey":"referral","Title":"14. The Referral Ask","Content":"<p>If they're not the right person or don't handle these cases, ask who is — this keeps the call productive.</p>\n<div class=\"script-box\">\n                    \"No problem at all. In your business, who normally handles bridging or secured loan enquiries when they come through? Would it be worth me having a quick chat with them?\"\n                </div>","SortOrder":14,"Active":true},{"SectionKey":"outcomes","Title":"15. Call Outcomes (Dispositions)","Content":"<p>Use the Power Dialer outcome buttons to log every call. Consistent dispositions make your numbers meaningful.</p>\n<ul>\n<li><strong>No Answer:</strong> rang out, no contact. Try again another day.</li>\n<li><strong>Voicemail:</strong> left a message. Note it so you don't double-leave.</li>\n<li><strong>Call Back:</strong> interested but not now. Set a callback date.</li>\n<li><strong>Not Interested:</strong> genuinely not a fit. Moves to Not Suitable so it stops appearing in your queue.</li>\n<li><strong>Appointment:</strong> a proper conversation booked. Gold.</li>\n<li><strong>Log Deal:</strong> a deal is live — opens the deal logger with the value and date.</li>\n</ul>","SortOrder":15,"Active":true},{"SectionKey":"kpis","Title":"16. Weekly KPI Targets","Content":"<p>Set your own numbers — these are a starting point for a cold-calling block. Track attempts, not just wins.</p>\n<ul>\n<li>Dials per day: <strong>40–60</strong></li>\n<li>Conversations per day: <strong>8–12</strong></li>\n<li>Voicemails per day: <strong>10–15</strong></li>\n<li>Appointments booked per week: <strong>3–5</strong></li>\n<li>Referrals asked for on every qualified call: <strong>100%</strong></li>\n</ul>\n<div class=\"tip-box\"><strong>Mindset:</strong> celebrate the dials, not just the deals. Activity beats anxiety.</div>","SortOrder":16,"Active":true},{"SectionKey":"compliance","Title":"17. Compliance-Safe Phrasing","Content":"<p>You're a broker-facing BDM, not advising borrowers directly. Keep language honest and non-advisory.</p>\n<ul>\n<li><strong>Say:</strong> \"We can look at…\" / \"We specialise in finding routes for…\" / \"We may be able to help with…\"</li>\n<li><strong>Don't say:</strong> \"We can definitely do it\" / \"Guaranteed approval\" / \"It'll complete in X days\" / \"Best rates guaranteed\"</li>\n<li><strong>On timelines:</strong> \"We can usually move quickly on these — I'll confirm realistic timelines once I've seen the case.\"</li>\n<li><strong>On rates/criteria:</strong> \"Rates and criteria depend on the case — I'll get you a clear indication once we have the detail.\"</li>\n<li><strong>If unsure:</strong> \"I don't want to guess — let me check with the desk and come back to you with something accurate.\"</li>\n</ul>","SortOrder":17,"Active":true}];
let callGuideSections = [];
let selectedGuideSectionKey = "";
function normaliseGuideSection(raw,index=1){ const key=String(raw?.SectionKey||raw?.key||('section_'+Date.now()+'_'+index)).trim().replace(/[^a-zA-Z0-9_-]/g,'_').toLowerCase(); return {SectionKey:key,Title:String(raw?.Title||raw?.title||'New Section'),Content:String(raw?.Content||raw?.content||'<p>Enter section content here.</p>'),SortOrder:Number(raw?.SortOrder||raw?.sortOrder||index)||index,Active:raw?.Active===false||['false','no','0'].includes(String(raw?.Active).toLowerCase())?false:true,Updated:raw?.Updated||'',UpdatedBy:raw?.UpdatedBy||''}; }
function sanitizeGuideHtml(input){ const doc=new DOMParser().parseFromString(String(input||''),'text/html'); doc.querySelectorAll('script,iframe,object,embed,style,link,form').forEach(n=>n.remove()); doc.querySelectorAll('*').forEach(el=>{ [...el.attributes].forEach(a=>{ const n=a.name.toLowerCase(),v=a.value||''; if(n.startsWith('on')||n==='srcdoc'||(['href','src'].includes(n)&&/^\s*javascript:/i.test(v))) el.removeAttribute(a.name); }); }); return doc.body.innerHTML; }
function getStoredCallGuide(){ try{const raw=localStorage.getItem(CALL_GUIDE_STORAGE_KEY);if(raw)return JSON.parse(raw).map(normaliseGuideSection);}catch(e){} return DEFAULT_CALL_GUIDE_SECTIONS.map(normaliseGuideSection); }
function saveStoredCallGuide(items){ callGuideSections=items.map(normaliseGuideSection).sort((a,b)=>a.SortOrder-b.SortOrder||a.Title.localeCompare(b.Title));localStorage.setItem(CALL_GUIDE_STORAGE_KEY,JSON.stringify(callGuideSections)); }
function getGuideSectionsSorted(){ return (callGuideSections||[]).filter(x=>x.Active).sort((a,b)=>a.SortOrder-b.SortOrder||a.Title.localeCompare(b.Title)); }
function renderCallGuide(){ const root=document.getElementById('guideDynamicRoot'),side=document.getElementById('guideSidebarList');if(!root||!side)return;const sections=getGuideSectionsSorted();side.innerHTML=sections.map(x=>`<li><a href="#${escapeHtml(x.SectionKey)}">${escapeHtml(x.Title)}</a></li>`).join('');root.innerHTML=sections.map(x=>`<section id="${escapeHtml(x.SectionKey)}" class="guide-section"><h3>${escapeHtml(x.Title)}</h3>${sanitizeGuideHtml(x.Content)}</section>`).join(''); }
async function cloudLoadCallGuide(){ const sid=await getSiteId(),listId=await resolveListId(getListName('callGuide'));let url=`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items?expand=fields&$top=500`,out=[];while(url){const data=await graphGet(url);(data.value||[]).forEach(it=>{const f=it.fields||{};out.push(normaliseGuideSection({SectionKey:f.SectionKey||f.Title,Title:f.Title,Content:f.Content||'',SortOrder:f.SortOrder,Active:f.Active,Updated:f.Updated,UpdatedBy:f.UpdatedBy,spId:it.id,etag:it['@odata.etag']}));});url=data['@odata.nextLink']||null;}return out.sort((a,b)=>a.SortOrder-b.SortOrder||a.Title.localeCompare(b.Title)); }
async function saveCallGuideCloud(section){ const sid=await getSiteId(),listId=await resolveListId(getListName('callGuide'));const data=await graphGet(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items?expand=fields&$top=500`);const found=(data.value||[]).find(it=>String(it.fields?.SectionKey||it.fields?.Title||'')===String(section.SectionKey));const fields={Title:section.Title,SectionKey:section.SectionKey,Content:sanitizeGuideHtml(section.Content),SortOrder:String(section.SortOrder),Active:section.Active===true?'Yes':'No',Updated:new Date().toISOString(),UpdatedBy:currentUser?.Email||''};if(found)await graphPatch(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items/${found.id}/fields`,fields,found['@odata.etag']||undefined);else await graphPost(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items`,{fields}); }
async function deleteCallGuideCloud(sectionKey){ const sid=await getSiteId(),listId=await resolveListId(getListName('callGuide'));const data=await graphGet(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items?expand=fields&$top=500`);const found=(data.value||[]).find(it=>String(it.fields?.SectionKey||it.fields?.Title||'')===String(sectionKey));if(found)await graphDelete(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items/${found.id}`); }
async function reloadCallGuideFromSource(showStatus=false){ try{ if(isDemoMode()||!m365Configured()){callGuideSections=getStoredCallGuide();renderCallGuide();renderCallGuideAdmin();if(showStatus)setGuideAdminStatus('Call Guide reloaded from Demo Mode.');return callGuideSections;}const cloud=await cloudLoadCallGuide();callGuideSections=cloud.length?cloud:getStoredCallGuide();if(!cloud.length)saveStoredCallGuide(callGuideSections);else localStorage.setItem(CALL_GUIDE_STORAGE_KEY,JSON.stringify(callGuideSections));renderCallGuide();renderCallGuideAdmin();if(showStatus)setGuideAdminStatus(cloud.length?'Shared Call Guide loaded from Microsoft List.':'No shared sections found yet — local defaults are displayed.');return callGuideSections;}catch(e){callGuideSections=getStoredCallGuide();renderCallGuide();renderCallGuideAdmin();if(showStatus)setGuideAdminStatus('Could not load shared Call Guide: '+(e.message||e));return callGuideSections;} }
function setGuideAdminStatus(msg){const el=document.getElementById('guideAdminStatus');if(el)el.textContent=msg||'';}
function renderCallGuideAdmin(){ const sel=document.getElementById('guideAdminSectionSelect');if(!sel)return;const arr=(callGuideSections||[]).slice().sort((a,b)=>a.SortOrder-b.SortOrder||a.Title.localeCompare(b.Title));sel.innerHTML=arr.map(x=>`<option value="${escapeHtml(x.SectionKey)}">${escapeHtml(x.Title)}${x.Active?'':' (hidden)'}</option>`).join('');if(selectedGuideSectionKey&&arr.some(x=>x.SectionKey===selectedGuideSectionKey))sel.value=selectedGuideSectionKey;else selectedGuideSectionKey=sel.value||'';const item=arr.find(x=>x.SectionKey===selectedGuideSectionKey);if(!item)return;document.getElementById('guideAdminTitle').value=item.Title;document.getElementById('guideAdminSort').value=item.SortOrder;document.getElementById('guideAdminContent').innerHTML=sanitizeGuideHtml(item.Content);document.getElementById('guideAdminActive').checked=item.Active!==false; }
function newCallGuideSection(){ const key='new_section_'+Date.now(),next=Math.max(0,...(callGuideSections||[]).map(x=>Number(x.SortOrder)||0))+1,item=normaliseGuideSection({SectionKey:key,Title:'New Call Guide Section',Content:'<p>Enter your new content here.</p>',SortOrder:next,Active:true});callGuideSections.push(item);selectedGuideSectionKey=key;saveStoredCallGuide(callGuideSections);renderCallGuide();renderCallGuideAdmin();document.getElementById('guideAdminTitle').focus();setGuideAdminStatus('New section created locally. Save it to publish it.'); }
async function saveCallGuideSection(){ const key=document.getElementById('guideAdminSectionSelect')?.value;if(!key)return;const idx=callGuideSections.findIndex(x=>x.SectionKey===key);if(idx<0)return;const item=callGuideSections[idx];item.Title=document.getElementById('guideAdminTitle').value.trim()||'Untitled Section';item.SortOrder=Math.max(1,parseInt(document.getElementById('guideAdminSort').value||idx+1,10));item.Content=sanitizeGuideHtml(document.getElementById('guideAdminContent').innerHTML);item.Active=document.getElementById('guideAdminActive').checked;item.Updated=new Date().toISOString();item.UpdatedBy=currentUser?.Email||'Demo User';saveStoredCallGuide(callGuideSections);renderCallGuide();renderCallGuideAdmin();if(isDemoMode()||!m365Configured()){setGuideAdminStatus('Section saved locally for Demo Mode.');return;}try{await saveCallGuideCloud(item);await reloadCallGuideFromSource(false);setGuideAdminStatus('Section saved to the shared Microsoft List.');}catch(e){setGuideAdminStatus('Saved locally, but Microsoft List update failed: '+(e.message||e));} }
async function deleteCallGuideSection(){ const key=document.getElementById('guideAdminSectionSelect')?.value;if(!key)return;const item=callGuideSections.find(x=>x.SectionKey===key);if(!item)return;if(!confirm('Delete the Call Guide section "'+item.Title+'"?'))return;callGuideSections=callGuideSections.filter(x=>x.SectionKey!==key);saveStoredCallGuide(callGuideSections);selectedGuideSectionKey=callGuideSections[0]?.SectionKey||'';renderCallGuide();renderCallGuideAdmin();if(!isDemoMode()&&m365Configured()){try{await deleteCallGuideCloud(key);setGuideAdminStatus('Section deleted from the shared Microsoft List.');}catch(e){setGuideAdminStatus('Deleted locally, but Microsoft List deletion failed: '+(e.message||e));}}else setGuideAdminStatus('Section deleted locally for Demo Mode.'); }
async function publishDefaultCallGuide(){ if(!confirm('Publish the built-in Call Guide defaults to the shared list? Existing sections with the same SectionKey will not be overwritten.'))return;const defaults=DEFAULT_CALL_GUIDE_SECTIONS.map(normaliseGuideSection);if(isDemoMode()||!m365Configured()){const existing=getStoredCallGuide(),keys=new Set(existing.map(x=>x.SectionKey));saveStoredCallGuide(existing.concat(defaults.filter(x=>!keys.has(x.SectionKey))));await reloadCallGuideFromSource(false);setGuideAdminStatus('Default sections added to Demo Mode.');return;}try{const existing=await cloudLoadCallGuide(),keys=new Set(existing.map(x=>x.SectionKey));for(const item of defaults)if(!keys.has(item.SectionKey))await saveCallGuideCloud(item);await reloadCallGuideFromSource(false);setGuideAdminStatus('Default sections published to the shared Microsoft List.');}catch(e){setGuideAdminStatus('Could not publish defaults: '+(e.message||e));} }

/* ---- Manual SharePoint list sources (admin override, per device) ----
   An admin can point the portal at a different list (e.g. a backup) without editing
   code. The override is stored in this browser only; to change it for the whole team,
   update M365_CONFIG.lists and redeploy. */
const LIST_OVERRIDE_KEY = "adt7j_list_override_v1";
function getListOverrides(){ try{ return JSON.parse(localStorage.getItem(LIST_OVERRIDE_KEY)) || {}; }catch(e){ return {}; } }
function saveListOverrides(ov){ localStorage.setItem(LIST_OVERRIDE_KEY, JSON.stringify(ov || {})); }
// Which list name to use for a data type? Override if set, else the M365_CONFIG default.
function getListName(key){
    const ov = getListOverrides();
    return (ov && ov[key]) ? ov[key] : M365_CONFIG.lists[key];
}
// Wipe every cached list id + id-map + snapshot so a switch to a different list is clean.
function clearListCaches(){
    for(const k in _listIdCache) delete _listIdCache[k];
    try{ setBrokerIdMap({}); }catch(e){}
    _brokerMeta = {};
    try{ setUserIdMap({}); }catch(e){}
    try{ setDealIdMap({}); }catch(e){}
    _brokerSnapshot = {};
    _userSnapshot = {};
}
// Try to resolve a list by display name; returns {ok, id, error}.
async function testListConnection(displayName){
    try{
        const id = await resolveListId(displayName);
        return { ok: true, id };
    }catch(e){
        return { ok: false, error: (e.message || String(e)) };
    }
}
async function resolveListId(displayName){
    if(_listIdCache[displayName]) return _listIdCache[displayName];
    const sid = await getSiteId();
    const filter = encodeURIComponent(`displayName eq '${displayName}'`);
    const data = await graphGet(`${GRAPH_BASE}/sites/${sid}/lists?$filter=${filter}&$select=id,displayName`);
    const found = (data.value||[]).find(l => (l.displayName||"").toLowerCase() === displayName.toLowerCase());
    if(!found){
        // Fallback: fetch all lists and match case-insensitively
        const all = await graphGet(`${GRAPH_BASE}/sites/${sid}/lists?$select=id,displayName`);
        const match = (all.value||[]).find(l => (l.displayName||"").toLowerCase() === displayName.toLowerCase());
        if(!match) throw new Error(`SharePoint list "${displayName}" not found. Create it on the site, or check the name in M365_CONFIG.lists.`);
        _listIdCache[displayName] = match.id;
        return match.id;
    }
    _listIdCache[displayName] = found.id;
    return found.id;
}

function getSharePointListUrl(listKey){
    const name=getListName(listKey); return `https://${M365_CONFIG.sharePointHost}${M365_CONFIG.sitePath}/Lists/${encodeURIComponent(name)}/AllItems.aspx`;
}
function getSharePointRecordUrl(listKey,itemId){
    if(!itemId) return getSharePointListUrl(listKey);
    return `${getSharePointListUrl(listKey)}?ID=${encodeURIComponent(itemId)}`;
}

function showCloudError(e){
    console.error("SharePoint sync error:", e);
    const el = document.getElementById("configWarning");
    if(el){
        el.style.display = "block";
        el.innerHTML = "<strong>SharePoint sync issue:</strong> " + (e.message||e) + " — your last change was saved locally and will retry. Check the list names and permissions in the setup guide.";
    }
}

/* ----------------------- Field mapping (broker <-> list item) ------------ */
function brokerToFields(b){
    const a=normaliseAssignedTo(b.AssignedTo);
    return {
        Title:b.Title||"", Company:b.Company||"", Phone:b.Phone||"", Email:b.Email||"", Website:b.Website||"",
        Address:b.Address||"", City:b.City||"", Notes:b.Notes||"", PrefComm:b.PrefComm||"",
        LoanTypes:(b.LoanTypes||[]).join("; "), Volume:b.Volume||"", Network:b.Network||"", Status:b.Status||"",
        NextFollowUp:b.NextFollowUp||"", LastContactDate:b.LastContactDate||"",
        AssignedTo:a ? ((a.Title||a.EMail||"") + (a.EMail ? " <"+a.EMail+">" : "")) : "",
        AssignedToEmail:a?.EMail||"", AssignedToName:a?.Title||"",
        IsNotSuitable:isNotSuitable(b)?"Yes":"No", PortalId:String(b.Id),
        NotSuitableReason:b.NotSuitableReason||"", DiallerOutcome:b.DiallerOutcome||"",
        DiallerOutcomeReason:b.DiallerOutcomeReason||"", DiallerOutcomeDate:b.DiallerOutcomeDate||"", NotSuitableSource:b.NotSuitableSource||"", WorkflowStateUpdatedAt:b.WorkflowStateUpdatedAt||"",
        ClaimedAt:b.ClaimedAt||"", ClaimedBy:b.ClaimedBy||"", ClaimExpiresAt:b.ClaimExpiresAt||""
    };
}
function fieldsToBroker(fields, fallbackId){
    const assignedTo=normaliseAssignedTo(fields.AssignedTo || fields.AssignedToEmail || fields.AssignedToUser);
    const localId=fields.PortalId ? (Number(fields.PortalId)||fallbackId) : fallbackId;
    return {
        Id:localId, Modified:fields.lastModifiedDateTime||fields.Modified||new Date().toISOString(), Title:fields.Title||"", Company:fields.Company||"",
        Phone:fields.Phone||"", Email:fields.Email||"", Website:fields.Website||"", Address:fields.Address||"", City:fields.City||"",
        Notes:fields.Notes||"", PrefComm:fields.PrefComm||"Phone", LoanTypes:(fields.LoanTypes?String(fields.LoanTypes).split(";").map(s=>s.trim()).filter(Boolean):["Residential Bridging"]),
        Volume:fields.Volume||"Under £1M", Network:fields.Network||"", Status:fields.Status||"Cold", NextFollowUp:fields.NextFollowUp||"",
        LastContactDate:fields.LastContactDate||fields.Modified||"", AssignedTo:assignedTo,
        IsNotSuitable:(String(fields.IsNotSuitable||"").toLowerCase()==="yes" || String(fields.IsNotSuitable||"").toLowerCase()==="true"),
        NotSuitableReason:fields.NotSuitableReason||"", DiallerOutcome:fields.DiallerOutcome||"", DiallerOutcomeReason:fields.DiallerOutcomeReason||"", DiallerOutcomeDate:fields.DiallerOutcomeDate||"",
        NotSuitableSource:fields.NotSuitableSource||"", WorkflowStateUpdatedAt:fields.WorkflowStateUpdatedAt||"", ClaimedAt:fields.ClaimedAt||"", ClaimedBy:fields.ClaimedBy||"", ClaimExpiresAt:fields.ClaimExpiresAt||""
    };
}

/* ----------------------- Local-id <-> SharePoint-item-id map ------------ */
const BROKER_IDMAP_KEY = "adt7j_broker_idmap_v1";
let _brokerMeta = {}; // local broker id -> {spId, etag}

function getBrokerIdMap(){ try{ return JSON.parse(localStorage.getItem(BROKER_IDMAP_KEY)) || {}; }catch(e){ return {}; } }
function setBrokerIdMap(m){ localStorage.setItem(BROKER_IDMAP_KEY, JSON.stringify(m)); }

/* ----------------------- Cloud: load / push / delete brokers ------------- */
async function cloudLoadBrokers(){
    const sid=await getSiteId(), listId=await resolveListId(getListName("brokers"));
    let url=`${GRAPH_BASE}/sites/${listId?`lists/${listId}`:`lists/${listId}`}/items?expand=fields&$top=500`;
    // The URL above intentionally uses the resolved list id; keep it explicit for readability.
    url=`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items?expand=fields&$top=500`;
    const brokers=[], idmap={}, meta={}; let nextId=Date.now();
    while(url){
        const data=await graphGet(url);
        (data.value||[]).forEach(it=>{
            const f=it.fields||{};
            const localId=f.PortalId ? (Number(f.PortalId)||nextId++) : nextId++;
            idmap[localId]=it.id; meta[localId]={spId:it.id,etag:it["@odata.etag"]||""};
            const b=fieldsToBroker(f,localId);
            if(!b.AssignedTo){ b.AssignedToEmail=String(f.AssignedToEmail||f.AssignedToEMail||"").toLowerCase(); b.AssignedToName=String(f.AssignedToName||""); b.AssignedToLookupId=f.AssignedToLookupId||f.AssignedToId||""; }
            brokers.push(b);
        });
        url=data["@odata.nextLink"]||null;
    }
    _brokerMeta=meta; setBrokerIdMap(idmap); return brokers;
}
async function cloudGetCurrentBroker(localId){
    const sid=await getSiteId(), listId=await resolveListId(getListName("brokers")), map=getBrokerIdMap();
    const spId=map[localId]||_brokerMeta[localId]?.spId; if(!spId) return null;
    const it=await graphGet(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items/${spId}?expand=fields`);
    const b=fieldsToBroker(it.fields||{},Number(localId));
    _brokerMeta[localId]={spId:it.id,etag:it["@odata.etag"]||""};
    return {broker:b,etag:it["@odata.etag"]||"",spId:it.id};
}
let _brokerColumnNames=null;
async function getBrokerColumnNames(){
    if(_brokerColumnNames) return _brokerColumnNames;
    try{ const sid=await getSiteId(), listId=await resolveListId(getListName("brokers")); const d=await graphGet(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/columns?$select=name,displayName`); _brokerColumnNames=new Set((d.value||[]).flatMap(c=>[c.name,c.displayName]).filter(Boolean)); }catch(e){ _brokerColumnNames=new Set(); }
    return _brokerColumnNames;
}
function filterBrokerFieldsForColumns(fields,columns){
    const out={}; const optional=new Set(["ClaimedAt","ClaimedBy","ClaimExpiresAt","AssignedToLookupId","WorkflowStateUpdatedAt"]); for(const [k,v] of Object.entries(fields)){ if(!optional.has(k) || columns.has(k)) out[k]=v; } return out;
}
async function cloudPushBroker(b){
    const sid=await getSiteId(), listId=await resolveListId(getListName("brokers")), map=getBrokerIdMap();
    const fields=filterBrokerFieldsForColumns(brokerToFields(b),await getBrokerColumnNames()); const spId=map[b.Id];
    if(spId){
        const etag=_brokerMeta[b.Id]?.etag||"";
        try{ await graphPatch(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items/${spId}/fields`,fields,etag||undefined); }
        catch(e){ if(e.status===412) throw new Error("This broker was changed by another user. Refresh the broker before saving your change."); throw e; }
    }else{
        const created=await graphPost(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items`,{fields}); map[b.Id]=created.id; setBrokerIdMap(map); _brokerMeta[b.Id]={spId:created.id,etag:created["@odata.etag"]||""};
    }
}
async function cloudDeleteBroker(b){
    const sid = await getSiteId();
    const listId = await resolveListId(getListName("brokers"));
    const map = getBrokerIdMap();
    const targetPortalId = String(b?.Id ?? "");
    const targetEmail = String(b?.Email || "").trim().toLowerCase();
    const targetCompany = String(b?.Company || "").trim().toLowerCase();
    const targetTitle = String(b?.Title || "").trim().toLowerCase();

    // First try the cached SharePoint item id/ETag. This is the normal path.
    let current = null;
    try{ current = await cloudGetCurrentBroker(b.Id); }catch(e){ current = null; }
    if(current){
        try{
            await graphDelete(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items/${current.spId}`, current.etag || undefined);
            delete map[b.Id]; setBrokerIdMap(map); delete _brokerMeta[b.Id];
            return;
        }catch(e){
            // A stale cache/ETag or a missing item should not prevent the admin fallback.
            if(e.status !== 404 && e.status !== 412) throw e;
        }
    }

    // Admin fallback: resolve the real SharePoint item directly from the list.
    // This is important for imported/legacy brokers whose PortalId or local id-map is missing.
    let url = `${GRAPH_BASE}/sites/${sid}/lists/${listId}/items?expand=fields&$top=500`;
    const matches = [];
    while(url){
        const data = await graphGet(url);
        for(const it of (data.value||[])){
            const f=it.fields||{};
            const portal=String(f.PortalId||"").trim();
            const email=String(f.Email||"").trim().toLowerCase();
            const company=String(f.Company||"").trim().toLowerCase();
            const title=String(f.Title||"").trim().toLowerCase();
            let match=false;
            if(targetPortalId && portal && portal===targetPortalId) match=true;
            else if(targetEmail && email && targetEmail===email && targetCompany && company===targetCompany) match=true;
            else if(!targetEmail && targetCompany && company===targetCompany && targetTitle && title===targetTitle) match=true;
            else if(!targetEmail && targetCompany && company===targetCompany && !targetTitle) match=true;
            if(match) matches.push(it);
        }
        url=data["@odata.nextLink"]||null;
    }

    // Last-resort exact name/company match for legacy Apex-style records.
    if(!matches.length && targetCompany){
        url = `${GRAPH_BASE}/sites/${sid}/lists/${listId}/items?expand=fields&$top=500`;
        while(url){
            const data=await graphGet(url);
            for(const it of (data.value||[])){
                const f=it.fields||{};
                const company=String(f.Company||"").trim().toLowerCase();
                const title=String(f.Title||"").trim().toLowerCase();
                if(company===targetCompany || title===targetCompany) matches.push(it);
            }
            url=data["@odata.nextLink"]||null;
        }
    }

    if(!matches.length){
        throw new Error(`Could not locate \"${b.Company || b.Title || 'this broker'}\" in the configured Broker List. The record may be in a different SharePoint list/source.`);
    }

    let deleted=0;
    for(const it of matches){
        try{
            await graphDelete(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items/${it.id}`, it["@odata.etag"] || undefined);
            deleted++;
        }catch(e){
            if(e.status===404) deleted++;
            else if(e.status===412){
                // The item changed since we read it. Re-fetch its current ETag once, then retry.
                const fresh=await graphGet(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items/${it.id}?expand=fields`);
                await graphDelete(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items/${it.id}`, fresh["@odata.etag"] || undefined);
                deleted++;
            }else throw e;
        }
    }
    delete map[b.Id]; setBrokerIdMap(map); delete _brokerMeta[b.Id];
    if(!deleted) throw new Error(`SharePoint returned no deleted records for \"${b.Company || b.Title || 'this broker'}\".`);
}

async function deleteBrokerAdmin(itemId, reason = "Manual removal by admin") {
    const target = String(itemId ?? "");
    const items = getStoredBrokers();

    // Resolve the record from the exact Admin card first, then by all stable identifiers.
    let local = items.find(i =>
        String(i.Id ?? "") === target ||
        String(i.PortalId ?? "") === target
    );

    if(!local && typeof allItems !== "undefined"){
        const rendered = allItems.find(i =>
            String(i.Id ?? "") === target ||
            String(i.PortalId ?? "") === target
        );
        if(rendered){
            const company = String(rendered.Company ?? "").trim().toLowerCase();
            const email = String(rendered.Email ?? "").trim().toLowerCase();
            local = items.find(i =>
                String(i.Id ?? "") === String(rendered.Id ?? "") ||
                String(i.PortalId ?? "") === String(rendered.PortalId ?? "") ||
                (company && email &&
                 String(i.Company ?? "").trim().toLowerCase() === company &&
                 String(i.Email ?? "").trim().toLowerCase() === email)
            ) || rendered;
        }
    }

    if(!local) throw new Error("Broker record could not be resolved in the current CRM data.");

    // Live Microsoft 365 mode: delete the actual SharePoint item.
    if(m365Configured() && !isDemoMode()){
        try{
            await cloudPushAudit({
                RecordId:Number(local.Id),
                Timestamp:new Date().toISOString(),
                User:currentUser?.Email||currentUser?.Title||'Admin',
                Action:'DELETE',
                RecordTitle:local.Title||'',
                Company:local.Company||'',
                Reason:reason,
                Trail:'auditAccount'
            });
        }catch(e){ console.warn('Delete audit could not be written:', e); }

        await cloudDeleteBroker(local);
        const fresh=await cloudLoadBrokers();
        localStorage.setItem(STORAGE_KEY,JSON.stringify(fresh));
        _brokerSnapshot=_snapshotBrokers(fresh);
        allItems=fresh;
        return;
    }

    // Demo/local mode: remove by every stable identity and create a permanent tombstone.
    const idSet=new Set(
        [itemId,local.Id,local.PortalId,local.PortalID]
        .filter(v=>v!==null && v!==undefined && String(v)!=="")
        .map(String)
    );
    const company=String(local.Company??"").trim().toLowerCase();
    const email=String(local.Email??"").trim().toLowerCase();
    const portal=String(local.PortalId??"").trim().toLowerCase();
    const key=(company||"")+"|"+(email||"");

    logAuditRecord(local,"DELETE",reason);

    const remaining=items.filter(i=>{
        const iCompany=String(i.Company??"").trim().toLowerCase();
        const iEmail=String(i.Email??"").trim().toLowerCase();
        const iKey=iCompany+"|"+iEmail;
        return !idSet.has(String(i.Id??"")) &&
               !idSet.has(String(i.PortalId??"")) &&
               !(key!=="|" && iKey===key);
    });

    // Preserve deleted identities so Demo Mode's seed cannot recreate this broker.
    const deleted=JSON.parse(localStorage.getItem("7J_DEMO_DELETED_BROKERS")||"[]");
    idSet.forEach(v=>{if(!deleted.includes(v))deleted.push(v);});
    localStorage.setItem("7J_DEMO_DELETED_BROKERS",JSON.stringify(deleted));

    const keys=JSON.parse(localStorage.getItem("7J_DEMO_DELETED_KEYS")||"[]");
    if(key!=="|" && !keys.includes(key))keys.push(key);
    if(portal && !keys.includes("portal:"+portal))keys.push("portal:"+portal);
    localStorage.setItem("7J_DEMO_DELETED_KEYS",JSON.stringify(keys));

    localStorage.setItem(STORAGE_KEY,JSON.stringify(remaining));
    allItems=remaining;
    _brokerSnapshot=_snapshotBrokers(remaining);
}

/* ----------------------- Cloud: load / push audit logs ------------------ */
function auditTrailKey(actionType){
    return /^CALL_/.test(actionType) || actionType === "DEAL_LOGGED" ? "auditDialer" : "auditAccount";
}
async function cloudLoadAuditFromList(listKey){
    const sid = await getSiteId();
    const listId = await resolveListId(getListName(listKey));
    let url = `${GRAPH_BASE}/sites/${sid}/lists/${listId}/items?expand=fields&$top=500`;
    const logs = [];
    while(url){
        const data = await graphGet(url);
        (data.value||[]).forEach(it=>{
            const f=it.fields||{};
            logs.push({Timestamp:f.Timestamp||it.createdDateTime||"",User:f.User||"",Action:f.Title||f.Action||"",RecordId:f.RecordId?Number(f.RecordId):0,RecordTitle:f.RecordTitle||"",Company:f.Company||"",Reason:f.Reason||"",Trail:listKey});
        });
        url=data["@odata.nextLink"]||null;
    }
    return logs;
}
async function cloudLoadAudit(){
    const results=await Promise.all([cloudLoadAuditFromList("auditAccount"),cloudLoadAuditFromList("auditDialer")]);
    return results.flat().sort((a,b)=>(b.Timestamp||"").localeCompare(a.Timestamp||""));
}
function hydrateBrokerOutcomeMetadata(brokers, logs){
    // SharePoint broker state is authoritative. Audit history is only used to fill
    // legacy records that do not yet carry the newer outcome fields. A broker state
    // change newer than an audit entry must never be overwritten by an older outcome.
    const latestById={};
    (logs||[]).forEach(l=>{
        const id=String(l.RecordId||""); if(!id) return;
        const action=String(l.Action||"").toUpperCase();
        if(!/^CALL_/.test(action) && action!=="MARKED_NOT_SUITABLE" && action!=="REVERTED_TO_OPEN") return;
        if(!latestById[id] || String(l.Timestamp||"")>String(latestById[id].Timestamp||"")) latestById[id]=l;
    });
    brokers.forEach(b=>{
        if(String(b.WorkflowState||"").toUpperCase()==="OPEN"){
            b.IsNotSuitable=false;
            b.NotSuitableReason="";
            b.NotSuitableSource="";
            b.DiallerOutcome="";
            b.DiallerOutcomeReason="";
            b.DiallerOutcomeDate="";
            return;
        }
        const l=latestById[String(b.Id)]; if(!l) return;
        const auditTime=Date.parse(l.Timestamp||"")||0;
        const stateTime=Date.parse(b.WorkflowStateUpdatedAt||b.Modified||"")||0;
        const explicitState = b.WorkflowStateUpdatedAt || b.IsNotSuitable===true || b.DiallerOutcome || b.NotSuitableReason;
        // If the broker itself was updated after the audit entry, trust SharePoint.
        // A small tolerance covers the audit row being written milliseconds after the broker update.
        if(stateTime && auditTime && stateTime >= auditTime-3000) return;
        const action=String(l.Action||"").toUpperCase();
        if(action==="REVERTED_TO_OPEN"){
            b.IsNotSuitable=false; b.NotSuitableReason=""; b.NotSuitableSource=""; b.DiallerOutcome=""; b.DiallerOutcomeReason=""; b.DiallerOutcomeDate=""; return;
        }
        if(explicitState && !b.DiallerOutcome && !b.NotSuitableReason && !b.IsNotSuitable) return;
        if(action==="MARKED_NOT_SUITABLE"){
            b.IsNotSuitable=true; b.NotSuitableReason=b.NotSuitableReason||l.Reason||"Manual"; b.NotSuitableSource=b.NotSuitableSource||"Manual"; b.DiallerOutcomeDate=b.DiallerOutcomeDate||l.Timestamp||""; return;
        }
        const o=action.replace(/^CALL_/i,"").toUpperCase();
        if(!b.DiallerOutcome) b.DiallerOutcome=o;
        if(!b.DiallerOutcomeDate) b.DiallerOutcomeDate=l.Timestamp||"";
        if(["NOT_INTERESTED","NO_PRODUCT","NO_SUITABLE_PRODUCT","OUT_OF_AREA","COMPLIANCE","COMPLIANCE_UNREGULATED","DUPLICATE","DUPLICATE_INACTIVE"].includes(o)){
            b.IsNotSuitable=true; b.NotSuitableReason=b.NotSuitableReason||l.Reason||o.replace(/_/g," "); b.NotSuitableSource=b.NotSuitableSource||"Dialler";
        }
    });
    return brokers;
}
async function cloudPushAudit(entry){
    const sid=await getSiteId();
    const trail=auditTrailKey(entry.Action);
    const listId=await resolveListId(getListName(trail));
    await graphPost(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items`,{fields:{
        Title:entry.Action, Timestamp:entry.Timestamp, User:entry.User,
        RecordTitle:entry.RecordTitle, Company:entry.Company, Reason:entry.Reason,
        RecordId:String(entry.RecordId), Trail:trail
    }});
}

/* ----------------------- Cloud: load / push users (roster + role) --------- */
// Users use the same id-map pattern as brokers: local Id -> SharePoint item id.
const USER_IDMAP_KEY = "adt7j_user_idmap_v1";
function getUserIdMap(){ try{ return JSON.parse(localStorage.getItem(USER_IDMAP_KEY)) || {}; }catch(e){ return {}; } }
function setUserIdMap(m){ localStorage.setItem(USER_IDMAP_KEY, JSON.stringify(m)); }

async function cloudLoadUsers(){
    const sid = await getSiteId();
    const listId = await resolveListId(getListName("users"));
    let url = `${GRAPH_BASE}/sites/${sid}/lists/${listId}/items?expand=fields&$top=200`;
    const users = []; const idmap = {};
    while(url){
        const data = await graphGet(url);
        (data.value||[]).forEach(it=>{
            const f = it.fields||{};
            const localId = f.PortalId ? (Number(f.PortalId) || it.id) : it.id;
            idmap[localId] = it.id;
            users.push({ Id: localId, Title: f.Title||f.Name||"", Email: f.Email||"", Role: f.Role||"BDM" });
        });
        url = data["@odata.nextLink"] || null;
    }
    setUserIdMap(idmap);
    return users;
}
async function cloudPushUser(u){
    const sid = await getSiteId();
    const listId = await resolveListId(getListName("users"));
    const map = getUserIdMap();
    const fields = { Title: u.Title||"", Email: u.Email||"", Role: u.Role||"BDM", PortalId: String(u.Id) };
    const spId = map[u.Id];
    if(spId){
        await graphPatch(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items/${spId}/fields`, fields);
    } else {
        const created = await graphPost(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items`, { fields });
        map[u.Id] = created.id; setUserIdMap(map);
    }
}
async function cloudDeleteUser(u){
    const sid = await getSiteId();
    const listId = await resolveListId(getListName("users"));
    const map = getUserIdMap();
    const spId = map[u.Id];
    if(spId){
        await graphDelete(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items/${spId}`);
        delete map[u.Id]; setUserIdMap(map);
    }
}
