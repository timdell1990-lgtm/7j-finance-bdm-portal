const USERS_STORAGE_KEY = "adt7j_users_db_v1";
const SESSION_KEY = "adt7j_current_session_v1";
const STORAGE_KEY = "adt7j_brokers_db_v3";
const AUDIT_STORAGE_KEY = "adt7j_audit_archive_v2";

const DEFAULT_USERS = [
    { Id: 1, Title: "Tim (Admin Demo)", Email: "admin.demo@local", Role: "Admin" },
    { Id: 2, Title: "Sarah (Demo BDM)", Email: "sarah.demo@local", Role: "BDM" },
    { Id: 3, Title: "James (Demo BDM)", Email: "james.demo@local", Role: "BDM" }
];

const DEFAULT_BROKERS = [
    {
        Id: 101,
        Modified: new Date().toISOString(),
        Title: "John Smith",
        Company: "Apex Mortgage Brokers",
        Phone: "020 7946 0912",
        Email: "john@apexbrokers.co.uk",
        Website: "https://apexbrokers.co.uk",
        Address: "12 Financial Way",
        City: "London",
        Notes: "Interested in quick turnaround times for residential bridging.",
        PrefComm: "Phone",
        LoanTypes: ["Residential Bridging", "Auction Finance"],
        Volume: "£1M–£5M",
        Network: "PMS",
        Status: "Warm",
        NextFollowUp: "",
        LastContactDate: new Date().toISOString(),
        AssignedTo: null,
        IsNotSuitable: false
    },
    {
        Id: 102,
        Modified: new Date().toISOString(),
        Title: "Sarah Jenkins",
        Company: "Meridian Finance Solutions",
        Phone: "0161 496 0145",
        Email: "s.jenkins@meridianfin.co.uk",
        Website: "https://meridianfin.co.uk",
        Address: "45 King Street",
        City: "Manchester",
        Notes: "Focuses heavily on commercial properties.",
        PrefComm: "Email",
        LoanTypes: ["Commercial Bridging", "Development Finance", "HMO / Multi-Unit Blocks"],
        Volume: "£5M+",
        Network: "TMA",
        Status: "Active Introducer",
        NextFollowUp: "",
        LastContactDate: new Date(Date.now() - 70 * 86400000).toISOString(),
        AssignedTo: { Title: "Sarah Jenkins", EMail: "other@broker.com" },
        IsNotSuitable: false
    }
];

function getStoredUsers() {
    const raw = localStorage.getItem(USERS_STORAGE_KEY);
    if (!raw) {
        localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(DEFAULT_USERS));
        return DEFAULT_USERS;
    }
    try { return JSON.parse(raw); } catch(e) { return DEFAULT_USERS; }
}

function saveStoredUsers(users) {
    localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(users));
    // Fire-and-forget sync of changed users to the BDM Users list
    cloudSyncUsers(users).catch(e => showCloudError(e));
}

// Snapshot used to detect which user records changed since last sync
let _userSnapshot = {};
function _snapshotUsers(users){ const s={}; users.forEach(u=> s[u.Id] = JSON.stringify(u)); return s; }
async function cloudSyncUsers(users){
    if(!m365Configured()) return;
    const snap = _snapshotUsers(users);
    const prev = _userSnapshot;
    // Deletions first
    for(const id in prev){
        if(!(id in snap)){ await cloudDeleteUser(JSON.parse(prev[id])); }
    }
    // Creates + updates — cloudPushUser decides PATCH vs POST using the id map
    for(const id in snap){
        if(!(id in prev) || prev[id] !== snap[id]){
            await cloudPushUser(JSON.parse(snap[id]));
        }
    }
    _userSnapshot = snap;
}

let GLOBAL_SETTINGS_CACHE = null;
const GLOBAL_SETTINGS_DEFAULTS = { companyName:'7J Finance', tenantId:'', sharePointHost:'7jfinance.sharepoint.com', sitePath:'/sites/BDM', authMode:'tenant', aiEnabled:true, aiKey:"", aiModel:"gpt-5.6-mini", aiThreshold:25, aiProxy:"", liveRefreshSeconds:30, claimMinutes:5 };
function getGlobalSettingValue(key, fallback){
    if(GLOBAL_SETTINGS_CACHE && GLOBAL_SETTINGS_CACHE[key] !== undefined) return GLOBAL_SETTINGS_CACHE[key];
    const map={companyName:'adt7j_company_name_v1',tenantId:'adt7j_tenant_id_v1',sharePointHost:'adt7j_sp_host_v1',sitePath:'adt7j_sp_path_v1',authMode:'adt7j_auth_mode_v1',releaseDays:'adt7j_broker_release_days_v1',aiEnabled:'adt7j_ai_enabled_v1',aiKey:'adt7j_ai_key_v1',aiModel:'adt7j_ai_model_v1',aiThreshold:'adt7j_ai_threshold_v1',aiProxy:'adt7j_ai_proxy_v1',liveRefreshSeconds:'adt7j_live_refresh_seconds_v1',claimMinutes:'adt7j_claim_minutes_v1'};
    const raw=map[key]?localStorage.getItem(map[key]):null;
    if(raw!==null) return raw; return fallback;
}
function getBrokerReleaseDays(){
    const n=parseInt(getGlobalSettingValue('releaseDays',42),10);
    return Number.isFinite(n)&&n>0?n:42;
}
function saveBrokerReleaseDays(){ return saveGlobalPortalSettings(); }
function resetBrokerReleaseDays(){
    document.getElementById('globalReleaseDays').value='42';
    saveGlobalPortalSettings();
}
function loadBrokerReleaseDaysSetting(){
    const el=document.getElementById('brokerReleaseDays'); if(el) el.value=String(getBrokerReleaseDays());
    const ge=document.getElementById('globalReleaseDays'); if(ge) ge.value=String(getBrokerReleaseDays());
}
async function loadGlobalPortalSettings(showStatus=false){
    let cfg={...GLOBAL_SETTINGS_DEFAULTS};
    if(isDemoMode() || !m365Configured()){
        cfg.companyName=localStorage.getItem('adt7j_company_name_v1')||'7J Finance'; cfg.tenantId=localStorage.getItem('adt7j_tenant_id_v1')||M365_CONFIG.tenantId; cfg.sharePointHost=localStorage.getItem('adt7j_sp_host_v1')||M365_CONFIG.sharePointHost; cfg.sitePath=localStorage.getItem('adt7j_sp_path_v1')||M365_CONFIG.sitePath; cfg.authMode=localStorage.getItem('adt7j_auth_mode_v1')||'tenant'; cfg.releaseDays=parseInt(localStorage.getItem('adt7j_broker_release_days_v1')||'42',10)||42;
        cfg.aiEnabled=localStorage.getItem('adt7j_ai_enabled_v1')!=='false';
        cfg.aiKey=localStorage.getItem('adt7j_ai_key_v1')||''; cfg.aiModel=localStorage.getItem('adt7j_ai_model_v1')||'gpt-5.6-mini'; cfg.aiThreshold=parseInt(localStorage.getItem('adt7j_ai_threshold_v1')||'25',10)||25; cfg.aiProxy=localStorage.getItem('adt7j_ai_proxy_v1')||''; cfg.liveRefreshSeconds=parseInt(localStorage.getItem('adt7j_live_refresh_seconds_v1')||'30',10)||30; cfg.claimMinutes=parseInt(localStorage.getItem('adt7j_claim_minutes_v1')||'5',10)||5;
    } else {
        try {
            const sid=await getSiteId(); const listId=await resolveListId(getListName('globalSettings'));
            const data=await graphGet(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items?expand=fields&$top=200`);
            (data.value||[]).forEach(it=>{const f=it.fields||{}; const k=f.SettingKey||f.Title; if(k) cfg[k]=f.SettingValue;});
        } catch(e){ if(showStatus) alert('Could not load global settings: '+(e.message||e)); }
    }
    // Global SharePoint source routing: the Global Portal Settings list is the
    // single source of truth for which Lists the whole team uses.
    if(cfg.listSources){
        try{
            const globalSources=typeof cfg.listSources==='string'?JSON.parse(cfg.listSources):cfg.listSources;
            if(globalSources && typeof globalSources==='object'){ saveListOverrides(globalSources); clearListCaches(); }
        }catch(e){ console.warn('Global list source settings could not be parsed:',e); }
    }
    GLOBAL_SETTINGS_CACHE={...cfg};
    if(cfg.tenantId){M365_CONFIG.tenantId=cfg.tenantId;} if(cfg.sharePointHost){M365_CONFIG.sharePointHost=cfg.sharePointHost;} if(cfg.sitePath){M365_CONFIG.sitePath=cfg.sitePath;} localStorage.setItem('adt7j_company_name_v1',cfg.companyName||'7J Finance'); localStorage.setItem('adt7j_tenant_id_v1',M365_CONFIG.tenantId||''); localStorage.setItem('adt7j_sp_host_v1',M365_CONFIG.sharePointHost||''); localStorage.setItem('adt7j_sp_path_v1',M365_CONFIG.sitePath||''); localStorage.setItem('adt7j_auth_mode_v1',cfg.authMode||'tenant'); applyCompanyName(cfg.companyName||'7J Finance');
    localStorage.setItem('adt7j_broker_release_days_v1',String(parseInt(cfg.releaseDays,10)||42));
    localStorage.setItem('adt7j_ai_enabled_v1',String(cfg.aiEnabled!==false && cfg.aiEnabled!=='false'));
    localStorage.setItem('adt7j_ai_key_v1',cfg.aiKey||''); localStorage.setItem('adt7j_ai_model_v1',cfg.aiModel||'gpt-5.6-mini'); localStorage.setItem('adt7j_ai_threshold_v1',String(Math.min(25,Math.max(1,parseInt(cfg.aiThreshold,10)||25)))); localStorage.setItem('adt7j_ai_proxy_v1',cfg.aiProxy||''); localStorage.setItem('adt7j_live_refresh_seconds_v1',String(Math.min(300,Math.max(15,parseInt(cfg.liveRefreshSeconds,10)||30)))); localStorage.setItem('adt7j_claim_minutes_v1',String(Math.min(30,Math.max(1,parseInt(cfg.claimMinutes,10)||5))));
    const cn=document.getElementById('portalCompanyName'); if(cn) cn.value=cfg.companyName||'7J Finance'; const te=document.getElementById('portalTenantId'); if(te) te.value=cfg.tenantId||M365_CONFIG.tenantId; const sh=document.getElementById('portalSharePointHost'); if(sh) sh.value=cfg.sharePointHost||M365_CONFIG.sharePointHost; const sp=document.getElementById('portalSitePath'); if(sp) sp.value=cfg.sitePath||M365_CONFIG.sitePath; const am=document.getElementById('portalAuthMode'); if(am) am.value=cfg.authMode||'tenant'; const cid=document.getElementById('portalClientId'); if(cid) cid.value=M365_CONFIG.clientId;
    const r=document.getElementById('globalReleaseDays'); if(r) r.value=String(parseInt(cfg.releaseDays,10)||42);
    const a=document.getElementById('globalAIEnabled'); if(a) { a.checked=!(cfg.aiEnabled===false||String(cfg.aiEnabled).trim().toLowerCase()==='false'); a.dispatchEvent(new Event('change')); }
    const k=document.getElementById('globalAIKey'); if(k) k.value=cfg.aiKey||''; const m=document.getElementById('globalAIModel'); if(m) m.value=cfg.aiModel||'gpt-5.6-mini'; const th=document.getElementById('globalAIThreshold'); if(th) th.value=String(Math.min(25,Math.max(1,parseInt(cfg.aiThreshold,10)||25))); const pr=document.getElementById('globalAIProxy'); if(pr) pr.value=cfg.aiProxy||''; const lr=document.getElementById('globalLiveRefresh'); if(lr) lr.value=String(Math.min(300,Math.max(15,parseInt(cfg.liveRefreshSeconds,10)||30))); const cm=document.getElementById('globalClaimMinutes'); if(cm) cm.value=String(Math.min(30,Math.max(1,parseInt(cfg.claimMinutes,10)||5)));
    updateAIBrokerSearchVisibility();
    if(showStatus){const st=document.getElementById('globalSettingsStatus');if(st)st.textContent='Global settings loaded.';}
    return cfg;
}
async function saveGlobalPortalSettings(){
    try{
      if(typeof notifyGlobalSettingsSaved==='function') notifyGlobalSettingsSaved('Saving global settings…','ok');
    }catch(_e){}
    const companyName=document.getElementById('portalCompanyName')?.value.trim()||getGlobalSettingValue('companyName','7J Finance'); const tenantId=document.getElementById('portalTenantId')?.value.trim()||M365_CONFIG.tenantId; const sharePointHost=document.getElementById('portalSharePointHost')?.value.trim()||M365_CONFIG.sharePointHost; const sitePath=document.getElementById('portalSitePath')?.value.trim()||M365_CONFIG.sitePath; const authMode=document.getElementById('portalAuthMode')?.value||'tenant';
    const release=Math.max(1,Math.min(3650,parseInt(document.getElementById('globalReleaseDays')?.value||'42',10)));
    const aiEnabled=document.getElementById('globalAIEnabled')?.checked===true;
    const aiKey=document.getElementById('globalAIKey')?.value||''; const aiModel=document.getElementById('globalAIModel')?.value||'gpt-5.6-mini'; const aiThreshold=Math.min(25,Math.max(1,parseInt(document.getElementById('globalAIThreshold')?.value||'25',10))); const aiProxy=document.getElementById('globalAIProxy')?.value||''; const liveRefreshSeconds=Math.min(300,Math.max(15,parseInt(document.getElementById('globalLiveRefresh')?.value||'30',10))); const claimMinutes=Math.min(30,Math.max(1,parseInt(document.getElementById('globalClaimMinutes')?.value||'5',10)));
    GLOBAL_SETTINGS_CACHE={...(GLOBAL_SETTINGS_CACHE||{}),companyName,tenantId,sharePointHost,sitePath,authMode,releaseDays:release,aiEnabled,aiKey,aiModel,aiThreshold,aiProxy,liveRefreshSeconds,claimMinutes};
    try{
      localStorage.setItem('adt7j_company_name_v1',companyName); localStorage.setItem('adt7j_tenant_id_v1',tenantId);
    }catch(storageErr){
      if(typeof notifyGlobalSettingsSaved==='function') notifyGlobalSettingsSaved('Could not save settings in this browser: '+(storageErr.message||storageErr),'error');
      return;
    } localStorage.setItem('adt7j_sp_host_v1',sharePointHost); localStorage.setItem('adt7j_sp_path_v1',sitePath); localStorage.setItem('adt7j_auth_mode_v1',authMode); M365_CONFIG.tenantId=tenantId; M365_CONFIG.sharePointHost=sharePointHost; M365_CONFIG.sitePath=sitePath; _siteId=null; _listIdCache={}; applyCompanyName(companyName); localStorage.setItem('adt7j_broker_release_days_v1',String(release)); localStorage.setItem('adt7j_ai_enabled_v1',String(aiEnabled)); localStorage.setItem('adt7j_ai_key_v1',aiKey); localStorage.setItem('adt7j_ai_model_v1',aiModel); localStorage.setItem('adt7j_ai_threshold_v1',String(aiThreshold)); localStorage.setItem('adt7j_ai_proxy_v1',aiProxy); localStorage.setItem('adt7j_live_refresh_seconds_v1',String(liveRefreshSeconds)); localStorage.setItem('adt7j_claim_minutes_v1',String(claimMinutes));
    if(isDemoMode() || !m365Configured()){ if(typeof syncAIAdminToggle==='function')syncAIAdminToggle(); updateAIBrokerSearchVisibility(); if(typeof notifyGlobalSettingsSaved==='function')notifyGlobalSettingsSaved('Global settings saved locally for Demo Mode.','ok'); renderAdmin(); return; }
    try{
        const sid=await getSiteId(); const listId=await resolveListId(getListName('globalSettings'));
        const data=await graphGet(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items?expand=fields&$top=200`); const existing={};
        (data.value||[]).forEach(it=>{const f=it.fields||{};const key=f.SettingKey||f.Title;if(key)existing[key]=it.id;});
        const settings={companyName:String(companyName),tenantId:String(tenantId),sharePointHost:String(sharePointHost),sitePath:String(sitePath),authMode:String(authMode),releaseDays:String(release),aiEnabled:String(aiEnabled),aiKey,aiModel,aiThreshold:String(aiThreshold),aiProxy,liveRefreshSeconds:String(liveRefreshSeconds),claimMinutes:String(claimMinutes),listSources:JSON.stringify(getListOverrides())};
        for(const [key,value] of Object.entries(settings)){
            const fields={Title:key,SettingKey:key,SettingValue:String(value),Updated:new Date().toISOString(),UpdatedBy:currentUser?.Email||''};
            if(existing[key]) await graphPatch(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items/${existing[key]}/fields`,fields);
            else await graphPost(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items`,{fields});
        }
        if(typeof syncAIAdminToggle==='function')syncAIAdminToggle(); updateAIBrokerSearchVisibility(); if(typeof notifyGlobalSettingsSaved==='function')notifyGlobalSettingsSaved('Global settings saved successfully to Microsoft 365.','ok'); renderAdmin();
    }catch(e){if(typeof notifyGlobalSettingsSaved==='function')notifyGlobalSettingsSaved('Global settings could not be saved: '+(e.message||e),'error');}
}
function isAIBrokerSearchEnabled(){
    let raw=null;
    try{ raw=localStorage.getItem('adt7j_ai_enabled_v1'); }catch(e){}
    // If the setting has not been loaded yet, fall back to the application's
    // current global cache/default. Once loaded, localStorage is synchronised
    // from the shared Global Portal Settings list.
    if(raw===null){
        try{ raw=getGlobalSettingValue('aiEnabled',true); }catch(e){ raw=true; }
    }
    return raw===true || String(raw).trim().toLowerCase()==='true' || String(raw).trim()==='1';
}

function updateAIBrokerSearchVisibility(){
    const p=document.getElementById('aiBrokerSearchPanel');
    if(!p) return;
    const adminToggle=document.getElementById('globalAIEnabled');
    const enabled=adminToggle ? adminToggle.checked===true : isAIBrokerSearchEnabled();

    p.classList.toggle('ai-admin-hidden',!enabled);
    p.setAttribute('aria-hidden',enabled?'false':'true');
    p.hidden=!enabled;
    p.style.setProperty('display',enabled?'block':'none','important');
}

window.addEventListener('storage',function(e){
    if(e.key==='adt7j_ai_enabled_v1'){
        GLOBAL_SETTINGS_CACHE=GLOBAL_SETTINGS_CACHE||{};
        delete GLOBAL_SETTINGS_CACHE.aiEnabled;
        updateAIBrokerSearchVisibility();
    }
});

function getStoredBrokers() {
    const raw = localStorage.getItem(STORAGE_KEY);
    let items = DEFAULT_BROKERS;
    if (raw) {
        try { items = JSON.parse(raw); } catch(e) { items = DEFAULT_BROKERS; }
    }
    
    // Normalize legacy single string LoanType to LoanTypes array
    items.forEach(item => {
        if (!item.LoanTypes && item.LoanType) {
            item.LoanTypes = [item.LoanType];
        } else if (!item.LoanTypes) {
            item.LoanTypes = ["Residential Bridging"];
        }
    });

    // Check configurable inactivity reversion.
    const releaseDays = getBrokerReleaseDays();
    const RELEASE_MS = releaseDays * 24 * 60 * 60 * 1000;
    const now = Date.now();
    let updated = false;

    items.forEach(item => {
        if (!item.IsNotSuitable && item.AssignedTo) {
            const lastRef = new Date(item.LastContactDate || item.Modified).getTime();
            if (now - lastRef > RELEASE_MS) {
                // Revert back to open brokers
                item.AssignedTo = null;
                item.Modified = new Date().toISOString();
                updated = true;
                logAuditRecord(item, "AUTO_REVERT", `Reverted to open brokers due to ${releaseDays} days of inactivity without update.`);
            }
        }
    });

    if (updated) {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(items));
    }
    return items;
}

// Broker cloud writes must be serialised. Bulk imports previously launched one
// fire-and-forget sync per row, so several overlapping syncs could all see the
// same broker as "new" and POST it to Microsoft Lists more than once.
let _brokerSyncQueue = Promise.resolve();
function queueBrokerCloudSync(items){
    // Capture this save state now so later local mutations do not change a queued job.
    const snapshot = JSON.parse(JSON.stringify(items || []));
    _brokerSyncQueue = _brokerSyncQueue
        .catch(() => {}) // a failed earlier sync must not permanently block later saves
        .then(() => cloudSyncBrokers(snapshot));
    _brokerSyncQueue.catch(e => showCloudError(e));
    return _brokerSyncQueue;
}

function saveStoredBrokers(items, sync=true) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(items));
    if(sync) queueBrokerCloudSync(items);
}

// Snapshot used to detect which broker records changed since last sync
let _brokerSnapshot = {};
let _brokerSyncInFlight = false;
function _snapshotBrokers(items){ const s={}; items.forEach(b=> s[b.Id] = JSON.stringify(b)); return s; }
async function cloudSyncBrokers(items){
    if(!m365Configured()) return;
    _brokerSyncInFlight=true;
    try{
        const snap=_snapshotBrokers(items), prev=_brokerSnapshot;
        for(const id in prev){ if(!(id in snap)) await cloudDeleteBroker(JSON.parse(prev[id])); }
        for(const id in snap){
            if(!(id in prev) || prev[id]!==snap[id]){
                await cloudPushBroker(JSON.parse(snap[id]));
                // Record each successful item immediately. If a later item fails,
                // a subsequent sync will PATCH the successful items rather than
                // POSTing them again. This makes partial imports idempotent.
                _brokerSnapshot[id]=snap[id];
            }
        }
        _brokerSnapshot=snap;
    } finally { _brokerSyncInFlight=false; }
}

/* ----------------------- Deals: storage + cloud sync -------------------- */
const DEALS_STORAGE_KEY = "adt7j_deals_db_v1";
const DEAL_IDMAP_KEY    = "adt7j_deals_idmap_v1";
function getDealIdMap(){ try{ return JSON.parse(localStorage.getItem(DEAL_IDMAP_KEY)) || {}; }catch(e){ return {}; } }
function setDealIdMap(m){ localStorage.setItem(DEAL_IDMAP_KEY, JSON.stringify(m)); }

function getStoredDeals() {
    const raw = localStorage.getItem(DEALS_STORAGE_KEY);
    if(!raw) return [];
    try{ return JSON.parse(raw); }catch(e){ return []; }
}
function saveStoredDeals(items){
    localStorage.setItem(DEALS_STORAGE_KEY, JSON.stringify(items));
    cloudSyncDeals(items).catch(e => showCloudError(e));
}
let _dealSnapshot = {};
function _snapshotDeals(items){ const s={}; items.forEach(d=> s[d.Id] = JSON.stringify(d)); return s; }
async function cloudSyncDeals(items){
    if(!m365Configured()) return;
    const snap = _snapshotDeals(items);
    const prev = _dealSnapshot;
    for(const id in prev){ if(!(id in snap)){ await cloudDeleteDeal(JSON.parse(prev[id])); } }
    for(const id in snap){ if(!(id in prev) || prev[id] !== snap[id]){ await cloudPushDeal(JSON.parse(snap[id])); } }
    _dealSnapshot = snap;
}

// Deal <-> SharePoint list fields
function dealToFields(d){
    return {
        Title:        d.Company || d.BrokerName || "(no broker)",
        Company:      d.Company || "",
        BrokerName:   d.BrokerName || "",
        BrokerId:     String(d.BrokerId || ""),
        DealValue:     Number(d.DealValue) || 0,
        DealDate:      d.DealDate || new Date().toISOString(),
        LoanType:      d.LoanType || "",
        BDM:           d.BDM || "",
        Notes:         d.Notes || "",
        PortalId:      String(d.Id)
    };
}
function fieldsToDeal(it, fields){
    const f = fields || {};
    return {
        Id: Number(f.PortalId) || Number(it.id),
        BrokerId: f.BrokerId || "",
        BrokerName: f.BrokerName || f.Title || "",
        Company: f.Company || f.Title || "",
        DealValue: Number(f.DealValue) || 0,
        DealDate: f.DealDate || it.createdDateTime || "",
        LoanType: f.LoanType || "",
        BDM: f.BDM || "",
        Notes: f.Notes || ""
    };
}
async function cloudLoadDeals(){
    const sid = await getSiteId();
    const listId = await resolveListId(getListName("deals"));
    let url = `${GRAPH_BASE}/sites/${sid}/lists/${listId}/items?expand=fields&$top=500`;
    const deals = []; const idmap = {};
    while(url){
        const data = await graphGet(url);
        (data.value||[]).forEach(it => {
            const f = it.fields || {};
            const d = fieldsToDeal(it, f);
            if(!d.Id) d.Id = Date.now() + Math.floor(Math.random()*1000);
            deals.push(d);
            if(f.PortalId) idmap[f.PortalId] = it.id;
        });
        url = data["@odata.nextLink"] || null;
    }
    // Sort newest first (client-side — avoids relying on the DealDate column being indexed)
    deals.sort((a,b) => new Date(b.DealDate || 0) - new Date(a.DealDate || 0));
    setDealIdMap(idmap);
    return deals;
}
async function cloudPushDeal(d){
    const sid = await getSiteId();
    const listId = await resolveListId(getListName("deals"));
    const map = getDealIdMap();
    const fields = dealToFields(d);
    const spId = map[d.Id];
    if(spId){
        await graphPatch(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items/${spId}/fields`, fields);
    } else {
        const created = await graphPost(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items`, { fields });
        if(created && created.id){ map[d.Id] = created.id; setDealIdMap(map); }
    }
}
async function cloudDeleteDeal(d){
    const sid = await getSiteId();
    const listId = await resolveListId(getListName("deals"));
    const map = getDealIdMap();
    const spId = map[d.Id];
    if(spId){
        await graphDelete(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items/${spId}`);
        delete map[d.Id]; setDealIdMap(map);
    }
}

function getStoredAuditLogs() {
    const raw = localStorage.getItem(AUDIT_STORAGE_KEY);
    if(!raw) return [];
    try { return JSON.parse(raw); } catch(e) { return []; }
}

function logAuditRecord(item, actionType, reason = "Updated record") {
    let logs = getStoredAuditLogs();
    const entry = {
        Timestamp: new Date().toISOString(),
        User: currentUser ? currentUser.Title + " (" + currentUser.Email + ")" : "System",
        Action: actionType,
        RecordId: item.Id,
        RecordTitle: item.Title || "(No name)",
        Company: item.Company || "(No company)",
        Reason: reason
    };
    logs.unshift(entry);
    localStorage.setItem(AUDIT_STORAGE_KEY, JSON.stringify(logs));
    // Fire-and-forget push to the BDM Audit Log list
    cloudPushAudit(entry).catch(e => showCloudError(e));
}

let currentUser = null;
let allItems = [];
let calendarCursor = new Date();
let calendarEvents = [];
let calendarEditingId = null;
let calendarCallbackContext = null;
let calendarViewMode = 'month';
const BROKER_RELEASE_DAYS_KEY = 'adt7j_broker_release_days_v1';
const CALENDAR_STORAGE_KEY = "adt7j_calendar_events_v1";
const DEMO_MODE_KEY = "adt7j_demo_mode_v1";
let currentView = "open";
let selectedItemId = null;
let parsedRows = [];
let columnMapping = {};

// --- AUTHENTICATION & LOGIN LOGIC (Microsoft 365 / Entra ID via MSAL.js) ---
function setLoginError(msg){
    const el = document.getElementById("loginError");
    if(!el) return;
    el.textContent = msg;
    el.style.display = "block";
}
function showLoginScreen(){
    const ov = document.getElementById("loginOverlay");
    if(ov) ov.style.display = "flex";
}
function hideLoginScreen(){
    const ov = document.getElementById("loginOverlay");
    if(ov) ov.style.display = "none";
}

// Called when the page loads — decides whether to sign in, finish a redirect, or run locally.
async function checkSession() {
    // When opened locally, enter the existing local demo environment automatically.
    // This makes both file:// and localhost builds self-contained and prevents the
    // production Microsoft 365 credentials embedded in the deployment build from
    // ever being used by a local browser session.
    if(isLocalRuntime() && !isDemoMode()){
        enterLocalDemoMode('admin');
        return;
    }
    if(isDemoMode()){
        const rawSession=sessionStorage.getItem(SESSION_KEY);
        if(rawSession){ try{currentUser=JSON.parse(rawSession);hideLoginScreen();setUserbox();initApp();return;}catch(e){} }
    }
    // Not configured yet: run in local-only mode so the file still works for testing.
    if(!m365Configured()){
        const warn = document.getElementById("configWarning");
        if(warn){ warn.style.display = "block"; warn.innerHTML = "<strong>Local mode:</strong> Microsoft 365 sign-in is not configured yet. Fill in the M365_CONFIG block at the top of the file (tenantId, clientId, SharePoint site) and create the three Microsoft Lists. See the setup guide."; }
        const demoLink = document.getElementById("localDemoLink");
        if(demoLink) demoLink.style.display = "block";
        // Restore last local session if present
        const rawSession = sessionStorage.getItem(SESSION_KEY);
        if(rawSession){ try{ currentUser = JSON.parse(rawSession); hideLoginScreen(); setUserbox(); initApp(); return; }catch(e){} }
        showLoginScreen();
        return;
    }

    try{
        await ensureMsal();
        // Handle the return from a Microsoft login redirect
        const resp = await _msalInstance.handleRedirectPromise();
        if(resp && resp.account){ _msalInstance.setActiveAccount(resp.account); }
    }catch(e){
        setLoginError("Sign-in error: " + (e.message || e));
        showLoginScreen();
        return;
    }

    const accounts = _msalInstance.getAllAccounts();
    if(accounts.length === 0){
        showLoginScreen();
        return;
    }
    _msalInstance.setActiveAccount(accounts[0]);
    await proceedAfterLogin();
}

// Begin the Microsoft 365 sign-in (redirect flow)
function openLoginOrganisationSetup(){
    const c=localStorage.getItem('adt7j_company_name_v1')||'7J Finance', t=localStorage.getItem('adt7j_tenant_id_v1')||M365_CONFIG.tenantId, h=localStorage.getItem('adt7j_sp_host_v1')||M365_CONFIG.sharePointHost, pth=localStorage.getItem('adt7j_sp_path_v1')||M365_CONFIG.sitePath, m=localStorage.getItem('adt7j_auth_mode_v1')||'tenant';
    [['loginCompanyName',c],['loginTenantId',t],['loginSharePointHost',h],['loginSitePath',pth]].forEach(([id,v])=>{const el=document.getElementById(id);if(el)el.value=v;}); const mt=document.getElementById('loginMultitenant'); if(mt)mt.checked=m==='multitenant'; const cid=document.getElementById('loginClientIdReadOnly'); if(cid)cid.textContent=M365_CONFIG.clientId; const modal=document.getElementById('loginOrganisationSetupModal'); if(modal)modal.classList.add('show');
}
function closeLoginOrganisationSetup(){const modal=document.getElementById('loginOrganisationSetupModal');if(modal)modal.classList.remove('show');}
function saveLoginOrganisationSetup(){
    const company=document.getElementById('loginCompanyName')?.value.trim()||'7J Finance', tenant=document.getElementById('loginTenantId')?.value.trim()||M365_CONFIG.tenantId, host=document.getElementById('loginSharePointHost')?.value.trim()||M365_CONFIG.sharePointHost, path=document.getElementById('loginSitePath')?.value.trim()||M365_CONFIG.sitePath, mode=document.getElementById('loginMultitenant')?.checked?'multitenant':'tenant';
    localStorage.setItem('adt7j_company_name_v1',company);localStorage.setItem('adt7j_tenant_id_v1',tenant);localStorage.setItem('adt7j_sp_host_v1',host);localStorage.setItem('adt7j_sp_path_v1',path);localStorage.setItem('adt7j_auth_mode_v1',mode);M365_CONFIG.tenantId=tenant;M365_CONFIG.sharePointHost=host;M365_CONFIG.sitePath=path;_siteId=null;_listIdCache={};applyCompanyName(company); const st=document.getElementById('loginOrganisationSetupStatus');if(st)st.textContent='Saved. If the tenant/sign-in mode changed, close this window and sign in again.'; setTimeout(closeLoginOrganisationSetup,900);
}

async function handleLogin(e) {
    if(e) e.preventDefault();
    if(!m365Configured()){ setLoginError("M365 not configured — see setup guide."); return; }
    try{
        await ensureMsal();
        await _msalInstance.loginRedirect({ scopes: M365_CONFIG.scopes });
    }catch(err){
        setLoginError("Could not start sign-in: " + (err.message || err));
    }
}

// Sign out of Microsoft 365
function handleLogout() {
    try {
        // Always clear the local/demo session first.
        sessionStorage.removeItem(SESSION_KEY);
        sessionStorage.removeItem(DEMO_MODE_KEY);
        localStorage.removeItem("7J_DEMO_CURRENT_USER");
        currentUser = null;
        window.dialerQueue = [];
        window.currentView = "home";

        // Hide the application and return to the front/login screen immediately.
        document.querySelectorAll(".app-view").forEach(v => {
            v.classList.remove("active-view");
            v.style.display = "none";
        });
        const header = document.querySelector("header");
        if(header) header.style.display = "none";
        const overlay = document.getElementById("loginOverlay");
        if(overlay) overlay.style.display = "flex";

        const badge = document.getElementById("roleBadge");
        if(badge) badge.style.display = "none";
        const ub = document.getElementById("userbox");
        if(ub) ub.textContent = "Microsoft 365";

        // In real Microsoft 365 mode, finish the server-side/Entra session too.
        if(m365Configured() && _msalInstance){
            try {
                const acc = _msalInstance.getActiveAccount();
                if(acc) {
                    _msalInstance.logoutRedirect({ account: acc });
                    return;
                }
            } catch(e) {
                console.warn("Microsoft logout redirect unavailable:", e);
            }
        }
    } catch(e) {
        console.error("Sign out failed:", e);
        window.location.reload();
    }
}

// After a successful sign-in: load the user profile + role + all data from SharePoint Lists
async function proceedAfterLogin() {
    try{
        // 1) Who is signed in?
        const me = await graphGet(`${GRAPH_BASE}/me`);
        const email = (me.userPrincipalName || me.mail || "").toLowerCase();

        // 2) Load the BDM Users roster (for roles + assign dropdown). Create the signed-in user if missing.
        let users = [];
        try{ users = await cloudLoadUsers(); }catch(e){ console.warn("Users list not loaded:", e.message); }
        let meUser = users.find(u => (u.Email || "").toLowerCase() === email);
        currentUser = {
            Id: meUser ? meUser.Id : 0,
            Title: me.displayName || "",
            Email: email,
            // Admin status comes from the M365_CONFIG admins list (by email),
            // not from a role field anyone could edit in SharePoint.
            Role: isAdmin(email) ? "Admin" : "BDM"
        };
        if(meUser && meUser.Role !== currentUser.Role){
            // Keep the SharePoint list in sync with the configured role
            try{ await cloudPushUser(currentUser); }catch(e){ console.warn("Could not sync user role:", e.message); }
        }
        if(!meUser){
            // First time this person signs in — add them to the BDM Users list as a BDM
            try{ await cloudPushUser(currentUser); }catch(e){ console.warn("Could not add user to list:", e.message); }
        }
        // Cache the roster locally so the assign/admin dropdowns work
        localStorage.setItem(USERS_STORAGE_KEY, JSON.stringify(users.length ? users : [currentUser]));
        _userSnapshot = _snapshotUsers(users.length ? users : [currentUser]);

        // 3) Load the global portal settings BEFORE any broker/list data.
        // This ensures every BDM uses the same centrally configured SharePoint Lists.
        try{ await loadGlobalPortalSettings(false); }catch(e){ console.warn('Global portal settings not loaded before data refresh:',e.message); }

        // 4) Load brokers + audit from SharePoint Lists into the local cache
        try{
            const brokers = await cloudLoadBrokers();
            localStorage.setItem(STORAGE_KEY, JSON.stringify(brokers));
            _brokerSnapshot = _snapshotBrokers(brokers);
        }catch(e){ showCloudError(e); }
        try{
            const logs = await cloudLoadAudit();
            localStorage.setItem(AUDIT_STORAGE_KEY, JSON.stringify(logs));
            const hydrated=hydrateBrokerOutcomeMetadata(getStoredBrokers(),logs);
            localStorage.setItem(STORAGE_KEY,JSON.stringify(hydrated)); allItems=hydrated; _brokerSnapshot=_snapshotBrokers(hydrated);
        }catch(e){
            // Keep the existing account audit data usable if the new dialler list has not yet been created.
            try{
                const legacy = await cloudLoadAuditFromList("auditAccount");
                localStorage.setItem(AUDIT_STORAGE_KEY, JSON.stringify(legacy));
            }catch(inner){ showCloudError(inner); }
        }
        try{
            const deals = await cloudLoadDeals();
            localStorage.setItem(DEALS_STORAGE_KEY, JSON.stringify(deals));
            _dealSnapshot = _snapshotDeals(deals);
        }catch(e){ showCloudError(e); }

        hideLoginScreen();
        setUserbox();
        initApp();
    }catch(e){
        setLoginError("Could not load your data: " + (e.message || e));
        showLoginScreen();
    }
}

// Fallback so the file stays usable before Microsoft 365 is configured.
function seedDemoDiallerActivity(){
    const existing=getStoredAuditLogs();
    if(existing.some(l=>l.Trail==="auditDialer" && /^CALL_/.test(l.Action))) return;
    const brokers=getStoredBrokers();
    const outcomes=["NO_ANSWER","CONNECTED","CALL_BACK","APPOINTMENT","NO_PRODUCT","NOT_INTERESTED","CONNECTED","APPOINTMENT","OUT_OF_AREA"];
    const now=Date.now();
    const demo=outcomes.map((o,i)=>{ const b=brokers[i % Math.max(1,brokers.length)] || {}; const d=new Date(now-i*86400000*2); return {Timestamp:d.toISOString(),User:"Tim (Demo BDM) (demo@local)",Action:"CALL_"+o,RecordId:b.Id||i+1,RecordTitle:b.Title||"Demo Broker "+(i+1),Company:b.Company||"Demo Company",Reason:"Demo dialler outcome",Trail:"auditDialer"}; });
    localStorage.setItem(AUDIT_STORAGE_KEY,JSON.stringify(demo.concat(existing)));
}
function ensureDemoBrokerViews(){
    if(!isDemoMode() || sessionStorage.getItem("7jDemoViewsSeeded")==="1") return;
    let brokers=getStoredBrokers();
    if(!brokers.length) return;
    let deleted=[];
    try{ deleted=JSON.parse(localStorage.getItem("7J_DEMO_DELETED_BROKERS") || "[]"); }catch(e){ deleted=[]; }
    // Deliberately deleted demo records must never be resurrected by seeding.
    const deletedKeys = JSON.parse(localStorage.getItem("7J_DEMO_DELETED_KEYS") || "[]");
    brokers=brokers.filter(b=>{
        const idMatch=deleted.includes(String(b.Id)) || deleted.includes(String(b.PortalId||""));
        const key=String(b.Company||"").trim().toLowerCase()+"|"+String(b.Email||"").trim().toLowerCase();
        return !idMatch && !deletedKeys.includes(key);
    });
    if(!brokers.length){
        localStorage.setItem(STORAGE_KEY,JSON.stringify([]));
        allItems=[];
        sessionStorage.setItem("7jDemoViewsSeeded","1");
        return;
    }
    brokers[0].AssignedTo={Title:"Tim (Demo BDM)",EMail:"demo@local"};
    brokers[0].IsNotSuitable=false;
    if(brokers.length>1){
        const ns=brokers[1]; ns.AssignedTo=null; ns.IsNotSuitable=true; ns.NotSuitableReason="Not Interested"; ns.NotSuitableSource="Dialler"; ns.DiallerOutcome="NOT_INTERESTED"; ns.DiallerOutcomeDate=ns.Modified||new Date().toISOString();
    }
    if(brokers.length>2){ brokers[2].AssignedTo=null; brokers[2].IsNotSuitable=false; }
    localStorage.setItem(STORAGE_KEY,JSON.stringify(brokers));
    allItems=brokers;
    sessionStorage.setItem("7jDemoViewsSeeded","1");
}

function enterLocalDemoMode(profile="admin"){
    // Explicit demo mode never calls Microsoft Graph. It uses localStorage/sessionStorage only.
    sessionStorage.setItem(DEMO_MODE_KEY, "1");
    const profiles={
        admin:{Id:1,Title:"Tim (Admin Demo)",Email:"admin.demo@local",Role:"Admin"},
        bdm1:{Id:2,Title:"Sarah (Demo BDM)",Email:"sarah.demo@local",Role:"BDM"},
        bdm2:{Id:3,Title:"James (Demo BDM)",Email:"james.demo@local",Role:"BDM"}
    };
    const chosen=profiles[profile]||profiles.admin;
    currentUser=Object.assign({},chosen);
    // Keep all demo identities in the local roster so Admin reassign/filter
    // controls behave exactly like the live user roster.
    const existingUsers=getStoredUsers();
    const demoUsers=Object.values(profiles);
    const mergedUsers=[...existingUsers];
    demoUsers.forEach(du=>{
        const ix=mergedUsers.findIndex(u=>String(u.Email||'').toLowerCase()===String(du.Email).toLowerCase());
        if(ix<0) mergedUsers.push(du);
        else mergedUsers[ix]=Object.assign({},mergedUsers[ix],du);
    });
    localStorage.setItem(USERS_STORAGE_KEY,JSON.stringify(mergedUsers));
    sessionStorage.setItem(SESSION_KEY,JSON.stringify(currentUser));

    seedDemoCalendarEvents();
    seedDemoDiallerActivity();
    ensureDemoBrokerViews();

    // Give each demo BDM a distinct view of the assigned sample brokers.
    let demoBrokers=getStoredBrokers();
    if(demoBrokers.length){
        demoBrokers.forEach(b=>{
            if(String(b.Company||"").toLowerCase().includes("apex")){
                b.AssignedTo={Title:"Sarah (Demo BDM)",EMail:"sarah.demo@local"};
                b.IsNotSuitable=false;
            } else if(String(b.Company||"").toLowerCase().includes("meridian")){
                b.AssignedTo={Title:"James (Demo BDM)",EMail:"james.demo@local"};
                b.IsNotSuitable=false;
            }
        });
        saveStoredBrokers(demoBrokers);
        allItems=demoBrokers;
        _brokerSnapshot=_snapshotBrokers(demoBrokers);
    }

    hideLoginScreen();
    setUserbox();
    const adminMenuBtn=document.getElementById("hamburgerAdminBtn");
    if(adminMenuBtn) adminMenuBtn.style.display=(currentUser.Role==="Admin")?"":"none";
    const w=document.getElementById("configWarning");
    if(w){w.style.display="block";w.innerHTML="<strong>Demo mode:</strong> "+escapeHtml(chosen.Title)+" is signed in. Microsoft 365 is disconnected; brokers, deals and calendar events are stored only in this browser.";}
    initApp();
}
function isDemoMode(){ return sessionStorage.getItem(DEMO_MODE_KEY)==="1"; }

function setUserbox(){
    document.body.classList.toggle("bdm-kpi-user", !!currentUser && currentUser.Role !== "Admin");
    const el = document.getElementById("userbox");
    if(el && currentUser){
        el.textContent = currentUser.Title + " (" + currentUser.Email + ")";
    }
    const badge = document.getElementById("roleBadge");
    if(badge && currentUser){
        if(currentUser.Role === "Admin"){
            badge.textContent = "Admin";
            badge.style.display = "inline-block";
        } else {
            badge.style.display = "none";
        }
    }
    // Admin Centre is visible in the hamburger only to administrators.
    const adminNav = document.getElementById("hamburgerAdminBtn");
    if(adminNav && currentUser){
        adminNav.style.display = (currentUser.Role === "Admin") ? "" : "none";
    }
}

// Passwords are no longer used — identity comes from Microsoft 365.
function openUserPasswordModal() {
    alert("Account access is managed by Microsoft 365. To change your password, use your normal Microsoft sign-in (account.microsoft.com). Roles for this portal are set by an admin in the User Management area.");
}

// Admin area is unlocked by role, not by a password.
function unlockAdminCentre() {
    if(!currentUser || currentUser.Role !== "Admin"){
        const err = document.getElementById("adminUnlockError");
        if(err){ err.textContent = "Only administrators can access this area."; err.style.display = "block"; }
        return;
    }
    document.getElementById("adminPasswordGate").style.display = "none";
    document.getElementById("adminUnlockedContent").style.display = "block";
    document.getElementById("adminUnlockError").style.display = "none";
    populateAdminUserFilterOptions();
    populateFilterOptions();
    populateAuditFilterOptions();
    renderUserManagementList();
    renderAuditLog();
    renderDialerAuditLog();
    renderAdmin();
    reloadCallGuideFromSource(false);
    prefillListSources();
    loadOrganisationSettings();
    loadGlobalPortalSettings(false);
}

/* ---- SharePoint list source override UI ---- */
function prefillListSources(){
    const ov = getListOverrides();
    const ids = [["listNameBrokers","brokers"],["listNameUsers","users"],["listNameAuditAccount","auditAccount"],["listNameAuditDialer","auditDialer"],["listNameDeals","deals"],["listNameGlobalSettings","globalSettings"],["listNameCallGuide","callGuide"],["listNamePerformanceReviews","performanceReviews"],["listNameKpiSnapshots","kpiSnapshots"],["listNameBackupManifest","backupManifest"]];
    const globalSources=(GLOBAL_SETTINGS_CACHE&&GLOBAL_SETTINGS_CACHE.listSources)?(typeof GLOBAL_SETTINGS_CACHE.listSources==='string'?(()=>{try{return JSON.parse(GLOBAL_SETTINGS_CACHE.listSources)}catch(e){return {}}})():GLOBAL_SETTINGS_CACHE.listSources):{}; ids.forEach(([id,key])=>{ const el=document.getElementById(id); if(el) el.value=globalSources[key] || ov[key] || M365_CONFIG.lists[key]; });
}
async function testListSources(){
    const vals = {brokers: document.getElementById("listNameBrokers").value.trim() || M365_CONFIG.lists.brokers,
                  users: document.getElementById("listNameUsers").value.trim() || M365_CONFIG.lists.users,
                  auditAccount: document.getElementById("listNameAuditAccount").value.trim() || M365_CONFIG.lists.auditAccount,
                  auditDialer: document.getElementById("listNameAuditDialer").value.trim() || M365_CONFIG.lists.auditDialer,
                  deals: document.getElementById("listNameDeals").value.trim() || M365_CONFIG.lists.deals,
                  globalSettings: document.getElementById("listNameGlobalSettings").value.trim() || M365_CONFIG.lists.globalSettings,
                  callGuide: document.getElementById("listNameCallGuide").value.trim() || M365_CONFIG.lists.callGuide,
                  performanceReviews: document.getElementById("listNamePerformanceReviews").value.trim() || M365_CONFIG.lists.performanceReviews,
                  kpiSnapshots: document.getElementById("listNameKpiSnapshots").value.trim() || M365_CONFIG.lists.kpiSnapshots,
                  backupManifest: document.getElementById("listNameBackupManifest").value.trim() || M365_CONFIG.lists.backupManifest};
    const el = document.getElementById("listSourceStatus");
    if(el){ el.style.display="block"; el.style.background="#fff4e5"; el.style.color="#8a5a00"; el.textContent = "Testing lists and columns…"; }
    const keys = Object.keys(vals);
    const results = await Promise.all(keys.map(k=>testListConnectionWithColumns(vals[k], k)));
    let allGood = true;
    const lines = keys.map((k,i)=>{
        const r = results[i]; const label = vals[k];
        if(!r.ok){ allGood = false; return "❌ " + label + ": not found"; }
        if(r.columnCheckError) return "✅ " + label + ": found (could not check columns — " + r.columnCheckError + ")";
        const missingReq = r.missingRequired || [];
        const missingOpt = r.missingOptional || [];
        if(!missingReq.length && !missingOpt.length) return "✅ " + label + ": found — all columns present";
        if(missingReq.length) allGood = false;
        let msg = (missingReq.length ? "⚠️ " : "✅ ") + label + ": found";
        if(missingReq.length) msg += " — missing columns: " + missingReq.join(", ");
        if(missingOpt.length) msg += " — missing optional: " + missingOpt.join(", ");
        return msg;
    });
    if(el){
        el.innerHTML = lines.map(l => escapeHtml(l)).join("<br/>");
        el.style.background = allGood ? "#e8f5e9" : "#fff4e5";
        el.style.color = allGood ? "#1b5e20" : "#8a5a00";
    }
}
async function saveListSources(){
    const fields = {brokers:'listNameBrokers', users:'listNameUsers', auditAccount:'listNameAuditAccount', auditDialer:'listNameAuditDialer', deals:'listNameDeals', globalSettings:'listNameGlobalSettings', callGuide:'listNameCallGuide', performanceReviews:'listNamePerformanceReviews', kpiSnapshots:'listNameKpiSnapshots', backupManifest:'listNameBackupManifest'};
    const ov = {};
    Object.entries(fields).forEach(([key,id])=>{ const v=document.getElementById(id)?.value.trim(); if(v) ov[key]=v; });
    // Keep a local cache as a fallback, but the authoritative copy is written to
    // the Global Portal Settings Microsoft List so every BDM receives the same routing.
    saveListOverrides(ov);
    clearListCaches();
    GLOBAL_SETTINGS_CACHE={...(GLOBAL_SETTINGS_CACHE||{}),listSources:JSON.stringify(ov)};
    if(isDemoMode() || !m365Configured()){
        const st=document.getElementById('listSourceStatus'); if(st) st.textContent='Saved globally for Microsoft 365 mode; Demo Mode is using local storage.';
        return;
    }
    try{
        const sid=await getSiteId(); const listId=await resolveListId(getListName('globalSettings'));
        const data=await graphGet(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items?expand=fields&$top=200`);
        const existing=(data.value||[]).find(it=>{const f=it.fields||{};return (f.SettingKey||f.Title)==='listSources'});
        const fieldsBody={Title:'listSources',SettingKey:'listSources',SettingValue:JSON.stringify(ov),Updated:new Date().toISOString(),UpdatedBy:currentUser?.Title||''};
        if(existing) await graphPatch(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items/${existing.id}/fields`,fieldsBody);
        else await graphPost(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items`,{fields:fieldsBody});
        const st=document.getElementById('listSourceStatus'); if(st) st.textContent='Saved globally ✓ — all BDMs will use these SharePoint Lists after their next refresh/login.';
        await loadGlobalPortalSettings(false);
        prefillListSources();
    }catch(e){
        const st=document.getElementById('listSourceStatus'); if(st) st.textContent='Saved locally, but global Microsoft List update failed: '+(e.message||e);
    }
}


/* ---- Reassign broker modal ---- */
let _reassignTargetId = null;
function openReassignModal(itemId){
    _reassignTargetId = itemId;
    const modal = document.getElementById("reassignModal");
    const sel = document.getElementById("reassignUserSelect");
    const emptyNote = document.getElementById("reassignEmptyNote");
    const confirmBtn = document.getElementById("reassignConfirmBtn");
    // Show the broker being reassigned
    const brokers = getStoredBrokers();
    const b = brokers.find(x => String(x.Id) === String(itemId));
    const nameEl = document.getElementById("reassignBrokerName");
    if(nameEl) nameEl.textContent = b ? (b.Company || b.Title || "(no name)") : "";
    let users = getStoredUsers();
    if(isDemoMode()){
        const demoRoster=[
            {Id:1,Title:"Tim (Admin Demo)",Email:"admin.demo@local",Role:"Admin"},
            {Id:2,Title:"Sarah (Demo BDM)",Email:"sarah.demo@local",Role:"BDM"},
            {Id:3,Title:"James (Demo BDM)",Email:"james.demo@local",Role:"BDM"}
        ];
        const byEmail=new Map(users.map(u=>[String(u.Email||'').toLowerCase(),u]));
        demoRoster.forEach(u=>byEmail.set(u.Email.toLowerCase(),u));
        users=Array.from(byEmail.values());
    }
    if(!users || users.length === 0){
        sel.innerHTML = "<option value=''>No team members yet</option>";
        sel.style.display = "none";
        if(emptyNote) emptyNote.style.display = "block";
        if(confirmBtn) confirmBtn.style.display = "none";
    } else {
        sel.style.display = "block";
        if(emptyNote) emptyNote.style.display = "none";
        if(confirmBtn) confirmBtn.style.display = "inline-block";
        sel.innerHTML = users.map(u => `<option value="${escapeHtml(u.Email)}">${escapeHtml(u.Title || u.Email)} (${escapeHtml(u.Email)})</option>`).join("");
        sel.value = users[0].Email || "";
    }
    modal.style.display = "flex";
}
function closeReassignModal(){
    const modal = document.getElementById("reassignModal");
    if(modal) modal.style.display = "none";
    _reassignTargetId = null;
}
async function confirmReassign(){
    if(_reassignTargetId == null) return;
    const sel = document.getElementById("reassignUserSelect");
    const email = sel ? sel.value : "";
    if(!email) return;
    const users = getStoredUsers();
    const chosenUser = users.find(u => (u.Email||"").toLowerCase() === email.toLowerCase());
    if(!chosenUser) return;
    allItems = getStoredBrokers();
    const idx = allItems.findIndex(i => String(i.Id) === String(_reassignTargetId));
    if(idx === -1){ closeReassignModal(); reload(); return; }

    const broker = allItems[idx];
    const assignedTitle = chosenUser.Title || chosenUser.Email;
    const assignedEmail = chosenUser.Email || "";
    try{
        if(m365Configured() && !isDemoMode()){
            // Reassignment updates the existing SharePoint item directly.
            // Do not use the generic local-save path here because a missing
            // local SP ID map can otherwise turn an update into a POST.
            const current = await cloudGetCurrentBroker(_reassignTargetId);
            if(!current) throw new Error("This broker could not be located in Microsoft Lists. Refresh the broker data and try again.");
            if(isNotSuitable(current.broker)) throw new Error("This broker is currently marked Not Suitable and cannot be assigned.");
            const now = new Date().toISOString();
            const patch = {
                AssignedTo: assignedTitle + (assignedEmail ? " <" + assignedEmail + ">" : ""),
                AssignedToEmail: assignedEmail,
                AssignedToName: assignedTitle,
                IsNotSuitable: "No",
                Modified: now,
                ClaimedAt: "",
                ClaimedBy: "",
                ClaimExpiresAt: ""
            };
            const sid = await getSiteId();
            const listId = await resolveListId(getListName("brokers"));
            await graphPatch(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items/${current.spId}/fields`, filterBrokerFieldsForColumns(patch, await getBrokerColumnNames()), current.etag || undefined);
            const fresh = await cloudLoadBrokers();
            localStorage.setItem(STORAGE_KEY, JSON.stringify(fresh));
            _brokerSnapshot = _snapshotBrokers(fresh);
            allItems = fresh;
            const updated = fresh.find(i => String(i.Id) === String(_reassignTargetId)) || Object.assign({}, broker, {AssignedTo:{Title:assignedTitle,EMail:assignedEmail},IsNotSuitable:false,Modified:now});
            logAuditRecord(updated, "ADMIN_REASSIGN", `Reassigned to ${assignedTitle} by admin`);
        }else{
            broker.AssignedTo = { Title: assignedTitle, EMail: assignedEmail };
            broker.IsNotSuitable = false;
            broker.Modified = new Date().toISOString();
            saveStoredBrokers(allItems);
            logAuditRecord(broker, "ADMIN_REASSIGN", `Reassigned to ${assignedTitle} by admin`);
        }
        closeReassignModal();
        reload();
    }catch(e){
        alert(e.message || "Unable to reassign this broker.");
        closeReassignModal();
        if(m365Configured() && !isDemoMode()){
            try{ const fresh=await cloudLoadBrokers(); localStorage.setItem(STORAGE_KEY,JSON.stringify(fresh)); _brokerSnapshot=_snapshotBrokers(fresh); allItems=fresh; }catch(_){}
        }
        reload();
    }
}

/* ---- Log Deal modal ---- */
let _logDealBrokerId = null;
function openLogDealModal(){
    const brokers = getStoredBrokers();
    const b = brokers.find(x => String(x.Id) === String(selectedItemId));
    if(!b){ alert("Open a broker record first, then log a deal."); return; }
    _logDealBrokerId = b.Id;
    const nameEl = document.getElementById("logDealBroker");
    if(nameEl) nameEl.textContent = "Broker: " + (b.Company || b.Title || "(no name)");
    const lt = document.getElementById("dealLoanType");
    const loans = ["Residential Bridging","Commercial Bridging","Development Finance","Buy-to-Let Mortgages","Second Charge / Secured Loans","HMO / Multi-Unit Blocks","Expat / Foreign National Mortgages","Auction Finance","Complex Refurbishment","Other"];
    lt.innerHTML = loans.map(l=>`<option value="${l}">${l}</option>`).join("");
    const bl = getLoanTypes(b);
    if(bl && bl[0]) lt.value = bl[0];
    document.getElementById("dealValue").value = "";
    document.getElementById("dealDate").value = new Date().toISOString().slice(0,10);
    document.getElementById("dealNotes").value = "";
    document.getElementById("logDealModal").classList.add("show");
}
function closeLogDealModal(){
    document.getElementById("logDealModal").classList.remove("show");
    _logDealBrokerId = null;
}
function saveDeal(){
    if(_logDealBrokerId == null) return;
    const val = parseFloat(document.getElementById("dealValue").value);
    if(isNaN(val) || val <= 0){ alert("Enter a deal value greater than £0."); return; }
    const date = document.getElementById("dealDate").value || new Date().toISOString().slice(0,10);
    const loanType = document.getElementById("dealLoanType").value;
    const notes = document.getElementById("dealNotes").value.trim();
    const brokers = getStoredBrokers();
    const b = brokers.find(x => String(x.Id) === String(_logDealBrokerId));
    const deal = {
        Id: Date.now(),
        BrokerId: String(_logDealBrokerId),
        BrokerName: b ? (b.Title || "") : "",
        Company: b ? (b.Company || b.Title || "") : "",
        DealValue: val,
        DealDate: date,
        LoanType: loanType,
        BDM: currentUser ? currentUser.Title : "",
        Notes: notes
    };
    const deals = getStoredDeals();
    deals.unshift(deal);
    saveStoredDeals(deals);
    logAuditRecord(b || {Company: deal.Company, Title: deal.BrokerName}, "DEAL_LOGGED", `Deal logged: £${val.toLocaleString()} (${loanType}) by ${deal.BDM}`);
    closeLogDealModal();
    // Refresh every KPI/deals surface immediately. This is especially important
    // for BDMs whose KPI view is scoped to their signed-in identity.
    populateDealsFilters();
    renderDeals();
    renderKPIDashboard();
    renderKPI();
    renderDialerKPI();
    if(typeof renderDialerRanking === "function") renderDialerRanking();
}

function renderUserManagementList() {
    const users = getStoredUsers();
    const container = document.getElementById("userTableList");
    if(!users || users.length === 0){
        container.innerHTML = `<div style="padding:12px; text-align:center; font-size:13px; color:var(--muted);">No team members have signed in yet. People appear here automatically after their first Microsoft 365 sign-in.</div>`;
        return;
    }
    let html = "<table class='preview' style='margin-top:10px;'><tr><th>Name</th><th>Email</th><th>Role</th></tr>";
    users.forEach(u => {
        // Role shown is decided by the M365_CONFIG admins list (by email)
        const role = isAdmin(u.Email) ? "Admin" : (u.Role || "BDM");
        html += `<tr>
            <td>${escapeHtml(u.Title)}</td>
            <td>${escapeHtml(u.Email)}</td>
            <td>${escapeHtml(role)}</td>
        </tr>`;
    });
    html += "</table>";
    container.innerHTML = html;
}

function renderAuditLog() {
    const searchVal = document.getElementById("auditSearchBox") ? document.getElementById("auditSearchBox").value.trim().toLowerCase() : "";
    const actionFilter = document.getElementById("auditFilterAction") ? document.getElementById("auditFilterAction").value : "all";
    const userFilter   = document.getElementById("auditFilterUser")   ? document.getElementById("auditFilterUser").value   : "all";
    let logs = getStoredAuditLogs().filter(l => !l.Trail || l.Trail === "auditAccount");
    
    if(actionFilter && actionFilter !== "all"){
        logs = logs.filter(l => (l.Action||"") === actionFilter);
    }
    if(userFilter && userFilter !== "all"){
        logs = logs.filter(l => (l.User||"") === userFilter);
    }
    if(searchVal) {
        logs = logs.filter(l => {
            const hay = [l.Timestamp, l.User, l.Action, l.RecordTitle, l.Company, l.Reason].join(" ").toLowerCase();
            return hay.includes(searchVal);
        });
    }

    const container = document.getElementById("auditLogContainer");
    if(logs.length === 0) {
        container.innerHTML = `<div style="padding:15px; text-align:center; font-size:13px; color:var(--muted);">No audit logs match your filters. <a onclick="clearAuditFilters()" style="color:var(--accent);cursor:pointer;">Clear filters</a></div>`;
        return;
    }
    let html = "<table class='preview' style='margin-top:0;'><tr><th>Timestamp (Immutable)</th><th>User</th><th>Action</th><th>Record</th><th>Reason / Detail</th></tr>";
    logs.forEach(l => {
        html += `<tr>
            <td><code>${escapeHtml(l.Timestamp)}</code></td>
            <td>${escapeHtml(l.User)}</td>
            <td><strong>${escapeHtml(l.Action)}</strong></td>
            <td>${escapeHtml(l.RecordTitle)} (${escapeHtml(l.Company)})</td>
            <td>${escapeHtml(l.Reason)}</td>
        </tr>`;
    });
    html += "</table>";
    container.innerHTML = html;
}

function addNewUser() {
    const name = document.getElementById("newUserName").value.trim();
    const email = document.getElementById("newUserEmail").value.trim();
    const role = document.getElementById("newUserRole").value;
    const statusEl = document.getElementById("userManageStatus");

    if(!name || !email) {
        statusEl.textContent = "Please enter a name and email address.";
        return;
    }

    let users = getStoredUsers();
    if(users.some(u => u.Email.toLowerCase() === email.toLowerCase())) {
        statusEl.textContent = "A user with this email already exists.";
        return;
    }

    users.push({
        Id: Date.now(),
        Title: name,
        Email: email,
        Role: role || "BDM"
    });
    saveStoredUsers(users);
    document.getElementById("newUserName").value = "";
    document.getElementById("newUserEmail").value = "";
    statusEl.textContent = "User added — they can now sign in with their Microsoft 365 account.";
    renderUserManagementList();
}

// Toggle a user's role between BDM and Admin
function promptUpdateRole(id) {
    let users = getStoredUsers();
    const idx = users.findIndex(u => u.Id === id);
    if(idx === -1) return;
    const next = users[idx].Role === "Admin" ? "BDM" : "Admin";
    if(!confirm("Set " + users[idx].Title + " to " + next + " role?")) return;
    users[idx].Role = next;
    saveStoredUsers(users);
    renderUserManagementList();
}

function deleteUser(id) {
    if(!confirm("Remove this user from the portal roster? They will lose access on next sign-in.")) return;
    let users = getStoredUsers();
    users = users.filter(u => u.Id !== id);
    saveStoredUsers(users);
    renderUserManagementList();
}

// Kept for compatibility — admin credentials are now managed by Microsoft 365.
function updateAdminCredentials() {
    const statusEl = document.getElementById("adminChangeStatus");
    if(statusEl) statusEl.textContent = "Administrator accounts are managed through Microsoft 365 / Entra ID.";
}

function toggleBrokerFullscreen(e){
    if(e){e.preventDefault();e.stopPropagation();}
    const panel=document.querySelector("#detailOverlay .panel");
    if(!panel)return;
    const isFullscreen=panel.classList.toggle("fullscreen-panel");
    const b=document.getElementById("detailFullscreenBtn");
    if(b){
        b.textContent="⛶";
        b.title=isFullscreen?"Exit full screen":"Full screen";
        b.setAttribute("aria-label",isFullscreen?"Exit full screen":"Full screen");
    }
}

// --- Unsaved broker edit protection ---
let detailInitialSnapshot = null;
let detailDirty = false;
function getDetailFormSnapshot(){
    const ids=['f_name','f_company','f_phone','f_email','f_website','f_address','f_city','f_prefComm','f_volume','f_network','f_status','f_nextFollowUp'];
    const out={};
    ids.forEach(id=>{out[id]=document.getElementById(id)?.value || '';});
    out.f_loanTypes=Array.from(document.querySelectorAll('#f_loanTypesContainer input[type=checkbox]:checked')).map(x=>x.value).sort();
    out.f_logContact=Boolean(document.getElementById('f_logContact')?.checked);
    return JSON.stringify(out);
}
function markDetailClean(){detailInitialSnapshot=getDetailFormSnapshot();detailDirty=false;}
function refreshDetailDirty(){detailDirty=Boolean(detailInitialSnapshot && getDetailFormSnapshot()!==detailInitialSnapshot);return detailDirty;}
function brokerHasUnsavedChanges(){return refreshDetailDirty();}

document.addEventListener('change',function(e){if(e.target&&e.target.id==='guideAdminSectionSelect'){selectedGuideSectionKey=e.target.value;renderCallGuideAdmin();}});

function loadOrganisationSettings(){const company=localStorage.getItem('adt7j_company_name_v1')||'7J Finance',tenant=localStorage.getItem('adt7j_tenant_id_v1')||M365_CONFIG.tenantId,host=localStorage.getItem('adt7j_sp_host_v1')||M365_CONFIG.sharePointHost,path=localStorage.getItem('adt7j_sp_path_v1')||M365_CONFIG.sitePath,mode=localStorage.getItem('adt7j_auth_mode_v1')||'tenant';M365_CONFIG.tenantId=tenant;M365_CONFIG.sharePointHost=host;M365_CONFIG.sitePath=path;[['portalCompanyName',company],['portalTenantId',tenant],['portalSharePointHost',host],['portalSitePath',path]].forEach(([id,v])=>{const e=document.getElementById(id);if(e)e.value=v;});const am=document.getElementById('portalAuthMode');if(am)am.value=mode;const cid=document.getElementById('portalClientId');if(cid)cid.value=M365_CONFIG.clientId;applyCompanyName(company);}
function applyCompanyName(name){
    const n=(name||'7J Finance').trim()||'7J Finance';
    document.querySelectorAll('[data-company-name]').forEach(e=>{e.textContent=n;});
    const t=document.querySelector('title');
    if(t)t.textContent=n+' - BDM Portal & Call Guide';
    localStorage.setItem('adt7j_company_name_v1',n);
}
async function saveOrganisationSettings(){const company=document.getElementById('portalCompanyName')?.value.trim()||'7J Finance',tenant=document.getElementById('portalTenantId')?.value.trim()||M365_CONFIG.tenantId,host=document.getElementById('portalSharePointHost')?.value.trim()||M365_CONFIG.sharePointHost,path=document.getElementById('portalSitePath')?.value.trim()||M365_CONFIG.sitePath,mode=document.getElementById('portalAuthMode')?.value||'tenant';localStorage.setItem('adt7j_company_name_v1',company);localStorage.setItem('adt7j_tenant_id_v1',tenant);localStorage.setItem('adt7j_sp_host_v1',host);localStorage.setItem('adt7j_sp_path_v1',path);localStorage.setItem('adt7j_auth_mode_v1',mode);M365_CONFIG.tenantId=tenant;M365_CONFIG.sharePointHost=host;M365_CONFIG.sitePath=path;_siteId=null;_listIdCache={};applyCompanyName(company);const st=document.getElementById('organisationSettingsStatus');if(st)st.textContent='Saved locally. Publishing to Global Portal Settings…';try{await saveGlobalPortalSettings();if(st)st.textContent='Organisation settings saved. Tenant/sign-in changes apply on the next Microsoft 365 sign-in.';}catch(e){if(st)st.textContent='Saved locally, but global publish failed: '+(e.message||e);}}
function resetOrganisationSettings(){['adt7j_company_name_v1','adt7j_tenant_id_v1','adt7j_sp_host_v1','adt7j_sp_path_v1','adt7j_auth_mode_v1'].forEach(k=>localStorage.removeItem(k));M365_CONFIG.tenantId='e35c0059-b98e-4fd5-abab-a72d3307c532';M365_CONFIG.sharePointHost='7jfinance.sharepoint.com';M365_CONFIG.sitePath='/sites/BDM';loadOrganisationSettings();const st=document.getElementById('organisationSettingsStatus');if(st)st.textContent='Reset to the current application defaults.';}

// --- PORTAL NAVIGATION & APP LOGIC ---
function toggleBrokerToolsMenu(e){ e?.stopPropagation(); const m=document.getElementById('brokerToolsMenu'); if(m)m.classList.toggle('open'); }
document.addEventListener('click',()=>{const m=document.getElementById('brokerToolsMenu');if(m)m.classList.remove('open');});

document.querySelectorAll('#brokerToolsMenu [data-view]').forEach(btn=>btn.addEventListener('click',()=>{ const m=document.getElementById('brokerToolsMenu');if(m)m.classList.remove('open'); }));

function switchMasterView(viewName) {
    const allViews=['homeView','portalView','calendarView','guideView'];
    allViews.forEach(id=>{const el=document.getElementById(id);if(el){el.classList.remove('active-view');el.style.display='none';}});
    ['listView','uploadView','adminView','dealsView','dialerView','kpiDashboardView'].forEach(id=>{const el=document.getElementById(id);if(el)el.style.display='none';});
    document.querySelectorAll('.master-nav button').forEach(b=>b.classList.remove('active'));
    document.getElementById('masterMenu')?.classList.remove('open');
    document.getElementById('displaySettingsPopover')?.classList.remove('open');
    const show=(id,display='flex')=>{const el=document.getElementById(id);if(el){el.classList.add('active-view');el.style.display=display;}return el;};
    const nav=id=>document.getElementById(id)?.classList.add('active');
    if(viewName==='home'){nav('navHome');show('homeView','flex');populateHome();return;}
    if(viewName==='portal'){show('portalView','flex');currentView='open';document.getElementById('listView').style.display='block';const sub=document.querySelector('#portalView .sub-nav');if(sub)sub.style.display='flex';const ns=document.getElementById('notSuitableOutcomeFilterWrap');if(ns)ns.style.display='none';document.querySelectorAll('.sub-nav button').forEach(b=>b.classList.remove('active'));const ob=document.querySelector('.sub-nav button[data-view="open"]');if(ob)ob.classList.add('active');reload();updateAIBrokerSearchVisibility();return;}
    if(viewName==='calendar'){show('calendarView','flex');calendarCursor=new Date(calendarCursor.getFullYear(),calendarCursor.getMonth(),1);setCalendarView('month');refreshCalendar();return;}
    if(viewName==='kpi'){show('portalView','flex');const sub=document.querySelector('#portalView .sub-nav');if(sub)sub.style.display='none';currentView='kpiDashboard';document.getElementById('kpiDashboardView').style.display='block';renderKPIDashboard();return;}
    if(viewName==='guide'){show('guideView','flex');reloadCallGuideFromSource(false);return;}
}

function closeDetailPanel(e){
    if(e){e.preventDefault();e.stopPropagation();e.stopImmediatePropagation();}
    if(brokerHasUnsavedChanges()){
        const leave=window.confirm('You have unsaved changes to this broker. Close without saving?');
        if(!leave) return false;
    }
    const overlay=document.getElementById('detailOverlay');
    if(overlay) overlay.classList.remove('show');
    const panel=document.querySelector('#detailOverlay .panel');
    if(panel) panel.classList.remove('fullscreen-panel');
    const b=document.getElementById('detailFullscreenBtn');
    if(b){b.textContent='⛶';b.title='Full screen';b.setAttribute('aria-label','Full screen');}
    detailInitialSnapshot=null; detailDirty=false;
    return false;
}

function toggleDisplaySettings(e){
    if(e){e.preventDefault();e.stopPropagation();}
    const p=document.getElementById('displaySettingsPopover');if(p)p.classList.toggle('open');
}
function applyDisplaySettings(){
    const theme=document.getElementById('displayTheme')?.value||'light';
    const size=document.getElementById('displayFontSize')?.value||'medium';
    const font=document.getElementById('displayFontStyle')?.value||'system';
    document.body.classList.remove('display-dark','display-light','font-small','font-medium','font-large','font-xlarge','font-system','font-inter','font-serif','font-mono');
    document.body.classList.add(theme==='dark'?'display-dark':'display-light','font-'+size,'font-'+font);
    localStorage.setItem('7jDisplaySettings',JSON.stringify({theme,size,font}));
}
function loadDisplaySettings(){
    try{const x=JSON.parse(localStorage.getItem('7jDisplaySettings')||'null')||{};
      const theme=x.theme||'light',size=x.size||'medium',font=x.font||'system';
      ['displayTheme','displayFontSize','displayFontStyle'].forEach(id=>{const el=document.getElementById(id);if(el)el.value=id==='displayTheme'?theme:id==='displayFontSize'?size:font;});
      applyDisplaySettings();
    }catch(_){applyDisplaySettings();}
}
function resetDisplaySettings(){localStorage.removeItem('7jDisplaySettings');document.getElementById('displayTheme').value='light';document.getElementById('displayFontSize').value='medium';document.getElementById('displayFontStyle').value='system';applyDisplaySettings();}

function generateCustomPitch() {
    const name = document.getElementById('pitchBrokerName').value.trim() || "[Name]";
    const company = document.getElementById('pitchCompanyName').value.trim() || "[Company]";
    const text = `"Hi ${name}, it's Tim calling from 7J Finance. I know I'm catching you completely out of the blue over at ${company}—did I happen to catch you at an okay time for a quick chat?"`;
    document.getElementById('customPitchOutput').textContent = text;
}

function referenceDate(item) {
    const lcd = item.LastContactDate;
    return lcd ? new Date(lcd) : new Date(item.Modified);
}

function normaliseAssignedTo(value){
    if(!value) return null;
    if(Array.isArray(value)) value=value[0];
    if(typeof value === "object"){
        const email=value.EMail||value.Email||value.email||value.mail||value.userPrincipalName||"";
        const title=value.Title||value.DisplayName||value.displayName||value.Name||email||"";
        return email||title ? {Title:String(title),EMail:String(email).toLowerCase()} : null;
    }
    const text=String(value).trim();
    if(!text) return null;
    const m=text.match(/<([^>]+)>/);
    const em=text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
    const email=m?m[1].trim().toLowerCase():(em?em[0].toLowerCase():"");
    const title=m?text.replace(/<[^>]*>/,"").trim():(email?text.replace(email,"").trim():text);
    return {Title:title||email,EMail:email};
}
function assignedEmail(item){ return normaliseAssignedTo(item?.AssignedTo)?.EMail || String(item?.AssignedToEmail||item?.AssignedToEMail||"").trim().toLowerCase(); }
function assignedName(item){ return (normaliseAssignedTo(item?.AssignedTo)?.Title || String(item?.AssignedToName||"")).trim().toLowerCase(); }
function currentUserEmails(){ return [currentUser?.Email,currentUser?.EMail].filter(Boolean).map(v=>String(v).trim().toLowerCase()); }
function currentUserNames(){ return [currentUser?.Title,currentUser?.displayName].filter(Boolean).map(v=>String(v).trim().toLowerCase()); }
function isMine(item) {
    if(isNotSuitable(item) || !currentUser) return false;
    const emails=currentUserEmails(), names=currentUserNames();
    const email=assignedEmail(item), name=assignedName(item);
    if(email && emails.includes(email)) return true;
    if(!email && name && names.includes(name)) return true;
    return false;
}
function hasAssignment(item){ return !!(assignedEmail(item) || assignedName(item) || item?.AssignedToLookupId || item?.AssignedToId); }
function isOpen(item){ return !isNotSuitable(item) && !hasAssignment(item); }
function isNotSuitable(item){
    if(!item) return false;
    if(String(item.WorkflowState||"").toUpperCase()==="OPEN") return false;
    const flag=item.IsNotSuitable;
    if(flag===true || flag===1 || ["yes","true","1"].includes(String(flag).trim().toLowerCase())) return true;
    const o=String(item.DiallerOutcome||item.DiallerOutcomeReason||"").toUpperCase().replace(/\s+/g,"_");
    return ["NOT_INTERESTED","NO_PRODUCT","NO_SUITABLE_PRODUCT","OUT_OF_AREA","COMPLIANCE","COMPLIANCE_UNREGULATED","DUPLICATE","DUPLICATE_INACTIVE"].includes(o);
}
function getNotSuitableOutcomeKey(item){
    if(item?.DiallerOutcome) return 'CALL_' + item.DiallerOutcome;
    return item?.NotSuitableReason ? 'MANUAL_' + item.NotSuitableReason : 'MANUAL_OTHER';
}
function getNotSuitableOutcomeLabel(item){
    if(item?.DiallerOutcome) return outcomeLabel('CALL_' + item.DiallerOutcome);
    return item?.NotSuitableReason ? 'Manual · ' + item.NotSuitableReason : 'Manual / Other';
}
function populateNotSuitableOutcomeFilter(){
    const sel=document.getElementById('notSuitableOutcomeFilter'); if(!sel) return;
    const current=sel.value||'all', seen=new Map();
    getStoredBrokers().filter(isNotSuitable).forEach(item=>{ const key=getNotSuitableOutcomeKey(item); if(!seen.has(key)) seen.set(key,getNotSuitableOutcomeLabel(item)); });
    sel.innerHTML='<option value="all">All not-suitable outcomes</option>'+Array.from(seen.entries()).sort((a,b)=>a[1].localeCompare(b[1])).map(([k,v])=>`<option value="${escapeHtml(k)}">${escapeHtml(v)}</option>`).join('');
    if(Array.from(sel.options).some(o=>o.value===current)) sel.value=current;
}

/* ----------------------- Shared filter helpers --------------------------- */
// Normalise LoanTypes to an array (handles legacy strings and cloud-restored values)
function getLoanTypes(item){
    if(Array.isArray(item.LoanTypes)) return item.LoanTypes;
    return String(item.LoanTypes || item.LoanType || "").split(/[;,]/).map(s => s.trim()).filter(Boolean);
}
// Distinct values for a field across the FULL dataset (so options don't vanish as you filter)
function distinctValues(field){
    const set = new Set();
    allItems.forEach(i => {
        let v = i[field];
        if(field === "LoanTypes"){
            getLoanTypes(i).forEach(x => { if(x) set.add(String(x).trim()); });
            return;
        }
        if(Array.isArray(v)){ v.forEach(x => { if(x) set.add(String(x).trim()); }); }
        else if(v) set.add(String(v).trim());
    });
    return Array.from(set).sort((a,b)=> a.localeCompare(b));
}
// (Re)build a select's options from a list, preserving the current selection
function repopulateSelect(selectEl, placeholderLabel, values){
    if(!selectEl) return;
    const current = selectEl.value;
    let html = `<option value="all">${placeholderLabel}</option>`;
    values.forEach(v => { html += `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`; });
    selectEl.innerHTML = html;
    // Restore selection if still present, otherwise reset to 'all'
    selectEl.value = Array.from(selectEl.options).some(o => o.value === current) ? current : "all";
}
// Populate every filter dropdown from the full dataset. Call after reload / cloud-load / imports.
function populateFilterOptions(){
    const statuses  = distinctValues("Status");
    const loans     = distinctValues("LoanTypes");
    const networks  = distinctValues("Network");
    const volumes   = distinctValues("Volume");
    const cities    = distinctValues("City");
    // Pipeline
    repopulateSelect(document.getElementById("filterStatus"),    "All Statuses",   statuses);
    repopulateSelect(document.getElementById("filterLoanType"),   "All Loan Types", loans);
    repopulateSelect(document.getElementById("filterNetwork"),  "All Networks",  networks);
    repopulateSelect(document.getElementById("filterVolume"),    "All Volumes",   volumes);
    repopulateSelect(document.getElementById("filterCity"),      "All Cities",    cities);
    // Admin
    repopulateSelect(document.getElementById("adminFilterStatus"),    "All Statuses",   statuses);
    repopulateSelect(document.getElementById("adminFilterLoanType"), "All Loan Types", loans);
    repopulateSelect(document.getElementById("adminFilterNetwork"), "All Networks",  networks);
    repopulateSelect(document.getElementById("adminFilterVolume"),   "All Volumes",   volumes);
    repopulateSelect(document.getElementById("adminFilterCity"),     "All Cities",    cities);
}
// Audit dropdowns come from the audit log dataset
function populateAuditFilterOptions(){
    const logs = getStoredAuditLogs();
    const actions = Array.from(new Set(logs.map(l => l.Action).filter(Boolean))).sort();
    const users   = Array.from(new Set(logs.map(l => l.User).filter(Boolean))).sort();
    repopulateSelect(document.getElementById("auditFilterAction"), "All Actions", actions);
    repopulateSelect(document.getElementById("auditFilterUser"),  "All Users",   users);
    const dialerLogs = logs.filter(l => l.Trail === "auditDialer");
    repopulateSelect(document.getElementById("dialerAuditFilterAction"), "All Actions", [...new Set(dialerLogs.map(l=>l.Action).filter(Boolean))].sort());
    repopulateSelect(document.getElementById("dialerAuditFilterUser"),  "All Users",   [...new Set(dialerLogs.map(l=>l.User).filter(Boolean))].sort());
}
function sortItems(items, sortVal){
    const arr = items.slice();
    if(sortVal === "contactNew"){
        arr.sort((a,b)=> referenceDate(b).getTime() - referenceDate(a).getTime());
    } else if(sortVal === "contactOld"){
        arr.sort((a,b)=> referenceDate(a).getTime() - referenceDate(b).getTime());
    } else if(sortVal === "name"){
        arr.sort((a,b)=> (a.Company || a.Title || "").localeCompare(b.Company || b.Title || ""));
    } else { // recent
        arr.sort((a,b)=> (b.Id || 0) - (a.Id || 0));
    }
    return arr;
}
function clearPipelineFilters(){
    document.getElementById("searchBox").value = "";
    ["filterStatus","filterLoanType","filterNetwork","filterVolume","filterCity"].forEach(id => {
        const el = document.getElementById(id); if(el) el.value = "all";
    });
    const sortEl = document.getElementById("filterSort"); if(sortEl) sortEl.value = "recent";
    render();
}
function renderDialerAuditLog(){
    renderFilteredAudit("dialerAuditLogContainer", "dialerAuditSearchBox", "dialerAuditFilterAction", "dialerAuditFilterUser", getStoredAuditLogs().filter(l=>l.Trail === "auditDialer"));
}
function renderFilteredAudit(containerId, searchId, actionId, userId, sourceLogs){
    const searchVal = document.getElementById(searchId)?.value.trim().toLowerCase() || "";
    const af = document.getElementById(actionId)?.value || "all";
    const uf = document.getElementById(userId)?.value || "all";
    let logs = sourceLogs.slice();
    if(af !== "all") logs=logs.filter(l=>l.Action===af);
    if(uf !== "all") logs=logs.filter(l=>l.User===uf);
    if(searchVal) logs=logs.filter(l=>[l.Timestamp,l.User,l.Action,l.RecordTitle,l.Company,l.Reason].join(" ").toLowerCase().includes(searchVal));
    const c=document.getElementById(containerId); if(!c) return;
    if(!logs.length){ c.innerHTML='<div style="padding:15px;text-align:center;font-size:13px;color:var(--muted);">No audit logs match your filters.</div>'; return; }
    c.innerHTML='<table class="preview" style="margin-top:0;"><tr><th>Timestamp</th><th>User</th><th>Action</th><th>Record</th><th>Reason / Detail</th></tr>'+logs.map(l=>`<tr><td><code>${escapeHtml(l.Timestamp)}</code></td><td>${escapeHtml(l.User)}</td><td><strong>${escapeHtml(l.Action)}</strong></td><td>${escapeHtml(l.RecordTitle)} (${escapeHtml(l.Company)})</td><td>${escapeHtml(l.Reason)}</td></tr>`).join('')+'</table>';
}
function clearDialerAuditFilters(){
    ["dialerAuditSearchBox"].forEach(id=>{const e=document.getElementById(id);if(e)e.value="";});
    ["dialerAuditFilterAction","dialerAuditFilterUser"].forEach(id=>{const e=document.getElementById(id);if(e)e.value="all";});
    renderDialerAuditLog();
}

function clearAuditFilters(){
    document.getElementById("auditSearchBox").value = "";
    const a = document.getElementById("auditFilterAction"); const u = document.getElementById("auditFilterUser");
    if(a) a.value = "all"; if(u) u.value = "all";
    renderAuditLog();
}
function clearAdminFilters(){
    adminPipelinePage = 1;
    document.getElementById("adminSearchBox").value = "";
    ["adminFilterSelect","adminUserFilterSelect","adminFilterStatus","adminFilterLoanType","adminFilterNetwork","adminFilterVolume","adminFilterCity"].forEach(id => {
        const el = document.getElementById(id); if(el) el.value = "all";
    });
    const sortEl = document.getElementById("adminFilterSort"); if(sortEl) sortEl.value = "recent";
    renderAdmin();
}
function setPipelineView(view){
    currentView=view;
    document.querySelectorAll(".sub-nav button").forEach(b=>b.classList.toggle("active", b.dataset.view===view));
    const nsFilterWrap=document.getElementById("notSuitableOutcomeFilterWrap");
    if(nsFilterWrap) nsFilterWrap.style.display=view==="notSuitable" ? "flex" : "none";
    if(view==="notSuitable") populateNotSuitableOutcomeFilter();
    ["open","mine","notSuitable"].forEach(v=>{ const el=document.getElementById("listView"); if(el) el.style.display="block"; });
    // Explicitly render the selected dataset; never fall back to the open pool.
    allItems=getStoredBrokers();
    populateFilterOptions();
    render();
    updateAIBrokerSearchVisibility();
}

function render() {
    const query = document.getElementById("searchBox").value.trim().toLowerCase();
    const fStatus    = document.getElementById("filterStatus");
    const fLoanType  = document.getElementById("filterLoanType");
    const fNetwork  = document.getElementById("filterNetwork");
    const fVolume    = document.getElementById("filterVolume");
    const fCity      = document.getElementById("filterCity");
    const fSort      = document.getElementById("filterSort");

    const vStatus   = fStatus   ? fStatus.value   : "all";
    const vLoan     = fLoanType ? fLoanType.value : "all";
    const vNetwork  = fNetwork  ? fNetwork.value  : "all";
    const vVolume   = fVolume   ? fVolume.value   : "all";
    const vCity     = fCity     ? fCity.value     : "all";
    const vSort     = fSort     ? fSort.value     : "recent";

    const totalInView = allItems.filter(i => {
        if(currentView === "mine") return isMine(i);
        if(currentView === "notSuitable") return isNotSuitable(i);
        return isOpen(i);
    }).length;

    let items = allItems.filter(i => {
        if(currentView === "mine") return isMine(i);
        if(currentView === "notSuitable") return isNotSuitable(i);
        return isOpen(i);
    });

    if(currentView === "notSuitable") {
        const outcomeFilter=document.getElementById("notSuitableOutcomeFilter")?.value || "all";
        if(outcomeFilter !== "all") items = items.filter(i => getNotSuitableOutcomeKey(i) === outcomeFilter);
    }

    if(vStatus !== "all")  items = items.filter(i => (i.Status||"") === vStatus);
    if(vLoan !== "all")   items = items.filter(i => getLoanTypes(i).includes(vLoan));
    if(vNetwork !== "all")items = items.filter(i => (i.Network||"") === vNetwork);
    if(vVolume !== "all") items = items.filter(i => (i.Volume||"") === vVolume);
    if(vCity !== "all")   items = items.filter(i => (i.City||"") === vCity);

    if(query){
        items = items.filter(i => {
            const loansStr = getLoanTypes(i).join(" ");
            const hay = [i.Title, i.Company, i.Email, loansStr, i.Status, i.City, i.Network].join(" ").toLowerCase();
            return hay.includes(query);
        });
    }

    items = sortItems(items, vSort);

    const list = document.getElementById("cardList");
    list.innerHTML = "";
    
    let viewDesc = "in the open pool";
    if(currentView === "mine") viewDesc = "assigned to you";
    else if(currentView === "notSuitable") viewDesc = "marked as not suitable";

    document.getElementById("statusMsg").textContent = items.length + " " + viewDesc;
    const countEl = document.getElementById("pipelineFilterCount");
    if(countEl) countEl.textContent = items.length + " of " + totalInView + " brokers shown";

    if(items.length === 0) {
        list.innerHTML = `<div style="padding:20px; text-align:center; color:var(--muted); background:#fff; border-radius:var(--radius); border:1px solid var(--border);">No records match your filters. <a onclick="clearPipelineFilters()" style="color:var(--accent);cursor:pointer;">Clear filters</a></div>`;
        return;
    }

    items.forEach(item => {
        const card = document.createElement("div");
        card.className = "card";
        let badge = "";
        if(currentView === "open"){
            badge = '<span class="badge open">Open</span>';
        } else if(currentView === "mine") {
            badge = '<span class="badge mine">Assigned to you</span>';
        } else {
            badge = `<span class="badge not-suitable">${escapeHtml(getNotSuitableOutcomeLabel(item))}</span>`;
        }

        const refDt = referenceDate(item);
        const loanSummary = (() => { const lt = getLoanTypes(item); return lt.length > 0 ? lt.join(", ") : "General"; })();

        // The 6-week reversion only applies to ASSIGNED brokers, so only show it for those
        let reversionHtml = "";
        if(currentView === "mine" && item.AssignedTo && !item.IsNotSuitable){
            const reversionDt = new Date(refDt.getTime() + (getBrokerReleaseDays() * 24 * 60 * 60 * 1000));
            reversionHtml = ` <strong style="color:var(--amber); margin-left:6px;">(Releases if unupdated: ${reversionDt.toLocaleDateString()})</strong>`;
        }

        card.innerHTML = `
            <div class="main-info">
                <strong>${escapeHtml(item.Company || item.Title || "(no name)")}</strong>
                <span>${escapeHtml(item.Title || "")} · ${escapeHtml(loanSummary)} (${escapeHtml(item.Status || "Cold")}) · last contact ${refDt.toLocaleDateString()}${reversionHtml}</span>
            </div>
            ${badge}
        `;
        card.addEventListener("click", () => openDetail(item));
        list.appendChild(card);
    });
}

function escapeHtml(s){
    return String(s || "").replace(/[&<>"']/g, c => ({
        "&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"
    }[c]));
}

function openDetail(item){
    selectedItemId = item.Id;
    document.getElementById("detailTitle").textContent = item.Company || item.Title || "Broker";
    document.getElementById("f_name").value = item.Title || "";
    document.getElementById("f_company").value = item.Company || "";
    
    const phoneVal = item.Phone || "";
    document.getElementById("f_phone").value = phoneVal;
    const nuacomBtn = document.getElementById("nuacomCallBtn");
    if(phoneVal.trim()) {
        nuacomBtn.href = `tel:${phoneVal.trim()}`;
        nuacomBtn.style.pointerEvents = "auto";
        nuacomBtn.style.opacity = "1";
    } else {
        nuacomBtn.href = "#";
        nuacomBtn.style.pointerEvents = "none";
        nuacomBtn.style.opacity = "0.5";
    }

    const emailVal = item.Email || "";
    document.getElementById("f_email").value = emailVal;
    const outlookBtn = document.getElementById("outlookEmailBtn");
    if(emailVal.trim()) {
        outlookBtn.href = `mailto:${emailVal.trim()}?subject=7J%20Finance%20-%20Enquiry`;
        outlookBtn.style.pointerEvents = "auto";
        outlookBtn.style.opacity = "1";
    } else {
        outlookBtn.href = "#";
        outlookBtn.style.pointerEvents = "none";
        outlookBtn.style.opacity = "0.5";
    }

    let webVal = item.Website || "";
    document.getElementById("f_website").value = webVal;
    const webBtn = document.getElementById("websiteLinkBtn");
    if(webVal.trim()) {
        let formattedUrl = webVal.trim();
        if(!formattedUrl.startsWith("http://") && !formattedUrl.startsWith("https://")) {
            formattedUrl = "https://" + formattedUrl;
        }
        webBtn.href = formattedUrl;
        webBtn.style.pointerEvents = "auto";
        webBtn.style.opacity = "1";
    } else {
        webBtn.href = "#";
        webBtn.style.pointerEvents = "none";
        webBtn.style.opacity = "0.5";
    }

    document.getElementById("f_address").value = item.Address || "";
    document.getElementById("f_city").value = item.City || "";
    document.getElementById("f_prefComm").value = item.PrefComm || "Phone";
    
    // Set checkboxes for multi-select loan types
    const container = document.getElementById("f_loanTypesContainer");
    const checkboxes = container.querySelectorAll("input[type=checkbox]");
    const currentLoans = Array.isArray(item.LoanTypes) ? item.LoanTypes : (item.LoanType ? [item.LoanType] : []);
    checkboxes.forEach(chk => {
        chk.checked = currentLoans.includes(chk.value);
    });

    document.getElementById("f_volume").value = item.Volume || "Under £1M";
    document.getElementById("f_network").value = item.Network || "";
    document.getElementById("f_status").value = item.Status || "Cold";
    document.getElementById("f_nextFollowUp").value = item.NextFollowUp || "";

    // Set notes preview
    updateNotesWidgetDisplay(item.Notes || "");

    const assignInfo = document.getElementById("assignInfo");
    const assignBtn = document.getElementById("assignBtn");
    const releaseBtn = document.getElementById("releaseBtn");
    const moveToNotSuitableBtn = document.getElementById("moveToNotSuitableBtn");
    const revertToOpenBtn = document.getElementById("revertToOpenBtn");
    
    if(item.IsNotSuitable) {
        assignInfo.textContent = "Status: Moved to Not Suitable" + (item.NotSuitableReason ? ` (Reason: ${item.NotSuitableReason})` : "");
        assignBtn.style.display = "none";
        releaseBtn.style.display = "none";
        moveToNotSuitableBtn.style.display = "none";
        revertToOpenBtn.style.display = "inline-block";
    } else if(item.AssignedTo){
        assignInfo.textContent = "Currently assigned to " + item.AssignedTo.Title;
        moveToNotSuitableBtn.style.display = "inline-block";
        revertToOpenBtn.style.display = "none";
        if(isMine(item)){
            assignBtn.style.display = "none";
            releaseBtn.style.display = "inline-block";
        } else {
            assignBtn.style.display = "inline-block";
            assignBtn.textContent = "Reassign to me";
            releaseBtn.style.display = "none";
        }
    } else {
        assignInfo.textContent = "Currently unassigned";
        assignBtn.style.display = "inline-block";
        assignBtn.textContent = "Assign to me";
        releaseBtn.style.display = "none";
        moveToNotSuitableBtn.style.display = "inline-block";
        revertToOpenBtn.style.display = "none";
    }

    document.getElementById("detailOverlay").classList.add("show");
    markDetailClean();
}

function updateNotesWidgetDisplay(notesText) {
    const previewEl = document.getElementById("notesReadOnlyPreview");
    if(!notesText || notesText.trim() === "") {
        previewEl.textContent = "No notes recorded yet.";
    } else {
        previewEl.textContent = notesText;
    }
}

function parseStructuredNotes(text){
    if(!text)return []; return String(text).split(/\n\n--------------------\n\n/).filter(Boolean).map(raw=>{const m=raw.match(/^\[([^·]+)·([^·]+)·([^\]]+)\]:\n([\s\S]*)$/);return m?{timestamp:m[1].trim(),author:m[2].trim(),category:m[3].trim(),text:m[4]}:{timestamp:'',author:'',category:'General',text:raw};});
}
function renderNotesHistory(){
    const item=getStoredBrokers().find(i=>i.Id===selectedItemId); if(!item)return; const filter=document.getElementById('notesHistoryFilter')?.value||'all'; const notes=parseStructuredNotes(item.Notes||'').filter(n=>filter==='all'||n.category===filter); const el=document.getElementById('notesHistoryDisplay');
    el.innerHTML=notes.length?notes.map(n=>`<div style="padding:9px 0;border-bottom:1px solid #e5e7eb"><div style="font-size:11px;color:#6264a7;font-weight:700">${escapeHtml(n.category)} · ${escapeHtml(n.author)} · ${escapeHtml(n.timestamp)}</div><div style="margin-top:4px;white-space:pre-wrap">${escapeHtml(n.text)}</div></div>`).join(''):'<div style="color:var(--muted)">No notes match this category.</div>';
}
function openNotesPopup() {
    allItems = getStoredBrokers(); const item = allItems.find(i => i.Id === selectedItemId); if(!item) return;
    const f=document.getElementById('notesHistoryFilter'); if(f)f.value='all'; document.getElementById('newNoteInput').value=''; const cat=document.getElementById('newNoteCategory'); if(cat)cat.value='General'; renderNotesHistory(); document.getElementById('notesPopupModal').classList.add('show');
}
function closeNotesPopup() {
    document.getElementById("notesPopupModal").classList.remove("show");
}

function appendNewNote() {
    const newNoteText = document.getElementById("newNoteInput").value.trim();
    if(!newNoteText) {
        alert("Please enter a note before saving.");
        return;
    }

    allItems = getStoredBrokers();
    const idx = allItems.findIndex(i => i.Id === selectedItemId);
    if(idx !== -1) {
        const timestamp = new Date().toLocaleString("en-GB");
        const author = currentUser ? (currentUser.Title || currentUser.Email || "System") : "Demo User";
        const category = document.getElementById("newNoteCategory")?.value || "General";
        const formattedEntry = `[${timestamp} · ${author} · ${category}]:\n${newNoteText}`;
        
        const existingNotes = allItems[idx].Notes || "";
        const combinedNotes = existingNotes ? existingNotes + "\n\n--------------------\n\n" + formattedEntry : formattedEntry;
        
        allItems[idx].Notes = combinedNotes;
        allItems[idx].Modified = new Date().toISOString();
        allItems[idx].LastContactDate = new Date().toISOString(); // Update last contact on note addition
        saveStoredBrokers(allItems);

        logAuditRecord(allItems[idx], "NOTE_ADDED", `Added note: "${newNoteText.slice(0, 40)}..."`);

        updateNotesWidgetDisplay(combinedNotes);
        renderNotesHistory();
        document.getElementById("newNoteInput").value = "";
        alert("Note successfully added and hard-saved to broker record.");
    }
    reload();
}

async function cloudFindBrokerForRevert(localId,local){const sid=await getSiteId(),listId=await resolveListId(getListName('brokers'));const data=await graphGet(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items?expand=fields&$top=500`);const target=String(localId);const found=(data.value||[]).find(it=>String(it.fields?.PortalId||'')===target)||(data.value||[]).find(it=>{const f=it.fields||{};return String(f.Email||'').toLowerCase()===String(local?.Email||'').toLowerCase()&&String(f.Company||'').toLowerCase()===String(local?.Company||'').toLowerCase();});if(!found)return null;const b=fieldsToBroker(found.fields||{},Number(localId));_brokerMeta[localId]={spId:found.id,etag:found['@odata.etag']||''};const map=getBrokerIdMap();map[localId]=found.id;setBrokerIdMap(map);return {broker:b,etag:found['@odata.etag']||'',spId:found.id};}
async function setBrokerOpenState(localId,reason){const local=getStoredBrokers().find(i=>String(i.Id)===String(localId));if(!local)throw new Error('Broker record could not be found.');if(m365Configured()&&!isDemoMode()){let current=await cloudGetCurrentBroker(localId);if(!current)current=await cloudFindBrokerForRevert(localId,local);if(!current)throw new Error('Could not locate this broker in the configured Microsoft List. Refresh the broker data and try again.');const patch={AssignedTo:null,AssignedToEmail:'',AssignedToName:'',AssignedToLookupId:null,AssignedToId:null,IsNotSuitable:false,NotSuitableReason:'',NotSuitableSource:'',DiallerOutcome:'',DiallerOutcomeReason:'',DiallerOutcomeDate:'',ClaimedAt:'',ClaimedBy:'',ClaimExpiresAt:'',WorkflowState:'OPEN',WorkflowStateUpdatedAt:new Date().toISOString(),Modified:new Date().toISOString()};await graphPatch(`${GRAPH_BASE}/sites/${await getSiteId()}/lists/${await resolveListId(getListName('brokers'))}/items/${current.spId}/fields`,filterBrokerFieldsForColumns(patch,await getBrokerColumnNames()),current.etag||undefined);const revertLog={RecordId:Number(localId),Timestamp:new Date().toISOString(),User:currentUser?.Email||currentUser?.Title||'System',Action:'REVERTED_TO_OPEN',RecordTitle:local.Title||'',Company:local.Company||'',Reason:reason||'Reverted broker back to Open Brokers',Trail:'auditAccount'}; try{await cloudPushAudit(revertLog);}catch(auditErr){console.warn('Revert audit could not be written:',auditErr);} const fresh=await cloudLoadBrokers(),check=fresh.find(i=>String(i.Id)===String(localId));if(check&&(isNotSuitable(check)||hasAssignment(check)))throw new Error('SharePoint still reports this broker as assigned or Not Suitable. The record may have a workflow rule or calculated field preventing the reset.');localStorage.setItem(STORAGE_KEY,JSON.stringify(fresh));_brokerSnapshot=_snapshotBrokers(fresh);allItems=fresh;}else{const items=getStoredBrokers(),idx=items.findIndex(i=>String(i.Id)===String(localId));if(idx<0)throw new Error('Broker record could not be found in local data.');Object.assign(items[idx],{AssignedTo:null,AssignedToEmail:'',AssignedToName:'',AssignedToLookupId:'',AssignedToId:'',IsNotSuitable:false,NotSuitableReason:'',NotSuitableSource:'',DiallerOutcome:'',DiallerOutcomeReason:'',DiallerOutcomeDate:'',ClaimedAt:'',ClaimedBy:'',ClaimExpiresAt:'',WorkflowState:'OPEN',WorkflowStateUpdatedAt:new Date().toISOString(),Modified:new Date().toISOString()});saveStoredBrokers(items);allItems=items;}const b=getStoredBrokers().find(i=>String(i.Id)===String(localId))||local;if(isDemoMode()||!m365Configured()) logAuditRecord(b,'REVERTED_TO_OPEN',reason||'Reverted broker back to Open Brokers');}

async function moveCurrentToNotSuitable(){
    const reason=prompt("Select or enter reason for marking Not Suitable:\n1. Out of Geographic Scope\n2. Does not handle Bridging/Secured\n3. Compliance / Unregulated\n4. Duplicate / Inactive Business\n5. Other","Out of Geographic Scope"); if(reason===null)return;
    const localId=selectedItemId;
    try{
        if(m365Configured()&&!isDemoMode()){
            const current=await cloudGetCurrentBroker(localId); if(!current)throw new Error("Broker is no longer available.");
            const now=new Date().toISOString(); const patch={WorkflowState:"NOT_SUITABLE",IsNotSuitable:"Yes",NotSuitableReason:reason,NotSuitableSource:"Manual",AssignedTo:"",AssignedToEmail:"",AssignedToName:"",AssignedToLookupId:null,DiallerOutcome:"",DiallerOutcomeReason:"",DiallerOutcomeDate:now,ClaimedAt:"",ClaimedBy:"",ClaimExpiresAt:"",Modified:now};
            await graphPatch(`${GRAPH_BASE}/sites/${await getSiteId()}/lists/${await resolveListId(getListName('brokers'))}/items/${current.spId}/fields`,filterBrokerFieldsForColumns(patch,await getBrokerColumnNames()),current.etag||undefined);
            const fresh=hydrateBrokerOutcomeMetadata(await cloudLoadBrokers(),getStoredAuditLogs()); localStorage.setItem(STORAGE_KEY,JSON.stringify(fresh)); _brokerSnapshot=_snapshotBrokers(fresh); allItems=fresh;
        }else{
            const items=getStoredBrokers(),idx=items.findIndex(i=>String(i.Id)===String(localId)); if(idx>=0){Object.assign(items[idx],{WorkflowState:"NOT_SUITABLE",IsNotSuitable:true,NotSuitableReason:reason,NotSuitableSource:"Manual",AssignedTo:null,DiallerOutcome:"",DiallerOutcomeReason:"",DiallerOutcomeDate:new Date().toISOString(),Modified:new Date().toISOString()});saveStoredBrokers(items);allItems=items;}
        }
        const b=getStoredBrokers().find(i=>String(i.Id)===String(localId))||local; logAuditRecord(b,"MARKED_NOT_SUITABLE",`Moved to Not Suitable. Reason: ${reason}`); reload(); detailInitialSnapshot=null; detailDirty=false; document.getElementById("detailOverlay").classList.remove("show");
    }catch(e){alert(e.message||"Could not move broker to Not Suitable.");}
}
async function revertCurrentToOpen(){
    try{await setBrokerOpenState(selectedItemId,"Reverted broker from Not Suitable back to Open Brokers");reload();detailInitialSnapshot=null;detailDirty=false;document.getElementById("detailOverlay").classList.remove("show");}catch(e){alert(e.message||"Could not revert broker to Open Brokers.");}
}

function openMicrosoftBrokerRecord(){
    const b=getStoredBrokers().find(x=>String(x.Id)===String(selectedItemId));
    if(!b){alert("Open a broker record first.");return;}
    const map=getBrokerIdMap(); const spId=map[b.Id];
    if(!spId){alert("This broker has not been linked to a Microsoft List item yet. Save the record first, then try again.");return;}
    window.open(getSharePointRecordUrl("brokers",spId),"_blank","noopener");
}

function openBrokerAuditModal() {
    allItems = getStoredBrokers();
    const item = allItems.find(i => i.Id === selectedItemId);
    if(!item) return;

    document.getElementById("brokerAuditTitle").textContent = `Audit Trail: ${item.Company || item.Title || "Broker"}`;
    const logs = getStoredAuditLogs().filter(l => l.RecordId === item.Id);
    
    const container = document.getElementById("brokerAuditContent");
    if(logs.length === 0) {
        container.innerHTML = `<p style="font-size:13px; color:var(--muted);">No specific audit logs recorded for this broker yet.</p>`;
    } else {
        let html = "<table class='preview'><tr><th>Timestamp</th><th>User</th><th>Action</th><th>Detail</th></tr>";
        logs.forEach(l => {
            html += `<tr>
                <td><code>${escapeHtml(l.Timestamp)}</code></td>
                <td>${escapeHtml(l.User)}</td>
                <td><strong>${escapeHtml(l.Action)}</strong></td>
                <td>${escapeHtml(l.Reason)}</td>
            </tr>`;
        });
        html += "</table>";
        container.innerHTML = html;
    }

    document.getElementById("brokerAuditModal").classList.add("show");
}

function closeBrokerAuditModal() {
    document.getElementById("brokerAuditModal").classList.remove("show");
}

document.getElementById("f_phone").addEventListener("input", e => {
    const val = e.target.value.trim();
    const btn = document.getElementById("nuacomCallBtn");
    if(val) {
        btn.href = `tel:${val}`;
        btn.style.pointerEvents = "auto";
        btn.style.opacity = "1";
    } else {
        btn.href = "#";
        btn.style.pointerEvents = "none";
        btn.style.opacity = "0.5";
    }
});

document.getElementById("f_email").addEventListener("input", e => {
    const val = e.target.value.trim();
    const btn = document.getElementById("outlookEmailBtn");
    if(val) {
        btn.href = `mailto:${val}?subject=7J%20Finance%20-%20Enquiry`;
        btn.style.pointerEvents = "auto";
        btn.style.opacity = "1";
    } else {
        btn.href = "#";
        btn.style.pointerEvents = "none";
        btn.style.opacity = "0.5";
    }
});

document.getElementById("f_website").addEventListener("input", e => {
    const val = e.target.value.trim();
    const btn = document.getElementById("websiteLinkBtn");
    if(val) {
        let formattedUrl = val;
        if(!formattedUrl.startsWith("http://") && !formattedUrl.startsWith("https://")) {
            formattedUrl = "https://" + formattedUrl;
        }
        btn.href = formattedUrl;
        btn.style.pointerEvents = "auto";
        btn.style.opacity = "1";
    } else {
        btn.href = "#";
        btn.style.pointerEvents = "none";
        btn.style.opacity = "0.5";
    }
});

function saveDetail(){
    allItems = getStoredBrokers();
    const idx = allItems.findIndex(i => i.Id === selectedItemId);
    if(idx !== -1) {
        const old = allItems[idx];
        
        // Gather selected loan types from checkboxes
        const container = document.getElementById("f_loanTypesContainer");
        const checkboxes = container.querySelectorAll("input[type=checkbox]:checked");
        const newLoanTypes = Array.from(checkboxes).map(chk => chk.value);
        const oldLoanTypes = Array.isArray(old.LoanTypes) ? old.LoanTypes : (old.LoanType ? [old.LoanType] : []);

        let changes = [];
        
        // Compare simple fields
        const fieldsToCheck = [
            { key: "Title", label: "Contact Name", el: "f_name" },
            { key: "Company", label: "Company", el: "f_company" },
            { key: "Phone", label: "Phone", el: "f_phone" },
            { key: "Email", label: "Email", el: "f_email" },
            { key: "Website", label: "Website", el: "f_website" },
            { key: "Address", label: "Address", el: "f_address" },
            { key: "City", label: "City", el: "f_city" },
            { key: "PrefComm", label: "Preferred Comm", el: "f_prefComm" },
            { key: "Volume", label: "Volume", el: "f_volume" },
            { key: "Network", label: "Network", el: "f_network" },
            { key: "Status", label: "Status", el: "f_status" },
            { key: "NextFollowUp", label: "Next Follow-Up", el: "f_nextFollowUp" }
        ];

        fieldsToCheck.forEach(f => {
            const newVal = document.getElementById(f.el).value;
            const oldVal = old[f.key] || "";
            if(newVal !== oldVal) {
                changes.push(`${f.label}: '${oldVal}' ➔ '${newVal}'`);
                allItems[idx][f.key] = newVal;
            }
        });

        // Compare loan types array
        if(JSON.stringify(newLoanTypes.sort()) !== JSON.stringify(oldLoanTypes.sort())) {
            changes.push(`Loan Types: '${oldLoanTypes.join(", ")}' ➔ '${newLoanTypes.join(", ")}'`);
            allItems[idx].LoanTypes = newLoanTypes;
            delete allItems[idx].LoanType; // Clean up legacy single property
        }

        allItems[idx].Modified = new Date().toISOString();

        if(document.getElementById("f_logContact").checked){
            allItems[idx].LastContactDate = new Date().toISOString();
        }
        saveStoredBrokers(allItems);

        const auditSummary = changes.length > 0 
            ? "Updated fields: " + changes.join(" | ") 
            : "Saved record (no field changes)";

        logAuditRecord(allItems[idx], "UPDATE", auditSummary);
    }
    reload();
    detailInitialSnapshot=null; detailDirty=false;
    document.getElementById("detailOverlay").classList.remove("show");
}

async function claimBrokerOnCloud(broker){
    if(isDemoMode() || !m365Configured()) return {ok:true};
    try{
        const sid=await getSiteId(); const listId=await resolveListId(getListName('brokers')); const map=getBrokerIdMap(); const spId=map[broker.Id];
        if(!spId) return {ok:true};
        const tok=await getGraphToken();
        const url=`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items/${spId}?expand=fields`;
        const r=await fetch(url,{headers:{Authorization:'Bearer '+tok}});
        if(!r.ok) throw new Error('Could not refresh broker before assignment ('+r.status+')');
        const latest=await r.json(); const f=latest.fields||{};
        if(f.IsNotSuitable && String(f.IsNotSuitable).toLowerCase()==='yes') return {ok:false,reason:'This broker has already been marked Not Suitable by another user.'};
        if(f.AssignedTo) return {ok:false,reason:'This broker has already been assigned to another BDM. The list was refreshed.'};
        const patch={AssignedTo:`${currentUser.Title||''} <${currentUser.Email||''}>`,LastContactDate:new Date().toISOString(),Status:broker.Status||'Cold'};
        const pr=await fetch(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items/${spId}/fields`,{method:'PATCH',headers:{Authorization:'Bearer '+tok,'Content-Type':'application/json','If-Match':latest['@odata.etag']||'*'},body:JSON.stringify(patch)});
        if(pr.status===412) return {ok:false,reason:'Another BDM claimed this broker first. Please refresh the queue.'};
        if(!pr.ok) throw new Error('Assignment update failed ('+pr.status+')');
        return {ok:true};
    }catch(e){ return {ok:false,reason:e.message||String(e)}; }
}

async function updateSingleBrokerCloud(localId,patch){
    const current=await cloudGetCurrentBroker(localId); if(!current)throw new Error("Broker is no longer available.");
    const sid=await getSiteId(),listId=await resolveListId(getListName("brokers"));
    await graphPatch(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items/${current.spId}/fields`,filterBrokerFieldsForColumns(patch,await getBrokerColumnNames()),current.etag||undefined);
    return await cloudGetCurrentBroker(localId);
}

async function assignToMe(){
    const localId=selectedItemId; if(localId==null) return;
    const btn=document.getElementById("assignBtn"); if(btn){btn.disabled=true;btn.textContent="Checking availability…";}
    try{
        if(m365Configured() && !isDemoMode()){
            const current=await cloudGetCurrentBroker(localId);
            if(!current) throw new Error("This broker is no longer available.");
            if(isNotSuitable(current.broker)){
                localStorage.setItem(STORAGE_KEY,JSON.stringify(await cloudLoadBrokers())); allItems=getStoredBrokers(); populateNotSuitableOutcomeFilter(); render();
                throw new Error("This broker has been moved to Not Suitable.");
            }
            if(hasAssignment(current.broker)){
                localStorage.setItem(STORAGE_KEY,JSON.stringify(await cloudLoadBrokers())); allItems=getStoredBrokers(); render();
                throw new Error("This broker has already been claimed by another BDM.");
            }
            const sid=await getSiteId(), listId=await resolveListId(getListName("brokers"));
            const now=new Date().toISOString();
            const claimMinutes=Math.min(30,Math.max(1,parseInt(getGlobalSettingValue("claimMinutes",5),10)||5)); const fields={AssignedTo:(currentUser.Title||currentUser.Email)+" <"+currentUser.Email+">",AssignedToEmail:currentUser.Email||"",AssignedToName:currentUser.Title||currentUser.Email||"",LastContactDate:now,Modified:now,IsNotSuitable:"No",ClaimedAt:now,ClaimedBy:currentUser.Email||currentUser.Title||"",ClaimExpiresAt:new Date(Date.now()+claimMinutes*60000).toISOString()};
            await graphPatch(`${GRAPH_BASE}/sites/${sid}/lists/${listId}/items/${current.spId}/fields`,filterBrokerFieldsForColumns(fields,await getBrokerColumnNames()),current.etag);
            // Refresh after the successful atomic claim so every local view uses the server state.
            const fresh=await cloudLoadBrokers(); localStorage.setItem(STORAGE_KEY,JSON.stringify(fresh)); _brokerSnapshot=_snapshotBrokers(fresh); allItems=fresh;
            logAuditRecord(fresh.find(b=>String(b.Id)===String(localId))||Object.assign({},current.broker,fields),"ASSIGN",`Assigned to ${currentUser.Title}`);
        }else{
            const brokers=getStoredBrokers(), idx=brokers.findIndex(b=>String(b.Id)===String(localId));
            if(idx!==-1){ brokers[idx].AssignedTo={Title:currentUser.Title,EMail:currentUser.Email}; brokers[idx].Modified=new Date().toISOString(); brokers[idx].LastContactDate=new Date().toISOString(); saveStoredBrokers(brokers); logAuditRecord(brokers[idx],"ASSIGN",`Assigned to ${currentUser.Title}`); }
        }
        reload(); detailInitialSnapshot=null; detailDirty=false; document.getElementById("detailOverlay").classList.remove("show");
    }catch(e){
        alert(e.message||"Unable to claim this broker.");
        if(m365Configured()&&!isDemoMode()){ try{const fresh=await cloudLoadBrokers();localStorage.setItem(STORAGE_KEY,JSON.stringify(fresh));_brokerSnapshot=_snapshotBrokers(fresh);allItems=fresh;}catch(_){} }
        reload();
    }finally{ if(btn){btn.disabled=false;btn.textContent="Assign to me";} }
}

async function release(){
    const localId=selectedItemId; if(localId==null)return;
    try{await setBrokerOpenState(localId,"Released from My Brokers back to Open Brokers");reload();document.getElementById("detailOverlay").classList.remove("show");detailInitialSnapshot=null;detailDirty=false;}catch(e){alert(e.message||"Could not release this broker.");}
}

function buildLocalBroker(body){
    return {
        Id: Date.now() + Math.floor(Math.random() * 1000000),
        Modified: new Date().toISOString(),
        Title: body.Title || "",
        Company: body.Company || "",
        Phone: body.Phone || "",
        Email: body.Email || "",
        Website: body.Website || "",
        Address: body.Address || "",
        City: body.City || "",
        Notes: body.Notes || "",
        PrefComm: body.PrefComm || "Phone",
        LoanTypes: Array.isArray(body.LoanTypes) ? body.LoanTypes : (body.LoanTypes ? body.LoanTypes.split(",").map(s => s.trim()).filter(Boolean) : ["Residential Bridging"]),
        Volume: body.Volume || "Under £1M",
        Network: body.Network || "",
        Status: body.Status || "Cold",
        NextFollowUp: body.NextFollowUp || "",
        LastContactDate: new Date().toISOString(),
        AssignedTo: null,
        IsNotSuitable: false
    };
}

function createItemLocal(body){
    allItems = getStoredBrokers();
    const newItem = buildLocalBroker(body);
    allItems.push(newItem);
    saveStoredBrokers(allItems);
    logAuditRecord(newItem, "CREATE", "Created new broker record");
    return newItem;
}

function deleteItemLocal(id, reason = "Deleted by admin"){
    // Demo/local deletion: IDs may arrive as strings or numbers. Keep a tombstone
    // so the demo seed cannot recreate a broker that an admin deliberately deleted.
    const targetId = String(id);
    let items = getStoredBrokers();
    const target = items.find(i => String(i.Id) === targetId || String(i.PortalId || "") === targetId);
    if(!target) throw new Error("Broker record could not be found in the local CRM data.");
    logAuditRecord(target, "DELETE", reason);
    items = items.filter(i => !(String(i.Id) === targetId || String(i.PortalId || "") === targetId));
    if(isDemoMode()){
        try{
            const deleted = JSON.parse(localStorage.getItem("7J_DEMO_DELETED_BROKERS") || "[]");
            if(!deleted.includes(targetId)) deleted.push(targetId);
            localStorage.setItem("7J_DEMO_DELETED_BROKERS", JSON.stringify(deleted));
        }catch(e){}
    }
    allItems = items;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(items));
    _brokerSnapshot = _snapshotBrokers(items);
}

function reload(){
    allItems = getStoredBrokers();
    populateFilterOptions();        // refresh broker dropdowns from the full dataset
    populateAuditFilterOptions();   // refresh audit dropdowns from the log
    populateDealsFilters();        // refresh deals filter dropdowns
    populateDialerLoanFilter();    // refresh dialer loan-type dropdown
    if(currentView === "admin"){
        if(document.getElementById("adminUnlockedContent").style.display === "block") {
            populateAdminUserFilterOptions();
            renderAdmin();
            renderAuditLog();
        }
    } else if(currentView === "deals"){
        renderDeals();
    } else if(currentView === "kpiDashboard"){
        renderKPIDashboard();
    } else if(currentView === "kpi"){
        renderKPI();
    } else if(currentView === "dialer"){
        renderDialer();
    } else {
        render();
    }
}

/* ---- Unified KPI Dashboard rendering ---- */
function clearKPIDashboardFilters(){
    const p=document.getElementById("dashboardPeriod"),u=document.getElementById("dashboardUser"),y=document.getElementById("dashboardYear");
    if(p)p.value="all"; if(u)u.value="all"; if(y)y.value=String(new Date().getFullYear());
    const ds=document.getElementById("dashboardStartDate"),de=document.getElementById("dashboardEndDate"); if(ds)ds.value=""; if(de)de.value=""; renderKPIDashboard();
}
function populateKPIDashboardFilters(){
    const users=new Set();
    getDialerLogs().forEach(l=>{if(l.User)users.add(l.User.split(" (")[0]);});
    getStoredDeals().forEach(d=>{if(d.BDM)users.add(d.BDM);});
    const u=document.getElementById("dashboardUser"); if(u){const cur=u.value;u.innerHTML='<option value="all">All BDMs</option>'+Array.from(users).sort().map(x=>`<option value="${escapeHtml(x)}">${escapeHtml(x)}</option>`).join("");u.value=Array.from(users).includes(cur)?cur:"all";}
    const years=new Set([new Date().getFullYear()]); getStoredDeals().forEach(d=>{const y=new Date(d.DealDate||0).getFullYear();if(y>2000)years.add(y);});
    const y=document.getElementById("dashboardYear"); if(y){const cur=Number(y.value)||new Date().getFullYear();y.innerHTML=Array.from(years).sort((a,b)=>b-a).map(v=>`<option value="${v}">${v}</option>`).join("");y.value=years.has(cur)?String(cur):String(new Date().getFullYear());}
}
function getKPIUserScope(){
    return (currentUser && currentUser.Role !== "Admin") ? String(currentUser.Title||currentUser.Email||"") : "all";
}
function normaliseKPIUserName(v){ return String(v||"").split(" (")[0].trim().toLowerCase(); }
function renderKPIDashboard(){
    if(!document.getElementById("kpiDashboardView"))return; populateKPIDashboardFilters();
    const period=document.getElementById("dashboardPeriod")?.value||"all";
    const selectedUser=document.getElementById("dashboardUser")?.value||"all";
    const user=(currentUser && currentUser.Role!=="Admin") ? (currentUser.Title||currentUser.Email||"") : selectedUser;
    const year=Number(document.getElementById("dashboardYear")?.value)||new Date().getFullYear();
    const customStart=document.getElementById("dashboardStartDate")?.value, customEnd=document.getElementById("dashboardEndDate")?.value;
    let deals, logs;
    if(customStart || customEnd){
        const cs=customStart?new Date(customStart+"T00:00:00"):new Date("1900-01-01T00:00:00");
        const ce=customEnd?new Date(customEnd+"T23:59:59.999"):new Date("2999-12-31T23:59:59.999");
        deals=getStoredDeals().filter(d=>{const dt=new Date(d.DealDate||0);return dt>=cs&&dt<=ce;});
        logs=getDialerLogs().filter(l=>{const dt=new Date(l.Timestamp||0);return dt>=cs&&dt<=ce;});
    } else {
        deals=filterByPeriod(getStoredDeals(),period,d=>d.DealDate);
        logs=filterByPeriod(getDialerLogs(),period,l=>l.Timestamp);
    }
    if(user!=="all")deals=deals.filter(d=>normaliseKPIUserName(d.BDM)===normaliseKPIUserName(user));
    if(user!=="all")logs=logs.filter(l=>normaliseKPIUserName(l.User)===normaliseKPIUserName(user));
    const dealValue=deals.reduce((s,d)=>s+(Number(d.DealValue)||0),0);
    const averageDeal=deals.length?dealValue/deals.length:0;
    const uniqueDealBrokers=new Set(deals.map(d=>String(d.BrokerId||d.Company||d.BrokerName||"")).filter(Boolean)).size;
    const monthStart=new Date(); monthStart.setDate(1); monthStart.setHours(0,0,0,0);
    const dealsThisMonth=deals.filter(d=>new Date(d.DealDate||0)>=monthStart).length;
    const cards=[["Deals",deals.length],["Deal value",fmtGBP(dealValue)],["Average deal",fmtGBP(averageDeal)],["Brokers with deals",uniqueDealBrokers],["Deals this month",dealsThisMonth]];
    document.getElementById("dashboardSummary").innerHTML=cards.map(x=>`<div><div class="kpi-label">${x[0]}</div><div class="kpi-value">${x[1]}</div></div>`).join("");
    const monthly=Array.from({length:12},(_,m)=>{const md=deals.filter(d=>{const dt=new Date(d.DealDate||0);return dt.getFullYear()===year&&dt.getMonth()===m});return {m,count:md.length,value:md.reduce((s,d)=>s+(Number(d.DealValue)||0),0)};});
    const max=Math.max(1,...monthly.map(x=>x.value));
    document.getElementById("dashboardMonthlyChart").innerHTML=monthly.map(x=>`<div class="month-col" title="${x.count} deals · ${fmtGBP(x.value)}"><span class="month-value">${fmtGBP(x.value)}</span><div class="month-bar" style="height:${Math.max(4,Math.round(x.value/max*150))}px"></div><span class="month-label">${new Date(year,x.m,1).toLocaleDateString("en-GB",{month:"short"})}</span></div>`).join("");
    const brokers={};deals.forEach(d=>{const k=String(d.BrokerId||d.BrokerName||d.Company||d.Id);if(!brokers[k])brokers[k]={name:d.BrokerName||"(No broker)",company:d.Company||"",deals:0,value:0};brokers[k].deals++;brokers[k].value+=Number(d.DealValue)||0;});
    const br=Object.values(brokers).sort((a,b)=>b.value-a.value||b.deals-a.deals).slice(0,10);
    document.getElementById("dashboardBrokerRanking").innerHTML=br.length?'<table class="kpi-table"><tr><th>#</th><th>Broker</th><th>Deals</th><th>Value</th></tr>'+br.map((r,i)=>`<tr><td class="kpi-rank">${i+1}</td><td><strong>${escapeHtml(r.name)}</strong><br><span style="font-size:10px;color:#737373">${escapeHtml(r.company)}</span></td><td>${r.deals}</td><td><strong>${fmtGBP(r.value)}</strong></td></tr>`).join('')+'</table>':'<div style="padding:10px;color:var(--muted)">No deals for this period.</div>';
    const bdms={};deals.forEach(d=>{const k=d.BDM||"(unassigned)";if(!bdms[k])bdms[k]={name:k,deals:0,value:0};bdms[k].deals++;bdms[k].value+=Number(d.DealValue)||0;});
    const bdm=Object.values(bdms).sort((a,b)=>b.value-a.value).slice(0,10);
    document.getElementById("dashboardBdmRanking").innerHTML=bdm.length?'<table class="kpi-table"><tr><th>#</th><th>BDM</th><th>Deals</th><th>Value</th></tr>'+bdm.map((r,i)=>`<tr><td class="kpi-rank">${i+1}</td><td>${escapeHtml(r.name)}</td><td>${r.deals}</td><td><strong>${fmtGBP(r.value)}</strong></td></tr>`).join('')+'</table>':'<div style="padding:10px;color:var(--muted)">No deals.</div>';
}

/* ---- KPI Tracker rendering ---- */
function renderKPI(){
    const sel=document.getElementById("kpiYear"); if(!sel)return; const all=getStoredDeals(); const years=new Set([new Date().getFullYear()]); all.forEach(d=>{const y=new Date(d.DealDate||0).getFullYear();if(y>2000)years.add(y)}); const current=Number(sel.value)||new Date().getFullYear(); sel.innerHTML=Array.from(years).sort((a,b)=>b-a).map(y=>`<option value="${y}">${y}</option>`).join(""); sel.value=years.has(current)?String(current):String(new Date().getFullYear());
    const rows=Array.from({length:12},(_,m)=>{const ds=all.filter(d=>{const dt=new Date(d.DealDate||0);return dt.getFullYear()===Number(sel.value)&&dt.getMonth()===m});return {month:new Date(Number(sel.value),m,1).toLocaleDateString("en-GB",{month:"short"}),count:ds.length,value:ds.reduce((s,d)=>s+(Number(d.DealValue)||0),0)}});
    const total=rows.reduce((s,r)=>s+r.value,0), count=rows.reduce((s,r)=>s+r.count,0), max=Math.max(1,...rows.map(r=>r.value));
    document.getElementById("kpiSummary").innerHTML=[['Deals',count],['Deal value',fmtGBP(total)],['Average deal',fmtGBP(count?total/count:0)],['Best month',rows.slice().sort((a,b)=>b.value-a.value)[0]?.month||'-']].map(x=>`<div style="background:#fff;border:1px solid var(--border);padding:14px;border-radius:var(--radius);"><div style="font-size:11px;color:var(--muted);font-weight:600;">${x[0]}</div><div style="font-size:20px;font-weight:700;color:var(--primary-dark);margin-top:4px;">${x[1]}</div></div>`).join('');
    document.getElementById("kpiMonthlyTable").innerHTML=`<table class="preview"><tr><th>Month</th><th>Deals</th><th>Deal value</th><th>Average</th></tr>${rows.map(r=>`<tr><td>${r.month}</td><td>${r.count}</td><td>${fmtGBP(r.value)}</td><td>${fmtGBP(r.count?r.value/r.count:0)}</td></tr>`).join('')}</table>`;
    document.getElementById("kpiMonthlyBars").innerHTML=rows.map(r=>`<div class="kpi-month-bar"><strong>${r.month}</strong><div class="kpi-bar-track"><div class="kpi-bar-fill" style="width:${Math.round(r.value/max*100)}%"></div></div><span>${fmtGBP(r.value)}</span></div>`).join('');
}

/* ---- Deals & Rankings rendering ---- */
function populateDealsFilters(){
    const deals = getStoredDeals();
    const bdms = Array.from(new Set(deals.map(d => d.BDM).filter(Boolean))).sort();
    const loans = Array.from(new Set(deals.map(d => d.LoanType).filter(Boolean))).sort();
    const bdmSel = document.getElementById("dealsFilterBDM");
    const loanSel = document.getElementById("dealsFilterLoan");
    if(bdmSel){ const v=bdmSel.value; bdmSel.innerHTML = `<option value="all">All BDMs</option>` + bdms.map(b=>`<option value="${escapeHtml(b)}">${escapeHtml(b)}</option>`).join(""); bdmSel.value=v; }
    if(loanSel){ const v=loanSel.value; loanSel.innerHTML = `<option value="all">All Loan Types</option>` + loans.map(l=>`<option value="${escapeHtml(l)}">${escapeHtml(l)}</option>`).join(""); loanSel.value=v; }
}
function clearDealsFilters(){
    const b=document.getElementById("dealsFilterBDM"); const l=document.getElementById("dealsFilterLoan"); const p=document.getElementById("dealsFilterPeriod");
    if(b) b.value="all"; if(l) l.value="all"; if(p) p.value="all";
    renderDeals();
}
function fmtGBP(n){ return "£" + Number(n||0).toLocaleString("en-GB", {maximumFractionDigits:0}); }
function renderDeals(){
    const all = getStoredDeals();
    const bdmF = (currentUser && currentUser.Role!=="Admin")
        ? (currentUser.Title||currentUser.Email||"")
        : (document.getElementById("dealsFilterBDM") ? document.getElementById("dealsFilterBDM").value : "all");
    const loanF = document.getElementById("dealsFilterLoan") ? document.getElementById("dealsFilterLoan").value : "all";
    const periodF = document.getElementById("dealsFilterPeriod") ? document.getElementById("dealsFilterPeriod").value : "all";
    let deals = all;
    if(bdmF !== "all") deals = deals.filter(d => normaliseKPIUserName(d.BDM) === normaliseKPIUserName(bdmF));
    if(loanF !== "all") deals = deals.filter(d => d.LoanType === loanF);
    deals = filterByPeriod(deals, periodF, d => d.DealDate);
    const countEl = document.getElementById("dealsFilterCount"); if(countEl) countEl.textContent = deals.length + " deals";

    const totalValue = deals.reduce((s,d) => s + (Number(d.DealValue)||0), 0);
    const now = new Date();
    const monthDeals = deals.filter(d => { const dt = new Date(d.DealDate); return dt.getMonth()===now.getMonth() && dt.getFullYear()===now.getFullYear(); });
    const monthValue = monthDeals.reduce((s,d) => s + (Number(d.DealValue)||0), 0);

    const sum = document.getElementById("dealsSummary");
    if(sum){
        const card = (label,val) => `<div style="background:#fff;border:1px solid var(--border);padding:14px;border-radius:var(--radius);"><div style="font-size:11px;color:var(--muted);font-weight:600;">${label}</div><div style="font-size:20px;font-weight:700;color:var(--primary-dark);margin-top:4px;">${val}</div></div>`;
        sum.innerHTML = card("Total Deals", deals.length) + card("Total Value", fmtGBP(totalValue)) + card("Deals This Month", monthDeals.length) + card("Value This Month", fmtGBP(monthValue));
    }

    // Broker ranking by total value
    const byBroker = {};
    deals.forEach(d => { const k = d.Company || d.BrokerName || "(unknown)"; if(!byBroker[k]) byBroker[k] = {name:k, count:0, value:0}; byBroker[k].count++; byBroker[k].value += Number(d.DealValue)||0; });
    const brokerRanked = Object.values(byBroker).sort((a,b)=>b.value-a.value).slice(0,10);
    const brEl = document.getElementById("brokerRanking");
    if(brEl){
        if(!brokerRanked.length){ brEl.innerHTML = `<div style="font-size:13px;color:var(--muted);padding:8px;">No deals yet.</div>`; }
        else {
            brEl.innerHTML = `<table class='preview'><tr><th>#</th><th>Broker</th><th>Deals</th><th>Value</th></tr>` +
                brokerRanked.map((r,i)=>`<tr><td>${i+1}</td><td>${escapeHtml(r.name)}</td><td>${r.count}</td><td>${fmtGBP(r.value)}</td></tr>`).join("") + `</table>`;
        }
    }

    // BDM ranking
    const byBDM = {};
    deals.forEach(d => { const k = d.BDM || "(unassigned)"; if(!byBDM[k]) byBDM[k] = {name:k, count:0, value:0}; byBDM[k].count++; byBDM[k].value += Number(d.DealValue)||0; });
    const bdmRanked = Object.values(byBDM).sort((a,b)=>b.value-a.value);
    const bdmEl = document.getElementById("bdmRanking");
    if(bdmEl){
        if(!bdmRanked.length){ bdmEl.innerHTML = `<div style="font-size:13px;color:var(--muted);padding:8px;">No deals yet.</div>`; }
        else {
            bdmEl.innerHTML = `<table class='preview'><tr><th>#</th><th>BDM</th><th>Deals</th><th>Value</th></tr>` +
                bdmRanked.map((r,i)=>`<tr><td>${i+1}</td><td>${escapeHtml(r.name)}</td><td>${r.count}</td><td>${fmtGBP(r.value)}</td></tr>`).join("") + `</table>`;
        }
    }

    // Recent deals table
    const recent = deals.slice().sort((a,b)=> new Date(b.DealDate||0) - new Date(a.DealDate||0)).slice(0,50);
    const tEl = document.getElementById("dealsTable");
    if(tEl){
        if(!recent.length){ tEl.innerHTML = `<div style="padding:16px;text-align:center;color:var(--muted);font-size:13px;">No deals logged yet. Open a broker record and click “Log Deal” to add one.</div>`; }
        else {
            tEl.innerHTML = `<table class='preview'><tr><th>Date</th><th>Broker</th><th>Loan Type</th><th>Value</th><th>BDM</th></tr>` +
                recent.map(d=>`<tr><td>${escapeHtml((d.DealDate||"").slice(0,10))}</td><td>${escapeHtml(d.Company || d.BrokerName || "")}</td><td>${escapeHtml(d.LoanType||"")}</td><td>${fmtGBP(d.DealValue)}</td><td>${escapeHtml(d.BDM||"")}</td></tr>`).join("") + `</table>`;
        }
    }
}

/* ---- Dialler KPI / Ranking helpers ---- */
function filterByPeriod(items, period, dateGetter){
    if(period === "all") return items.slice();
    const now = new Date();
    const start = new Date(now); start.setHours(0,0,0,0);
    if(period === "week"){
        const day = start.getDay() || 7;
        start.setDate(start.getDate() - day + 1);
    } else if(period === "month"){
        start.setDate(1);
    } else if(period === "year"){
        start.setMonth(0,1);
    }
    return items.filter(x => { const dt = new Date(dateGetter(x)||0); return dt >= start && dt <= now; });
}
function getDialerLogs(){ return getStoredAuditLogs().filter(l => l.Trail === "auditDialer" && (/^CALL_/.test(l.Action) || l.Action === "DEAL_LOGGED")); }
function populateDialerKpiUsers(){
    const logs=getDialerLogs(); const users=[...new Set(logs.map(l=>l.User).filter(Boolean))].sort();
    ["dialerKpiUser","dialerRankUser"].forEach(id=>{ const e=document.getElementById(id); if(!e)return; const v=e.value; e.innerHTML='<option value="all">All users</option>'+users.map(u=>`<option value="${escapeHtml(u)}">${escapeHtml(u)}</option>`).join(''); e.value=users.includes(v)?v:'all'; });
}
function outcomeLabel(a){ return String(a||'').replace(/^CALL_/,'').replace('DEAL_LOGGED','LOG_DEAL').replace(/_/g,' ').toLowerCase().replace(/\b\w/g,c=>c.toUpperCase()); }
function outcomeScore(a){ return ({CALL_LOG_DEAL:5,DEAL_LOGGED:5,CALL_APPOINTMENT:4,CALL_CALL_BACK:2,CALL_CONNECTED:1,CALL_NO_ANSWER:0,CALL_VOICEMAIL:0,CALL_NOT_INTERESTED:0,CALL_NO_PRODUCT:0,CALL_OUT_OF_AREA:0,CALL_COMPLIANCE:0,CALL_DUPLICATE:0}[a] ?? 0); }
function getUnifiedDialerKpiFilters(){
    const period=document.getElementById('dashboardPeriod')?.value||'all';
    const selectedUser=document.getElementById('dashboardUser')?.value||'all';
    const user=getKPIUserScope()!=='all'?getKPIUserScope():selectedUser;
    const start=document.getElementById('dashboardStartDate')?.value||'';
    const end=document.getElementById('dashboardEndDate')?.value||'';
    return {period,user,start,end};
}
function clearDialerKPIFilters(){clearKPIDashboardFilters();}
function getDialerFilteredLogs(){
    const {period,user,start,end}=getUnifiedDialerKpiFilters();
    let logs=getDialerLogs();
    if(start||end){const a=start?new Date(start+'T00:00:00'):new Date('1900-01-01'),b=end?new Date(end+'T23:59:59'):new Date('2999-12-31');logs=logs.filter(l=>{const x=new Date(l.Timestamp||0);return x>=a&&x<=b;});}
    else logs=filterByPeriod(logs,period,l=>l.Timestamp);
    if(user!=='all')logs=logs.filter(l=>normaliseKPIUserName(l.User)===normaliseKPIUserName(user));
    return logs;
}
function renderDialerKPI(){
    const logs=getDialerFilteredLogs();
    const counts={}; logs.forEach(l=>counts[l.Action]=(counts[l.Action]||0)+1);
    const total=logs.length, positive=logs.filter(l=>['CALL_CONNECTED','CALL_APPOINTMENT','CALL_LOG_DEAL','DEAL_LOGGED','CALL_CALL_BACK'].includes(l.Action)).length, appointments=counts.CALL_APPOINTMENT||0, deals=(counts.CALL_LOG_DEAL||0)+(counts.DEAL_LOGGED||0), unsuitable=logs.filter(l=>['CALL_NOT_INTERESTED','CALL_NO_PRODUCT','CALL_OUT_OF_AREA','CALL_COMPLIANCE','CALL_DUPLICATE'].includes(l.Action)).length;
    const cards=[['Calls',total],['Positive contacts',positive],['Appointments',appointments],['Deals',deals],['Not suitable',unsuitable],['Appointment rate',total?((appointments/total)*100).toFixed(1)+'%':'0%']];
    const summary=document.getElementById('dialerKpiSummary'); if(summary)summary.innerHTML=cards.map(x=>`<div class="dialer-kpi-stat"><div>${x[0]}</div><strong>${x[1]}</strong></div>`).join('');
    const order=['CALL_NO_ANSWER','CALL_VOICEMAIL','CALL_CALL_BACK','CALL_CONNECTED','CALL_APPOINTMENT','CALL_LOG_DEAL','DEAL_LOGGED','CALL_NOT_INTERESTED','CALL_NO_PRODUCT','CALL_OUT_OF_AREA','CALL_COMPLIANCE','CALL_DUPLICATE'];
    const outcome=document.getElementById('dialerOutcomeTable'); if(outcome)outcome.innerHTML='<table class="preview"><tr><th>Outcome</th><th>Calls</th><th>%</th></tr>'+order.map(a=>`<tr><td>${outcomeLabel(a)}</td><td>${counts[a]||0}</td><td>${total?(((counts[a]||0)/total)*100).toFixed(1):'0.0'}%</td></tr>`).join('')+'</table>';
    const weeks={}; logs.forEach(l=>{const d=new Date(l.Timestamp);if(isNaN(d))return;const day=d.getDay()||7,st=new Date(d);st.setHours(0,0,0,0);st.setDate(st.getDate()-day+1);const k=st.toISOString().slice(0,10);if(!weeks[k])weeks[k]={calls:0,appointments:0,deals:0,unsuitable:0};weeks[k].calls++;if(l.Action==='CALL_APPOINTMENT')weeks[k].appointments++;if(l.Action==='CALL_LOG_DEAL'||l.Action==='DEAL_LOGGED')weeks[k].deals++;if(['CALL_NOT_INTERESTED','CALL_NO_PRODUCT','CALL_OUT_OF_AREA','CALL_COMPLIANCE','CALL_DUPLICATE'].includes(l.Action))weeks[k].unsuitable++;});
    const wr=Object.entries(weeks).sort((a,b)=>b[0].localeCompare(a[0])).slice(0,12),weekly=document.getElementById('dialerWeeklyTable'); if(weekly)weekly.innerHTML=wr.length?'<table class="preview"><tr><th>Week commencing</th><th>Calls</th><th>Appointments</th><th>Deals</th><th>Not suitable</th></tr>'+wr.map(([k,v])=>`<tr><td>${k}</td><td>${v.calls}</td><td>${v.appointments}</td><td>${v.deals}</td><td>${v.unsuitable}</td></tr>`).join('')+'</table>':'<div style="padding:15px;color:var(--muted);">No dialler activity for this period.</div>';
}
function renderDialerRanking(){
    const logs=getDialerFilteredLogs();
    const brokers={}; logs.forEach(l=>{const k=(l.RecordTitle||'')+'||'+(l.Company||'');if(!brokers[k])brokers[k]={name:l.RecordTitle||'(No name)',company:l.Company||'',calls:0,appointments:0,deals:0,callbacks:0,score:0};brokers[k].calls++;if(l.Action==='CALL_APPOINTMENT')brokers[k].appointments++;if(l.Action==='CALL_LOG_DEAL'||l.Action==='DEAL_LOGGED')brokers[k].deals++;if(l.Action==='CALL_CALL_BACK')brokers[k].callbacks++;brokers[k].score+=outcomeScore(l.Action);});
    const ranked=Object.values(brokers).sort((a,b)=>b.score-a.score||b.deals-a.deals||b.appointments-a.appointments).slice(0,25),total=logs.length,rs=document.getElementById('dialerRankingSummary');
    if(rs)rs.innerHTML=[['Calls',total],['Brokers contacted',ranked.length],['Appointments',logs.filter(l=>l.Action==='CALL_APPOINTMENT').length],['Deals',logs.filter(l=>l.Action==='CALL_LOG_DEAL'||l.Action==='DEAL_LOGGED').length]].map(x=>`<div class="dialer-kpi-stat"><div>${x[0]}</div><strong>${x[1]}</strong></div>`).join('');
    const table=document.getElementById('dialerBrokerRanking');if(table)table.innerHTML=ranked.length?'<table class="preview"><tr><th>#</th><th>Broker</th><th>Company</th><th>Calls</th><th>Callbacks</th><th>Appointments</th><th>Deals</th><th>Score</th></tr>'+ranked.map((r,i)=>`<tr><td>${i+1}</td><td>${escapeHtml(r.name)}</td><td>${escapeHtml(r.company)}</td><td>${r.calls}</td><td>${r.callbacks}</td><td>${r.appointments}</td><td>${r.deals}</td><td><strong>${r.score}</strong></td></tr>`).join('')+'</table>':'<div style="padding:15px;color:var(--muted);">No dialler activity for this period.</div>';
}



/* ---- Power Dialer / Call Queue ---- */
let dialerQueue = [];
let dialerIndex = 0;

function dialerSanitizePhone(p){
    if(!p) return "";
    // Preserve a leading +, strip spaces, brackets, dashes and dots
    let s = String(p).trim();
    const hasPlus = s.charAt(0) === "+";
    s = s.replace(/[\s()\-.]/g, "");
    if(hasPlus) s = "+" + s;
    return s;
}
function populateDialerLoanFilter(){
    const sel = document.getElementById("dialerFilterLoan");
    if(!sel) return;
    const cur = sel.value;
    const loans = Array.from(new Set(getStoredBrokers().flatMap(b => getLoanTypes(b)))).sort();
    sel.innerHTML = `<option value="all">All Loan Types</option>` + loans.map(l=>`<option value="${escapeHtml(l)}">${escapeHtml(l)}</option>`).join("");
    sel.value = cur || "all";
}
function buildDialerQueue(){
    const scope = document.getElementById("dialerFilterScope") ? document.getElementById("dialerFilterScope").value : "open";
    const loanF = document.getElementById("dialerFilterLoan") ? document.getElementById("dialerFilterLoan").value : "all";
    let brokers = getStoredBrokers().slice();
    // Open = not marked not-suitable. "mine" = assigned to current user.
    brokers = brokers.filter(b => !isNotSuitable(b));
    if(scope === "mine"){
        const me = currentUser ? (currentUser.Title||"").toLowerCase() : "";
        brokers = brokers.filter(b => b.AssignedTo && (b.AssignedTo.Title||"").toLowerCase() === me);
    } else if(scope === "notcontacted"){
        const cutoff = Date.now() - (7*24*60*60*1000);
        brokers = brokers.filter(b => !b.LastContactDate || new Date(b.LastContactDate).getTime() < cutoff);
    }
    if(loanF !== "all") brokers = brokers.filter(b => getLoanTypes(b).includes(loanF));
    const withPhone = brokers.filter(b => dialerSanitizePhone(b.Phone));
    const missingCount = brokers.length - withPhone.length;
    // Least recently contacted first; never-contacted first of all
    withPhone.sort((a,b) => {
        const ad = a.LastContactDate ? new Date(a.LastContactDate).getTime() : 0;
        const bd = b.LastContactDate ? new Date(b.LastContactDate).getTime() : 0;
        return ad - bd;
    });
    dialerQueue = withPhone;
    dialerIndex = 0;
    const prog = document.getElementById("dialerProgress");
    if(prog) prog.textContent = dialerQueue.length + " callable" + (missingCount ? " · " + missingCount + " missing number" : "");
    renderDialer();
}
function renderDialer(){
    const card=document.getElementById("dialerCard"); if(!card)return;
    if(dialerQueue.length===0){ card.innerHTML=`<div class="dialer-modern-card" style="text-align:center;color:var(--muted)"><div style="font-size:32px;margin-bottom:8px">☎️</div><p style="font-weight:700;color:var(--primary-dark);margin-bottom:5px">No brokers in the queue</p><p style="font-size:13px">Try changing the scope/filter, or add brokers with phone numbers first.</p></div>`; return; }
    const item=dialerQueue[dialerIndex]; if(!item){dialerIndex=0;renderDialer();return;} selectedItemId=item.Id;
    const phone=dialerSanitizePhone(item.Phone), loans=getLoanTypes(item), lastContact=item.LastContactDate?new Date(item.LastContactDate).toLocaleDateString("en-GB"):"Never";
    const prog=document.getElementById("dialerProgress"); if(prog)prog.textContent=(dialerIndex+1)+" of "+dialerQueue.length+" callable";
    card.innerHTML=`<div class="dialer-modern-card">
      <div class="dialer-modern-head"><div><div class="dialer-company">${escapeHtml(item.Company||item.Title||"(no name)")}</div><div class="dialer-meta">${escapeHtml(item.Title||"")} · ${escapeHtml(item.City||"")||"No city"} · ${escapeHtml(item.Network||"")||"No network"}</div></div><button type="button" class="btn secondary" onclick="openDetail(getStoredBrokers().find(b=>b.Id===${item.Id}))">Open full record</button></div>
      <div class="dialer-phone"><div style="flex:1;min-width:170px"><div class="dialer-section-title" style="margin-bottom:3px">Phone</div><div class="dialer-phone-number">${escapeHtml(phone)}</div></div><a id="dialerCallLink" href="tel:${escapeHtml(phone)}" class="btn" style="font-size:14px;padding:11px 18px;text-decoration:none">📞 Dial with Nuacom</a><button type="button" class="btn secondary" onclick="dialerCopyNumber('${escapeHtml(phone)}')">Copy</button></div>
      <div class="dialer-section"><div class="dialer-section-title">Broker details</div><div style="display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:9px 18px;font-size:13px"><div><strong>Email</strong><br>${escapeHtml(item.Email||"")||"—"}</div><div><strong>Website</strong><br>${item.Website?`<a href="${escapeHtml(item.Website)}" target="_blank" rel="noopener">${escapeHtml(item.Website)}</a>`:"—"}</div><div><strong>Loan types</strong><br>${loans.length?escapeHtml(loans.join(", ")):"—"}</div><div><strong>Volume</strong><br>${escapeHtml(item.Volume||"")||"—"}</div><div><strong>Status</strong><br>${escapeHtml(item.Status||"")}</div><div><strong>Last contact</strong><br>${lastContact}</div></div></div>
      <div class="dialer-section"><div class="dialer-section-title">Notes</div><div style="font-size:13px;color:var(--muted)">${escapeHtml((item.Notes||"").slice(0,220))||"No notes yet."}</div></div>
      <div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:12px"><button type="button" class="btn secondary" onclick="dialerPrev()" ${dialerIndex===0?"disabled":""}>◀ Previous</button><button type="button" class="btn secondary" onclick="dialerSkip()">Skip ▶</button></div>
      <div class="dialer-section"><div class="dialer-section-title">Contact outcome</div><div class="dialer-outcomes"><button type="button" class="btn secondary dialer-outcome-btn" data-dialer-outcome="NO_ANSWER">No Answer</button><button type="button" class="btn secondary dialer-outcome-btn" data-dialer-outcome="VOICEMAIL">Voicemail</button><button type="button" class="btn secondary dialer-outcome-btn" data-dialer-outcome="CALL_BACK">Callback</button><button type="button" class="btn secondary dialer-outcome-btn" data-dialer-outcome="CONNECTED">Connected</button></div></div>
      <div class="dialer-section"><div class="dialer-section-title">Positive</div><div class="dialer-outcomes"><button type="button" class="btn dialer-outcome-btn" style="background:#eefaf2;color:#1a6b3a;border-color:#bfe6cd" data-dialer-outcome="APPOINTMENT">Appointment</button><button type="button" class="btn dialer-outcome-btn" style="background:#eefaf2;color:#1a6b3a;border-color:#bfe6cd" data-dialer-outcome="LOG_DEAL">Log Deal £</button></div></div>
      <div class="dialer-section"><div class="dialer-section-title" style="color:var(--red)">Not suitable — removed from dialler</div><div class="dialer-outcomes"><button type="button" class="btn danger dialer-outcome-btn" data-dialer-outcome="NOT_INTERESTED">Not Interested</button><button type="button" class="btn danger dialer-outcome-btn" data-dialer-outcome="NO_PRODUCT">No Suitable Product</button><button type="button" class="btn danger dialer-outcome-btn" data-dialer-outcome="OUT_OF_AREA">Out of Area</button><button type="button" class="btn danger dialer-outcome-btn" data-dialer-outcome="COMPLIANCE">Compliance / Unregulated</button><button type="button" class="btn danger dialer-outcome-btn" data-dialer-outcome="DUPLICATE">Duplicate / Inactive</button></div></div>
    </div>`;

    // Bind dialler outcome taps directly to the freshly-rendered buttons.
    // This mirrors the robust Admin delete-button fix and avoids relying on
    // inline onclick handlers in embedded/mobile file viewers.
    card.querySelectorAll('.dialer-outcome-btn').forEach(btn => {
        const outcome = btn.getAttribute('data-dialer-outcome');
        btn.onclick = async function(ev){
            ev.preventDefault();
            ev.stopPropagation();
            if(!outcome || btn.disabled) return;
            btn.disabled = true;
            const original = btn.textContent;
            btn.textContent = 'Saving…';
            try {
                await dialerOutcome(outcome);
            } catch(e) {
                console.error('Dialler outcome failed:', e);
                alert(e && e.message ? e.message : 'Could not save dialler outcome.');
                btn.disabled = false;
                btn.textContent = original;
            }
        };
    });
}

function dialerCopyNumber(phone){
    try{ navigator.clipboard.writeText(phone); }catch(e){}
}
function dialerPrev(){
    if(dialerIndex > 0){ dialerIndex--; renderDialer(); }
}
function dialerSkip(){
    if(dialerIndex < dialerQueue.length - 1){ dialerIndex++; renderDialer(); }
}
function dialerDial(){
    // Trigger the visible tel: link (keeps it within the user gesture)
    const link = document.getElementById("dialerCallLink");
    if(link) link.click();
}
async function dialerOutcome(outcome){
    if(outcome === "LOG_DEAL"){openLogDealModal();return;}
    const item=dialerQueue[dialerIndex]; if(!item)return;
    if(outcome === "CALL_BACK"){
        calendarCallbackContext={
            brokerId:item.Id,
            subject:"Callback — "+(item.Company||item.Title||"Broker"),
            notes:"Callback scheduled from Power Dialer.",
            dialerRecordId:item.Id
        };
    }
    const unsuitable=["NOT_INTERESTED","NO_PRODUCT","OUT_OF_AREA","COMPLIANCE","DUPLICATE"];
    const reason=outcome.replace(/_/g," ").toLowerCase().replace(/\b\w/g,c=>c.toUpperCase());
    try{
        if(m365Configured()&&!isDemoMode()){
            const current=await cloudGetCurrentBroker(item.Id); if(!current)throw new Error("Broker is no longer available.");
            const now=new Date().toISOString(); const patch={DiallerOutcome:outcome,DiallerOutcomeReason:reason,DiallerOutcomeDate:now,LastContactDate:now,Modified:now};
            if(unsuitable.includes(outcome)){Object.assign(patch,{WorkflowState:"NOT_SUITABLE",IsNotSuitable:"Yes",NotSuitableReason:reason,NotSuitableSource:"Dialler",AssignedTo:"",AssignedToEmail:"",AssignedToName:"",AssignedToLookupId:null,ClaimedAt:"",ClaimedBy:"",ClaimExpiresAt:""});}
            await graphPatch(`${GRAPH_BASE}/sites/${await getSiteId()}/lists/${await resolveListId(getListName('brokers'))}/items/${current.spId}/fields`,filterBrokerFieldsForColumns(patch,await getBrokerColumnNames()),current.etag||undefined);
            const fresh=hydrateBrokerOutcomeMetadata(await cloudLoadBrokers(),getStoredAuditLogs());localStorage.setItem(STORAGE_KEY,JSON.stringify(fresh));_brokerSnapshot=_snapshotBrokers(fresh);allItems=fresh;const fb=fresh.find(b=>String(b.Id)===String(item.Id))||item;logAuditRecord(fb,"CALL_"+outcome,"Dialler outcome: "+reason);
        }else{
            const brokers=getStoredBrokers(),idx=brokers.findIndex(b=>String(b.Id)===String(item.Id)); if(idx!==-1){const b=brokers[idx],now=new Date().toISOString();b.DiallerOutcome=outcome;b.DiallerOutcomeReason=reason;b.DiallerOutcomeDate=now;b.LastContactDate=now;b.Modified=now;if(unsuitable.includes(outcome)){b.WorkflowState="NOT_SUITABLE";b.IsNotSuitable=true;b.NotSuitableReason=reason;b.NotSuitableSource="Dialler";b.AssignedTo=null;}saveStoredBrokers(brokers);logAuditRecord(b,"CALL_"+outcome,"Dialler outcome: "+reason);}
        }
        buildDialerQueue();

        // Refresh dialler KPI data immediately after the outcome is logged.
        // getDialerLogs() reads the same audit store written above, so the
        // KPI cards/table/ranking reflect the new disposition without a
        // page refresh or leaving/re-entering the KPI tab.
        try{
            if(typeof populateDialerKpiUsers === "function") populateDialerKpiUsers();
            if(typeof renderDialerKPI === "function") renderDialerKPI();
            if(typeof renderDialerRanking === "function") renderDialerRanking();
        }catch(kpiErr){
            console.warn("Dialler KPI refresh failed:", kpiErr);
        }

        if(outcome === "CALL_BACK"){
            openCalendarEventModal(item.Id);
        }
    }catch(e){
        calendarCallbackContext=null;
        alert(e.message||"Could not save dialler outcome.");
        buildDialerQueue();
    }
}

const dropZone = document.getElementById("dropZone");
const fileInput = document.getElementById("fileInput");

["dragenter","dragover"].forEach(evt =>
    dropZone.addEventListener(evt, e => { e.preventDefault(); dropZone.classList.add("drag"); }));
["dragleave","drop"].forEach(evt =>
    dropZone.addEventListener(evt, e => { e.preventDefault(); dropZone.classList.remove("drag"); }));
dropZone.addEventListener("drop", e => {
    if(e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener("change", e => {
    if(e.target.files.length) handleFile(e.target.files[0]);
});

function handleFile(file){
    const reader = new FileReader();
    reader.onload = evt => {
        const wb = XLSX.read(evt.target.result, { type: "binary" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        parsedRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        if(parsedRows.length === 0){
            document.getElementById("uploadStatus").textContent = "No rows found in that file.";
            return;
        }
        buildMappingUI(Object.keys(parsedRows[0]));
        buildPreview();
    };
    reader.readAsBinaryString(file);
}

function buildMappingUI(sourceColumns){
    const targets = [
        ["name","Broker Name"], ["company","Company Name"], ["phone","Contact Number"],
        ["email","Email"], ["website","Website"], ["address","Address"], ["city","City"],
        ["notes","Notes"], ["prefComm","Preferred Communication"], ["loanTypes","Loan Types Handled"],
        ["volume","Estimated Volume"], ["network","Network / Affiliation"], ["status","BDM Status"],
        ["nextFollowUp","Next Follow-Up Date"]
    ];
    const area = document.getElementById("mappingArea");
    area.style.display = "block";
    area.innerHTML = "<h3 style='font-size:14px;'>Map columns</h3>";
    targets.forEach(([key,label]) => {
        const row = document.createElement("div");
        row.className = "mapping-row";
        const options = sourceColumns.map(c =>
            `<option value="${escapeHtml(c)}" ${c.toLowerCase().includes(key.toLowerCase())?"selected":""}>${escapeHtml(c)}</option>`
        ).join("");
        row.innerHTML = `<label>${label}</label><select data-target="${key}"><option value="">— skip —</option>${options}</select>`;
        area.appendChild(row);
        columnMapping[key] = row.querySelector("select").value;
        row.querySelector("select").addEventListener("change", e => {
            columnMapping[key] = e.target.value;
        });
    });
    document.getElementById("submitUploadBtn").style.display = "inline-block";
}

function buildPreview(){
    const area = document.getElementById("previewArea");
    const cols = Object.keys(parsedRows[0]);
    let html = "<table class='preview'><tr>" + cols.map(c => `<th>${escapeHtml(c)}</th>`).join("") + "</tr>";
    parsedRows.slice(0, 5).forEach(row => {
        html += "<tr>" + cols.map(c => `<td>${escapeHtml(row[c])}</td>`).join("") + "</tr>";
    });
    html += "</table>";
    if(parsedRows.length > 5){
        html += `<p style="font-size:12px;color:var(--muted);">Showing 5 of ${parsedRows.length} rows.</p>`;
    }
    area.innerHTML = html;
}

document.getElementById("submitUploadBtn").addEventListener("click", async () => {
    const status = document.getElementById("uploadStatus");
    const btn = document.getElementById("submitUploadBtn");
    if(btn.disabled) return;

    // Build the import plan first. Nothing is written to local storage or SharePoint
    // until the user explicitly confirms the validated batch.
    let existingBrokers = getStoredBrokers();
    const emailKeys = new Set(existingBrokers.map(b => String(b.Email||"").trim().toLowerCase()).filter(Boolean));
    const phoneKeys = new Set(existingBrokers.map(b => dialerSanitizePhone(b.Phone||"")).filter(Boolean));
    const rowKeys = new Set();
    const pending = [];
    let duplicatesSkipped = 0, blankSkipped = 0;

    parsedRows.forEach(row => {
        const email = columnMapping.email ? String(row[columnMapping.email] ?? "").trim().toLowerCase() : "";
        const phone = columnMapping.phone ? dialerSanitizePhone(String(row[columnMapping.phone] ?? "")) : "";
        const name = columnMapping.name ? String(row[columnMapping.name] ?? "").trim() : "";
        const company = columnMapping.company ? String(row[columnMapping.company] ?? "").trim() : "";
        if(!name && !company && !email && !phone){ blankSkipped++; return; }

        const composite = [email, phone, name.toLowerCase(), company.toLowerCase()].join("|");
        const duplicate = (email && emailKeys.has(email)) || (phone && phoneKeys.has(phone)) || rowKeys.has(composite);
        if(duplicate){ duplicatesSkipped++; return; }

        const rawLoanTypes = columnMapping.loanTypes ? String(row[columnMapping.loanTypes] ?? "") : "Residential Bridging";
        pending.push(buildLocalBroker({
            Title:name, Company:company,
            Phone:columnMapping.phone ? String(row[columnMapping.phone] ?? "") : "",
            Email:columnMapping.email ? String(row[columnMapping.email] ?? "") : "",
            Website:columnMapping.website ? String(row[columnMapping.website] ?? "") : "",
            Address:columnMapping.address ? String(row[columnMapping.address] ?? "") : "",
            City:columnMapping.city ? String(row[columnMapping.city] ?? "") : "",
            Notes:columnMapping.notes ? String(row[columnMapping.notes] ?? "") : "",
            PrefComm:columnMapping.prefComm ? String(row[columnMapping.prefComm] ?? "Phone") : "Phone",
            LoanTypes:rawLoanTypes,
            Volume:columnMapping.volume ? String(row[columnMapping.volume] ?? "Under £1M") : "Under £1M",
            Network:columnMapping.network ? String(row[columnMapping.network] ?? "") : "",
            Status:columnMapping.status ? String(row[columnMapping.status] ?? "Cold") : "Cold",
            NextFollowUp:columnMapping.nextFollowUp ? String(row[columnMapping.nextFollowUp] ?? "") : ""
        }));
        if(email) emailKeys.add(email);
        if(phone) phoneKeys.add(phone);
        rowKeys.add(composite);
    });

    if(!pending.length){
        status.textContent = `Nothing to import. Skipped ${duplicatesSkipped} duplicate row(s)${blankSkipped ? ` and ${blankSkipped} blank row(s)` : ""}.`;
        return;
    }

    const cloudMode = m365Configured() && !isDemoMode();
    const confirmed = window.confirm(
        `Import ${pending.length} new broker record(s)?\n\n` +
        `Duplicates skipped: ${duplicatesSkipped}\n` +
        `Blank rows skipped: ${blankSkipped}\n\n` +
        (cloudMode
            ? "The records will be saved locally first, then uploaded to Microsoft Lists once. They will NOT be deleted automatically."
            : "The records will be saved locally only in Demo/Local mode.")
    );
    if(!confirmed) return;

    btn.disabled = true;
    btn.textContent = "Importing…";
    status.textContent = `Preparing ${pending.length} broker record(s)…`;

    try{
        // Commit locally once. The cloud write is deliberately separate so a failed
        // Graph request cannot trigger a hidden retry or a second local insert.
        existingBrokers = existingBrokers.concat(pending);
        allItems = existingBrokers;
        saveStoredBrokers(existingBrokers, false);
        pending.forEach(item => logAuditRecord(item, "CREATE", "Created new broker record via Excel import"));

        if(cloudMode){
            status.textContent = `Uploading ${pending.length} broker record(s) to Microsoft Lists…`;
            await queueBrokerCloudSync(existingBrokers);

            // Verify every imported broker received a SharePoint item id. Do not
            // retry failed records here: retrying a POST blindly is exactly what
            // caused the duplicate-record problem this importer is designed to avoid.
            const map = getBrokerIdMap();
            const missing = pending.filter(item => !map[item.Id]);
            if(missing.length){
                throw new Error(`${missing.length} imported broker record(s) were saved locally but did not receive a Microsoft Lists item ID. No automatic retry was performed.`);
            }
            status.textContent = `Import complete: ${pending.length} broker record(s) created and verified in Microsoft Lists. Skipped ${duplicatesSkipped} duplicate row(s)${blankSkipped ? ` and ${blankSkipped} blank row(s)` : ""}.`;
        }else{
            status.textContent = `Import complete: ${pending.length} broker record(s) saved locally. Skipped ${duplicatesSkipped} duplicate row(s)${blankSkipped ? ` and ${blankSkipped} blank row(s)` : ""}.`;
        }
        reload();
    }catch(e){
        console.error("Excel broker import failed:", e);
        status.textContent = `Import saved locally, but Microsoft Lists sync was not fully completed: ${e.message||e}. No automatic retry or deletion was performed.`;
        reload();
    }finally{
        btn.disabled = false;
        btn.textContent = "Create broker records";
    }
});

function downloadListTemplate(){
    if(typeof XLSX === "undefined"){ alert("Excel support is not loaded. Please reload the portal and try again."); return; }
    const wb=XLSX.utils.book_new();
    const sheets={
        "BDM Brokers":["Title","Company","Phone","Email","Website","Address","City","Notes","PrefComm","LoanTypes","Volume","Network","Status","NextFollowUp","LastContactDate","AssignedTo","IsNotSuitable","PortalId"],
        "BDM Users":["Title","Email","Role","PortalId"],
        "BDM Audit - Account":["Title","Timestamp","User","RecordTitle","Company","Reason","RecordId","Trail"],
        "BDM Audit - Dialler":["Title","Timestamp","User","RecordTitle","Company","Reason","RecordId","Trail"],
        "BDM Deals":["Title","Company","BrokerName","BrokerId","DealValue","DealDate","LoanType","BDM","Notes","PortalId"],
        "BDM Global Settings":["Title","SettingKey","SettingValue","Updated","UpdatedBy"]
    };
    Object.entries(sheets).forEach(([name,headers])=>{const ws=XLSX.utils.aoa_to_sheet([headers]);ws['!cols']=headers.map(h=>({wch:Math.max(12,Math.min(28,h.length+3))}));XLSX.utils.book_append_sheet(wb,ws,name);});
    XLSX.writeFile(wb,`7J_Microsoft_Lists_Template_${new Date().toISOString().slice(0,10)}.xlsx`);
}
const templateBtn=document.getElementById("downloadListTemplateBtn");
if(templateBtn) templateBtn.addEventListener("click",downloadListTemplate);

document.getElementById("exportExcelBtn").addEventListener("click", () => {
    const brokersData = getStoredBrokers().map(b => ({
        ID: b.Id,
        "Contact Name": b.Title,
        "Company Name": b.Company,
        "Phone": b.Phone,
        "Email": b.Email,
        "Website": b.Website,
        "City": b.City,
        "Address": b.Address,
        "Loan Types": Array.isArray(b.LoanTypes) ? b.LoanTypes.join(", ") : (b.LoanType || ""),
        "Estimated Volume": b.Volume,
        "Network": b.Network,
        "BDM Status": b.Status,
        "Preferred Comm": b.PrefComm,
        "Assigned To": b.AssignedTo ? b.AssignedTo.Title : "Unassigned",
        "Is Not Suitable": b.IsNotSuitable ? "Yes" : "No",
        "Not Suitable Reason": b.NotSuitableReason || "",
        "Last Contact Date": b.LastContactDate,
        "Notes": b.Notes
    }));

    const auditData = getStoredAuditLogs().map(l => ({
        Timestamp: l.Timestamp,
        User: l.User,
        Action: l.Action,
        "Record ID": l.RecordId,
        "Record Name": l.RecordTitle,
        Company: l.Company,
        "Reason / Detail": l.Reason
    }));

    const workbook = XLSX.utils.book_new();
    const brokerSheet = XLSX.utils.json_to_sheet(brokersData);
    const auditSheet = XLSX.utils.json_to_sheet(auditData);

    XLSX.utils.book_append_sheet(workbook, brokerSheet, "Brokers & Pipeline");
    XLSX.utils.book_append_sheet(workbook, auditSheet, "Audit Trail & Changes");

    XLSX.writeFile(workbook, `7J_Full_System_Export_${new Date().toISOString().slice(0,10)}.xlsx`);
});

document.getElementById("exportJsonBtn").addEventListener("click", () => {
    const data = {version:"V13",backupDate:new Date().toISOString(),brokers:getStoredBrokers(),users:getStoredUsers(),auditLogs:getStoredAuditLogs(),deals:getStoredDeals(),performanceReviews:getStoredPerformanceReviews(),kpiSnapshots:getStoredKpiSnapshots(),callGuide:getStoredCallGuide(),listSources:getListOverrides(),globalSettings:GLOBAL_SETTINGS_CACHE||{}};
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `7J_Full_System_Backup_${new Date().toISOString().slice(0,10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
});

document.getElementById("restoreExcelInput").addEventListener("change", e => {
    const file = e.target.files[0];
    if(!file) return;
    const reader = new FileReader();
    reader.onload = evt => {
        try {
            const wb = XLSX.read(evt.target.result, { type: "binary" });
            const targetSheetName = wb.SheetNames.includes("Brokers & Pipeline") ? "Brokers & Pipeline" : wb.SheetNames[0];
            const sheet = wb.Sheets[targetSheetName];
            const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
            
            if(rows.length > 0) {
                const restoredItems = rows.map((row, idx) => {
                    let assignedObj = null;
                    const assignedVal = row["AssignedTo"] || row["Assigned To"] || "";
                    if(assignedVal && assignedVal !== "Unassigned") {
                        assignedObj = { Title: assignedVal, EMail: "" };
                    }

                    const rawLoans = row["Loan Types"] || row["LoanTypes"] || row["Primary Loan Type"] || "Residential Bridging";
                    const loanTypesArr = typeof rawLoans === "string" ? rawLoans.split(",").map(s => s.trim()).filter(Boolean) : ["Residential Bridging"];

                    return {
                        Id: Number(row.ID || row.Id) || (Date.now() + idx),
                        Modified: row.Modified || new Date().toISOString(),
                        Title: row.Title || row["Contact Name"] || "",
                        Company: row.Company || row["Company Name"] || "",
                        Phone: row.Phone || row["Contact Number"] || "",
                        Email: row.Email || "",
                        Website: row.Website || "",
                        Address: row.Address || "",
                        City: row.City || "",
                        Notes: row.Notes || "",
                        PrefComm: row.PrefComm || row["Preferred Comm"] || "Phone",
                        LoanTypes: loanTypesArr,
                        Volume: row.Volume || row["Estimated Volume"] || "Under £1M",
                        Network: row.Network || "",
                        Status: row.Status || row["BDM Status"] || "Cold",
                        NextFollowUp: row.NextFollowUp || "",
                        LastContactDate: row.LastContactDate || row["Last Contact Date"] || new Date().toISOString(),
                        AssignedTo: assignedObj,
                        IsNotSuitable: row.IsNotSuitable === true || row.IsNotSuitable === "true" || row["Is Not Suitable"] === "Yes",
                        NotSuitableReason: row.NotSuitableReason || row["Not Suitable Reason"] || ""
                    };
                });

                saveStoredBrokers(restoredItems);
                document.getElementById("backupStatusMsg").textContent = `Successfully restored ${restoredItems.length} records from Excel backup.`;
                reload();
            } else {
                alert("The Excel file appears to be empty.");
            }
        } catch(err) {
            alert("Error parsing Excel backup file.");
        }
    };
    reader.readAsBinaryString(file);
});

document.getElementById("restoreFileInput").addEventListener("change", e => {
    const file = e.target.files[0];
    if(!file) return;
    const reader = new FileReader();
    reader.onload = evt => {
        try {
            const parsed = JSON.parse(evt.target.result);
            if(Array.isArray(parsed)) {
                saveStoredBrokers(parsed);
                document.getElementById("backupStatusMsg").textContent = `Successfully restored ${parsed.length} records from JSON backup.`;
                reload();
            } else if(parsed && Array.isArray(parsed.brokers)) {
                saveStoredBrokers(parsed.brokers);
                if(Array.isArray(parsed.auditLogs)) localStorage.setItem(AUDIT_STORAGE_KEY,JSON.stringify(parsed.auditLogs));
                if(Array.isArray(parsed.users)) localStorage.setItem(USERS_STORAGE_KEY,JSON.stringify(parsed.users));
                if(Array.isArray(parsed.deals)) saveStoredDeals(parsed.deals);
                if(Array.isArray(parsed.performanceReviews)) saveStoredPerformanceReviews(parsed.performanceReviews);
                if(Array.isArray(parsed.kpiSnapshots)) saveStoredKpiSnapshots(parsed.kpiSnapshots);
                if(Array.isArray(parsed.callGuide)) saveStoredCallGuide(parsed.callGuide);
                if(parsed.listSources&&typeof parsed.listSources==="object") saveListOverrides(parsed.listSources);
                document.getElementById("backupStatusMsg").textContent = `Successfully restored full system backup.`;
                reload();
            } else {
                alert("Invalid backup file format.");
            }
        } catch(err) {
            alert("Error parsing backup JSON file.");
        }
    };
    reader.readAsText(file);
});

function populateAdminUserFilterOptions() {
    const userSelect = document.getElementById("adminUserFilterSelect");
    if(!userSelect) return;
    const currentVal = userSelect.value;
    let users = getStoredUsers();
    if(isDemoMode()){
        const demoRoster=[
            {Id:1,Title:"Tim (Admin Demo)",Email:"admin.demo@local",Role:"Admin"},
            {Id:2,Title:"Sarah (Demo BDM)",Email:"sarah.demo@local",Role:"BDM"},
            {Id:3,Title:"James (Demo BDM)",Email:"james.demo@local",Role:"BDM"}
        ];
        const byTitle=new Map(users.map(u=>[String(u.Title||'').toLowerCase(),u]));
        demoRoster.forEach(u=>byTitle.set(u.Title.toLowerCase(),u));
        users=Array.from(byTitle.values());
    }
    
    let html = `<option value="all">Filter Pipeline by User (All Users)</option>`;
    users.forEach(u => {
        html += `<option value="${escapeHtml(u.Title)}">${escapeHtml(u.Title)} (${escapeHtml(u.Email)})</option>`;
    });
    userSelect.innerHTML = html;
    userSelect.value = currentVal;
}

async function runAIBrokerSearch(){
    const q=document.getElementById('aiBrokerSearchInput')?.value.trim(); const st=document.getElementById('aiBrokerSearchStatus'); const out=document.getElementById('aiBrokerSearchResults');
    if(!q){if(st)st.textContent='Enter a search description first.';return;}
    const threshold=Math.min(25,Math.max(1,parseInt(getGlobalSettingValue('aiThreshold',25),10)||25));
    const proxy=getGlobalSettingValue('aiProxy',''); const key=getGlobalSettingValue('aiKey','');
    if(!proxy && !key){if(st)st.textContent='AI search is enabled but no Azure proxy endpoint has been configured. An API key is supported for legacy/demo testing only.';return;}
    if(st)st.textContent=`Searching for up to ${threshold} new broker prospects…`; if(out)out.innerHTML='';
    try{
        let results=[];
        if(proxy){
            const response=await fetch(proxy,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({query:q,maxResults:threshold})});
            if(!response.ok) throw new Error('AI proxy request failed ('+response.status+')');
            const data=await response.json(); results=Array.isArray(data)?data:(data.results||[]);
        } else {
            const model=getGlobalSettingValue('aiModel','gpt-5.6-mini');
            const response=await fetch('https://api.openai.com/v1/responses',{method:'POST',headers:{'Content-Type':'application/json','Authorization':'Bearer '+key},body:JSON.stringify({model,input:`Find at least ${threshold} and up to ${threshold} new UK mortgage/finance broker prospects matching this request: ${q}. Use web search. Return JSON only as an array with name, company, phone, email, website, city, reason. Do not invent contact details; leave unknown fields blank. Prefer distinct real businesses and avoid duplicates.`,tools:[{type:'web_search_preview'}]})});
            if(!response.ok) throw new Error('AI search request failed ('+response.status+')');
            const data=await response.json(); let text=data.output_text||''; try{results=JSON.parse(text.match(/\[[\s\S]*\]/)?.[0]||'[]');}catch(e){}
        }
        results=Array.isArray(results)?results.slice(0,threshold):[];
        if(!results.length){out.innerHTML=`<div class="ai-result"><div>No broker prospects were returned.</div></div>`;if(st)st.textContent='Search complete — 0 prospects.';return;}
        results.forEach((r,idx)=>{const el=document.createElement('div');el.className='ai-result';el.innerHTML=`<div><strong>${escapeHtml(r.name||r.company||'Unnamed prospect')}</strong><small>${escapeHtml(r.company||'')} ${r.city?' · '+escapeHtml(r.city):''}</small><small>${escapeHtml(r.reason||'')}</small></div><button type="button" class="btn" data-ai-add="${idx}">Add</button>`;el.querySelector('button').addEventListener('click',()=>addAIProspect(r));out.appendChild(el);});
        if(st)st.textContent=results.length+' prospect'+(results.length===1?'':'s')+' found.';
    }catch(e){if(st)st.textContent='AI search error: '+(e.message||e);}
}
function addAIProspect(r){
    const items=getStoredBrokers(); const email=(r.email||'').toLowerCase().trim(); const phone=(r.phone||'').replace(/\D/g,'');
    const duplicate=items.some(b=>(email&&String(b.Email||'').toLowerCase()===email)||(phone&&String(b.Phone||'').replace(/\D/g,'')===phone)||(r.name&&r.company&&String(b.Title||'').toLowerCase()===String(r.name).toLowerCase()&&String(b.Company||'').toLowerCase()===String(r.company).toLowerCase()));
    if(duplicate){alert('That broker already exists in the portal.');return;}
    const body={Title:r.name||'',Company:r.company||'',Phone:r.phone||'',Email:r.email||'',Website:r.website||'',City:r.city||'',Notes:r.reason?`AI prospect source: ${r.reason}`:'AI prospect',Status:'Cold',AssignedTo:null,IsNotSuitable:false,LoanTypes:['Residential Bridging']};
    createItemLocal(body); reload();
}

let adminPipelinePage = 1;

function renderAdmin(){
    loadBrokerReleaseDaysSetting();
    updateAIBrokerSearchVisibility();
    const query = document.getElementById("adminSearchBox").value.trim().toLowerCase();
    const filterVal = document.getElementById("adminFilterSelect") ? document.getElementById("adminFilterSelect").value : "all";
    const userFilterVal = document.getElementById("adminUserFilterSelect") ? document.getElementById("adminUserFilterSelect").value : "all";
    const aStatus  = document.getElementById("adminFilterStatus")    ? document.getElementById("adminFilterStatus").value    : "all";
    const aLoan    = document.getElementById("adminFilterLoanType")   ? document.getElementById("adminFilterLoanType").value   : "all";
    const aNetwork = document.getElementById("adminFilterNetwork")   ? document.getElementById("adminFilterNetwork").value   : "all";
    const aVolume  = document.getElementById("adminFilterVolume")     ? document.getElementById("adminFilterVolume").value     : "all";
    const aCity    = document.getElementById("adminFilterCity")       ? document.getElementById("adminFilterCity").value       : "all";
    const aSort    = document.getElementById("adminFilterSort")       ? document.getElementById("adminFilterSort").value       : "recent";
    
    let items = allItems;

    if(filterVal === "assigned") {
        items = items.filter(i => i.AssignedTo);
    } else if(filterVal === "unassigned") {
        items = items.filter(i => !i.AssignedTo);
    }

    if(userFilterVal && userFilterVal !== "all") {
        items = items.filter(i => i.AssignedTo && i.AssignedTo.Title.toLowerCase() === userFilterVal.toLowerCase());
    }
    if(aStatus !== "all")   items = items.filter(i => (i.Status||"") === aStatus);
    if(aLoan !== "all")     items = items.filter(i => getLoanTypes(i).includes(aLoan));
    if(aNetwork !== "all")  items = items.filter(i => (i.Network||"") === aNetwork);
    if(aVolume !== "all")  items = items.filter(i => (i.Volume||"") === aVolume);
    if(aCity !== "all")     items = items.filter(i => (i.City||"") === aCity);

    if(query){
        items = items.filter(i => {
            const hay = [i.Title, i.Company, i.Email, i.Status, i.City, i.Network].join(" ").toLowerCase();
            return hay.includes(query);
        });
    }

    items = sortItems(items, aSort);

    const totalFiltered = items.length;
    const pageSizeEl = document.getElementById("adminPageSize");
    const pageSize = Math.max(1, parseInt(pageSizeEl?.value || "20", 10) || 20);
    const totalPages = Math.max(1, Math.ceil(totalFiltered / pageSize));
    if(adminPipelinePage > totalPages) adminPipelinePage = totalPages;
    const pageStart = (adminPipelinePage - 1) * pageSize;
    const pageItems = items.slice(pageStart, pageStart + pageSize);

    const total = allItems.length;
    const openCount = allItems.filter(isOpen).length;
    const assignedCount = allItems.filter(i => i.AssignedTo && !i.IsNotSuitable).length;
    
    const statsEl = document.getElementById("adminStats");
    statsEl.innerHTML = `
        <div class="stat-pills">
            <span class="stat-pill">${total} total</span>
            <span class="stat-pill">${openCount} open</span>
            <span class="stat-pill">${assignedCount} assigned</span>
        </div>`;
    const countEl = document.getElementById("adminFilterCount");
    if(countEl){
      const shownFrom = totalFiltered ? pageStart + 1 : 0;
      const shownTo = Math.min(pageStart + pageSize, totalFiltered);
      countEl.textContent = shownFrom + "–" + shownTo + " of " + totalFiltered + " brokers";
    }

    const pagination = document.getElementById("adminPagination");
    if(pagination){
      pagination.innerHTML = "";
      if(totalPages > 1){
        const prev = document.createElement("button");
        prev.type="button"; prev.textContent="‹ Previous"; prev.disabled=adminPipelinePage===1;
        prev.onclick=()=>{adminPipelinePage--;renderAdmin();};
        pagination.appendChild(prev);

        const info = document.createElement("span");
        info.textContent = "Page " + adminPipelinePage + " of " + totalPages;
        pagination.appendChild(info);

        const next = document.createElement("button");
        next.type="button"; next.textContent="Next ›"; next.disabled=adminPipelinePage===totalPages;
        next.onclick=()=>{adminPipelinePage++;renderAdmin();};
        pagination.appendChild(next);
      }
    }

    const list = document.getElementById("adminList");
    list.innerHTML = "";
    if(totalFiltered === 0){
        list.innerHTML = `<div style="padding:20px; text-align:center; color:var(--muted);">No records match your filters. <a onclick="clearAdminFilters()" style="color:var(--accent);cursor:pointer;">Clear filters</a></div>`;
        return;
    }
    pageItems.forEach(item => {
        const card = document.createElement("div");
        card.className = "admin-card";
        const assignedLabel = item.IsNotSuitable
            ? "Not Suitable" + (item.NotSuitableReason ? ` (${item.NotSuitableReason})` : "")
            : (item.AssignedTo ? `Assigned to ${escapeHtml(item.AssignedTo.Title)}` : "Unassigned");

        const refDt = referenceDate(item);
        // Reversion only applies to assigned active brokers
        let reversionHtml = "";
        if(currentView === "mine" && item.AssignedTo && !item.IsNotSuitable){
            const reversionDt = new Date(refDt.getTime() + (getBrokerReleaseDays() * 24 * 60 * 60 * 1000));
            reversionHtml = ` <strong style="color:var(--amber); margin-left:6px;">(Releases if unupdated: ${reversionDt.toLocaleDateString()})</strong>`;
        }

        card.innerHTML = `
            <div class="main-info">
                <strong>${escapeHtml(item.Company || item.Title || "(no name)")}</strong><br>
                <span style="font-size:12px;color:var(--muted);">${assignedLabel} · last contact ${refDt.toLocaleDateString()}${reversionHtml}</span>
            </div>
            <div class="actions">
                <button type="button" data-act="reassign">Reassign…</button>
                <button type="button" data-act="release">Release</button>
                <button type="button" data-act="delete" class="danger">Delete</button>
            </div>
        `;
        card.querySelector('[data-act="reassign"]').addEventListener("click", () => openReassignModal(item.Id));
        card.querySelector('[data-act="release"]').addEventListener("click", () => {
            allItems = getStoredBrokers();
            const idx = allItems.findIndex(i => String(i.Id ?? "") === String(item.Id ?? "") || String(i.PortalId ?? "") === String(item.PortalId ?? ""));
            if(idx !== -1) {
                allItems[idx].AssignedTo = null;
                allItems[idx].IsNotSuitable = false;
                saveStoredBrokers(allItems);
                logAuditRecord(allItems[idx], "ADMIN_RELEASE", "Released to open pool by admin");
            }
            reload();
        });
        const deleteBtn = card.querySelector('[data-act="delete"]');
        deleteBtn.onclick = function(ev){
            ev.preventDefault();
            ev.stopPropagation();
            const btn = this;
            if(btn.dataset.confirmDelete !== "1"){
                btn.dataset.confirmDelete = "1";
                btn.dataset.originalText = btn.textContent;
                btn.textContent = "Confirm delete";
                btn.style.background = "#b42318";
                btn.style.color = "#fff";
                btn.style.borderColor = "#b42318";
                clearTimeout(btn._deleteTimer);
                btn._deleteTimer = setTimeout(()=>{
                    btn.dataset.confirmDelete = "0";
                    btn.textContent = btn.dataset.originalText || "Delete";
                    btn.style.background = "";
                    btn.style.color = "";
                    btn.style.borderColor = "";
                }, 5000);
                return;
            }

            clearTimeout(btn._deleteTimer);
            btn.disabled = true;
            btn.textContent = "Deleting…";

            const reason = "Manual removal by admin";
            (async()=>{
                try{
                    await deleteBrokerAdmin(item.Id, reason);
                    btn.textContent = "Deleted";
                    btn.style.background = "#176b37";
                    btn.style.color = "#fff";
                    btn.style.borderColor = "#176b37";

                    // Force the current local view to use the newly persisted dataset.
                    allItems = getStoredBrokers();
                    _brokerSnapshot = _snapshotBrokers(allItems);

                    setTimeout(()=>reload(), 350);
                }catch(e){
                    console.error("Broker deletion failed:", e);
                    btn.disabled = false;
                    btn.dataset.confirmDelete = "0";
                    btn.textContent = "Delete";
                    btn.style.background = "";
                    btn.style.color = "";
                    btn.style.borderColor = "";
                    alert("The broker could not be deleted.\n\n" + (e.message || e));
                }
            })();
        };
        list.appendChild(card);
    });
}

document.getElementById("adminSearchBox").addEventListener("input", renderAdmin);

document.getElementById("addNewBrokerBtn").addEventListener("click", () => {
    selectedItemId = Date.now();
    document.getElementById("detailTitle").textContent = "New Broker Record";
    document.getElementById("f_name").value = "";
    document.getElementById("f_company").value = "";
    document.getElementById("f_phone").value = "";
    document.getElementById("nuacomCallBtn").href = "#";
    document.getElementById("nuacomCallBtn").style.pointerEvents = "none";
    document.getElementById("nuacomCallBtn").style.opacity = "0.5";
    document.getElementById("f_email").value = "";
    document.getElementById("outlookEmailBtn").href = "#";
    document.getElementById("outlookEmailBtn").style.pointerEvents = "none";
    document.getElementById("outlookEmailBtn").style.opacity = "0.5";
    document.getElementById("f_website").value = "";
    document.getElementById("websiteLinkBtn").href = "#";
    document.getElementById("websiteLinkBtn").style.pointerEvents = "none";
    document.getElementById("websiteLinkBtn").style.opacity = "0.5";
    document.getElementById("f_address").value = "";
    document.getElementById("f_city").value = "";
    updateNotesWidgetDisplay("");
    document.getElementById("f_prefComm").value = "Phone";
    
    // Clear checkboxes
    const container = document.getElementById("f_loanTypesContainer");
    container.querySelectorAll("input[type=checkbox]").forEach(chk => chk.checked = false);

    document.getElementById("f_volume").value = "Under £1M";
    document.getElementById("f_network").value = "";
    document.getElementById("f_status").value = "Cold";
    document.getElementById("f_nextFollowUp").value = "";

    document.getElementById("assignInfo").textContent = "New item (unassigned)";
    document.getElementById("assignBtn").style.display = "inline-block";
    document.getElementById("assignBtn").textContent = "Assign to me";
    document.getElementById("releaseBtn").style.display = "none";
    document.getElementById("moveToNotSuitableBtn").style.display = "inline-block";
    document.getElementById("revertToOpenBtn").style.display = "none";
    
    const saveBtn = document.getElementById("saveDetailBtn");
    saveBtn.onclick = () => {
        const checkboxes = container.querySelectorAll("input[type=checkbox]:checked");
        const loanTypesArr = Array.from(checkboxes).map(chk => chk.value);

        const body = {
            Title: document.getElementById("f_name").value,
            Company: document.getElementById("f_company").value,
            Phone: document.getElementById("f_phone").value,
            Email: document.getElementById("f_email").value,
            Website: document.getElementById("f_website").value,
            Address: document.getElementById("f_address").value,
            City: document.getElementById("f_city").value,
            Notes: "",
            PrefComm: document.getElementById("f_prefComm").value,
            LoanTypes: loanTypesArr,
            Volume: document.getElementById("f_volume").value,
            Network: document.getElementById("f_network").value,
            Status: document.getElementById("f_status").value,
            NextFollowUp: document.getElementById("f_nextFollowUp").value
        };
        createItemLocal(body);
        reload();
        document.getElementById("detailOverlay").classList.remove("show");
        saveBtn.onclick = () => saveDetail();
    };

    document.getElementById("detailOverlay").classList.add("show");
});

document.querySelectorAll(".sub-nav button").forEach(btn => {
    btn.addEventListener("click", (ev) => { ev.preventDefault(); ev.stopPropagation(); setPipelineView(btn.dataset.view); });
});

document.getElementById("searchBox").addEventListener("input", render);
document.getElementById("notSuitableOutcomeFilter")?.addEventListener("change", render);
document.getElementById("closeDetail")?.addEventListener("click", closeDetailPanel, {capture:true});
loadDisplaySettings();
window.addEventListener('beforeunload', function(e){
    if(document.getElementById('detailOverlay')?.classList.contains('show') && brokerHasUnsavedChanges()){
        e.preventDefault(); e.returnValue='';
    }
});
document.getElementById("saveDetailBtn").addEventListener("click", () =>
    saveDetail());
document.getElementById("assignBtn").addEventListener("click", () =>
    assignToMe());
document.getElementById("releaseBtn").addEventListener("click", () =>
    release());

let liveBrokerRefreshTimer=null;
let _dataHealth={lastRefresh:null,lastError:"",source:"Local",count:0,open:0,mine:0,notSuitable:0,inFlight:false};
function renderDataHealth(){const g=document.getElementById('dataHealthGrid');if(!g)return;const b=getStoredBrokers();_dataHealth.count=b.length;_dataHealth.open=b.filter(isOpen).length;_dataHealth.mine=b.filter(isMine).length;_dataHealth.notSuitable=b.filter(isNotSuitable).length;const last=_dataHealth.lastRefresh?new Date(_dataHealth.lastRefresh).toLocaleString('en-GB'):'Never';const source=isDemoMode()?'Demo / local':(m365Configured()?'Microsoft 365 / SharePoint':'Local');g.innerHTML=[['Source',source],['Last refresh',last],['Broker records',b.length],['Open / Mine / Not suitable',`${_dataHealth.open} / ${_dataHealth.mine} / ${_dataHealth.notSuitable}`]].map(x=>`<div style="background:#f8fafc;border:1px solid var(--border);border-radius:12px;padding:12px"><div style="font-size:10px;color:var(--muted);font-weight:700;text-transform:uppercase;letter-spacing:.06em">${x[0]}</div><div style="font-size:15px;font-weight:800;color:var(--primary-dark);margin-top:4px">${escapeHtml(x[1])}</div></div>`).join('');const m=document.getElementById('dataHealthMessage');if(m)m.innerHTML=_dataHealth.lastError?`<strong style="color:var(--red)">Last sync error:</strong> ${escapeHtml(_dataHealth.lastError)}`:`<span style="color:var(--green)">✓ SharePoint is the source of truth. Live refresh is active while the portal is visible.</span>`;}
async function refreshDataHealthNow(){await refreshBrokersFromCloud(true);renderDataHealth();}
async function refreshBrokersFromCloud(force=false){if(isDemoMode()||!m365Configured()||_brokerSyncInFlight){renderDataHealth();return;}if(!force&&document.getElementById('detailOverlay')?.classList.contains('show')&&brokerHasUnsavedChanges()){renderDataHealth();return;}if(_dataHealth.inFlight&&!force)return;_dataHealth.inFlight=true;try{let brokers=await cloudLoadBrokers();try{const logs=await cloudLoadAudit();localStorage.setItem(AUDIT_STORAGE_KEY,JSON.stringify(logs));}catch(_){}brokers=hydrateBrokerOutcomeMetadata(brokers,getStoredAuditLogs());localStorage.setItem(STORAGE_KEY,JSON.stringify(brokers));_brokerSnapshot=_snapshotBrokers(brokers);allItems=brokers;_dataHealth.lastRefresh=new Date().toISOString();_dataHealth.lastError="";if(['open','mine','notSuitable'].includes(currentView))render();if(currentView==='dialer')buildDialerQueue();}catch(e){_dataHealth.lastError=e.message||String(e);}finally{_dataHealth.inFlight=false;renderDataHealth();}}
function startLiveBrokerRefresh(){if(liveBrokerRefreshTimer)clearInterval(liveBrokerRefreshTimer);const seconds=Math.min(300,Math.max(15,parseInt(getGlobalSettingValue('liveRefreshSeconds',30),10)||30));if(isDemoMode()||!m365Configured()){renderDataHealth();return;}refreshBrokersFromCloud(true);liveBrokerRefreshTimer=setInterval(()=>{if(document.visibilityState==='visible')refreshBrokersFromCloud(false);},seconds*1000);}
document.addEventListener('visibilitychange',()=>{if(document.visibilityState==='visible'&&m365Configured()&&!isDemoMode())refreshBrokersFromCloud(true);});window.addEventListener('focus',()=>{if(m365Configured()&&!isDemoMode())refreshBrokersFromCloud(true);});
async function initApp(){
    await loadGlobalPortalSettings(false);
    try{await loadPerformanceReviews();}catch(e){console.warn("Performance review load failed:",e.message||e);}
    renderDataHealth();
    startLiveBrokerRefresh();
    populateHome();
    switchMasterView('home');
}

checkSession();
