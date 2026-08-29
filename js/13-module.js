function getCallbackEvents(){
    const now=Date.now();
    return (calendarEvents||[])
      .filter(e=>!!e.isCallback || /^callback\b/i.test(String(e.subject||"")))
      .filter(e=>new Date(e.start).getTime() >= now-3600000)
      .sort((a,b)=>new Date(a.start)-new Date(b.start));
}
function renderCallbackCount(){
    const el=document.getElementById("callbackCountBadge");
    if(el) el.textContent=String(getCallbackEvents().length);
}
function openCallbackList(){
    const modal=document.getElementById("callbackListModal"), body=document.getElementById("callbackListContent");
    if(!modal||!body)return;
    const events=getCallbackEvents();
    body.innerHTML=events.length ? `<div class="callback-list">${events.map(e=>{
        const dt=new Date(e.start);
        const broker=eventBrokerName(e)||"Unlinked broker";
        return `<div class="callback-row">
          <div><strong>${escapeHtml(e.subject||"Callback")}</strong><div style="color:var(--muted);font-size:13px;margin-top:3px">${escapeHtml(broker)}</div><div class="callback-date" style="margin-top:5px">${dt.toLocaleDateString("en-GB",{weekday:"short",day:"2-digit",month:"short",year:"numeric"})} · ${dt.toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"})}</div></div>
          <button class="btn secondary" type="button" onclick="openCallbackInDialler('${String(e.id)}')">Open in Dialler</button>
        </div>`;
    }).join("")}</div>` : `<div style="padding:24px;text-align:center;color:var(--muted)">No scheduled callbacks.</div>`;
    modal.classList.add("show");
}
function closeCallbackList(){document.getElementById("callbackListModal")?.classList.remove("show");}

function openCallbackInDialler(eventId){
    const e=(calendarEvents||[]).find(x=>String(x.id)===String(eventId));
    if(!e)return;
    const brokerId=String(e.brokerId||"");
    const broker=brokerId ? getStoredBrokers().find(b=>String(b.Id)===brokerId) : null;
    closeCallbackList();
    openPortalUtility("dialer");
    if(!broker){
        alert("The callback is not linked to a broker record, so the dialler details could not be opened.");
        return;
    }
    const existingIndex=(dialerQueue||[]).findIndex(b=>String(b.Id)===brokerId);
    if(existingIndex>=0){
        dialerIndex=existingIndex;
        selectedItemId=broker.Id;
        renderDialer();
        return;
    }
    const phone=dialerSanitizePhone(broker.Phone||broker.PhoneNumber||broker.Mobile||broker.Telephone||"");
    if(!phone){
        alert("This callback is linked to a broker, but the broker has no usable telephone number for the dialler.");
        return;
    }
    dialerQueue=[broker];
    dialerIndex=0;
    selectedItemId=broker.Id;
    renderDialer();
}

// Keep the badge in sync whenever the calendar is rendered/refreshed.
const _callbackRenderCalendar = window.renderCalendar;
window.renderCalendar = function(){
    if(typeof _callbackRenderCalendar==="function") _callbackRenderCalendar.apply(this,arguments);
    renderCallbackCount();
};
