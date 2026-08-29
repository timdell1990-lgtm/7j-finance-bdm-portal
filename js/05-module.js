function toggleMasterMenu(e){e?.stopPropagation();document.getElementById('masterMenu')?.classList.toggle('open');}
document.addEventListener('click',function(e){const t=e.target;if(!t||typeof t.closest!=='function')return;if(!t.closest('.master-menu')&&!t.closest('.master-menu-btn'))document.getElementById('masterMenu')?.classList.remove('open');});
function populateHome(){if(!currentUser)return;const w=document.getElementById('homeWelcome'),i=document.getElementById('homeUserInfo'),r=document.getElementById('homeRoleBadge');if(w)w.textContent='Welcome, '+(currentUser.Title||'User');if(i)i.textContent=(currentUser.Email||'')+' · 7J Finance workspace';if(r)r.textContent=currentUser.Role||'BDM';document.querySelectorAll('.admin-home-tile').forEach(x=>x.style.display=currentUser.Role==='Admin'?'':'none');}
function openPortalUtility(v){
    document.getElementById('masterMenu')?.classList.remove('open');
    if(v==='admin'&&currentUser?.Role!=='Admin'){alert('Admin Centre is restricted to administrators.');return;}
    switchMasterView('portal');currentView=v;const sub=document.querySelector('#portalView .sub-nav');if(sub)sub.style.display='none';
    document.getElementById('listView')?.style.setProperty('display','none');document.getElementById('uploadView')?.style.setProperty('display',v==='upload'?'block':'none');document.getElementById('adminView')?.style.setProperty('display',v==='admin'?'block':'none');document.getElementById('dialerView')?.style.setProperty('display',v==='dialer'?'block':'none');document.getElementById('dealsView')?.style.setProperty('display','none');document.getElementById('kpiDashboardView')?.style.setProperty('display','none');
    if(v==='upload'&&typeof window.populateUploadUI==='function')window.populateUploadUI();
    if(v==='dialer'){if(window.dialerQueue?.length)renderDialer();else buildDialerQueue();}
    if(v==='admin'){document.getElementById('adminPasswordGate').style.display='block';document.getElementById('adminUnlockedContent').style.display='none';}
}

function setKPIDashboardTab(tab){
    document.querySelectorAll('.kpi-section-tab').forEach(b=>b.classList.toggle('active',b.dataset.kpiTab===tab));
    document.body.classList.toggle('kpi-nonoverview-mode',tab!=='overview');
    const summary=document.getElementById('dashboardSummary');
    if(summary)summary.style.display=tab==='overview'?'grid':'none';
    const deals=document.getElementById('dashboardDealsPanel'),reviews=document.getElementById('dashboardReviewsPanel'),dialler=document.getElementById('dashboardDiallerPanel');
    if(deals)deals.style.display=tab==='deals'?'block':'none';
    if(reviews)reviews.style.display=tab==='reviews'?'block':'none';
    if(dialler)dialler.style.display=tab==='dialler'?'block':'none';
    document.body.classList.toggle('kpi-rankings-mode',tab==='rankings');
    document.body.classList.toggle('kpi-reviews-mode',tab==='reviews');
    document.body.classList.toggle('kpi-dialler-mode',tab==='dialler');
    renderKPIDashboard();
    if(tab==='reviews')renderPerformanceReviews();
    if(tab==='dialler'){renderDialerKPI();renderDialerRanking();}
}
// Render the selected KPI tab while keeping the same global filters for every view.
const _oldRenderKPI=window.renderKPIDashboard;
window.renderKPIDashboard=function(){
    const tab=document.querySelector('.kpi-section-tab.active')?.dataset.kpiTab||'overview';
    document.body.classList.toggle('kpi-nonoverview-mode',tab!=='overview');
    const summary=document.getElementById('dashboardSummary');
    if(summary)summary.style.display=tab==='overview'?'grid':'none';
    if(typeof _oldRenderKPI==='function')_oldRenderKPI();
    if(summary)summary.style.display=tab==='overview'?'grid':'none';
    const panel=document.getElementById('dashboardDealsPanel'),reviews=document.getElementById('dashboardReviewsPanel'),dialler=document.getElementById('dashboardDiallerPanel');
    if(!panel)return;
    if(reviews)reviews.style.display=tab==='reviews'?'block':'none';
    if(dialler)dialler.style.display=tab==='dialler'?'block':'none';
    try{saveKpiSnapshotLocal();}catch(e){}
    if(tab==='reviews')renderPerformanceReviews();
    if(tab==='dialler'){try{renderDialerKPI();renderDialerRanking();}catch(e){console.warn('Dialler KPI render failed:',e);}}
    if(tab==='deals'){
        const period=document.getElementById('dashboardPeriod')?.value||'all',selectedUser=document.getElementById('dashboardUser')?.value||'all',user=(currentUser&&currentUser.Role!=='Admin')?(currentUser.Title||currentUser.Email||''):selectedUser,start=document.getElementById('dashboardStartDate')?.value,end=document.getElementById('dashboardEndDate')?.value;let ds=getStoredDeals();
        if(start||end){const a=start?new Date(start+'T00:00:00'):new Date('1900-01-01'),b=end?new Date(end+'T23:59:59'):new Date('2999-12-31');ds=ds.filter(d=>{const x=new Date(d.DealDate||0);return x>=a&&x<=b;});}else ds=filterByPeriod(ds,period,d=>d.DealDate);
        if(user!=='all')ds=ds.filter(d=>normaliseKPIUserName(d.BDM)===normaliseKPIUserName(user));
        document.getElementById('dashboardDealsTable').innerHTML=ds.length?'<table class="kpi-table"><tr><th>Date</th><th>Broker</th><th>Company</th><th>BDM</th><th>Loan type</th><th>Value</th></tr>'+ds.sort((a,b)=>new Date(b.DealDate)-new Date(a.DealDate)).map(d=>'<tr><td>'+escapeHtml(d.DealDate||'')+'</td><td>'+escapeHtml(d.BrokerName||'')+'</td><td>'+escapeHtml(d.Company||'')+'</td><td>'+escapeHtml(d.BDM||'')+'</td><td>'+escapeHtml(d.LoanType||'')+'</td><td><strong>'+fmtGBP(d.DealValue)+'</strong></td></tr>').join('')+'</table>':'<div style="padding:12px;color:var(--muted)">No deals for the selected filters.</div>';
    }
};
const _oldSetUserbox=window.setUserbox;window.setUserbox=function(){if(typeof _oldSetUserbox==='function')_oldSetUserbox();populateHome();};
// Hide release date text in open broker cards if an older renderer inserts it.
const _oldRender=window.render;window.render=function(){if(typeof _oldRender==='function')_oldRender();document.querySelectorAll('#cardList [data-release-date],#cardList .release-date').forEach(e=>e.remove());};
