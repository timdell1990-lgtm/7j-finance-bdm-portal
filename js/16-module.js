(function(){
  let lastKey="";
  const previous=window.renderKPIDashboard;
  window.renderKPIDashboard=function(){
    if(typeof previous==="function") previous.apply(this,arguments);
    try{
      if(isDemoMode()||!m365Configured()) return;
      const s=makeKpiSnapshot(), key=s.Id+"|"+(currentUser?.Email||"");
      if(key===lastKey) return;
      lastKey=key;
      cloudSaveKpiSnapshot().catch(e=>console.warn("KPI snapshot sync failed:",e.message||e));
    }catch(e){console.warn("KPI snapshot preparation failed:",e.message||e)}
  };
})();
