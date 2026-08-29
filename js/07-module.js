(function(){
  'use strict';

  // ---------- Hidden Demo Mode ----------
  // Desktop: Ctrl+D while Home is visible.
  // Mobile: tap the Home icon 5 times within 2 seconds.
  // The visible login demo link remains hidden in production.
  let demoTapCount = 0;
  let demoTapTimer = null;

  function homeIsVisible(){
    const h = document.getElementById('homeView');
    return !!(h && (h.classList.contains('active-view') || h.style.display === 'flex'));
  }

  function launchHiddenDemo(){
    try{
      if(typeof enterLocalDemoMode === 'function'){
        enterLocalDemoMode();
        const b = document.getElementById('launchSyncBanner');
        if(b){
          b.className = 'show ok';
          b.textContent = 'Demo Mode enabled — Microsoft 365 is disconnected and data is local to this browser.';
          setTimeout(()=>{ b.className=''; }, 4500);
        }
      }
    }catch(e){ console.error('Hidden Demo Mode failed', e); }
  }

  document.addEventListener('keydown', function(e){
    if(homeIsVisible() && e.ctrlKey && !e.shiftKey && !e.altKey && String(e.key).toLowerCase()==='d'){
      e.preventDefault();
      launchHiddenDemo();
    }
  });

  document.addEventListener('click', function(e){
    const homeBtn = e.target.closest?.('#navHome');
    if(!homeBtn || !homeIsVisible()) return;
    demoTapCount++;
    clearTimeout(demoTapTimer);
    demoTapTimer = setTimeout(()=>{ demoTapCount=0; }, 2000);
    if(demoTapCount >= 5){
      demoTapCount=0;
      clearTimeout(demoTapTimer);
      launchHiddenDemo();
    }
  }, true);

  // ---------- Network / sync health ----------
  const banner = ()=>document.getElementById('launchSyncBanner');

  function setBanner(kind, text, autoHide){
    const b=banner();
    if(!b) return;
    b.className='show '+kind;
    b.textContent=text;
    if(autoHide) setTimeout(()=>{ if(b.textContent===text) b.className=''; }, autoHide);
  }

  window.addEventListener('offline', ()=>setBanner('offline','Offline — changes cannot be confirmed against Microsoft 365 until the connection returns.',0));
  window.addEventListener('online', ()=>{
    setBanner('ok','Connection restored — refreshing broker data…',3000);
    try{
      if(typeof refreshBrokersFromCloud==='function' && typeof isDemoMode==='function' && !isDemoMode() && typeof m365Configured==='function' && m365Configured()){
        refreshBrokersFromCloud(true);
      }
    }catch(_){}
  });

  // Turn the existing fire-and-forget broker sync into a visible failure instead of a silent success.
  if(typeof window.cloudSyncBrokers === 'function'){
    const originalCloudSyncBrokers = window.cloudSyncBrokers;
    window.cloudSyncBrokers = async function(items){
      try{
        const result = await originalCloudSyncBrokers(items);
        if(navigator.onLine !== false) setBanner('ok','Saved to Microsoft 365.',2200);
        return result;
      }catch(err){
        setBanner('error','Save not confirmed by Microsoft 365 — '+(err?.message||String(err))+' Use Refresh now before retrying.',0);
        throw err;
      }
    };
  }

  // If the existing saveStoredBrokers catches sync errors, surface them rather than silently swallowing them.
  if(typeof window.showCloudError === 'function'){
    const originalShowCloudError=window.showCloudError;
    window.showCloudError=function(err){
      setBanner('error','Microsoft 365 sync error — '+(err?.message||String(err)),0);
      return originalShowCloudError(err);
    };
  }
  // Admin menu item remains visible; access is enforced by openPortalUtility() and the admin gate.


  // ---------- Extra stale-state protection ----------
  // When a window/tab becomes visible, refresh live data before the user continues.
  document.addEventListener('visibilitychange', function(){
    if(document.visibilityState==='visible'){
      try{
        if(typeof isDemoMode==='function' && !isDemoMode() && typeof m365Configured==='function' && m365Configured() && typeof refreshBrokersFromCloud==='function'){
          refreshBrokersFromCloud(true);
        }
      }catch(_){}
    }
  });

  // Show a small initial connection state.
  setTimeout(()=>{
    if(navigator.onLine===false) setBanner('offline','Offline — Microsoft 365 changes cannot be confirmed.',0);
  },500);
})();
