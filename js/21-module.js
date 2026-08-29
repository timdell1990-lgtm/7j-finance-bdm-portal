(function(){
  function syncAIAdminToggle(){
    const input=document.getElementById('globalAIEnabled');
    const state=document.getElementById('globalAIEnabledState');
    if(!input) return;
    const on=input.checked===true;
    if(state) state.textContent=on?'Enabled':'Hidden';
    input.setAttribute('aria-label',on?'AI broker search enabled':'AI broker search hidden');
  }

  window.notifyGlobalSettingsSaved=function(message,kind){
    const status=document.getElementById('globalSettingsStatus');
    if(status){
      status.textContent=message;
      status.className='status-msg '+(kind||'ok');
    }
    let toast=document.getElementById('globalSettingsToast');
    if(!toast){
      toast=document.createElement('div');
      toast.id='globalSettingsToast';
      toast.className='global-settings-toast';
      document.body.appendChild(toast);
    }
    toast.textContent=message;
    toast.className='global-settings-toast show '+(kind||'ok');
    clearTimeout(window.__globalSettingsToastTimer);
    window.__globalSettingsToastTimer=setTimeout(function(){toast.classList.remove('show');},3200);
  };

  document.addEventListener('change',function(e){
    if(e.target && e.target.id==='globalAIEnabled'){
      syncAIAdminToggle();

      // Apply the visibility change immediately when the administrator toggles it.
      // The previous version only refreshed visibility after Save, so the panel
      // could remain visible until the settings were saved.
      try{
        localStorage.setItem('adt7j_ai_enabled_v1', e.target.checked ? 'true' : 'false');
      }catch(_e){}
      if(typeof updateAIBrokerSearchVisibility==='function'){
        updateAIBrokerSearchVisibility();
      }
    }
  });
  document.addEventListener('DOMContentLoaded',syncAIAdminToggle);
  window.syncAIAdminToggle=syncAIAdminToggle;
})();
