(function(){
  function wireToggles(){
    document.querySelectorAll('input[type="checkbox"]').forEach(function(input){
      if(input.dataset.toggleWired || input.id==='loginMultitenant' || input.id==='guideAdminActive' || input.closest('.admin-toggle')) return;
      const label=input.closest('label');
      if(!label) return;
      const text=(label.textContent||'').trim().toLowerCase();
      if(!/(enable|enabled|disable|disabled|show|hide|active|ai|multitenant|calendar|refresh|allow|demo)/.test(text)) return;
      input.dataset.toggleWired='1';
      input.classList.add('toggle-native-hidden');
      const sw=document.createElement('span'); sw.className='toggle-switch';
      input.parentNode.insertBefore(sw,input.nextSibling);
      const state=document.createElement('span'); state.className='toggle-state';
      state.textContent=input.checked?'ON':'OFF';
      label.appendChild(state);
      function sync(){state.textContent=input.checked?'ON':'OFF'}
      input.addEventListener('change',sync);
    });
  }
  document.addEventListener('DOMContentLoaded',function(){
    wireToggles();
    setTimeout(wireToggles,500);
    setTimeout(wireToggles,1500);
  });
  window.wirePortalToggles=wireToggles;
})();
