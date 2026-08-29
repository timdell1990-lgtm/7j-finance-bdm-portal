(function(){
  function sync(){
    if(typeof updateAIBrokerSearchVisibility === 'function') updateAIBrokerSearchVisibility();
  }

  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded', sync, {once:true});
  else sync();

  window.addEventListener('storage', function(e){
    if(e.key==='adt7j_ai_enabled_v1') sync();
  });

  // Do not use a self-observing MutationObserver here. The previous implementation
  // observed the exact attributes (class/hidden/style/aria-hidden) that it then
  // modified, which can create a MutationObserver feedback loop and freeze Chrome
  // on local/demo builds. The authoritative update function is called explicitly
  // whenever the setting changes instead.
})();
