(function(){
  function desktopOnly(){
    return window.matchMedia && window.matchMedia('(min-width: 769px)').matches;
  }

  function toggleDemoAccounts(){
    if(!desktopOnly()) return;

    const panel = document.getElementById('demoLoginPanel');
    if(!panel) return;

    const isVisible = panel.getAttribute('data-demo-visible') === 'true';
    const show = !isVisible;

    panel.hidden = !show;
    panel.setAttribute('data-demo-visible', show ? 'true' : 'false');
    panel.setAttribute('aria-hidden', show ? 'false' : 'true');

    // The portal deliberately hides this card with !important, so the
    // reveal must also use !important. This does not invoke demo login.
    panel.style.setProperty('display', show ? 'block' : 'none', 'important');
  }

  window.addEventListener('keydown', function(e){
    if(!desktopOnly()) return;
    if(e.ctrlKey && !e.altKey && !e.shiftKey && (e.key === 'd' || e.key === 'D' || e.keyCode === 68)){
      e.preventDefault();
      e.stopImmediatePropagation();
      toggleDemoAccounts();
    }
  }, true);
})();
