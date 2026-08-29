(function(){
  function keepMainHeaderStable(){
    const header=document.querySelector('header');
    if(!header) return;
    if(header.style.display==='none') header.style.removeProperty('display');
    if(header.hidden) header.hidden=false;
    if(header.getAttribute('aria-hidden')!=='false') header.setAttribute('aria-hidden','false');
  }
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded', keepMainHeaderStable, {once:true});
  else keepMainHeaderStable();
  window.addEventListener('load', keepMainHeaderStable, {once:true});
})();
