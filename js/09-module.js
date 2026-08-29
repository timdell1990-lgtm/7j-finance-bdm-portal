function frontPageHome(){
  try{
    // On the login/front page there is no active workspace yet; keep the user at the
    // branded front page and clear any stale session view.
    sessionStorage.removeItem("7JCurrentView");
    const btn=document.getElementById("frontHomeBtn");
    if(btn){btn.animate([{transform:"scale(.94)"},{transform:"scale(1)"}],{duration:160});}
  }catch(e){}
}
document.addEventListener('DOMContentLoaded',function(){
  const btn=document.getElementById('frontHomeBtn');
  if(btn && window.matchMedia('(max-width: 700px)').matches) btn.style.display='block';
  const lm=document.getElementById('loginMultitenant');
  const ga=document.getElementById('guideAdminActive');
  [lm,ga].forEach(function(i){
    if(!i) return;
    function sync(){
      const s=i.parentElement?.querySelector('.toggle-state');
      if(s) s.textContent=i.checked?'ON':'OFF';
    }
    i.addEventListener('change',sync); sync();
  });
});
