(function(){
  function bindLogout(){
    const b=document.getElementById('logoutBtn');
    if(!b || b.dataset.bound==='1') return;
    b.dataset.bound='1';
    b.addEventListener('click',function(e){
      e.preventDefault();
      e.stopPropagation();
      try{
        handleLogout();
      }catch(err){
        console.error('Sign out failed:',err);
        sessionStorage.removeItem(SESSION_KEY);
        sessionStorage.removeItem(DEMO_MODE_KEY);
        currentUser=null;
        window.location.reload();
      }
    },true);
  }
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',bindLogout);
  else bindLogout();
})();
