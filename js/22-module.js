(function(){
  document.addEventListener('input',function(e){
    if(e.target && e.target.id==='adminSearchBox'){
      if(typeof adminPipelinePage!=='undefined') adminPipelinePage=1;
      if(typeof renderAdmin==='function') renderAdmin();
    }
  });
  document.addEventListener('change',function(e){
    if(e.target && /^admin(Filter|User|Search|PageSize)/.test(e.target.id||'')){
      if(e.target.id!=='adminPageSize' && typeof adminPipelinePage!=='undefined') adminPipelinePage=1;
    }
  });
})();
