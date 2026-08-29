(function(){
  function refreshKpiUserSelectors(){
    try{
      if(typeof populateKPIUserSelectors === 'function'){
        populateKPIUserSelectors();
      }
      if(typeof populateKPIUserSelect === 'function'){
        populateKPIUserSelect();
      }
    }catch(e){
      console.warn('KPI user selector refresh failed:', e && (e.message || e));
    }
  }
  window.addEventListener('load', function(){
    setTimeout(refreshKpiUserSelectors, 250);
    setTimeout(refreshKpiUserSelectors, 1200);
  });
})();
