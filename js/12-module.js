(function(){
  function cleanAdmin(){
    const root=document.getElementById('adminUnlockedContent');
    if(!root || root.dataset.accordionDone==='1') return;
    const children=Array.from(root.children);
    let count=0;
    children.forEach(function(card){
      if(card.tagName!=='DIV' || card.classList.contains('admin-accordion') || !card.querySelector('h3')) return;
      const heading=card.querySelector('h3');
      const title=(heading.textContent||'Admin section').trim();
      const details=document.createElement('details');
      details.className='admin-accordion';
      const summary=document.createElement('summary'); summary.textContent=title;
      const body=document.createElement('div'); body.className='admin-accordion-body';
      // Move the complete original card inside the accordion so no controls are lost.
      root.insertBefore(details,card);
      details.appendChild(summary); details.appendChild(body); body.appendChild(card);
      count++;
    });
    if(count){ root.dataset.accordionDone='1'; }
  }
  document.addEventListener('DOMContentLoaded',function(){setTimeout(cleanAdmin,500);setTimeout(cleanAdmin,1500);});
  window.cleanAdminAccordion=cleanAdmin;
})();
