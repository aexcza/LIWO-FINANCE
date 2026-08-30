/* ===== Original inline script 0 | id: liwo-initial-loader ===== */
(function(){
  function bootLoader(){
    var el=document.getElementById("liwoLoadingOverlay");
    if(!el)return;
    el.classList.add("show");
    el.setAttribute("aria-hidden","false");
    var txt=el.querySelector(".lw-loader-text");
    if(txt)txt.textContent="Preparing LIWO Finance…";
    window.setTimeout(function(){
      if(window.sessionStorage && sessionStorage.getItem("cf_session")){
        /* The restored session will keep the loader visible through the API call. */
      }else{
        el.classList.remove("show");
        el.setAttribute("aria-hidden","true");
      }
    },900);
  }
  if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",bootLoader,{once:true});
  else bootLoader();
})();

/* ===== Original inline script 1 | id: liwo-desktop-mobile-routing ===== */
(function(){
  /* The desktop index remains desktop. Phones use the dedicated mobile.html frontend. */
  try{
    if(window.matchMedia && window.matchMedia('(max-width: 900px)').matches && !location.pathname.endsWith('/mobile.html')){
      var target=new URL('mobile.html',location.href);
      target.search=location.search;
      target.hash=location.hash;
      location.replace(target.href);
    }
  }catch(e){}
})();
