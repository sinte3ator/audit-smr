/* Аудит СМР — установка, обновление, сеть */
(function(){
  if(window.__smrPWA)return; window.__smrPWA=true;
  var deferred=null,reg=null,reloading=false;
  function el(id){return document.getElementById(id);}
  function show(e){if(e){e.classList.remove('hidden');e.style.display='flex';}}
  function hide(e){if(e){e.classList.add('hidden');e.style.display='';}}
  window.promptInstall=function(){
    if(deferred){var d=deferred;deferred=null;d.prompt();d.userChoice.then(function(){hide(el('installBtn'));},function(){hide(el('installBtn'));});}
    else alert('Откройте меню браузера (⋮) и выберите «Установить приложение» / «Добавить на главный экран». На iPhone: Safari → Поделиться → На экран «Домой».');
  };
  window.addEventListener('beforeinstallprompt',function(e){e.preventDefault();deferred=e;show(el('installBtn'));});
  window.addEventListener('appinstalled',function(){deferred=null;hide(el('installBtn'));});
  window.applyUpdate=function(){if(reg&&reg.waiting)reg.waiting.postMessage({type:'SKIP_WAITING'});else location.reload();};
  function showUpdate(){show(el('updateBtn'));}
  if('serviceWorker'in navigator&&location.protocol!=='file:'){
    navigator.serviceWorker.register('./sw.js').then(function(r){reg=r;if(r.waiting)showUpdate();
      r.addEventListener('updatefound',function(){var w=r.installing;if(!w)return;w.addEventListener('statechange',function(){if(w.state==='installed'&&navigator.serviceWorker.controller)showUpdate();});});
    }).catch(function(){});
    navigator.serviceWorker.addEventListener('controllerchange',function(){if(!reloading){reloading=true;location.reload();}});
  }
  function net(){var c=el('netChip');if(!c)return;var on=navigator.onLine;c.textContent=on?'Онлайн':'Офлайн';c.style.color=on?'#10b981':'#f59e0b';}
  window.addEventListener('online',net);window.addEventListener('offline',net);net();
})();
