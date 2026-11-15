/* Аудит СМР — установка, обновление, офлайн-статус */
(function () {
  if (window.__smrPWA) return; window.__smrPWA = true;
  var deferred = null, reg = null, reloading = false;
  var box = document.createElement('div');
  box.style.cssText = 'position:fixed;left:12px;bottom:calc(12px + env(safe-area-inset-bottom));z-index:99999;display:flex;flex-direction:column;gap:8px;font-family:system-ui,sans-serif;max-width:calc(100vw - 24px);';
  document.body.appendChild(box);
  function pill(text, bg, fn) {
    var b = document.createElement('button'); b.type = 'button'; b.textContent = text;
    b.style.cssText = 'border:none;border-radius:999px;padding:12px 18px;font-size:14px;font-weight:700;color:#fff;cursor:pointer;box-shadow:0 6px 20px rgba(0,0,0,.35);background:' + bg + ';text-align:left;';
    b.onclick = fn; box.appendChild(b); return b;
  }
  var installBtn = null, updateBtn = null, offlineChip = null;
  function doInstall() {
    if (deferred) { var d = deferred; deferred = null; d.prompt(); d.userChoice.then(hideInstall, hideInstall); }
    else alert('Откройте меню браузера (⋮) и выберите «Установить приложение» или «Добавить на главный экран».');
  }
  function hideInstall() { if (installBtn) { installBtn.remove(); installBtn = null; } }
  window.addEventListener('beforeinstallprompt', function (e) {
    e.preventDefault(); deferred = e;
    var ex = document.getElementById('installBtn');
    if (ex) { ex.classList.remove('hidden'); ex.style.display = 'inline-flex'; ex.onclick = doInstall; return; }
    if (!installBtn && !matchMedia('(display-mode: standalone)').matches) installBtn = pill('📲 Установить приложение', 'linear-gradient(135deg,#b83a9e,#5f1152)', doInstall);
  });
  window.addEventListener('appinstalled', function () { hideInstall(); deferred = null; });
  function applyUpdate() { if (reg && reg.waiting) reg.waiting.postMessage({ type: 'SKIP_WAITING' }); else location.reload(); }
  function showUpdate() {
    var ex = document.getElementById('updateBtn');
    if (ex) { ex.classList.remove('hidden'); ex.style.display = 'inline-flex'; ex.onclick = applyUpdate; return; }
    if (!updateBtn) updateBtn = pill('⬆ Обновить приложение', '#2fbf71', applyUpdate);
  }
  if ('serviceWorker' in navigator && location.protocol !== 'file:') {
    navigator.serviceWorker.register('./sw.js').then(function (r) {
      reg = r; if (r.waiting) showUpdate();
      r.addEventListener('updatefound', function () {
        var w = r.installing; if (!w) return;
        w.addEventListener('statechange', function () { if (w.state === 'installed' && navigator.serviceWorker.controller) showUpdate(); });
      });
    }).catch(function () {});
    navigator.serviceWorker.addEventListener('controllerchange', function () { if (!reloading) { reloading = true; location.reload(); } });
  }
  function net() {
    if (!navigator.onLine && !offlineChip) {
      offlineChip = document.createElement('div'); offlineChip.textContent = '⚡ Офлайн — всё работает';
      offlineChip.style.cssText = 'border-radius:999px;padding:8px 14px;font-size:12px;font-weight:600;color:#fff;background:rgba(42,7,37,.92);border:1px solid rgba(255,255,255,.25);';
      box.appendChild(offlineChip);
    }
    if (navigator.onLine && offlineChip) { offlineChip.remove(); offlineChip = null; }
  }
  window.addEventListener('online', net); window.addEventListener('offline', net); net();
  if (/iphone|ipad|ipod/i.test(navigator.userAgent) && !matchMedia('(display-mode: standalone)').matches) {
    pill('📲 Как установить', 'linear-gradient(135deg,#b83a9e,#5f1152)', function () {
      alert('Safari: кнопка «Поделиться» → «На экран „Домой“».');
    });
  }
})();
