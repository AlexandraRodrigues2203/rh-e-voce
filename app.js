(() => {
  'use strict';
  const cfg = window.APP_CONFIG || {};
  const portalUrl = cfg.portalUrl;
  const openBtn = document.getElementById('openPortal');
  const installBtn = document.getElementById('installBtn');
  const iosHelp = document.getElementById('iosHelp');
  const status = document.getElementById('status');
  let deferredPrompt = null;

  const isIos = /iphone|ipad|ipod/i.test(navigator.userAgent);
  const isStandalone = window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;

  if (isIos && !isStandalone) iosHelp.hidden = false;

  openBtn.addEventListener('click', () => {
    if (!portalUrl) {
      status.textContent = 'O endereço do Portal RH não está configurado.';
      return;
    }
    window.location.assign(portalUrl);
  });

  window.addEventListener('beforeinstallprompt', (event) => {
    event.preventDefault();
    deferredPrompt = event;
    installBtn.hidden = false;
  });

  installBtn.addEventListener('click', async () => {
    if (!deferredPrompt) return;
    deferredPrompt.prompt();
    await deferredPrompt.userChoice;
    deferredPrompt = null;
    installBtn.hidden = true;
  });

  if ('serviceWorker' in navigator) {
    window.addEventListener('load', async () => {
      try {
        const registrations = await navigator.serviceWorker.getRegistrations();
        await Promise.all(registrations.map((r) => {
          if (!r.active || !r.active.scriptURL.includes('service-worker.js')) return r.unregister();
          return Promise.resolve();
        }));
        await navigator.serviceWorker.register('./service-worker.js?v=8', { scope: './' });
      } catch (error) {
        console.error('Falha ao registrar o Service Worker:', error);
      }
    });
  }
})();
