(() => {
  const cfg = window.APP_CONFIG;
  const openBtn = document.getElementById('openPortal');
  const installBtn = document.getElementById('installApp');
  const status = document.getElementById('status');
  let deferredPrompt = null;

  if (!cfg || !cfg.portalUrl) {
    status.textContent = 'Configuração do portal não encontrada.';
    openBtn.disabled = true;
    return;
  }

  openBtn.addEventListener('click', () => {
    status.textContent = 'Abrindo o Portal RH...';
    window.location.href = cfg.portalUrl;
  });

  window.addEventListener('beforeinstallprompt', event => {
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
    window.addEventListener('load', () => navigator.serviceWorker.register('./service-worker.js'));
  }
})();
