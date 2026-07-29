(() => {
  const installButton = document.getElementById('installButton');
  const iosDialog = document.getElementById('iosDialog');
  let deferredPrompt = null;

  const isIos = /iphone|ipad|ipod/i.test(navigator.userAgent);
  const isStandalone = window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;

  if ('serviceWorker' in navigator) {
    window.addEventListener('load', async () => {
      try {
        const reg = await navigator.serviceWorker.register('./service-worker.js?v=10', { scope: './' });
        reg.update();
      } catch (err) {
        console.error('Falha ao registrar o service worker:', err);
      }
    });
  }

  window.addEventListener('beforeinstallprompt', (event) => {
    event.preventDefault();
    deferredPrompt = event;
    installButton.disabled = false;
  });

  window.addEventListener('appinstalled', () => {
    deferredPrompt = null;
    installButton.style.display = 'none';
  });

  if (isStandalone) installButton.style.display = 'none';

  installButton.addEventListener('click', async () => {
    if (deferredPrompt) {
      deferredPrompt.prompt();
      await deferredPrompt.userChoice;
      deferredPrompt = null;
      return;
    }

    if (isIos && typeof iosDialog.showModal === 'function') {
      iosDialog.showModal();
      return;
    }

    alert('No Android, abra o menu do Chrome (⋮) e toque em “Instalar aplicativo” ou “Adicionar à tela inicial”.');
  });
})();
