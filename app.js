(() => {
  const config = window.PORTAL_RH_CONFIG || {};
  const portal = document.getElementById('openPortal');
  const installBtn = document.getElementById('installBtn');
  const message = document.getElementById('installMessage');
  let deferredPrompt = null;

  portal.href = config.portalUrl || '#';

  const showMessage = (html) => {
    message.innerHTML = html;
    message.hidden = false;
  };

  window.addEventListener('beforeinstallprompt', (event) => {
    event.preventDefault();
    deferredPrompt = event;
    installBtn.hidden = false;
  });

  installBtn.addEventListener('click', async () => {
    if (deferredPrompt) {
      deferredPrompt.prompt();
      const result = await deferredPrompt.userChoice;
      deferredPrompt = null;
      if (result.outcome === 'accepted') {
        showMessage('Aplicativo instalado com sucesso.');
        installBtn.hidden = true;
      }
      return;
    }

    const ua = navigator.userAgent.toLowerCase();
    const isIOS = /iphone|ipad|ipod/.test(ua);
    if (isIOS) {
      showMessage('<strong>Para instalar no iPhone:</strong><br>abra esta página no Safari, toque em Compartilhar e escolha <em>Adicionar à Tela de Início</em>.');
    } else {
      showMessage('<strong>Para instalar no Android:</strong><br>abra o menu do Chrome (⋮) e escolha <em>Instalar aplicativo</em> ou <em>Adicionar à tela inicial</em>.');
    }
  });

  window.addEventListener('appinstalled', () => {
    installBtn.hidden = true;
    showMessage('Aplicativo instalado com sucesso.');
  });

  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('./service-worker.js?v=final2026')
        .catch(() => showMessage('A instalação automática não foi ativada. Use o menu do navegador para adicionar à tela inicial.'));
    });
  }
})();
