(() => {
  "use strict";
  const cfg = window.PORTAL_CONFIG;
  const installBtn = document.getElementById("installBtn");
  const openPortal = document.getElementById("openPortal");
  const status = document.getElementById("installStatus");
  const dialog = document.getElementById("installDialog");
  const dialogContent = document.getElementById("dialogContent");
  let deferredPrompt = null;

  const ua = navigator.userAgent.toLowerCase();
  const isIOS = /iphone|ipad|ipod/.test(ua);
  const isAndroid = /android/.test(ua);
  const isStandalone = window.matchMedia("(display-mode: standalone)").matches || window.navigator.standalone === true;

  openPortal.addEventListener("click", () => {
    window.location.href = cfg.portalUrl;
  });

  window.addEventListener("beforeinstallprompt", (event) => {
    event.preventDefault();
    deferredPrompt = event;
    status.textContent = "O aplicativo está pronto para instalação.";
  });

  window.addEventListener("appinstalled", () => {
    deferredPrompt = null;
    installBtn.hidden = true;
    status.textContent = "Aplicativo instalado com sucesso!";
  });

  function showInstructions() {
    if (isIOS) {
      dialogContent.innerHTML = `<ol><li>Abra esta página no <strong>Safari</strong>.</li><li>Toque em <strong>Compartilhar</strong> (quadrado com seta).</li><li>Selecione <strong>Adicionar à Tela de Início</strong>.</li><li>Confirme o nome <strong>RH | e Você</strong> e toque em <strong>Adicionar</strong>.</li></ol>`;
    } else if (isAndroid) {
      dialogContent.innerHTML = `<ol><li>Abra esta página no <strong>Google Chrome</strong>.</li><li>Toque nos <strong>três pontos</strong> do navegador.</li><li>Escolha <strong>Instalar aplicativo</strong> ou <strong>Adicionar à tela inicial</strong>.</li><li>Confirme a instalação.</li></ol>`;
    } else {
      dialogContent.innerHTML = `<p>Abra este endereço no celular. No Android, use o Chrome. No iPhone, use o Safari e escolha <strong>Adicionar à Tela de Início</strong>.</p>`;
    }
    if (typeof dialog.showModal === "function") dialog.showModal();
    else alert(dialogContent.textContent);
  }

  installBtn.addEventListener("click", async () => {
    if (isStandalone) {
      status.textContent = "O aplicativo já está instalado.";
      return;
    }
    if (deferredPrompt) {
      deferredPrompt.prompt();
      const choice = await deferredPrompt.userChoice;
      status.textContent = choice.outcome === "accepted" ? "Instalação iniciada." : "Instalação cancelada.";
      deferredPrompt = null;
      return;
    }
    showInstructions();
  });

  if (isStandalone) {
    installBtn.hidden = true;
    status.textContent = "Aplicativo instalado.";
  }

  if ("serviceWorker" in navigator) {
    window.addEventListener("load", () => {
      navigator.serviceWorker.register("service-worker.js?v=9", {scope:"./"}).catch(() => {
        status.textContent = "Abra pelo link oficial para instalar o aplicativo.";
      });
    });
  }
})();
