const PORTAL_KEY='rhPortalUrl';
const portalParam=new URLSearchParams(location.search).get('portal');
if(portalParam && /^https:\/\//i.test(portalParam)){localStorage.setItem(PORTAL_KEY,portalParam);history.replaceState({},'',location.pathname);}
const getPortal=()=>localStorage.getItem(PORTAL_KEY)||'';
const dialog=document.getElementById('configDialog');
const input=document.getElementById('portalInput');
function openPortal(){const url=getPortal();if(url){location.href=url;return;} input.value='';dialog.classList.add('show');input.focus();}
function savePortal(){const value=input.value.trim();if(!/^https:\/\//i.test(value)||!value.includes('/exec')){alert('Cole o link publicado do Portal RH, começando com https:// e terminando em /exec.');return;}localStorage.setItem(PORTAL_KEY,value);dialog.classList.remove('show');location.href=value;}
function closeDialog(){dialog.classList.remove('show');}
let deferredPrompt;
window.addEventListener('beforeinstallprompt',e=>{e.preventDefault();deferredPrompt=e;document.getElementById('installBtn').hidden=false;});
async function installApp(){if(deferredPrompt){deferredPrompt.prompt();await deferredPrompt.userChoice;deferredPrompt=null;return;}alert('No iPhone: toque em Compartilhar e depois em “Adicionar à Tela de Início”.');}
if('serviceWorker' in navigator){window.addEventListener('load',()=>navigator.serviceWorker.register('./service-worker.js').catch(()=>{}));}
document.getElementById('portalStatus').textContent=getPortal()?'Portal configurado':'Portal ainda não configurado';
