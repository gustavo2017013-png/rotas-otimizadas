// Rotas Otimizadas — funciona sem internet depois de instalado.
// Só guarda e entrega os arquivos do próprio app. Nunca envia nada e não acessa nenhum outro site.
// Ao mudar qualquer arquivo do app, aumente o número da versão abaixo.
const VERSION = 'rotas-v7';
const SHELL = [
  './',
  './index.html',
  './app.js',
  './app.css',
  './manifest.webmanifest',
  './icon.svg',
  './icon-192.png',
  './icon-512.png',
  './icon-maskable-512.png',
  './icon-180.png'
];
const WAIT_MS = 3000; // com internet lenta, depois disso abre a cópia guardada

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(SHELL.map(u => new Request(u, { cache: 'reload' })))).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  // apaga as versões antigas guardadas
  e.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.filter(k => k !== VERSION).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const url = new URL(req.url);
  if (url.origin !== location.origin) return; // nada de outros sites passa por aqui

  // Com internet: pega a versão mais nova na hora (e guarda). Sem internet ou sinal fraco: usa a cópia guardada.
  const key = req.mode === 'navigate' ? './index.html' : url.origin + url.pathname;
  e.respondWith(caches.open(VERSION).then(async cache => {
    const fresh = fetch(req, { cache: 'no-cache' }).then(res => {
      if (res && res.ok) cache.put(key, res.clone());
      return res;
    }).catch(() => null);
    e.waitUntil(fresh);
    const timeout = new Promise(r => setTimeout(() => r(null), WAIT_MS));
    const net = await Promise.race([fresh, timeout]);
    if (net && net.ok) return net;
    const hit = await cache.match(key, { ignoreSearch: true });
    if (hit) return hit;
    return (await fresh) || new Response('Sem internet e sem cópia guardada.', { status: 503 });
  }));
});
