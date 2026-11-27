// Rotas Otimizadas — funciona sem internet depois de instalado.
// Só guarda e entrega os arquivos do próprio app. Nunca envia nada e não acessa nenhum outro site.
// Ao mudar qualquer arquivo do app, aumente o número da versão abaixo.
const VERSION = 'rotas-v6';
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

self.addEventListener('install', e => {
  e.waitUntil(caches.open(VERSION).then(c => c.addAll(SHELL)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', e => {
  // apaga as versões antigas guardadas (inclusive a antiga cópia das fontes)
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

  // Abre na hora com o que está guardado e atualiza em segundo plano (sinal fraco na rua não trava o app).
  const key = req.mode === 'navigate' ? './index.html' : url.origin + url.pathname;
  e.respondWith(caches.open(VERSION).then(async cache => {
    const hit = await cache.match(key, { ignoreSearch: true });
    const fresh = fetch(req).then(res => {
      if (res && res.ok) cache.put(key, res.clone());
      return res;
    }).catch(() => null);
    if (hit) { e.waitUntil(fresh); return hit; }
    return (await fresh) || new Response('Sem internet e sem cópia guardada.', { status: 503 });
  }));
});
