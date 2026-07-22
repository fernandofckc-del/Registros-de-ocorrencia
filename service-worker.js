var CACHE_NAME = 'registro-irrigacao-v6';
var ASSETS = [
  './',
  './index.html',
  './manifest.json',
  './icon-192.png',
  './icon-512.png'
];

self.addEventListener('install', function (event) {
  event.waitUntil(
    caches.open(CACHE_NAME).then(function (cache) {
      // Guarda cada arquivo separadamente: se um falhar, os outros ainda são salvos.
      // (cache.addAll falha tudo-ou-nada se UM arquivo der erro - isso causava o app
      // não funcionar offline de jeito nenhum.)
      return Promise.all(
        ASSETS.map(function (url) {
          return cache.add(url).catch(function (err) {
            console.log('[SW] Falha ao guardar em cache:', url, err);
          });
        })
      );
    }).then(function () {
      return self.skipWaiting();
    })
  );
});

self.addEventListener('activate', function (event) {
  event.waitUntil(
    caches.keys().then(function (keys) {
      return Promise.all(
        keys.filter(function (k) { return k !== CACHE_NAME; })
            .map(function (k) { return caches.delete(k); })
      );
    })
  );
  self.clients.claim();
});

self.addEventListener('fetch', function (event) {
  var url = event.request.url;

  // Nunca cachear chamadas ao backend (Google Apps Script) - sempre tentar rede real
  if (url.indexOf('script.google.com') > -1 || url.indexOf('googleusercontent.com') > -1) {
    return;
  }

  // O documento principal (index.html) sempre tenta a rede primeiro, pra nunca
  // ficar preso numa versão antiga depois de uma atualização. Só usa o cache
  // se estiver de fato offline.
  if (event.request.mode === 'navigate' || url.indexOf('index.html') > -1) {
    event.respondWith(
      fetch(event.request).then(function (fresh) {
        caches.open(CACHE_NAME).then(function (cache) { cache.put(event.request, fresh.clone()); });
        return fresh;
      }).catch(function () {
        return caches.match(event.request).then(function (cached) {
          return cached || caches.match('./index.html');
        });
      })
    );
    return;
  }

  // Outros arquivos (ícones, manifest): cache primeiro, com fallback pra rede
  event.respondWith(
    caches.match(event.request).then(function (cached) {
      return cached || fetch(event.request);
    })
  );
});
