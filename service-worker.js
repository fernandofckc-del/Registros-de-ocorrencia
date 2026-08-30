var CACHE_NAME = 'registro-irrigacao-v49';
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
        // event.waitUntil() é essencial aqui: sem ele, o navegador pode
        // encerrar o service worker assim que a resposta é entregue,
        // ANTES dessa gravação no cache terminar - aí a atualização nunca
        // chega a ficar salva, e a versão antiga continua sendo servida
        // quando o celular fica offline (foi exatamente esse bug que
        // fazia uma aba já removida "voltar" sem internet).
        event.waitUntil(
          caches.open(CACHE_NAME).then(function (cache) {
            return cache.put(event.request, fresh.clone());
          })
        );
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

// ================= ENVIO EM SEGUNDO PLANO (Background Sync) =================
// Quando o navegador pega internet de novo, ele dispara este evento sozinho
// - mesmo com o app FECHADO. Aqui a gente lê a "gaveta" (IndexedDB) onde o
// index.html guardou os registros que não conseguiram sair na hora, e tenta
// enviar cada um de novo. Se algum falhar (ex.: internet caiu de novo no
// meio), o navegador tenta de novo mais tarde automaticamente - não precisa
// tratar isso aqui, é o próprio contrato do Background Sync.
var OUTBOX_DB = 'irrigacao_outbox_v1', OUTBOX_STORE = 'fila';

function abrirOutboxSW_() {
  return new Promise(function (resolve, reject) {
    var req = indexedDB.open(OUTBOX_DB, 1);
    req.onupgradeneeded = function () {
      req.result.createObjectStore(OUTBOX_STORE, { keyPath: 'chave', autoIncrement: true });
    };
    req.onsuccess = function () { resolve(req.result); };
    req.onerror = function () { reject(req.error); };
  });
}

function listarOutboxSW_() {
  return abrirOutboxSW_().then(function (db) {
    return new Promise(function (resolve, reject) {
      var tx = db.transaction(OUTBOX_STORE, 'readonly');
      var req = tx.objectStore(OUTBOX_STORE).getAll();
      req.onsuccess = function () { resolve(req.result || []); };
      req.onerror = function () { reject(req.error); };
    });
  });
}

function apagarDoOutboxSW_(chave) {
  return abrirOutboxSW_().then(function (db) {
    return new Promise(function (resolve) {
      var tx = db.transaction(OUTBOX_STORE, 'readwrite');
      tx.objectStore(OUTBOX_STORE)['delete'](chave);
      tx.oncomplete = function () { resolve(); };
      tx.onerror = function () { resolve(); };
    });
  });
}

function enviarFilaPendente_() {
  return listarOutboxSW_().then(function (itens) {
    // Envia em sequência (não em paralelo) - o Apps Script não lida bem
    // com várias requisições simultâneas na mesma planilha.
    var promessa = Promise.resolve();
    itens.forEach(function (item) {
      promessa = promessa.then(function () {
        return fetch(item.url, {
          method: 'POST',
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify(item.payload)
        }).then(function (r) {
          if (!r.ok) throw new Error('HTTP ' + r.status);
          return r.text();
        }).then(function () {
          return apagarDoOutboxSW_(item.chave);
        }).catch(function (err) {
          console.log('[SW] Falha ao reenviar item da fila, tenta de novo depois:', err);
          // não apaga - fica na fila pra próxima tentativa
        });
      });
    });
    return promessa;
  });
}

self.addEventListener('sync', function (event) {
  if (event.tag === 'enviar-fila-registros') {
    event.waitUntil(enviarFilaPendente_());
  }
});
