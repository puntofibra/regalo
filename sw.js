/* =========================================================
   REGALO · Service Worker (PWA instalable)
   ---------------------------------------------------------
   Sube la VERSION cada vez que cambies index/admin/estilos
   para que los móviles se traigan la versión nueva.
   ========================================================= */
var VERSION = 'regalo-v3';
var SHELL = [
  './',
  './index.html',
  './admin.html',
  './estilos.css',
  './api.js',
  './manifest.webmanifest',
  './icono.svg',
  './favicon-32.png',
  './icon-180.png',
  './icon-192.png',
  './icon-512.png',
  './icon-192-maskable.png',
  './icon-512-maskable.png'
];

self.addEventListener('install', function(e){
  e.waitUntil(
    caches.open(VERSION).then(function(c){
      /* uno a uno: si algún archivo falta, la instalación no se rompe */
      return Promise.all(SHELL.map(function(u){
        return c.add(new Request(u, {cache:'reload'})).catch(function(){});
      }));
    }).then(function(){ return self.skipWaiting(); })
  );
});

self.addEventListener('activate', function(e){
  e.waitUntil(
    caches.keys().then(function(ks){
      return Promise.all(ks.map(function(k){ return k === VERSION ? null : caches.delete(k); }));
    }).then(function(){ return self.clients.claim(); })
  );
});

self.addEventListener('message', function(e){
  if (e.data === 'skipWaiting') self.skipWaiting();
});

self.addEventListener('fetch', function(e){
  var req = e.request;
  if (req.method !== 'GET') return;

  var url;
  try { url = new URL(req.url); } catch(err){ return; }

  /* La API de Apps Script y las fuentes van siempre a la red: nunca se cachean */
  if (url.origin !== self.location.origin) return;

  /* HTML: red primero (así ves los cambios al instante), caché si no hay internet */
  if (req.mode === 'navigate' || (req.headers.get('accept')||'').indexOf('text/html') > -1){
    e.respondWith(
      /* no-store: evita que GitHub Pages nos sirva un HTML viejo de su caché */
      fetch(req.url, { cache: 'no-store', credentials: 'same-origin' }).then(function(r){
        var copia = r.clone();
        caches.open(VERSION).then(function(c){ c.put(req, copia); });
        return r;
      }).catch(function(){
        return caches.match(req).then(function(r){ return r || caches.match('./index.html'); });
      })
    );
    return;
  }

  /* Resto (css, js, iconos): caché primero y se refresca por detrás */
  e.respondWith(
    caches.match(req).then(function(cached){
      var red = fetch(new Request(req.url, { cache: 'no-store' })).then(function(r){
        var copia = r.clone();
        caches.open(VERSION).then(function(c){ c.put(req, copia); });
        return r;
      }).catch(function(){ return cached; });
      return cached || red;
    })
  );
});
