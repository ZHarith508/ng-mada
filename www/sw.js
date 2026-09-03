// ============================================================
//  SERVICE WORKER - NG-MADA
// ============================================================

const CACHE_NAME = 'ngmada-v2';

// Fichiers à mettre en cache
const FILES_TO_CACHE = [
  '/www/',
  '/www/index.html',
  '/www/manifest.json',
  '/www/PhotoUpload.html',
  '/www/ResetPassword.html'
];

// Installation
self.addEventListener('install', function(e) {
  e.waitUntil(
    caches.open(CACHE_NAME)
      .then(function(cache) {
        console.log('📦 Cache ouvert');
        return cache.addAll(FILES_TO_CACHE);
      })
      .catch(function(err) {
        console.error('❌ Erreur cache:', err);
      })
  );
  self.skipWaiting();
});

// Activation
self.addEventListener('activate', function(e) {
  e.waitUntil(
    caches.keys().then(function(cacheNames) {
      return Promise.all(
        cacheNames.map(function(cacheName) {
          if (cacheName !== CACHE_NAME) {
            console.log('🗑️ Suppression ancien cache:', cacheName);
            return caches.delete(cacheName);
          }
        })
      );
    }).then(function() {
      return self.clients.claim();
    })
  );
});

// Interception des requêtes
self.addEventListener('fetch', function(e) {
  // Ne pas intercepter les requêtes vers Google
  if (e.request.url.includes('google.com') || 
      e.request.url.includes('drive.google.com') ||
      e.request.url.includes('quickchart.io') ||
      e.request.url.includes('api.qrserver.com')) {
    e.respondWith(fetch(e.request));
    return;
  }

  e.respondWith(
    caches.match(e.request)
      .then(function(response) {
        // Si trouvé dans le cache, on le retourne
        if (response) {
          return response;
        }
        
        // Sinon, on va chercher sur le réseau
        return fetch(e.request)
          .then(function(networkResponse) {
            // Mettre en cache la réponse pour plus tard
            return caches.open(CACHE_NAME)
              .then(function(cache) {
                cache.put(e.request, networkResponse.clone());
                return networkResponse;
              });
          })
          .catch(function() {
            // En cas d'erreur réseau, retourner une page hors-ligne
            return new Response('Page non disponible hors-ligne', {
              status: 503,
              statusText: 'Service Unavailable'
            });
          });
      })
  );
});