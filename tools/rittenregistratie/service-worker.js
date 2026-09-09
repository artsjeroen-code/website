const CACHE_NAME = 'rittenregistratie-shell-v13';
const APP_SHELL = [
  './',
  './index.html',
  './style.css',
  './app.js',
  './geocode.js',
  './routecheck.js',
  './audit.js',
  './pwa.js',
  './quick-capture.css',
  './quick-capture.js',
  './address-presets.css',
  './address-presets.js',
  './date-picker.css',
  './date-picker.js',
  './registration-policy.js',
  './manifest.webmanifest',
  './icons/icon-192.svg',
  './icons/icon-512.svg',
  '../../google_palette.css'
];

self.addEventListener('install', (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL))
  );
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(
      keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key))
    ))
  );
  self.clients.claim();
});

self.addEventListener('fetch', (event) => {
  const request = event.request;
  const url = new URL(request.url);

  if (request.method !== 'GET') return;

  if (url.origin === self.location.origin && (
    url.pathname.includes('/tools/rittenregistratie/api/') ||
    url.pathname.includes('/tools/rittenregistratie/quick/') ||
    url.pathname.endsWith('/tools/rittenregistratie/login.html') ||
    url.pathname.endsWith('/tools/rittenregistratie/passkey.js')
  )) {
    return;
  }

  if (url.origin !== self.location.origin) return;

  if (request.mode === 'navigate') {
    event.respondWith(
      fetch(request)
        .then((response) => {
          if (!response.redirected && response.ok) {
            const copy = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put('./index.html', copy));
          }
          return response;
        })
        .catch(() => caches.match('./index.html'))
    );
    return;
  }

  const isFrontendAsset = /\.(?:js|css)$/.test(url.pathname);
  if (isFrontendAsset) {
    event.respondWith(
      fetch(request)
        .then((response) => {
          if (response && response.status === 200 && response.type === 'basic') {
            const copy = response.clone();
            caches.open(CACHE_NAME).then((cache) => cache.put(request, copy));
          }
          return response;
        })
        .catch(() => caches.match(request))
    );
    return;
  }

  event.respondWith(
    caches.match(request).then((cached) => {
      if (cached) return cached;
      return fetch(request).then((response) => {
        if (!response || response.status !== 200 || response.type !== 'basic') return response;
        const copy = response.clone();
        caches.open(CACHE_NAME).then((cache) => cache.put(request, copy));
        return response;
      });
    })
  );
});
