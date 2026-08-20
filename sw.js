// Bump this version whenever any shell file below changes — the browser only
// re-installs the service worker (and thus refreshes the cache) when sw.js
// itself differs byte-for-byte from what's already registered. Leaving this
// string unchanged means updates silently never reach installed PWAs.
const CACHE_NAME = "kakeibo-shell-v14";
const SHELL_FILES = ["./", "./index.html", "./app.js", "./manifest.json", "./icon.svg"];

self.addEventListener("install", (e) => {
  e.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES))
  );
  self.skipWaiting();
});

self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k)))
    )
  );
  self.clients.claim();
});

// Network-first for API calls (Sheets data must stay fresh), cache-first for shell files.
self.addEventListener("fetch", (e) => {
  const url = e.request.url;
  if (url.includes("googleapis.com") || url.includes("accounts.google.com")) {
    return; // never intercept auth/data calls
  }
  e.respondWith(
    caches.match(e.request).then((cached) => cached || fetch(e.request))
  );
});
