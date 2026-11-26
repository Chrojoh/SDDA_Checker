const CACHE_NAME = "sdda-ledger-offline-ledger-brand-20261126";
const CORE_ASSETS = [
  "./",
  "./index.html",
  "./mobile-v2.html",
  "./app.js",
  "./styles.css",
  "./games.css",
  "./mobile-dashboard.css",
  "./vendor/xlsx.full.min.js",
  "./vendor/pdf-lib.min.js",
  "./manifest.webmanifest",
  "./icons/icon.svg",
  "./icons/ledger-192.png",
  "./icons/ledger-512.png"
];

self.addEventListener("install", event => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE_NAME);
    await cache.addAll(CORE_ASSETS);
    await self.skipWaiting();
  })());
});

self.addEventListener("activate", event => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter(key => key.startsWith("sdda-ledger-") && key !== CACHE_NAME).map(key => caches.delete(key)));
    await self.clients.claim();
  })());
});

self.addEventListener("fetch", event => {
  const request = event.request;
  if (request.method !== "GET") return;
  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;

  event.respondWith((async () => {
    const cache = await caches.open(CACHE_NAME);
    const fileFallback = new URL("./" + url.pathname.split("/").pop(), self.registration.scope).href;

    if (request.mode === "navigate") {
      try {
        const response = await fetch(request);
        if (response.ok) await cache.put(request, response.clone());
        return response;
      } catch (_) {
        return (await caches.match(request)) ||
          (await caches.match(url.pathname.endsWith("/mobile-v2.html") ? "./mobile-v2.html" : "./index.html")) ||
          (await caches.match("./"));
      }
    }

    if (request.destination === "script" || request.destination === "style") {
      try {
        const response = await fetch(request);
        if (response.ok) await cache.put(request, response.clone());
        return response;
      } catch (_) {
        return (await caches.match(request)) || (await caches.match(fileFallback)) || Response.error();
      }
    }

    const cached = await caches.match(request);
    if (cached) return cached;
    try {
      const response = await fetch(request);
      if (response.ok) await cache.put(request, response.clone());
      return response;
    } catch (_) {
      return Response.error();
    }
  })());
});
