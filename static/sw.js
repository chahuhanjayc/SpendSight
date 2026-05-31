const SPENDSIGHT_CACHE = "spendsight-v1";

self.addEventListener("install", event => {
  event.waitUntil(
    caches.open(SPENDSIGHT_CACHE).then(cache => cache.addAll([
      "/",
      "/static/manifest.webmanifest"
    ])).catch(() => null)
  );
  self.skipWaiting();
});

self.addEventListener("activate", event => {
  event.waitUntil(
    caches.keys().then(keys => Promise.all(
      keys.filter(key => key !== SPENDSIGHT_CACHE).map(key => caches.delete(key))
    ))
  );
  self.clients.claim();
});

self.addEventListener("fetch", event => {
  const request = event.request;
  if (request.method !== "GET") return;
  const url = new URL(request.url);
  if (url.origin !== self.location.origin) return;

  event.respondWith(
    fetch(request)
      .then(response => {
        const copy = response.clone();
        caches.open(SPENDSIGHT_CACHE).then(cache => cache.put(request, copy)).catch(() => null);
        return response;
      })
      .catch(() => caches.match(request))
  );
});
