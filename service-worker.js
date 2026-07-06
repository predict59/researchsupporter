const CACHE_NAME = "mw-price-survey-v24";
const BASE_PATH = new URL(self.registration.scope).pathname.replace(/\/$/, "");
const path = (value) => `${BASE_PATH}${value}`;
const APP_SHELL = [path("/"), path("/index.html"), path("/manifest.json"), path("/pwa-icon.svg")];

self.addEventListener("install", (event) => {
  event.waitUntil(caches.open(CACHE_NAME).then((cache) => cache.addAll(APP_SHELL)));
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((key) => key !== CACHE_NAME).map((key) => caches.delete(key)))),
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;
  const requestUrl = new URL(event.request.url);
  const isSameOrigin = requestUrl.origin === self.location.origin;
  if (!isSameOrigin) {
    event.respondWith(fetch(event.request));
    return;
  }
  event.respondWith(
    fetch(event.request)
        .then((response) => {
          const copy = response.clone();
          caches.open(CACHE_NAME).then((cache) => cache.put(event.request, copy));
          return response;
        })
        .catch(() => caches.match(event.request).then((cached) => {
          if (cached) return cached;
          if (isSameOrigin && event.request.mode === "navigate") return caches.match(path("/index.html"));
          return Response.error();
        })),
  );
});
