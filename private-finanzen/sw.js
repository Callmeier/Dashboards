const CACHE = "finance-cockpit-v2";
const SHEETJS = "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js";
const MSAL = "https://alcdn.msftauth.net/browser/2.35.0/js/msal-browser.min.js";
const LOCAL_ASSETS = [
  "./",
  "./index.html",
  "./styles.css",
  "./onedrive.css",
  "./app.js",
  "./onedrive.js",
  "./manifest.webmanifest",
  "./icon.svg"
];

self.addEventListener("install", (event) => {
  event.waitUntil((async () => {
    const cache = await caches.open(CACHE);
    await cache.addAll(LOCAL_ASSETS);
    for (const asset of [SHEETJS, MSAL]) {
      try {
        const response = await fetch(asset, { mode: "no-cors" });
        await cache.put(asset, response);
      } catch (_) {
        // Externe Bibliotheken werden beim nächsten Online-Aufruf erneut geladen.
      }
    }
    self.skipWaiting();
  })());
});

self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    const keys = await caches.keys();
    await Promise.all(keys.filter((key) => key !== CACHE).map((key) => caches.delete(key)));
    self.clients.claim();
  })());
});

self.addEventListener("fetch", (event) => {
  if (event.request.method !== "GET") return;
  event.respondWith((async () => {
    const cached = await caches.match(event.request);
    if (cached) return cached;
    try {
      const response = await fetch(event.request);
      if (event.request.url.startsWith(self.location.origin)) {
        const cache = await caches.open(CACHE);
        cache.put(event.request, response.clone());
      }
      return response;
    } catch (error) {
      if (event.request.mode === "navigate") {
        const fallback = await caches.match("./index.html");
        if (fallback) return fallback;
      }
      throw error;
    }
  })());
});
