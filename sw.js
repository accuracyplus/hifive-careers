// Minimal service worker for the HiFive Careers page — enables
// "Add to Home Screen" install prompts and caches the page shell.
// Live job listings and applications always go straight to the network.
const CACHE_NAME = "hifive-careers-shell-v1";
const SHELL_FILES = ["/", "/index.html", "/logo.png", "/hero.jpg", "/manifest.json"];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES)).catch(() => {})
  );
  self.skipWaiting();
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
    )
  );
  self.clients.claim();
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  // Never cache API/worker calls — job listings and applications must always be live
  if (url.hostname.includes("workers.dev")) {
    return;
  }

  if (event.request.mode === "navigate") {
    event.respondWith(
      fetch(event.request).catch(() => caches.match("/index.html"))
    );
    return;
  }

  event.respondWith(
    caches.match(event.request).then((cached) => cached || fetch(event.request))
  );
});
