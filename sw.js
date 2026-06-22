// Service worker with explicit update flow — enables "Add to Home Screen"
// install prompts, caches the app shell for instant offline-capable loads,
// and surfaces new versions to the user instead of silently auto-updating.
//
// IMPORTANT: bump this string on every deploy that should trigger the
// "Update Available" prompt. The browser only checks for a new service worker
// when this file's bytes differ from what's currently installed — bumping the
// version string is the simplest reliable way to guarantee that.
const SW_VERSION  = "2026-06-22.1";
const CACHE_NAME  = `hifive-hr-shell-${SW_VERSION}`;
const SHELL_FILES = ["/", "/index.html", "/logo.png", "/manifest.json"];

self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(CACHE_NAME).then((cache) => cache.addAll(SHELL_FILES)).catch(() => {})
  );
  // Deliberately NOT calling self.skipWaiting() here — the new worker stays
  // "waiting" until the page explicitly tells it to activate (via the
  // SKIP_WAITING message below), which is what makes the update prompt
  // meaningful rather than the page silently changing underneath the user.
});

self.addEventListener("activate", (event) => {
  event.waitUntil(
    caches.keys().then((names) =>
      Promise.all(names.filter((n) => n !== CACHE_NAME).map((n) => caches.delete(n)))
    )
  );
  self.clients.claim();
});

// Page sends this once the user taps "Update Now" in the prompt.
self.addEventListener("message", (event) => {
  if (event.data === "SKIP_WAITING") {
    self.skipWaiting();
  }
});

self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  // Never cache API/worker calls — recruitment data must always be live
  if (url.hostname.includes("workers.dev") || url.pathname.startsWith("/api")) {
    return;
  }

  // Network-first for navigation, falling back to cached shell if offline
  if (event.request.mode === "navigate") {
    event.respondWith(
      fetch(event.request).catch(() => caches.match("/index.html"))
    );
    return;
  }

  // Cache-first for static shell assets
  event.respondWith(
    caches.match(event.request).then((cached) => cached || fetch(event.request))
  );
});
