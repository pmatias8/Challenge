// Guarda os arquivos do app para abrir mesmo sem internet.
// (As consultas à API sempre vão para a rede.)
const VERSAO = "macros-v3";
const ARQUIVOS = ["./", "index.html", "estilo.css", "manifest.webmanifest", "icone.svg", "icone-180.png",
  "js/app.js", "js/xlsx.js", "js/vendas.js", "js/macros.js", "js/demo.js", "js/onedrive.js", "js/vendor/jszip.min.js", "js/vendor/msal-browser.min.js"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(VERSAO).then((c) => c.addAll(ARQUIVOS)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((ks) => Promise.all(ks.filter((k) => k !== VERSAO).map((k) => caches.delete(k))))
    .then(() => self.clients.claim()));
});
self.addEventListener("fetch", (e) => {
  const u = new URL(e.request.url);
  if (e.request.method !== "GET" || u.origin !== location.origin) return;
  // rede primeiro (pega atualizações), cache se estiver offline
  e.respondWith(fetch(e.request).then((r) => {
    if (r.ok) { const copia = r.clone(); caches.open(VERSAO).then((c) => c.put(e.request, copia)); }
    return r;
  }).catch(() => caches.match(e.request).then((r) => r || caches.match("index.html"))));
});
