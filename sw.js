// Substitui o service worker do app antigo: limpa a copia guardada no celular e se desliga.
// Vistorias pendentes (localStorage/IndexedDB) NAO sao apagadas.
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (event) => {
  event.waitUntil((async () => {
    const nomes = await caches.keys();
    await Promise.all(nomes.map((n) => caches.delete(n)));
    await self.registration.unregister();
    const abas = await self.clients.matchAll({ type: "window" });
    abas.forEach((c) => c.navigate(c.url));
  })());
});
