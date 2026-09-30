/* Agenda de Trabalhos — recebe as notificações push mesmo com a app fechada */
const CFG_URL = new URL('./__cfg', self.location.href).href;
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));

self.addEventListener('push', e => e.waitUntil(mostrar()));
async function mostrar() {
  let titulo = 'Agenda de Trabalhos', corpo = 'Há novidades na agenda. Toque para abrir.';
  try {
    const r = await (await caches.open('agenda-cfg')).match(CFG_URL);
    if (r) {
      const cfg = await r.json();
      const j = await fetch(cfg.api, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ pin: cfg.pin, acao: 'aviso', dados: {} }) }).then(x => x.json());
      if (j.ok && j.r && j.r.titulo) { titulo = j.r.titulo; corpo = j.r.corpo; }
    }
  } catch (e) { /* sem rede: mostra o aviso genérico */ }
  return self.registration.showNotification(titulo, {
    body: corpo, icon: 'icon-192.png', badge: 'icon-192.png', tag: 'agenda-' + Date.now(),
    vibrate: [300, 100, 300, 100, 300], requireInteraction: true, data: { url: self.registration.scope }
  });
}

self.addEventListener('notificationclick', e => {
  e.notification.close();
  e.waitUntil(self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(ws => {
    for (const w of ws) if ('focus' in w) return w.focus();
    return self.clients.openWindow(e.notification.data.url);
  }));
});
