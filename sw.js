// Service worker mínimo — necessário para o app ser instalável.
// Não faz cache do index.html para nunca servir versão velha do app.
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', () => {});
