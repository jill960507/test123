// 山餃 PWA Service Worker
// 更新網站內容或圖示時，把 VERSION 改掉，訪客下次開啟就會換成新版。
const VERSION = 'v1';
const SHELL = `sanjiao-shell-${VERSION}`;
const RUNTIME = `sanjiao-runtime-${VERSION}`;
const SHELL_FILES = [
    './',
    './index.html',
    './manifest.webmanifest',
    './icons/icon-192.png',
    './icons/icon-512.png',
    './icons/icon-maskable-512.png',
    './icons/apple-touch-icon.png',
    './icons/favicon-32.png'
];

self.addEventListener('install', event => {
    event.waitUntil(caches.open(SHELL).then(cache => cache.addAll(SHELL_FILES)).then(() => self.skipWaiting()));
});

self.addEventListener('activate', event => {
    event.waitUntil(
        caches.keys()
            .then(keys => Promise.all(keys.filter(k => k !== SHELL && k !== RUNTIME).map(k => caches.delete(k))))
            .then(() => self.clients.claim())
    );
});

// 先回傳快取，同時在背景更新
const staleWhileRevalidate = async request => {
    const cache = await caches.open(RUNTIME);
    const cached = await caches.match(request);
    const network = fetch(request)
        .then(res => {
            if (res && (res.ok || res.type === 'opaque')) cache.put(request, res.clone());
            return res;
        })
        .catch(() => cached);
    return cached || network;
};

self.addEventListener('fetch', event => {
    const { request } = event;
    if (request.method !== 'GET') return;
    const url = new URL(request.url);
    if (!url.protocol.startsWith('http')) return;

    // 頁面：優先抓最新版，離線時用快取
    if (request.mode === 'navigate') {
        event.respondWith(
            fetch(request)
                .then(res => {
                    if (res.ok) {
                        const copy = res.clone();
                        caches.open(SHELL).then(cache => cache.put('./index.html', copy));
                    }
                    return res;
                })
                .catch(() => caches.match('./index.html'))
        );
        return;
    }

    // 其他資源（圖片、字型、CDN 程式庫）
    event.respondWith(staleWhileRevalidate(request));
});
