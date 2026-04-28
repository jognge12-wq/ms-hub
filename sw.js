/* ================================================================
   施工管理ポータル Service Worker
   オフライン対応 + 弱電波環境での高速起動
   ================================================================

   キャッシュ戦略:
   - HTML/CSS/JS (静的アセット): Network First
       → まずネットを試みて、失敗したらキャッシュから返す
       → コード修正が即座に反映される（古いまま固まらない）
   - 画像/フォント: Cache First
       → キャッシュが優先、なければネットから取得
       → 表示が速い
   - GAS API (script.google.com/macros): キャッシュしない
       → リアルタイムデータなので常に最新を取得
   - Notion / Google 各種外部リンク: キャッシュしない

   更新方法:
   - 下の CACHE_VERSION を上げると旧キャッシュが一掃される
   - コードを大きく変えたらここのバージョンを上げること
================================================================ */

const CACHE_VERSION = 'v2-2026-04-28a';
const CACHE_NAME    = `portal-${CACHE_VERSION}`;

// 初回インストール時にプリキャッシュする最小セット
const PRECACHE_URLS = [
    './',
    './index.html',
    './task-management.html',
    './manifest.json',
    './icon.png',
    './icon-192.png',
    './apple-touch-icon.png',
    './assets/treeing.png',
    './tools/inspection.html',
    './tools/payment.html',
    './tools/property.html',
    './tools/_tool-common.css',
    './tools/_tool-common.js',
];

// ---- インストール: プリキャッシュ ----
self.addEventListener('install', event => {
    event.waitUntil(
        caches.open(CACHE_NAME).then(cache => {
            // 失敗しても全体を止めないように addAll ではなく個別に add
            return Promise.all(
                PRECACHE_URLS.map(url =>
                    cache.add(url).catch(err => {
                        console.warn('[sw] precache miss:', url, err && err.message);
                    })
                )
            );
        }).then(() => self.skipWaiting())
    );
});

// ---- アクティベート: 旧キャッシュを削除 ----
self.addEventListener('activate', event => {
    event.waitUntil(
        caches.keys().then(keys => Promise.all(
            keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k))
        )).then(() => self.clients.claim())
    );
});

// ---- fetch: リクエスト横取り ----
self.addEventListener('fetch', event => {
    const req = event.request;

    // GET 以外はそのまま通す
    if (req.method !== 'GET') return;

    const url = new URL(req.url);

    // --- 除外ルール: GAS / 外部 API はキャッシュしない ---
    if (
        url.hostname.includes('script.google.com') ||
        url.hostname.includes('script.googleusercontent.com') ||
        url.hostname.includes('notion.so') ||
        url.hostname.includes('notion.com') ||
        url.hostname.includes('calendar.google.com') ||
        url.hostname.includes('tasks.google.com') ||
        url.hostname.includes('api.notion.com')
    ) {
        return; // Service Worker で介入しない (ブラウザ標準の fetch)
    }

    // --- 画像 / フォント: Cache First ---
    const isImage = /\.(png|jpg|jpeg|gif|svg|webp|ico)$/i.test(url.pathname);
    const isFont  = /\.(woff2?|ttf|otf)$/i.test(url.pathname) ||
                    url.hostname.includes('fonts.gstatic.com');

    if (isImage || isFont) {
        event.respondWith(
            caches.match(req).then(cached => {
                if (cached) return cached;
                return fetch(req).then(res => {
                    if (res && res.status === 200) {
                        const copy = res.clone();
                        caches.open(CACHE_NAME).then(c => c.put(req, copy));
                    }
                    return res;
                }).catch(() => cached); // 最後のフォールバック
            })
        );
        return;
    }

    // --- それ以外 (HTML/CSS/JS など): Network First ---
    // navigate (HTML) は HTTP キャッシュをバイパスして常に最新を取得
    const fetchReq = req.mode === 'navigate'
        ? new Request(req, { cache: 'no-cache' })
        : req;
    event.respondWith(
        fetch(fetchReq).then(res => {
            if (res && res.status === 200 && url.origin === self.location.origin) {
                const copy = res.clone();
                caches.open(CACHE_NAME).then(c => c.put(req, copy));
            }
            return res;
        }).catch(() => caches.match(req).then(cached => {
            if (cached) return cached;
            // HTML リクエストがキャッシュにもネットにもなければ index.html で代用
            if (req.mode === 'navigate') {
                return caches.match('./index.html');
            }
            return new Response('', { status: 504, statusText: 'Offline and not cached' });
        }))
    );
});
