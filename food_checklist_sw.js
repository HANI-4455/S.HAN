/**
 * FOOD 제품 생산 Check-list — 서비스 워커
 *
 * 이 파일은 저장소 최상위에 있어 범위(scope)가 /S.HAN/ 전체다.
 * 허브의 다른 도구까지 가로채면 안 되므로, 이 앱이 쓰는 주소만 처리하고
 * 나머지 요청은 손대지 않고 그대로 통과시킨다.
 *
 * 앱을 고쳤는데 반영이 안 되면 CACHE 버전을 올린다.
 */
const CACHE = 'food-checklist-v1';

/* 이 앱의 자원만 캐시한다 */
const OWN = [
  'food_checklist.html',
  'food_checklist.webmanifest',
  'food_checklist_icon_180.png',
  'food_checklist_icon_192.png',
  'food_checklist_icon_512.png',
  'food_checklist_icon_maskable.png'
];

/* 폰트 CDN — 바뀌지 않으므로 캐시 우선 */
const FONT_HOSTS = ['cdn.jsdelivr.net', 'fastly.jsdelivr.net'];

self.addEventListener('install', e => {
  e.waitUntil(
    caches.open(CACHE)
      .then(c => c.addAll(OWN.map(f => new Request(f, { cache: 'reload' }))))
      .catch(() => {})           // 파일 하나가 없어도 설치는 계속한다
      .then(() => self.skipWaiting())
  );
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys()
      .then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;                     // API 는 POST — 손대지 않는다
  const url = new URL(req.url);

  if (FONT_HOSTS.includes(url.hostname)) {
    e.respondWith(cacheFirst(req));
    return;
  }
  if (url.origin !== self.location.origin) return;      // Apps Script 등 외부 요청은 통과

  const file = url.pathname.split('/').pop();
  if (!OWN.includes(file)) return;                      // 허브의 다른 도구는 통과

  // 화면은 항상 최신을 먼저 시도하고, 실패하면 캐시로 띄운다
  e.respondWith(networkFirst(req));
});

async function networkFirst(req) {
  try {
    const res = await fetch(req);
    if (res && res.ok) (await caches.open(CACHE)).put(req, res.clone());
    return res;
  } catch (err) {
    const hit = await caches.match(req);
    if (hit) return hit;
    throw err;
  }
}

async function cacheFirst(req) {
  const hit = await caches.match(req);
  if (hit) return hit;
  const res = await fetch(req);
  if (res && (res.ok || res.type === 'opaque')) (await caches.open(CACHE)).put(req, res.clone());
  return res;
}
