/**
 * DC Fleet Service Worker
 * - 기본 PWA 기능 (오프라인 캐시 일부)
 * - 푸시 알림 수신 (FCM 통합 시 사용)
 * - 백그라운드 동기화 (추후 사용)
 *
 * 캐시 정책 — Stale-while-revalidate
 *   • HTML/CSS/JS는 항상 네트워크 우선 (최신 코드)
 *   • 이미지는 캐시 우선 (성능)
 *   • API 호출은 캐시 안 함 (실시간 데이터)
 */

const CACHE_NAME = 'dc-fleet-v1';
const STATIC_ASSETS = [
  // GitHub Pages 환경에서 절대경로
  // 실제 캐시는 첫 방문 후 자동으로
];

// 설치 — 즉시 활성화
self.addEventListener('install', (event) => {
  console.log('[SW] Install');
  self.skipWaiting();
});

// 활성화 — 오래된 캐시 삭제
self.addEventListener('activate', (event) => {
  console.log('[SW] Activate');
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys.filter((k) => k !== CACHE_NAME).map((k) => caches.delete(k))
      )
    ).then(() => self.clients.claim())
  );
});

// 네트워크 요청 처리 (현재는 통과만 — 캐시 비활성)
self.addEventListener('fetch', (event) => {
  const url = new URL(event.request.url);
  // GAS API 호출은 캐시 안 함
  if (url.hostname.includes('script.google.com')) return;
  // GitHub raw/api 호출도 캐시 안 함
  if (url.hostname.includes('github')) return;
  // 일반 정적 자원 — 네트워크 우선, 실패 시 캐시
  if (event.request.method === 'GET' && url.origin === self.location.origin) {
    event.respondWith(
      fetch(event.request).catch(() => caches.match(event.request))
    );
  }
});

// 푸시 알림 수신 (FCM 통합 후 작동)
self.addEventListener('push', (event) => {
  if (!event.data) return;
  let data = {};
  try { data = event.data.json(); }
  catch (e) { data = { title: 'DC Fleet', body: event.data.text() }; }

  const title = data.title || '🚌 DC Fleet 알림';
  const options = {
    body: data.body || '',
    icon: data.icon || '/dongchoi-driver/icon-192.png',
    badge: data.badge || '/dongchoi-driver/icon-192.png',
    tag: data.tag || 'dc-fleet',
    requireInteraction: data.requireInteraction || false,
    data: data.data || {},
  };
  event.waitUntil(self.registration.showNotification(title, options));
});

// 알림 클릭 처리
self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  const url = event.notification.data?.url || '/dongchoi-driver/';
  event.waitUntil(
    clients.matchAll({ type: 'window' }).then((clientList) => {
      // 이미 열린 탭 있으면 포커스
      for (const client of clientList) {
        if (client.url.includes('dongchoi-driver') && 'focus' in client) {
          return client.focus();
        }
      }
      // 없으면 새 창
      if (clients.openWindow) return clients.openWindow(url);
    })
  );
});

// 메시지 (앱에서 보낸 명령 처리)
self.addEventListener('message', (event) => {
  if (event.data?.type === 'SKIP_WAITING') self.skipWaiting();
  if (event.data?.type === 'CLEAR_CACHE') {
    caches.keys().then((keys) => keys.forEach((k) => caches.delete(k)));
  }
});
