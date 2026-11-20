// LUXHOME app tổng — service worker (shell cache cho cài đặt PWA)
const CACHE = "luxhome-v44";
const SHELL = ["./index.html", "./luutru.html", "./bill.html", "./nv.html", "./nvadmin.html", "./thongbao.html", "./noibo.html", "./hopdong.html", "./congno.html", "./catalog.html", "./manifest.webmanifest", "./manifest-nv.webmanifest", "./icon-192.png", "./icon-512.png"];

// Cài đặt: chiếm quyền NGAY, nạp shell từng file (một file lỗi do mạng yếu không làm hỏng cập nhật)
self.addEventListener("install", e => {
  self.skipWaiting();
  e.waitUntil(caches.open(CACHE).then(c => Promise.all(SHELL.map(u => c.add(u).catch(() => {})))));
});
self.addEventListener("activate", e => {
  e.waitUntil(
    caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});
// Luôn lấy bản mới nhất từ mạng (bỏ qua cache HTTP), chỉ dùng cache khi offline
self.addEventListener("fetch", e => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.origin !== location.origin) return;
  e.respondWith(
    fetch(e.request, { cache: "no-store" }).then(r => {
      const cp = r.clone();
      caches.open(CACHE).then(c => c.put(e.request, cp)).catch(() => {});
      return r;
    }).catch(() => caches.match(e.request).then(m => m || caches.match("./index.html")))
  );
});
