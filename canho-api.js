// ===== MÁY CHỦ "QUẢN LÝ CĂN HỘ" (Apps Script riêng — xem apps-script/HUONG-DAN.md) =====
// Dán URL .../exec của máy chủ vào đây. Để trống = các trang dùng cách cũ (đọc Sheet công khai).
window.LUX_CANHO_API = "https://script.google.com/macros/s/AKfycbzaTUFKx22skc9TsRs1dXl8cKMgeBrDCsPBihN3bMtJ9hLvVecgnC7G0925BWYqbQp5gw/exec";

// Gọi máy chủ. GET cho lệnh đọc, POST (JSON dạng text/plain, không bị chặn CORS) cho lệnh ghi / tải ảnh.
window.luxApi = function (a, params, post, keepalive) {
  var API = window.LUX_CANHO_API;
  if (!API) return Promise.reject(new Error("noapi"));
  params = Object.assign({ a: a }, params || {});
  var req;
  if (post) {
    req = fetch(API, { method: "POST", body: JSON.stringify(params), keepalive: !!keepalive });   // keepalive: vẫn gửi xong khi tắt app (chỉ dùng cho lệnh nhỏ)
  } else {
    var qs = Object.keys(params).map(function (k) { return encodeURIComponent(k) + "=" + encodeURIComponent(params[k] == null ? "" : params[k]); }).join("&");
    req = fetch(API + "?" + qs + "&t=" + Date.now());
  }
  return req.then(function (r) { return r.text(); }).then(function (t) {
    var i = t.indexOf("{"), j = t.lastIndexOf("}");
    return JSON.parse(t.slice(i, j + 1));
  });
};

// ===== NHẬT KÝ NHÂN VIÊN =====
// Gọi sau mỗi thao tác thay đổi / gửi của NHÂN VIÊN (chỉ ghi khi trang đang chạy bằng token app phụ).
window.luxNvLog = function (type, detail, tok) {
  try {
    // chỉ ghi khi trang được mở từ app nhân viên (?tok=… trên đường dẫn) — tránh ghi nhầm khi chủ dùng máy từng đăng nhập NV
    tok = new URLSearchParams(location.search).get('tok') || '';
    if (!tok || !window.LUX_CANHO_API) return;
    var dev = '', ua = navigator.userAgent || '';
    try { dev = localStorage.getItem('nv_devid') || ''; } catch (e) {}
    var dn = (/iPhone/.test(ua) ? 'iPhone' : /iPad/.test(ua) ? 'iPad' : /Android/.test(ua) ? 'Android' : (navigator.platform || 'Máy'))
      + (/CriOS|Chrome/.test(ua) ? ' · Chrome' : /Safari/.test(ua) ? ' · Safari' : /Edg/.test(ua) ? ' · Edge' : '');
    window.luxApi('nvlog', { tok: tok, dev: dev, devname: dn, type: type, detail: String(detail || '').slice(0, 500),
      page: (location.pathname.split('/').pop() || '') }, true, true).catch(function () {});
  } catch (e) {}
};
