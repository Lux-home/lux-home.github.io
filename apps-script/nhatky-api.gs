/**
 * LUXHOME — NHẬT KÝ NHÂN VIÊN (cùng dự án với canho-api.gs)
 * Mỗi lần nhân viên (mở app phụ bằng token) thay đổi / gửi gì, app gọi 'nvlog'.
 * Máy chủ xác minh token với máy chủ chính (ai là nhân viên), rồi ghi vào bảng
 * "LUXHOME - NHẬT KÝ NHÂN VIÊN" (thời gian, mã NV, tên NV, loại việc, nội dung, trang, thiết bị).
 */
const MAIN_EXEC = 'https://script.google.com/macros/s/AKfycbyFhO_xjflBHiofIY_XuQCcUNKd1VCwTUKhW5Jvaya51WBD5IId5owYfZM_Q9DY0aZoWQ/exec';
const NK_NAME = 'LUXHOME - NHẬT KÝ NHÂN VIÊN';
const NK_HDR = ['Thời gian', 'Mã NV', 'Tên NV', 'Loại việc', 'Nội dung', 'Trang', 'Thiết bị'];

function nkSheet_() {
  const P = PropertiesService.getScriptProperties();
  let id = P.getProperty('nk_sheet'), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) {} }
  if (!ss) {
    ss = SpreadsheetApp.create(NK_NAME);
    const sh = ss.getSheets()[0]; sh.appendRow(NK_HDR); sh.setFrozenRows(1);
    P.setProperty('nk_sheet', ss.getId());
  }
  return ss.getSheets()[0];
}
/** Hỏi máy chủ chính: token này là nhân viên nào (nhớ 30 phút). */
function nvWho_(tok, dev, devname) {
  tok = String(tok || ''); if (!tok) return null;
  const cache = CacheService.getScriptCache(), key = 'nvw_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, tok + '|' + dev));
  const hit = cache.get(key); if (hit) return hit === '-' ? null : JSON.parse(hit);
  let who = null;
  try {
    const t = UrlFetchApp.fetch(MAIN_EXEC + '?nvme=1&callback=cb&tok=' + encodeURIComponent(tok) + '&dev=' + encodeURIComponent(dev || '') + '&devname=' + encodeURIComponent(devname || ''), { muteHttpExceptions: true }).getContentText();
    const d = JSON.parse(t.slice(t.indexOf('(') + 1, t.lastIndexOf(')')));
    if (d && d.ok && !d.admin) who = { ma: d.ma, ten: d.ten || d.ma };
  } catch (e) {}
  cache.put(key, who ? JSON.stringify(who) : '-', who ? 1800 : 60);
  return who;
}
function nvLog_(p) {
  const who = nvWho_(p.tok, p.dev, p.devname);
  if (!who) return { ok: false, err: 'badtok' };
  const clip = (s, n) => String(s == null ? '' : s).replace(/[\r\n]+/g, ' ').slice(0, n);
  nkSheet_().appendRow([Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss'), who.ma, who.ten,
    clip(p.type, 60), "'" + clip(p.detail, 500), clip(p.page, 40), clip(p.devname, 60)]);
  return { ok: true };
}
/** Cho Admin xem: tối đa 3000 dòng mới nhất. */
function nvLogList_() {
  const id = PropertiesService.getScriptProperties().getProperty('nk_sheet');
  if (!id) return { ok: true, rows: [], url: '' };
  const ss = SpreadsheetApp.openById(id), sh = ss.getSheets()[0], n = sh.getLastRow();
  if (n < 2) return { ok: true, rows: [], url: ss.getUrl() };
  const from = Math.max(2, n - 2999);
  const v = sh.getRange(from, 1, n - from + 1, NK_HDR.length).getDisplayValues();
  return { ok: true, rows: v.reverse(), url: ss.getUrl(), total: n - 1 };
}
