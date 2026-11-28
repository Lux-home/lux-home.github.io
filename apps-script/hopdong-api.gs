/**
 * LUXHOME — KHO HỢP ĐỒNG (cùng dự án với canho-api.gs …)
 * Mục 3 "Hợp đồng PDF" (check-in) nạp file → app tải file PDF gốc lên Drive:
 *   LUXHOME - HỢP ĐỒNG / <Tòa> / Tháng <m>-<yyyy> (tháng check-in) / P<phòng> / <tên file>.pdf   (vd T121 / Tháng 11-2026 / P301)
 * Ô "Tạo mới + lưu trữ HĐ" → "Lưu HĐ + PL" duyệt / tìm / mở các file này.
 */
const HD_ROOT_NAME = 'LUXHOME - HỢP ĐỒNG';
const CK_EXEC = 'https://script.google.com/macros/s/AKfycbylsYx2qOuQgp92bb5o_Q0Pg7K3I6FYfjvhlyBxUvBYgcE5yOOBth3PyniM3wgGtT4y/exec';

function hdRoot_() {
  const P = PropertiesService.getScriptProperties(), id = P.getProperty('hd_root');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const it = DriveApp.getFoldersByName(HD_ROOT_NAME), f = it.hasNext() ? it.next() : DriveApp.createFolder(HD_ROOT_NAME);
  P.setProperty('hd_root', f.getId());
  return f;
}
/** Ai đang gọi: chủ (mật khẩu Admin hoặc mật khẩu duyệt check-in) / nhân viên (token app phụ). */
function hdAuth_(p, need) {
  if (roleOf_(p.pw) === 'admin') return { ok: true, owner: true, ma: 'CHỦ', ten: 'Chủ' };
  if (p.ckpw) {
    const cache = CacheService.getScriptCache(), key = 'ck_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(p.ckpw)));
    let ok = cache.get(key);
    if (ok == null) {
      try {
        const t = UrlFetchApp.fetch(CK_EXEC + '?me=1&pw=' + encodeURIComponent(p.ckpw), { muteHttpExceptions: true }).getContentText();
        const d = JSON.parse(t.slice(t.indexOf('{'), t.lastIndexOf('}') + 1));
        ok = d && d.ok && d.owner ? '1' : '0';
      } catch (e) { ok = '0'; }
      cache.put(key, ok, ok === '1' ? 1800 : 60);
    }
    if (ok === '1') return { ok: true, owner: true, ma: 'CHỦ', ten: 'Chủ' };
  }
  if (p.tok) {
    const w = nvWho_(p.tok, p.dev, p.devname);
    if (w && (w.perms || []).some(x => need.indexOf(x) >= 0)) return { ok: true, owner: false, ma: w.ma, ten: w.ten };
  }
  return null;
}
function hdChild_(parent, name) { const it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); }
function hdClean_(s) { return String(s || '').replace(/[\\/:*?"<>|]/g, '').trim(); }

/** Mở phiên tải PDF lên đúng thư mục; nếu đã có file cùng tên thì báo đã có. */
function hdSession_(p) {
  const who = hdAuth_(p, ['checkin', 'checkinout', 'hopdong']);
  if (!who) return { ok: false, err: 'noperm', msg: 'Không có quyền lưu hợp đồng' };
  const toa = hdClean_(p.toa).toUpperCase(), thang = hdClean_(p.thang), phong = hdClean_(p.phong).toUpperCase(), name = hdClean_(p.name);
  if (!toa || !/^\d{4}-\d{2}$/.test(thang) || !phong || !name) return { ok: false, err: 'need', msg: 'Thiếu tòa / tháng / phòng / tên file' };
  const tm = thang.split('-'), lv2 = 'Tháng ' + (+tm[1]) + '-' + tm[0], lv3 = /^P/.test(phong) ? phong : 'P' + phong;
  const f = hdChild_(hdChild_(hdChild_(hdRoot_(), toa), lv2), lv3);
  const ex = f.getFilesByName(name);
  let tt = '', dup = null;
  try { const rr = tdRecord_(name, String(p.room || (phong + '-' + toa)).toUpperCase(), toa, lv2, p.start || '', who.owner ? 'Chủ' : who.ten); tt = rr.st; dup = rr.dup; } catch (e) {}
  while (ex.hasNext()) { const x = ex.next(); if (!x.isTrashed() && (!p.size || x.getSize() === +p.size)) return { ok: true, exists: true, id: x.getId(), path: toa + ' / ' + lv2 + ' / ' + lv3, tt: tt, dup: dup }; }
  const res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&supportsAllDrives=true', {
    method: 'post', contentType: 'application/json; charset=UTF-8', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'X-Upload-Content-Type': 'application/pdf', 'X-Upload-Content-Length': String(p.size || ''),
      Origin: p.origin || 'https://lux-home.github.io' },
    payload: JSON.stringify({ name: name, parents: [f.getId()] })
  });
  const loc = res.getAllHeaders()['Location'] || res.getAllHeaders()['location'];
  if (res.getResponseCode() !== 200 || !loc) return { ok: false, err: 'session', msg: res.getContentText().slice(0, 200) };
  if (!who.owner) { try { nkAppend_(who.ma, who.ten, 'Lưu hợp đồng PDF', toa + ' / ' + lv2 + ' / ' + lv3 + ' · ' + name, 'checkin.html', p.devname); } catch (e) {} }
  return { ok: true, url: loc, path: toa + ' / ' + lv2 + ' / ' + lv3, tt: tt, dup: dup };
}
/** Duyệt kho: path = [] (tòa) | [tòa] (tháng) | [tòa, tháng] (phòng) | [tòa, tháng, phòng] (file) */
function hdList_(p) {
  const who = hdAuth_(p, ['hopdong']);
  if (!who) return { ok: false, err: 'noperm', msg: 'Chưa được cấp quyền Tạo mới + lưu trữ HĐ' };
  let f = hdRoot_();
  const path = Array.isArray(p.path) ? p.path : String(p.path || '').split('/').filter(String);
  for (const name of path) { const it = f.getFoldersByName(name); if (!it.hasNext()) return { ok: false, err: 'notfound' }; f = it.next(); }
  const folders = [], files = [];
  const fi = f.getFolders(); while (fi.hasNext()) { const x = fi.next(); folders.push({ name: x.getName(), id: x.getId() }); }
  const fl = f.getFiles(); while (fl.hasNext()) { const x = fl.next(); if (!x.isTrashed()) files.push({ name: x.getName(), id: x.getId(), size: x.getSize(), at: Utilities.formatDate(x.getDateCreated(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm') }); }
  // cấp "Tháng 11-2026": tháng mới nhất trước; tòa / phòng: theo số
  const mk = n => { const m = String(n).match(/^Tháng (\d+)-(\d{4})$/); return m ? +m[2] * 100 + (+m[1]) : 0; };
  folders.sort((a, b) => path.length === 1 ? mk(b.name) - mk(a.name) : a.name.localeCompare(b.name, 'vi', { numeric: true }));
  files.sort((a, b) => a.name.localeCompare(b.name));
  return { ok: true, path: path, folders: folders, files: files };
}
/** Tìm file trong kho theo tên (phòng, tòa, môi giới…). */
function hdSearch_(p) {
  const who = hdAuth_(p, ['hopdong']);
  if (!who) return { ok: false, err: 'noperm' };
  const q = String(p.q || '').replace(/'/g, '').trim(); if (q.length < 2) return { ok: true, files: [] };
  const rootId = hdRoot_().getId(), out = [];
  const it = DriveApp.searchFiles("title contains '" + q + "' and trashed = false and mimeType = 'application/pdf'");
  while (it.hasNext() && out.length < 60) {
    const x = it.next(), chain = [];
    let par = x.getParents(), ok = false;
    for (let d = 0; d < 4 && par.hasNext(); d++) { const pf = par.next(); if (pf.getId() === rootId) { ok = true; break; } chain.unshift(pf.getName()); par = pf.getParents(); }
    if (ok) out.push({ name: x.getName(), id: x.getId(), size: x.getSize(), path: chain.join(' / '), at: Utilities.formatDate(x.getDateCreated(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy') });
  }
  return { ok: true, files: out };
}
/** Mở 1 file (chỉ file nằm trong kho hợp đồng). */
function hdGet_(p) {
  const who = hdAuth_(p, ['hopdong']);
  if (!who) return { ok: false, err: 'noperm' };
  const x = DriveApp.getFileById(p.id), rootId = hdRoot_().getId();
  let par = x.getParents(), inRoot = false;
  for (let d = 0; d < 4 && par.hasNext() && !inRoot; d++) { const pf = par.next(); if (pf.getId() === rootId) inRoot = true; else par = pf.getParents(); }
  if (!inRoot) return { ok: false, err: 'noperm' };
  if (x.getSize() > 30 * 1024 * 1024) return { ok: false, err: 'big', msg: 'File lớn hơn 30MB — mở trên Google Drive' };
  if (!who.owner) { try { nkAppend_(who.ma, who.ten, 'Xem hợp đồng', x.getName(), 'hdkho.html', p.devname); } catch (e) {} }
  return { ok: true, name: x.getName(), mime: 'application/pdf', data: Utilities.base64Encode(x.getBlob().getBytes()) };
}

/* ======================= THEO DÕI THANH TOÁN ĐỢT 1 CỦA HỢP ĐỒNG =======================
 * Mỗi hợp đồng nạp ở "Hợp đồng checkin PDF" được ghi vào bảng "LUXHOME - THEO DÕI THANH TOÁN ĐỢT 1"
 * với trạng thái "Chưa kiểm tra". Khi chủ chạy mục 3 "Kiểm tra thanh toán đợt 1 khi checkin",
 * kết quả từng phòng được ghi lại: Đạt / Chưa đúng (kèm lý do) / Tạm đạt theo bill (chưa có sao kê).
 */
const TD_NAME = 'LUXHOME - THEO DÕI THANH TOÁN ĐỢT 1';
const TD_HDR = ['Tên file', 'Phòng', 'Tòa', 'Tháng check-in', 'Ngày vào', 'Lưu lúc', 'Người nạp', 'Trạng thái', 'Chi tiết', 'Kiểm tra lúc', 'Bổ sung file QLKH', 'Dữ liệu hợp đồng'];
function tdSheet_() {
  const P = PropertiesService.getScriptProperties();
  let id = P.getProperty('td_sheet'), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) {} }
  if (!ss) { ss = SpreadsheetApp.create(TD_NAME); const sh = ss.getSheets()[0]; sh.appendRow(TD_HDR); sh.setFrozenRows(1); P.setProperty('td_sheet', ss.getId()); }
  const sh = ss.getSheets()[0];
  if (sh.getLastColumn() < TD_HDR.length) sh.getRange(1, 1, 1, TD_HDR.length).setValues([TD_HDR]);   // thêm cột mới cho bảng cũ
  return sh;
}
const tdNow_ = () => Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm');
/** Ghi nhật ký nạp hợp đồng. Trả { st, dup } — dup: lần nạp trước (trùng file, hoặc trùng phòng + ngày vào) */
function tdRecord_(name, room, toa, thang, start, who) {
  const sh = tdSheet_(), n = sh.getLastRow(), v = n > 1 ? sh.getRange(2, 1, n - 1, TD_HDR.length).getDisplayValues() : [];
  const info = r => ({ file: r[0], room: r[1], thang: r[3], start: r[4], at: r[5], by: r[6], st: r[7] || 'Chưa kiểm tra', qlkh: r[10] || '' });
  for (let i = 0; i < v.length; i++) if (v[i][0] === name) return { st: v[i][7] || 'Chưa kiểm tra', dup: info(v[i]) };   // đã có → giữ trạng thái
  const same = v.find(r => String(r[1]).toUpperCase() === String(room).toUpperCase() && start && r[4] === start);
  sh.appendRow(["'" + name, "'" + room, "'" + toa, "'" + thang, "'" + (start || ''), tdNow_(), who, 'Chưa kiểm tra', '', '', '', '']);
  return { st: 'Chưa kiểm tra', dup: same ? info(same) : null };
}
/** Đánh dấu hợp đồng đã bổ sung vào file Quản lý khách hàng */
function tdMarkAdd_(file, text) {
  try {
    const sh = tdSheet_(), n = sh.getLastRow(); if (n < 2 || !file) return;
    const v = sh.getRange(2, 1, n - 1, 1).getDisplayValues();
    for (let i = v.length - 1; i >= 0; i--) if (v[i][0] === file) { sh.getRange(i + 2, 11).setValue(text); return; }
  } catch (e) {}
}
function tdList_(p) {
  const who = hdAuth_(p, ['checkin', 'checkinout', 'hopdong']);
  if (!who) return { ok: false, err: 'noperm' };
  const sh = tdSheet_(), n = sh.getLastRow();
  const rows = n > 1 ? sh.getRange(2, 1, n - 1, TD_HDR.length).getDisplayValues().map(r => ({
    file: r[0], room: r[1], toa: r[2], thang: r[3], start: r[4], at: r[5], by: r[6], st: r[7] || 'Chưa kiểm tra', detail: r[8], checked: r[9], qlkh: r[10] || '' })) : [];
  return { ok: true, rows: rows.reverse(), url: SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('td_sheet')).getUrl() };
}
/** Chủ gửi kết quả kiểm tra: results = [{room, ngay, ok, detail, sk}] */
function tdSave_(p) {
  const who = hdAuth_(p, ['checkin', 'checkinout', 'hopdong']);     // nhân viên cũng chạy được bước 3 (đối chiếu sao kê)
  if (!who) return { ok: false, err: 'noperm' };
  const sh = tdSheet_(), n = sh.getLastRow(); if (n < 2) return { ok: true, n: 0 };
  const v = sh.getRange(2, 1, n - 1, TD_HDR.length).getDisplayValues();
  const norm = s => String(s || '').toUpperCase().replace(/\s/g, '');
  const day = s => { const m = String(s || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/); return m ? (+m[3]) * 10000 + (+m[2]) * 100 + (+m[1]) : 0; };
  let cnt = 0;
  (p.results || []).forEach(r => {
    const room = norm(r.room), d = day(r.ngay);
    let idx = -1;
    for (let i = 0; i < v.length; i++) if (norm(v[i][1]) === room && (!d || !day(v[i][4]) || day(v[i][4]) === d)) { idx = i; if (day(v[i][4]) === d) break; }
    if (idx < 0) return;
    const st = !r.ok ? 'Chưa đúng' : (r.sk ? 'Đạt' : 'Tạm đạt (theo bill, chưa có sao kê)');
    if (v[idx][7] === 'Đạt' && st !== 'Đạt' && !r.sk) return;          // đã đạt theo sao kê thì không hạ xuống vì lần kiểm tra thiếu sao kê
    v[idx][7] = st; v[idx][8] = String(r.detail || '').slice(0, 480); v[idx][9] = tdNow_();
    sh.getRange(idx + 2, 8, 1, 3).setValues([[v[idx][7], v[idx][8], v[idx][9]]]);
    cnt++;
  });
  return { ok: true, n: cnt };
}

/* ======================= EMAIL BÁO CÁO HẰNG NGÀY: THANH TOÁN ĐỢT 1 =======================
 * Giống email "Hạn visa / lưu trú": danh sách người nhận + các giờ gửi trong ngày (mặc định 07:00).
 * Trình kích hoạt chạy mỗi 15 phút (tdTick) và gửi khi tới giờ đã đặt (mỗi giờ 1 lần/ngày).
 */
/** Lấy danh sách email + giờ gửi của email cảnh báo Lưu trú / visa (máy chủ chính) — dùng chung 1 nơi quản lý.
 *  Cần thuộc tính dự án lt_cfg_pw (Cài đặt dự án → Thuộc tính tập lệnh), KHÔNG ghi trong mã. */
function tdLuuTruCfg_() {
  const P = PropertiesService.getScriptProperties(), C = CacheService.getScriptCache(), k = 'td_lt_cfg'; let v = C.get(k);
  if (v) return JSON.parse(v);
  const pw = P.getProperty('lt_cfg_pw'), out = {};
  if (pw) try {
    ['get', 'getsched'].forEach(a => {
      const t = UrlFetchApp.fetch(MAIN_EXEC + '?cfg=' + a + '&pw=' + encodeURIComponent(pw), { muteHttpExceptions: true }).getContentText();
      const d = JSON.parse(t.replace(/^[^(]*\(/, '').replace(/\);?\s*$/, ''));
      if (d && d.ok) { if (d.emails) out.emails = d.emails; if (d.times) out.times = d.times; }
    });
  } catch (e) {}
  if (out.emails || out.times) { C.put(k, JSON.stringify(out), 600); P.setProperty('td_lt_last', JSON.stringify(out)); return out; }
  try { return JSON.parse(P.getProperty('td_lt_last') || '{}'); } catch (e) { return {}; }
}
function tdCfg_() {
  const P = PropertiesService.getScriptProperties();
  if (P.getProperty('td_follow') !== '0') {   // mặc định: theo email + giờ của báo cáo Lưu trú / visa
    const lt = tdLuuTruCfg_();
    return { follow: true, linked: !!lt.emails, emails: lt.emails || [Session.getEffectiveUser().getEmail()], times: (lt.times && lt.times.length) ? lt.times : ['07:00', '21:00'],
      trigger: ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'tdTick') };
  }
  let em = []; try { em = JSON.parse(P.getProperty('td_mail') || 'null'); } catch (e) {}
  if (!em) em = [Session.getEffectiveUser().getEmail()].filter(String);
  let tm = []; try { tm = JSON.parse(P.getProperty('td_times') || 'null'); } catch (e) {}
  if (!tm || !tm.length) tm = ['07:00'];
  return { emails: em, times: tm, trigger: ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'tdTick') };
}
function tdMailCfg_(p) {
  const who = hdAuth_(p, []);
  if (!who || !who.owner) return { ok: false, err: 'noperm' };
  const P = PropertiesService.getScriptProperties();
  if (p.follow != null) P.setProperty('td_follow', p.follow ? '1' : '0');
  if (p.emails) P.setProperty('td_mail', JSON.stringify([].concat(p.emails).map(s => String(s).trim()).filter(s => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(s)).slice(0, 30)));
  if (p.times) P.setProperty('td_times', JSON.stringify([].concat(p.times).map(s => String(s).trim()).filter(s => /^\d{2}:\d{2}$/.test(s)).sort().slice(0, 12)));
  if (p.test) { const r = tdSendReport_(true); return Object.assign({ ok: true, sent: r }, tdCfg_()); }
  return Object.assign({ ok: true }, tdCfg_());
}
/** Chạy 1 lần trong trình soạn thảo: tạo trình kích hoạt 15 phút + gửi email thử. */
function tdCaiDatEmail() {
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'tdTick'))
    ScriptApp.newTrigger('tdTick').timeBased().everyMinutes(15).create();
  Logger.log(tdSendReport_(true));
}
function tdTick() {
  const P = PropertiesService.getScriptProperties(), cfg = tdCfg_();
  const now = new Date(), tz = 'Asia/Ho_Chi_Minh', day = Utilities.formatDate(now, tz, 'yyyy-MM-dd'), hm = Utilities.formatDate(now, tz, 'HH:mm');
  let sent = {}; try { sent = JSON.parse(P.getProperty('td_sent') || '{}'); } catch (e) {}
  if (sent.day !== day) sent = { day: day, t: [] };
  const due = cfg.times.filter(t => t <= hm && sent.t.indexOf(t) < 0);
  if (!due.length) return;
  sent.t = sent.t.concat(due); P.setProperty('td_sent', JSON.stringify(sent));
  tdSendReport_(false);
}
function tdSendReport_(force) {
  const cfg = tdCfg_(); if (!cfg.emails.length) return 'Chưa có email nhận';
  const sh = tdSheet_(), n = sh.getLastRow();
  const rows = n > 1 ? sh.getRange(2, 1, n - 1, TD_HDR.length).getDisplayValues() : [];
  const R = rows.map(r => ({ file: r[0], room: r[1], start: r[4], at: r[5], by: r[6], st: r[7] || 'Chưa kiểm tra', detail: r[8], checked: r[9] }));
  const sai = R.filter(x => /^Chưa đúng/.test(x.st)), chua = R.filter(x => x.st === 'Chưa kiểm tra'), tam = R.filter(x => /^Tạm/.test(x.st));
  const tong = sai.length + chua.length + tam.length;
  if (!tong && !force) return 'Không có hợp đồng cần báo';
  const e = s => String(s || '').replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
  const ngayNap = s => { const m = String(s).match(/^(\d{2})\/(\d{2})\/(\d{4})/); return m ? Math.floor((Date.now() - new Date(+m[3], +m[2] - 1, +m[1])) / 864e5) : ''; };
  const bang = (ds, mau, cotCT) => '<table style="border-collapse:collapse;font-size:13px;width:100%"><tr style="background:#f1f5f9">'
    + '<th style="text-align:left;padding:5px;border:1px solid #ddd">Phòng</th><th style="text-align:left;padding:5px;border:1px solid #ddd">Ngày vào</th><th style="text-align:left;padding:5px;border:1px solid #ddd">Nạp lúc / bởi</th><th style="text-align:left;padding:5px;border:1px solid #ddd">' + cotCT + '</th></tr>'
    + ds.map(x => '<tr><td style="padding:5px;border:1px solid #ddd;font-weight:bold;color:' + mau + '">' + e(x.room) + '</td><td style="padding:5px;border:1px solid #ddd">' + e(x.start) + '</td><td style="padding:5px;border:1px solid #ddd">' + e(x.at) + (x.by ? ' · ' + e(x.by) : '') + '</td><td style="padding:5px;border:1px solid #ddd">'
      + (cotCT === 'Đã chờ' ? (ngayNap(x.at) !== '' ? ngayNap(x.at) + ' ngày' : '') : e(x.detail)) + '</td></tr>').join('') + '</table>';
  let h = '<div style="font-family:Arial,sans-serif;max-width:760px"><h2 style="color:#b8860b;margin:0 0 4px">LUXHOME · Theo dõi thanh toán đợt 1 khi check-in</h2>'
    + '<p style="margin:0 0 12px;color:#555">' + Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm') + ' — ' + (tong ? '<b>' + tong + '</b> hợp đồng chưa đạt kiểm tra' : '✅ Tất cả hợp đồng đã đạt') + '</p>';
  if (sai.length) h += '<h3 style="color:#b91c1c">⛔ Thanh toán chưa đúng (' + sai.length + ')</h3>' + bang(sai, '#b91c1c', 'Lý do');
  if (chua.length) h += '<h3 style="color:#b45309">⚠️ Chưa qua bước kiểm tra thanh toán đợt 1 (' + chua.length + ')</h3>' + bang(chua, '#b45309', 'Đã chờ');
  if (tam.length) h += '<h3 style="color:#a16207">⚠️ Tạm đạt theo bill, chưa đối chiếu sao kê (' + tam.length + ')</h3>' + bang(tam, '#a16207', 'Chi tiết');
  h += '<p style="margin-top:14px"><a href="https://lux-home.github.io/checkin.html?mode=tt">Mở app → Kiểm tra thanh toán đợt 1 khi checkin</a></p></div>';
  MailApp.sendEmail({ to: cfg.emails.join(','), subject: 'LUXHOME · Thanh toán đợt 1: ' + (tong ? sai.length + ' chưa đúng, ' + chua.length + ' chưa kiểm tra' + (tam.length ? ', ' + tam.length + ' tạm đạt' : '') : 'tất cả đạt'), htmlBody: h, name: 'LUXHOME' });
  return 'Đã gửi tới ' + cfg.emails.join(', ');
}


/* ===== Hợp đồng nhân viên nạp hằng ngày: lưu dữ liệu đã đọc (khách, tiền, bill) để chủ kiểm tra sau khi có sao kê ===== */
function tdSaveData_(p) {
  const who = hdAuth_(p, ['checkin', 'checkinout', 'hopdong']);
  if (!who) return { ok: false, err: 'noperm' };
  const file = String(p.file || '').trim(), data = String(p.data || '');
  if (!file || !data) return { ok: false, err: 'need' };
  if (data.length > 45000) return { ok: false, err: 'big', msg: 'Dữ liệu hợp đồng quá lớn' };
  const sh = tdSheet_(), n = sh.getLastRow(), v = n > 1 ? sh.getRange(2, 1, n - 1, 1).getDisplayValues() : [];
  for (let i = v.length - 1; i >= 0; i--) if (v[i][0] === file) {
    sh.getRange(i + 2, 12).setValue(data);
    if (String(sh.getRange(i + 2, 8).getValue()) === 'Đã xoá') sh.getRange(i + 2, 8).setValue('Chưa kiểm tra');   // nạp lại sau khi xoá
    return { ok: true };
  }
  const room = String(p.room || '').toUpperCase(), toa = (room.split('-')[1] || '');
  const now = new Date(Date.now()), thang = 'Tháng ' + (+Utilities.formatDate(now, 'Asia/Ho_Chi_Minh', 'M')) + '-' + Utilities.formatDate(now, 'Asia/Ho_Chi_Minh', 'yyyy');
  sh.appendRow(["'" + file, "'" + room, "'" + toa, "'" + thang, "'" + (p.start || ''), tdNow_(), who.owner ? 'Chủ' : who.ten, 'Chưa kiểm tra', '', '', '', data]);
  return { ok: true, added: true };
}
/** Xoá hợp đồng khỏi danh sách chờ (nhân viên / chủ): đánh dấu "Đã xoá" + bỏ dữ liệu — file PDF gốc trong kho vẫn giữ */
function tdDel_(p) {
  const who = hdAuth_(p, ['checkin', 'checkinout', 'hopdong']);
  if (!who) return { ok: false, err: 'noperm' };
  const files = [].concat(p.files || []).map(String).filter(Boolean); if (!files.length) return { ok: false, err: 'need' };
  const sh = tdSheet_(), n = sh.getLastRow(); if (n < 2) return { ok: true, n: 0 };
  const v = sh.getRange(2, 1, n - 1, 1).getDisplayValues(); let k = 0;
  v.forEach((r, i) => { if (files.indexOf(r[0]) >= 0) { sh.getRange(i + 2, 8).setValue('Đã xoá'); sh.getRange(i + 2, 9).setValue('Xoá bởi ' + (who.owner ? 'Chủ' : who.ten) + ' lúc ' + tdNow_()); sh.getRange(i + 2, 12).setValue(''); k++; } });
  try { nkAppend_(who.ma || 'CHỦ', who.ten || 'Chủ', 'Xoá hợp đồng khỏi danh sách check-in', files.length + ' hợp đồng: ' + files.slice(0, 8).join(', ') + (files.length > 8 ? '…' : ''), 'checkin.html', p.devname); } catch (e) {}
  return { ok: true, n: k };
}
/** Các hợp đồng còn chờ kiểm tra (chưa Đạt hoặc chưa bổ sung file QLKH), có dữ liệu — 150 ngày gần nhất */
function tdPending_(p) {
  const who = hdAuth_(p, ['checkin', 'checkinout', 'hopdong']);
  if (!who) return { ok: false, err: 'noperm' };
  const sh = tdSheet_(), n = sh.getLastRow(); if (n < 2) return { ok: true, rows: [] };
  const v = sh.getRange(2, 1, n - 1, TD_HDR.length).getDisplayValues(), han = Date.now() - 150 * 864e5;
  const day = s => { const m = String(s || '').match(/^(\d{2})\/(\d{2})\/(\d{4})/); return m ? new Date(+m[3], +m[2] - 1, +m[1]).getTime() : 0; };
  const rows = v.filter(r => r[11] && (r[7] !== 'Đạt' || !r[10]) && day(r[5]) >= han)
    .map(r => ({ file: r[0], room: r[1], start: r[4], at: r[5], by: r[6], st: r[7] || 'Chưa kiểm tra', qlkh: r[10] || '', data: r[11] }));
  return { ok: true, rows: rows };
}
