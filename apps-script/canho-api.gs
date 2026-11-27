/**
 * LUXHOME — MÁY CHỦ "QUẢN LÝ CĂN HỘ" (Apps Script, dự án RIÊNG)
 * ------------------------------------------------------------------
 * App (lux-home.github.io) gọi máy chủ này để:
 *   - đăng nhập (mật khẩu kiểm tra TẠI MÁY CHỦ, không lộ ra ngoài)
 *   - đọc danh mục căn hộ / tòa / số phòng (catalog + trang Cập nhật căn hộ)
 *   - thêm / sửa / xóa căn, sửa tòa, quản lý mật khẩu (chỉ Admin)
 *   - tải ảnh / video lên đúng thư mục Drive  KHO ẢNH CĂN HỘ / <Tòa> / <Mã căn>
 * Nhờ vậy Google Sheet "KHO ẢNH CĂN HỘ - CSKH" KHÔNG cần chia sẻ công khai nữa.
 *
 * CÀI ĐẶT: xem apps-script/HUONG-DAN.md
 */

const SHEET_ID   = '1CqPzUBDTMwa8XY7ACdJywwbxMIgcuE9mTV7V7zdsXZ0';
const SH_UNIT    = 'DANH MỤC CĂN HỘ';
const SH_TOA     = 'DANH MỤC TÒA';
const SH_ROOM    = 'DANH MỤC PHÒNG';
const SH_PW      = 'MẬT KHẨU';
const ROOT_NAME  = 'KHO ẢNH CĂN HỘ';
const STAFF_MASTER = 'LUXADMIN2026';      // giống catalog.html (mật khẩu nhân viên khẩn cấp)
const WEEK_SECRET  = 'Lux#2026$rotate';   // giống catalog.html (mã khách theo tuần)

/* ======================= CỔNG VÀO ======================= */
function doGet(e) {
  const p = (e && e.parameter) || {};
  return out_(safe_(() => handle_(p.a || '', p)), p.callback);
}
function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) {}
  return out_(safe_(() => handle_(body.a || '', body)), null);
}
function out_(obj, cb) {
  const s = JSON.stringify(obj);
  if (cb && /^[\w.$]+$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + s + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}
function safe_(fn) {
  try { return fn(); } catch (err) { return { ok: false, err: 'server', msg: String(err && err.message || err) }; }
}

function handle_(a, p) {
  const role = roleOf_(p.pw);
  if (a === 'ping') return { ok: true, v: 1 };
  if (a === 'login') return role ? { ok: true, role: role } : { ok: false, err: 'badpw' };
  if (a === 'nvlog') return nvLog_(p);   // nhân viên ghi nhật ký (xác minh bằng token app phụ)
  // kho hợp đồng: tự xác thực (Admin / mật khẩu duyệt check-in / token nhân viên)
  if (a === 'hdsession') return hdSession_(p);
  if (a === 'hdlist')    return hdList_(p);
  if (a === 'hdsearch')  return hdSearch_(p);
  if (a === 'hdget')     return hdGet_(p);
  if (a === 'tdlist')    return tdList_(p);
  if (a === 'khadd')     return khAddRoute_(p);
  if (a === 'setavail')  return withLock_(() => setAvail_(p));
  if (a === 'khaddst')   return khAddStatus_(p);
  if (a === 'tdsave')    return tdSave_(p);
  if (a === 'tdmail')    return tdMailCfg_(p);
  // nhân viên (token app phụ) — chỉ xem / gửi đề xuất, chủ duyệt
  if (!role && p.tok && ['khlist', 'khpropose', 'khmyreqs', 'khexportget'].indexOf(a) >= 0) return staffRoute_(a, p);
  if (!role) return { ok: false, err: 'badpw' };

  // ---- đọc (mọi vai trò đã đăng nhập) ----
  if (a === 'list')  return list_(role);
  if (a === 'files') return files_(p.ma);

  // ---- ghi (chỉ Admin) ----
  if (role !== 'admin') return { ok: false, err: 'noperm' };
  if (a === 'save')         return withLock_(() => saveUnit_(p.unit || {}, p.orig || ''));
  if (a === 'delunit')      return withLock_(() => delUnit_(p.ma));
  if (a === 'upload')       return withLock_(() => upload_(p));
  if (a === 'delfile')      return withLock_(() => delFile_(p.ma, p.id));
  if (a === 'setcover')     return withLock_(() => setCover_(p.ma, p.id));
  if (a === 'videosession') return withLock_(() => videoSession_(p));
  if (a === 'savetoa')      return withLock_(() => saveToa_(p.toa || {}, p.orig || ''));
  if (a === 'uploadfacade') return withLock_(() => uploadFacade_(p));
  if (a === 'nvloglist')    return nvLogList_();
  if (a === 'pwlist')       return pwList_();
  if (a === 'pwsave')       return withLock_(() => pwSave_(p.item || {}));
  // ---- khách thuê (file QUAN LY KHACH HANG.xlsm) — chỉ Admin
  if (a === 'khlist')       { if (khPending_().length) withLock_(() => khApply_()); return Object.assign(khList_(), { errors: khErrors_(), pend: kpPendingByRow_() }); }
  if (a === 'khproplist')   return kpList_();
  if (a === 'khpropcount')  return kpCount_();
  if (a === 'khdecide')     return kpDecide_(p);
  if (a === 'khqueue')      return khQueue_(p);
  if (a === 'khapply')      return withLock_(() => khApply_());
  if (a === 'khstatus')     return khStatus_();
  if (a === 'khclearerr')   return khClearErr_();
  if (a === 'khexport')     return khExport_();
  if (a === 'khoutlist')    return khOutList_();
  return { ok: false, err: 'unknown' };
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { ok: false, err: 'busy' };
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ======================= TIỆN ÍCH ======================= */
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }
function norm_(s) {
  return String(s == null ? '' : s).toLowerCase().normalize('NFD')
    .replace(/[̀-ͯ]/g, '').replace(/đ/g, 'd').replace(/[^a-z0-9]/g, '');
}
function digits_(s) { const d = String(s == null ? '' : s).replace(/[^\d]/g, ''); return d ? Number(d) : ''; }
function driveId_(s) {
  s = String(s || '').trim(); if (!s) return '';
  let m = s.match(/\/folders\/([\w-]{10,})/) || s.match(/\/d\/([\w-]{10,})/) || s.match(/[?&]id=([\w-]{10,})/);
  if (m) return m[1];
  return /^[\w-]{20,}$/.test(s) ? s : '';
}
/** Tìm dòng tiêu đề (chứa cột bắt buộc) và map tên cột -> chỉ số (0-based). */
function header_(values, must) {
  for (let r = 0; r < Math.min(values.length, 15); r++) {
    const H = values[r].map(norm_);
    // dòng tiêu đề = có ô NGẮN chứa tên cột bắt buộc (bỏ qua dòng hướng dẫn dài)
    if (H.some(h => h.indexOf(must) >= 0 && h.length <= must.length + 30)) {
      const find = (...keys) => {
        for (const k of keys) { const i = H.findIndex(h => h === k); if (i >= 0) return i; }
        for (const k of keys) { const i = H.findIndex(h => h.indexOf(k) >= 0); if (i >= 0) return i; }
        return -1;
      };
      return { row: r, find: find };
    }
  }
  return null;
}

/* ======================= MẬT KHẨU / VAI TRÒ ======================= */
function weekCode_() {
  const DAY = 86400000;
  const vnDay = Math.floor((Date.now() + 7 * 3600 * 1000) / DAY);
  const weekIndex = Math.floor((vnDay - 4) / 7);
  let h = 2166136261; const s = WEEK_SECRET + ':' + weekIndex;
  for (let i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); }
  h = h >>> 0;
  const AL = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let code = '', x = h; for (let i = 0; i < 6; i++) { code += AL[x % AL.length]; x = Math.floor(x / AL.length); }
  return code;
}
function pwTable_() {
  const sh = ss_().getSheetByName(SH_PW); if (!sh) return null;
  const v = sh.getDataRange().getDisplayValues();
  const h = header_(v, 'matkhau'); if (!h) return null;
  return { sh: sh, v: v, hr: h.row, cLoai: h.find('loai'), cPw: h.find('matkhau'), cTt: h.find('trangthai'), cNote: h.find('ghichu') };
}
function roleOf_(pw) {
  pw = String(pw || '').trim(); if (!pw) return null;
  const cache = CacheService.getScriptCache();
  const key = 'r_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw));
  const hit = cache.get(key); if (hit) return hit === '-' ? null : hit;
  let role = null;
  const t = pwTable_();
  if (t && t.cPw >= 0) {
    for (let r = t.hr + 1; r < t.v.length; r++) {
      const row = t.v[r];
      if (String(row[t.cPw] || '').trim() !== pw) continue;
      const st = t.cTt >= 0 ? norm_(row[t.cTt]) : '';
      if (st === 'tat' || st === 'off' || st === 'khoa') continue;
      const l = norm_(t.cLoai >= 0 ? row[t.cLoai] : '');
      role = /admin|quanly|chu/.test(l) ? 'admin' : /nhanvien|staff/.test(l) ? 'staff' : 'guest';
      break;
    }
  }
  if (!role && pw === STAFF_MASTER) role = 'staff';
  if (!role && pw === weekCode_()) role = 'guest';
  cache.put(key, role || '-', 60);
  return role;
}
function pwList_() {
  const t = pwTable_(); if (!t) return { ok: false, err: 'nopwsheet' };
  const items = [];
  for (let r = t.hr + 1; r < t.v.length; r++) {
    const row = t.v[r]; const pw = String(row[t.cPw] || '').trim(); if (!pw) continue;
    items.push({ row: r + 1, loai: row[t.cLoai] || '', pw: pw, tt: t.cTt >= 0 ? row[t.cTt] : '', note: t.cNote >= 0 ? row[t.cNote] : '' });
  }
  return { ok: true, items: items, week: weekCode_() };
}
function pwSave_(it) {
  const t = pwTable_(); if (!t) return { ok: false, err: 'nopwsheet' };
  const pw = String(it.pw || '').trim(); if (!pw) return { ok: false, err: 'empty' };
  let row = Number(it.row) || 0;
  if (row && row <= t.hr + 1) return { ok: false, err: 'badrow' };
  // trùng mật khẩu với dòng khác?
  for (let r = t.hr + 1; r < t.v.length; r++) {
    if (r + 1 !== row && String(t.v[r][t.cPw] || '').trim() === pw) return { ok: false, err: 'dup' };
  }
  const isAdmin = l => /admin|quanly|chu/.test(norm_(l));
  const on = s => { const n = norm_(s); return !(n === 'tat' || n === 'off' || n === 'khoa'); };
  // không cho tắt / đổi loại Admin cuối cùng
  const admins = [];
  for (let r = t.hr + 1; r < t.v.length; r++) {
    const rw = t.v[r]; if (String(rw[t.cPw] || '').trim() && isAdmin(rw[t.cLoai]) && on(rw[t.cTt])) admins.push(r + 1);
  }
  if (row && admins.length === 1 && admins[0] === row && (!isAdmin(it.loai) || !on(it.tt))) return { ok: false, err: 'lastadmin' };
  if (it.del) {
    if (!row) return { ok: false, err: 'badrow' };
    if (admins.length === 1 && admins[0] === row) return { ok: false, err: 'lastadmin' };
    t.sh.deleteRow(row);
  } else {
    if (!row) { row = t.sh.getLastRow() + 1; }
    const set = (c, v) => { if (c >= 0) t.sh.getRange(row, c + 1).setValue(v); };
    set(t.cLoai, it.loai || 'Khách hàng'); set(t.cPw, pw); set(t.cTt, it.tt || 'Bật'); set(t.cNote, it.note || '');
    // giữ mật khẩu dạng chữ (không để Sheet đổi "0123" thành số 123)
    if (t.cPw >= 0) t.sh.getRange(row, t.cPw + 1).setNumberFormat('@').setValue(pw);
  }
  // (vai trò đã kiểm tra được nhớ 60 giây — mật khẩu vừa tắt sẽ hết tác dụng sau tối đa 1 phút)
  return pwList_();
}

/* ======================= ĐỌC DANH MỤC ======================= */
function unitTable_() {
  const sh = ss_().getSheetByName(SH_UNIT);
  const v = sh.getDataRange().getDisplayValues();
  const h = header_(v, 'macanho');
  if (!h) throw new Error('Không thấy dòng tiêu đề MÃ CĂN HỘ');
  const f = h.find;
  const C = {
    toa: f('matoa'), so: f('socanho'), ma: f('macanho'), tt: f('tinhtrang'), ngay: f('ngaydukienco', 'ngaydukien'),
    loai: f('loaican'), khu: f('khuvuc'), gia: f('giathue'), dien: f('dongiadien'), nuoc: f('dongianuoc'),
    dv: f('phidichvu'), nt: f('noithat'), ntEn: f('motatienganh', 'motaen'), anh: f('linkanhvideo', 'linkanh'),
    anhdd: f('anhdaidien'), map: f('linkbando'), msg: f('tinnhan'), facade: f('anhmattien')
  };
  return { sh: sh, v: v, hr: h.row, C: C };
}
function unitObj_(t, r) {
  const row = t.v[r], C = t.C, g = c => c >= 0 ? String(row[c] || '').trim() : '';
  return { row: r + 1, ma: g(C.ma), toa: g(C.toa), so: g(C.so), tt: g(C.tt), ngay: g(C.ngay), loai: g(C.loai), khu: g(C.khu),
    gia: g(C.gia), dien: g(C.dien), nuoc: g(C.nuoc), dv: g(C.dv), nt: g(C.nt), ntEn: g(C.ntEn),
    anh: g(C.anh), anhdd: g(C.anhdd), map: g(C.map), facade: g(C.facade) };
}
function findUnit_(t, ma) {
  ma = String(ma || '').trim().toUpperCase(); if (!ma) return -1;
  for (let r = t.hr + 1; r < t.v.length; r++) if (String(t.v[r][t.C.ma] || '').trim().toUpperCase() === ma) return r;
  return -1;
}
function validationList_(sh, a1) {
  try {
    const dv = sh.getRange(a1).getDataValidation(); if (!dv) return [];
    const type = dv.getCriteriaType(), vals = dv.getCriteriaValues();
    if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return (vals[0] || []).map(String).filter(Boolean);
    if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      return [].concat.apply([], vals[0].getDisplayValues()).map(s => String(s).trim()).filter(Boolean);
    }
  } catch (e) {}
  return [];
}
function toaTable_() {
  const sh = ss_().getSheetByName(SH_TOA); if (!sh) return null;
  const v = sh.getDataRange().getDisplayValues();
  const h = header_(v, 'matoa'); if (!h) return null;
  return { sh: sh, v: v, hr: h.row, cMa: h.find('matoa'), cDc: h.find('diachi'), cMap: h.find('linkgooglemap', 'map'), cFac: h.find('anhmattien', 'mattien') };
}
function list_(role) {
  const t = unitTable_();
  const units = [];
  for (let r = t.hr + 1; r < t.v.length; r++) {
    const u = unitObj_(t, r); if (!u.ma) continue;
    if (role === 'guest' && /khoa/.test(norm_(u.tt))) continue;
    units.push(u);
  }
  // "rows" giữ đúng dạng cũ (tiêu đề + dữ liệu) để catalog.html dùng lại hàm rowsToData
  const rows = [t.v[t.hr]].concat(t.v.slice(t.hr + 1).filter(row => {
    const ma = String(row[t.C.ma] || '').trim(); if (!ma) return false;
    return !(role === 'guest' && /khoa/.test(norm_(row[t.C.tt])));
  }));
  const res = { ok: true, role: role, rows: rows, units: units };
  if (role === 'guest') return res;

  const toas = [];
  const tt = toaTable_();
  if (tt) for (let r = tt.hr + 1; r < tt.v.length; r++) {
    const ma = String(tt.v[r][tt.cMa] || '').trim(); if (!ma) continue;
    const g = c => c >= 0 ? String(tt.v[r][c] || '').trim() : '';
    toas.push({ row: r + 1, ma: ma, diachi: g(tt.cDc), map: g(tt.cMap), facade: g(tt.cFac) });
  }
  const rooms = [];
  const rs = ss_().getSheetByName(SH_ROOM);
  if (rs) {
    const rv = rs.getDataRange().getDisplayValues(); const h = header_(rv, 'socanho');
    if (h) { const c = h.find('socanho'); for (let r = h.row + 1; r < rv.length; r++) { const s = String(rv[r][c] || '').trim(); if (s) rooms.push(s); } }
  }
  const first = t.hr + 2;   // dòng dữ liệu đầu tiên (1-based)
  const colA1 = c => c >= 0 ? t.sh.getRange(first, c + 1).getA1Notation() : null;
  const distinct = c => c < 0 ? [] : Array.from(new Set(units.map(u => String(t.v[u.row - 1][c] || '').trim()).filter(Boolean)));
  const opt = (c) => { const l = colA1(c) ? validationList_(t.sh, colA1(c)) : []; return l.length ? l : distinct(c); };
  res.toas = toas; res.rooms = rooms;
  res.opts = { tt: opt(t.C.tt), loai: opt(t.C.loai), khu: opt(t.C.khu) };
  res.hasEn = t.C.ntEn >= 0;
  return res;
}

/* ======================= THƯ MỤC DRIVE ======================= */
function rootFolder_() {
  const P = PropertiesService.getScriptProperties();
  const id = P.getProperty('ROOT_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  let root = null;
  // 1) suy ra từ link thư mục của một căn đã có: <căn> -> <tòa> -> <KHO ẢNH CĂN HỘ>
  try {
    const t = unitTable_();
    for (let r = t.hr + 1; r < t.v.length && !root; r++) {
      const fid = driveId_(t.v[r][t.C.anh]); if (!fid) continue;
      try {
        const f = DriveApp.getFolderById(fid); const p1 = f.getParents(); if (!p1.hasNext()) continue;
        const p2 = p1.next().getParents(); if (p2.hasNext()) root = p2.next();
      } catch (e) {}
    }
  } catch (e) {}
  // 2) tìm theo tên; 3) tạo mới
  if (!root) { const it = DriveApp.getFoldersByName(ROOT_NAME); root = it.hasNext() ? it.next() : DriveApp.createFolder(ROOT_NAME); }
  P.setProperty('ROOT_ID', root.getId());
  return root;
}
function child_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function shareView_(f) {
  try { f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
}
/** Lấy (hoặc tạo) thư mục của căn; ghi link vào cột LINK ẢNH & VIDEO nếu còn trống. */
function unitFolder_(t, r, create) {
  const u = unitObj_(t, r);
  const fid = driveId_(u.anh);
  if (fid) { try { return DriveApp.getFolderById(fid); } catch (e) {} }
  if (!create) return null;
  const f = child_(child_(rootFolder_(), u.toa || 'KHÁC'), u.ma);
  shareView_(f);
  if (t.C.anh >= 0) t.sh.getRange(r + 1, t.C.anh + 1).setValue(f.getUrl());
  return f;
}
function fileObj_(f) {
  const id = f.getId();
  return { id: id, name: f.getName(), mime: f.getMimeType(), size: f.getSize(),
    thumb: 'https://drive.google.com/thumbnail?id=' + id + '&sz=w400', url: 'https://drive.google.com/file/d/' + id + '/view' };
}
function files_(ma) {
  const t = unitTable_(); const r = findUnit_(t, ma); if (r < 0) return { ok: false, err: 'notfound' };
  const f = unitFolder_(t, r, false); if (!f) return { ok: true, files: [], cover: '' };
  const files = []; const it = f.getFiles();
  while (it.hasNext()) { const x = it.next(); if (!x.isTrashed()) files.push(fileObj_(x)); }
  files.sort((a, b) => a.name.localeCompare(b.name));
  return { ok: true, files: files, cover: driveId_(unitObj_(t, r).anhdd), folder: f.getUrl() };
}

/* ======================= GHI CĂN HỘ ======================= */
/** Nhân viên + chủ: chỉ đổi TÌNH TRẠNG + NGÀY TRỐNG (ngày dự kiến có) của 1 căn — không đụng giá, ảnh, mô tả. */
function setAvail_(p) {
  const who = hdAuth_(p, ['quanlycanho', 'canho', 'catalog']);
  if (!who) return { ok: false, err: 'noperm', msg: 'Không có quyền đổi tình trạng căn' };
  const t = unitTable_(), C = t.C, r = findUnit_(t, p.ma);
  if (r < 0) return { ok: false, err: 'notfound' };
  const ngay = String(p.ngay || '').trim();
  if (ngay && !/^\d{2}\/\d{2}\/\d{4}$/.test(ngay)) return { ok: false, err: 'baddate', msg: 'Ngày dạng dd/mm/yyyy' };
  const kind = String(p.kind || (ngay ? 'soon' : '')), label = kind === 'open' ? 'Đang có sẵn' : kind === 'lock' ? 'Đang Khóa' : kind === 'soon' ? 'Dự kiến có' : '';
  if (!label) return { ok: false, err: 'need' };
  const old = { tt: String(t.v[r][C.tt] || ''), ngay: C.ngay >= 0 ? String(t.v[r][C.ngay] || '') : '' };
  const fm = t.sh.getRange(r + 1, 1, 1, t.sh.getLastColumn()).getFormulas()[0];
  if (C.tt >= 0 && !fm[C.tt]) t.sh.getRange(r + 1, C.tt + 1).setValue(label);
  if (C.ngay >= 0 && !fm[C.ngay]) t.sh.getRange(r + 1, C.ngay + 1).setNumberFormat('@').setValue(kind === 'soon' ? ngay : '');
  SpreadsheetApp.flush();
  if (!who.owner) { try { nkAppend_(who.ma, who.ten, 'Cập nhật ngày trống / tình trạng', p.ma + ': ' + (old.tt || '?') + (old.ngay ? ' ' + old.ngay : '') + ' → ' + label + (kind === 'soon' ? ' ' + ngay : ''), 'canho.html', p.devname); } catch (e) {} }
  return { ok: true, ma: p.ma, tt: label, ngay: kind === 'soon' ? ngay : '' };
}
function saveUnit_(u, orig) {
  const t = unitTable_(), C = t.C, sh = t.sh;
  const toa = String(u.toa || '').trim().toUpperCase(), so = String(u.so || '').trim().toUpperCase();
  if (!toa || !so) return { ok: false, err: 'need', msg: 'Thiếu Mã tòa hoặc Số căn hộ' };
  const newMa = so + toa;
  let r = orig ? findUnit_(t, orig) : -1;
  if (orig && r < 0) return { ok: false, err: 'notfound' };
  const dup = findUnit_(t, newMa);
  if (dup >= 0 && dup !== r) return { ok: false, err: 'exists', msg: 'Mã ' + newMa + ' đã có' };

  let rowNum;
  if (r < 0) {
    // thêm dòng mới: chép nguyên dòng mẫu (công thức, định dạng, ô chọn) rồi điền dữ liệu
    let last = t.hr + 1;
    for (let i = t.v.length - 1; i > t.hr; i--) { if (String(t.v[i][C.toa] || t.v[i][C.ma] || '').trim()) { last = i + 1; break; } }
    rowNum = last + 1;
    const tpl = t.hr + 2;
    if (rowNum > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 1);
    sh.getRange(tpl, 1, 1, sh.getLastColumn()).copyTo(sh.getRange(rowNum, 1, 1, sh.getLastColumn()));
    // xóa dữ liệu (không phải công thức) vừa chép theo
    const fm = sh.getRange(rowNum, 1, 1, sh.getLastColumn()).getFormulas()[0];
    fm.forEach((f, i) => { if (!f) sh.getRange(rowNum, i + 1).clearContent(); });
  } else rowNum = r + 1;

  const formulas = sh.getRange(rowNum, 1, 1, sh.getLastColumn()).getFormulas()[0];
  const set = (c, v) => { if (c >= 0 && !formulas[c]) sh.getRange(rowNum, c + 1).setValue(v); };
  set(C.toa, toa); set(C.so, so);
  if (C.ma >= 0 && !formulas[C.ma]) set(C.ma, newMa);
  set(C.tt, u.tt || ''); set(C.ngay, /du\s*kien|dukien/.test(norm_(u.tt)) ? (u.ngay || '') : '');
  set(C.loai, u.loai || ''); set(C.khu, u.khu || '');
  set(C.gia, digits_(u.gia)); set(C.dien, digits_(u.dien)); set(C.nuoc, digits_(u.nuoc)); set(C.dv, digits_(u.dv));
  set(C.nt, u.nt || ''); if (C.ntEn >= 0) set(C.ntEn, u.ntEn || '');
  if (u.anhdd !== undefined) set(C.anhdd, u.anhdd || '');
  SpreadsheetApp.flush();

  // thêm số phòng mới vào DANH MỤC PHÒNG (nguồn ô chọn trong Excel/Sheet)
  try {
    const rs = ss_().getSheetByName(SH_ROOM);
    if (rs) {
      const rv = rs.getDataRange().getDisplayValues(); const h = header_(rv, 'socanho');
      if (h) {
        const c = h.find('socanho');
        const has = rv.slice(h.row + 1).some(x => String(x[c] || '').trim().toUpperCase() === so);
        if (!has) rs.getRange(rs.getLastRow() + 1, c + 1).setNumberFormat('@').setValue(so);
      }
    }
  } catch (e) {}

  // thư mục ảnh: tạo nếu chưa có; đổi tên nếu đổi mã
  const t2 = unitTable_(); const r2 = findUnit_(t2, newMa);
  if (r2 >= 0) {
    const f = unitFolder_(t2, r2, true);
    try { if (f && orig && f.getName() !== newMa && f.getName().toUpperCase() === String(orig).toUpperCase()) f.setName(newMa); } catch (e) {}
    return { ok: true, unit: unitObj_(unitTable_(), r2) };
  }
  return { ok: true, unit: { ma: newMa } };
}
function delUnit_(ma) {
  const t = unitTable_(); const r = findUnit_(t, ma); if (r < 0) return { ok: false, err: 'notfound' };
  t.sh.deleteRow(r + 1);   // chỉ xóa dòng; thư mục ảnh trên Drive được GIỮ LẠI
  return { ok: true };
}

/* ======================= ẢNH / VIDEO ======================= */
function upload_(p) {
  const t = unitTable_(); const r = findUnit_(t, p.ma); if (r < 0) return { ok: false, err: 'notfound' };
  const f = unitFolder_(t, r, true);
  const blob = Utilities.newBlob(Utilities.base64Decode(p.data), p.mime || 'image/jpeg', p.name || ('anh_' + Date.now() + '.jpg'));
  const file = f.createFile(blob);
  const u = unitObj_(t, r);
  if (!driveId_(u.anhdd) && /^image\//.test(file.getMimeType()) && t.C.anhdd >= 0) t.sh.getRange(r + 1, t.C.anhdd + 1).setValue(file.getId());
  return { ok: true, file: fileObj_(file) };
}
function inFolder_(file, folder) {
  const it = file.getParents(); while (it.hasNext()) if (it.next().getId() === folder.getId()) return true; return false;
}
function delFile_(ma, id) {
  const t = unitTable_(); const r = findUnit_(t, ma); if (r < 0) return { ok: false, err: 'notfound' };
  const f = unitFolder_(t, r, false); if (!f) return { ok: false, err: 'nofolder' };
  const file = DriveApp.getFileById(id);
  if (!inFolder_(file, f)) return { ok: false, err: 'noperm' };
  file.setTrashed(true);   // vào Thùng rác Drive (khôi phục được trong 30 ngày)
  if (driveId_(unitObj_(t, r).anhdd) === id && t.C.anhdd >= 0) t.sh.getRange(r + 1, t.C.anhdd + 1).clearContent();
  return { ok: true };
}
function setCover_(ma, id) {
  const t = unitTable_(); const r = findUnit_(t, ma); if (r < 0) return { ok: false, err: 'notfound' };
  if (t.C.anhdd < 0) return { ok: false, err: 'nocol' };
  t.sh.getRange(r + 1, t.C.anhdd + 1).setValue(id || '');
  return { ok: true };
}
/** Video lớn: máy chủ mở "phiên tải lên" trên Drive, điện thoại gửi thẳng file vào phiên đó. */
function videoSession_(p) {
  const t = unitTable_(); const r = findUnit_(t, p.ma); if (r < 0) return { ok: false, err: 'notfound' };
  const f = unitFolder_(t, r, true);
  const res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&supportsAllDrives=true', {
    method: 'post', contentType: 'application/json; charset=UTF-8', muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
      'X-Upload-Content-Type': p.mime || 'video/mp4',
      'X-Upload-Content-Length': String(p.size || ''),
      Origin: p.origin || 'https://lux-home.github.io'
    },
    payload: JSON.stringify({ name: p.name || ('video_' + Date.now() + '.mp4'), parents: [f.getId()] })
  });
  const loc = res.getAllHeaders()['Location'] || res.getAllHeaders()['location'];
  if (res.getResponseCode() !== 200 || !loc) return { ok: false, err: 'session', msg: res.getContentText().slice(0, 300) };
  return { ok: true, url: loc };
}

/* ======================= TÒA NHÀ ======================= */
function saveToa_(x, orig) {
  const t = toaTable_(); if (!t) return { ok: false, err: 'notoasheet' };
  const ma = String(x.ma || '').trim().toUpperCase(); if (!ma) return { ok: false, err: 'need' };
  const find = m => { for (let r = t.hr + 1; r < t.v.length; r++) if (String(t.v[r][t.cMa] || '').trim().toUpperCase() === m) return r; return -1; };
  let r = find(String(orig || ma).toUpperCase());
  const dup = find(ma); if (dup >= 0 && dup !== r) return { ok: false, err: 'exists', msg: 'Tòa ' + ma + ' đã có' };
  let row;
  if (r < 0) {
    let last = t.hr + 1;
    for (let i = t.v.length - 1; i > t.hr; i--) if (String(t.v[i][t.cMa] || '').trim()) { last = i + 1; break; }
    row = last + 1;
  } else row = r + 1;
  const set = (c, v) => { if (c >= 0) t.sh.getRange(row, c + 1).setValue(v); };
  set(t.cMa, ma); set(t.cDc, x.diachi || ''); set(t.cMap, x.map || '');
  if (x.facade !== undefined) set(t.cFac, x.facade || '');
  return { ok: true };
}
function uploadFacade_(p) {
  const t = toaTable_(); if (!t) return { ok: false, err: 'notoasheet' };
  const ma = String(p.toa || '').trim().toUpperCase();
  let row = -1; for (let r = t.hr + 1; r < t.v.length; r++) if (String(t.v[r][t.cMa] || '').trim().toUpperCase() === ma) { row = r + 1; break; }
  if (row < 0) return { ok: false, err: 'notfound' };
  const folder = child_(rootFolder_(), ma);
  const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(p.data), p.mime || 'image/jpeg', p.name || ('mattien_' + ma + '.jpg')));
  shareView_(file);
  if (t.cFac >= 0) t.sh.getRange(row, t.cFac + 1).setValue(file.getId());
  return { ok: true, file: fileObj_(file) };
}

/* Chạy 1 lần trong trình soạn thảo để cấp quyền (Drive, Sheets, gọi mạng). */
function capQuyen() {
  Logger.log(ss_().getName());
  Logger.log(rootFolder_().getName());
  UrlFetchApp.fetch('https://www.googleapis.com/discovery/v1/apis', { muteHttpExceptions: true });
  Logger.log('Mã khách tuần này: ' + weekCode_());
}
