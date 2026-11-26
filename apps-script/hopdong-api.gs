/**
 * LUXHOME — KHO HỢP ĐỒNG (cùng dự án với canho-api.gs …)
 * Mục 3 "Hợp đồng PDF" (check-in) nạp file → app tải file PDF gốc lên Drive:
 *   LUXHOME - HỢP ĐỒNG / <Tòa>-Tháng <m>-<yyyy> (tháng check-in) / P<phòng> / <tên file>.pdf   (vd T121-Tháng 11-2026 / P301)
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
  const tm = thang.split('-'), lv1 = toa + '-Tháng ' + (+tm[1]) + '-' + tm[0], lv2 = /^P/.test(phong) ? phong : 'P' + phong;
  const f = hdChild_(hdChild_(hdRoot_(), lv1), lv2);
  const ex = f.getFilesByName(name);
  while (ex.hasNext()) { const x = ex.next(); if (!x.isTrashed() && (!p.size || x.getSize() === +p.size)) return { ok: true, exists: true, id: x.getId(), path: lv1 + ' / ' + lv2 }; }
  const res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&supportsAllDrives=true', {
    method: 'post', contentType: 'application/json; charset=UTF-8', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken(), 'X-Upload-Content-Type': 'application/pdf', 'X-Upload-Content-Length': String(p.size || ''),
      Origin: p.origin || 'https://lux-home.github.io' },
    payload: JSON.stringify({ name: name, parents: [f.getId()] })
  });
  const loc = res.getAllHeaders()['Location'] || res.getAllHeaders()['location'];
  if (res.getResponseCode() !== 200 || !loc) return { ok: false, err: 'session', msg: res.getContentText().slice(0, 200) };
  if (!who.owner) { try { nkAppend_(who.ma, who.ten, 'Lưu hợp đồng PDF', lv1 + ' / ' + lv2 + ' · ' + name, 'checkin.html', p.devname); } catch (e) {} }
  return { ok: true, url: loc, path: lv1 + ' / ' + lv2 };
}
/** Duyệt kho: path = [] (Tòa-Tháng) | [Tòa-Tháng] (phòng) | [Tòa-Tháng, phòng] (file) */
function hdList_(p) {
  const who = hdAuth_(p, ['hopdong']);
  if (!who) return { ok: false, err: 'noperm', msg: 'Chưa được cấp quyền Tạo mới + lưu trữ HĐ' };
  let f = hdRoot_();
  const path = Array.isArray(p.path) ? p.path : String(p.path || '').split('/').filter(String);
  for (const name of path) { const it = f.getFoldersByName(name); if (!it.hasNext()) return { ok: false, err: 'notfound' }; f = it.next(); }
  const folders = [], files = [];
  const fi = f.getFolders(); while (fi.hasNext()) { const x = fi.next(); folders.push({ name: x.getName(), id: x.getId() }); }
  const fl = f.getFiles(); while (fl.hasNext()) { const x = fl.next(); if (!x.isTrashed()) files.push({ name: x.getName(), id: x.getId(), size: x.getSize(), at: Utilities.formatDate(x.getDateCreated(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm') }); }
  // cấp 1 "T121-Tháng 11-2026": xếp theo tòa, rồi tháng mới nhất trước
  const key = n => { const m = String(n).match(/^(.*)-Tháng (\d+)-(\d{4})$/); return m ? [m[1], +m[3] * 100 + (+m[2])] : [n, 0]; };
  folders.sort((a, b) => {
    if (path.length) return a.name.localeCompare(b.name, 'vi', { numeric: true });
    const x = key(a.name), y = key(b.name); return x[0] === y[0] ? y[1] - x[1] : x[0].localeCompare(y[0], 'vi', { numeric: true });
  });
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
