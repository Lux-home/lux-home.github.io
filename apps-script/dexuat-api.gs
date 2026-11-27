/**
 * LUXHOME — ĐỀ XUẤT CỦA NHÂN VIÊN (cùng dự án với canho-api.gs / khach-api.gs / nhatky-api.gs)
 * Nhân viên (app phụ, có quyền Check in - Check out) được xem danh sách khách, nhưng mọi
 * SỬA / OUT / XUẤT FILE chỉ là ĐỀ XUẤT. Chủ (Admin) duyệt thì mới ghi vào file / mới được tải file.
 * Mọi đề xuất + quyết định đều ghi vào Nhật ký nhân viên.
 */
const KP_PREFIX = 'khp_';
const KP_KEEP_DAYS = 7;          // đề xuất đã xử lý giữ lại để xem trong 7 ngày
const KP_EXPORT_HOURS = 24;      // xuất file đã duyệt: tải được trong 24 giờ, 1 lần

function staffWho_(p) {
  const w = nvWho_(p.tok, p.dev, p.devname);
  if (!w) return null;
  const ps = w.perms || [];
  if (ps.indexOf('checkin') < 0 && ps.indexOf('checkinout') < 0) return null;
  return w;
}
function kpAll_() {
  const P = PropertiesService.getScriptProperties(), all = P.getProperties(), now = Date.now(), out = [];
  Object.keys(all).forEach(k => {
    if (k.indexOf(KP_PREFIX) !== 0) return;
    let o; try { o = JSON.parse(all[k]); } catch (e) { return; }
    if (o.status !== 'cho' && now - (o.decidedAt || o.t) > KP_KEEP_DAYS * 86400000) { P.deleteProperty(k); return; }
    o.id = k; out.push(o);
  });
  return out.sort((a, b) => b.t - a.t);
}
function kpSave_(o) { const c = Object.assign({}, o); delete c.id; PropertiesService.getScriptProperties().setProperty(o.id, JSON.stringify(c)); }
function kpLabel_(o) { return o.type === 'out' ? 'Đề xuất OUT' : o.type === 'export' ? 'Yêu cầu xuất Excel' : o.type === 'add' ? 'Đề xuất bổ sung khách check-in' : 'Đề xuất sửa khách'; }
function kpDetail_(o) {
  if (o.type === 'export') return 'Xuất file QUAN LY KHACH HANG';
  if (o.type === 'add') return (o.add.ma || '') + ' · ' + (o.add.guests || []).map(g => g.ten).join(', ') + (o.add.file ? ' · HĐ ' + o.add.file : '') + (o.kq ? ' · thanh toán: ' + o.kq : '');
  const g = o.expect || {};
  let d = (g.ma || '') + ' · ' + (g.ten || '(dòng trống)') + ' · dòng ' + o.row;
  if (o.type === 'save') d += ' · ' + Object.keys(o.fields || {}).map(k => k + ': "' + ((o.before || {})[k] || '') + '" → "' + o.fields[k] + '"').join('; ');
  if (o.type === 'out' && o.note) d += ' · ghi chú: ' + o.note;
  return d;
}
function kpPendingByRow_() {
  const m = {}; kpAll_().forEach(o => { if (o.status === 'cho' && o.row) m[o.row] = (m[o.row] || 0) + 1; }); return m;
}

/* ---------- nhân viên ---------- */
function staffRoute_(a, p) {
  const who = staffWho_(p);
  if (!who) return { ok: false, err: 'notok', msg: 'Tài khoản nhân viên không có quyền Check in - Check out hoặc máy chưa được duyệt.' };
  if (a === 'khlist') {
    if (khPending_().length) withLock_(() => khApply_());
    return Object.assign(khList_(), { staff: true, who: who, pend: kpPendingByRow_() });
  }
  if (a === 'khpropose') {
    const type = ['save', 'out', 'export'].indexOf(p.type) >= 0 ? p.type : 'save';
    const o = { t: Date.now(), status: 'cho', who: { ma: who.ma, ten: who.ten }, type: type };
    if (type !== 'export') {
      o.row = +p.row; if (!(o.row > 8)) return { ok: false, err: 'badrow' };
      o.expect = { ma: (p.expect || {}).ma || '', ten: (p.expect || {}).ten || '' };
    }
    if (type === 'save') {
      const f = {}, b = {}; Object.keys(p.fields || {}).forEach(k => { if (K_EDIT.indexOf(k) >= 0) { f[k] = String(p.fields[k]); b[k] = String((p.before || {})[k] || ''); } });
      if (!Object.keys(f).length) return { ok: false, err: 'nothing' };
      o.fields = f; o.before = b;
    }
    if (type === 'out') { o.guest = p.guest || {}; o.note = String(p.note || '').slice(0, 300); o.openUnit = !!p.openUnit; }
    o.id = KP_PREFIX + o.t + '_' + Math.random().toString(36).slice(2, 7);
    kpSave_(o);
    try { nkAppend_(who.ma, who.ten, kpLabel_(o), kpDetail_(o), 'khach.html', p.devname); } catch (e) {}
    return { ok: true, id: o.id };
  }
  if (a === 'khmyreqs') {
    return { ok: true, items: kpAll_().filter(o => o.who && o.who.ma === who.ma).slice(0, 50).map(o => ({
      id: o.id, t: o.t, type: o.type, status: o.status, row: o.row, expect: o.expect, fields: o.fields, before: o.before, note: o.note,
      reason: o.reason || '', err: o.err || '', decidedAt: o.decidedAt || 0,
      canDownload: o.type === 'export' && o.status === 'duyet' && !o.used && Date.now() - o.decidedAt < KP_EXPORT_HOURS * 3600000 })) };
  }
  if (a === 'khexportget') {
    const o = kpAll_().find(x => x.id === p.id);
    if (!o || o.type !== 'export' || !o.who || o.who.ma !== who.ma) return { ok: false, err: 'notfound' };
    if (o.status !== 'duyet') return { ok: false, err: 'notapproved', msg: 'Yêu cầu chưa được duyệt' };
    if (o.used || Date.now() - o.decidedAt > KP_EXPORT_HOURS * 3600000) return { ok: false, err: 'expired', msg: 'Lượt tải đã dùng hoặc đã hết hạn — gửi yêu cầu mới' };
    const f = khExport_(); if (!f.ok) return f;
    o.used = Date.now(); kpSave_(o);
    try { nkAppend_(who.ma, who.ten, 'Tải file Excel (đã duyệt)', f.name, 'khach.html', p.devname); } catch (e) {}
    return f;
  }
  return { ok: false, err: 'noperm' };
}

/* ---------- chủ (Admin) ---------- */
function kpList_() {
  const all = kpAll_();
  return { ok: true, pending: all.filter(o => o.status === 'cho'), recent: all.filter(o => o.status !== 'cho').slice(0, 30) };
}
function kpDecide_(p) {
  const o = kpAll_().find(x => x.id === p.id);
  if (!o) return { ok: false, err: 'notfound' };
  if (o.status !== 'cho') return { ok: false, err: 'done', msg: 'Đề xuất này đã được xử lý' };
  o.decidedAt = Date.now();
  if (!p.approve) {
    o.status = 'tuchoi'; o.reason = String(p.reason || '').slice(0, 300); kpSave_(o);
    try { nkAppend_('CHỦ', 'Chủ', 'Từ chối: ' + kpLabel_(o), (o.who ? o.who.ten + ' · ' : '') + kpDetail_(o) + (o.reason ? ' · lý do: ' + o.reason : ''), 'khach.html', ''); } catch (e) {}
    return { ok: true, status: o.status };
  }
  if (o.type === 'export') {
    o.status = 'duyet'; kpSave_(o);
  } else {
    // đưa vào hàng chờ ghi file và ghi ngay
    const q = o.type === 'add' ? khQueue_(Object.assign({}, o.add, { by: 'NV ' + (o.who ? o.who.ten : '') })) : khQueue_(o.type === 'out'
      ? { type: 'out', row: o.row, guest: Object.assign({}, o.guest, { ma: o.expect.ma, ten: o.expect.ten }), note: (o.note || '') + ' (NV ' + (o.who ? o.who.ten : '') + ' đề xuất)', openUnit: o.openUnit }
      : { type: 'save', row: o.row, expect: o.expect, fields: o.fields });
    if (!q.ok) return { ok: false, err: 'queue', msg: 'Không đưa được vào hàng chờ' };
    o.status = 'duyet'; kpSave_(o);            // đánh dấu ngay để không bị duyệt 2 lần
    let r = null;
    try { r = withLock_(() => khApply_()); } catch (err) { r = { ok: false, msg: String(err.message || err) }; }
    const e = ((r && r.errors) || []).find(x => x.t === q.t);
    if (e) { o.status = 'loi'; o.err = e.msg; kpSave_(o); }
    else if (!r || !r.ok) { o.err = 'Đã duyệt — đang chờ ghi vào file (sẽ tự ghi ở lần kế tiếp)'; kpSave_(o); }
  }
  try { nkAppend_('CHỦ', 'Chủ', 'Duyệt: ' + kpLabel_(o), (o.who ? o.who.ten + ' · ' : '') + kpDetail_(o) + (o.status === 'loi' ? ' · LỖI: ' + o.err : ''), 'khach.html', ''); } catch (e) {}
  return { ok: true, status: o.status, err: o.err || '' };
}
function kpCount_() { return { ok: true, n: kpAll_().filter(o => o.status === 'cho').length }; }
