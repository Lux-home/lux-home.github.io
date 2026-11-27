/**
 * LUXHOME — QUẢN LÝ KHÁCH (thêm vào CÙNG dự án Apps Script với canho-api.gs)
 * ------------------------------------------------------------------
 * File QUAN LY KHACH HANG.xlsm trên Drive vẫn là KHO GỐC (các file Excel khác trên máy tính vẫn đọc nó),
 * nhưng không ai cần mở file nữa: app đọc danh sách và sửa từng khách qua máy chủ này.
 *
 *  - Đọc: mở file bằng Google Sheets (như máy chủ Lưu trú), trả về các dòng khách.
 *  - Ghi: sửa thẳng nội dung .xlsm (XML) — giữ nguyên macro, công thức, định dạng, nút bấm;
 *         chỉ đụng các ô được sửa của đúng 1 dòng; KHÔNG bao giờ ghi vào cột công thức (AI..AN).
 *  - Trước khi ghi: tự lưu bản sao vào thư mục "LUXHOME BACKUP QLKH" (tối đa 1 bản / 30 phút).
 *  - Khách trả phòng (OUT): lưu thông tin khách vào bảng "LUXHOME - KHÁCH ĐÃ TRẢ PHÒNG",
 *    rồi xoá trắng thông tin khách của dòng (E..AC) và tô ĐỎ ô Tên khách.
 */

const QLKH_ID     = '1olOswDQsPC8CAsaUiitbAePTeZm83ugV';
const QLKH_SHEET  = 'DANH SACH LUU TRU';
const QLKH_BACKUP = 'LUXHOME BACKUP QLKH';
const QLKH_MIME   = 'application/vnd.ms-excel.sheet.macroEnabled.12';
const OUT_LOG_NAME = 'LUXHOME - KHÁCH ĐÃ TRẢ PHÒNG';
let KH_TARGET = QLKH_ID;
let KH_TEST = false;        // hàm thử trên bản sao: không ghi nhật ký OUT, không đổi catalog   // file đang ghi (hàm thử đổi sang bản sao)

// Trường app  ->  cột trong sheet (theo dòng tiêu đề 8 của file)
const KF = {
  phong: 'B', toa: 'C', ma: 'D',
  ten: 'E', ns: 'F', qt: 'G', gt: 'H', cccd: 'I', hc: 'J', dc: 'K',
  bd: 'L', kt: 'M', lt: 'N', coc: 'O', tp: 'P', sn: 'Q', dck: 'R', nuoc: 'S', bt: 'T', ttnp: 'U',
  sale: 'V', lh: 'W', sdt: 'X',
  tnbd: 'Y', tnkt: 'Z', dvbd: 'AB', dvkt: 'AC', dcu: 'AD', dmoi: 'AE', khac: 'AF', gc: 'AG', gdien: 'AH',
  anh: 'AO', ngayanh: 'AP'
};
const K_DATE = ['bd', 'kt', 'lt', 'tnbd', 'tnkt', 'dvbd', 'dvkt'];          // ghi dạng ngày thật (công thức cần)
const K_NUM  = ['coc', 'tp', 'sn', 'dck', 'nuoc', 'bt', 'dcu', 'dmoi', 'khac', 'gdien'];
const K_EDIT = ['ten', 'ns', 'qt', 'gt', 'cccd', 'hc', 'dc', 'bd', 'kt', 'lt', 'coc', 'tp', 'sn', 'dck', 'nuoc', 'bt', 'ttnp',
                'sale', 'lh', 'sdt', 'tnbd', 'tnkt', 'dvbd', 'dvkt', 'dcu', 'dmoi', 'khac', 'gc', 'gdien'];
// Kiểm tra cấu trúc file trước khi ghi — nếu ai đó chèn/xoá cột thì DỪNG, không ghi sai chỗ
const K_CHECK = { D: 'maphong', E: 'tenkhach', L: 'batdauhd', N: 'thoihanluutru', Y: 'tienthuenha', AD: 'chisodien', AH: 'dongiadien' };

/* ---------------- tiện ích cột ---------------- */
function colNum_(s) { let n = 0; for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n; }
function colName_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
function xUnesc_(s) { return String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (a, b) => String.fromCharCode(+b)).replace(/&amp;/g, '&'); }
function xEsc_(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
// ô tiêu đề "Tên khách" thật (dòng 8) — bỏ dòng mẫu "[Tên khách]" (dòng 7)
function isHdrTen_(c) { c = String(c == null ? '' : c).trim(); return c.charAt(0) !== '[' && norm_(c) === 'tenkhach'; }
function cmp_(s) { return String(s == null ? '' : s).replace(/[\s ]+/g, ' ').trim().toUpperCase(); }

/* ======================= ĐỌC DANH SÁCH ======================= */
function khList_() {
  const X = xlsmOpen_(), s = X.s;
  // định dạng ngày theo style
  const st = X.get('xl/styles.xml'), fmtCode = {};
  const nfM = st.match(/<numFmts count="(\d+)">([\s\S]*?)<\/numFmts>/);
  if (nfM) nfM[2].replace(/<numFmt numFmtId="(\d+)" formatCode="([^"]*)"\/>/g, (m, id, code) => { fmtCode[id] = xUnesc_(code); return m; });
  const xfs = (st.match(/<cellXfs count="\d+">([\s\S]*?)<\/cellXfs>/)[1].match(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g)) || [];
  const isDate = sIdx => { const x = xfs[+sIdx] || ''; const id = +(x.match(/ numFmtId="(\d+)"/) || [0, 0])[1];
    if ((id >= 14 && id <= 22) || (id >= 45 && id <= 47)) return true; const c = fmtCode[id]; return !!c && /[dy]/i.test(c.replace(/"[^"]*"|\[[^\]]*\]/g, '')) && !/[#0]/.test(c); };
  const show = c => {
    if (!c) return '';
    const v = X.cellVal(c);
    if (c.t === '' && v !== '' && /^-?\d+(\.\d+)?$/.test(v) && isDate(c.s)) {
      const d = new Date(Date.UTC(1899, 11, 30) + Math.round(+v) * 86400000);
      return ('0' + d.getUTCDate()).slice(-2) + '/' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '/' + d.getUTCFullYear();
    }
    return /^-?\d+\.0+$/.test(v) ? v.replace(/\.0+$/, '') : v;
  };
  const rows = [], all = {};
  const sdA = s.indexOf('<sheetData'), sdB = s.indexOf('>', sdA) + 1, sdE = s.indexOf('</sheetData>');
  s.slice(sdB, sdE).replace(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g, r => { const n = +(r.match(/ r="(\d+)"/) || [])[1]; all[n] = r; return r; });
  let hr = 0;
  Object.keys(all).map(Number).sort((a, b) => a - b).some(n => { if (n > 20) return true; const cs = X.parseRow(all[n]); if (Object.keys(cs).some(k => isHdrTen_(X.cellVal(cs[k])))) { hr = n; return true; } return false; });
  if (!hr) throw new Error('Không thấy dòng tiêu đề "Tên khách"');
  const hc = X.parseRow(all[hr]);
  for (const L in K_CHECK) if (norm_(X.cellVal(hc[colNum_(L)])) !== K_CHECK[L]) return { ok: false, err: 'layout', msg: 'Cột ' + L + ' của file không đúng mẫu — không đọc để tránh sai.' };
  Object.keys(all).map(Number).sort((a, b) => a - b).forEach(n => {
    if (n <= hr) return;
    const cs = X.parseRow(all[n]), o = { row: n };
    for (const k in KF) o[k] = String(show(cs[colNum_(KF[k])])).trim();
    if (!o.ma && !o.ten && !o.toa) return;
    rows.push(o);
  });
  return { ok: true, rows: rows, at: Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'HH:mm dd/MM/yyyy') };
}

/* ======================= GHI VÀO FILE .XLSM ======================= */
function drive_(path, opt) {
  opt = opt || {};
  opt.headers = Object.assign({ Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, opt.headers || {});
  opt.muteHttpExceptions = true;
  return UrlFetchApp.fetch('https://www.googleapis.com' + path, opt);
}
function backupFolder_() {
  const q = encodeURIComponent("name='" + QLKH_BACKUP + "' and mimeType='application/vnd.google-apps.folder' and trashed=false");
  const r = JSON.parse(drive_('/drive/v3/files?q=' + q + '&fields=files(id)&supportsAllDrives=true&includeItemsFromAllDrives=true').getContentText());
  if (r.files && r.files.length) return r.files[0].id;
  return JSON.parse(drive_('/drive/v3/files?fields=id', { method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ name: QLKH_BACKUP, mimeType: 'application/vnd.google-apps.folder' }) }).getContentText()).id;
}
function maybeBackup_(name, force) {
  const P = PropertiesService.getScriptProperties(), last = +(P.getProperty('kh_lastbk') || 0);
  if (!force && Date.now() - last < 30 * 60 * 1000) return null;   // xoá dòng: luôn sao lưu trước
  const cp = drive_('/drive/v3/files/' + QLKH_ID + '/copy?supportsAllDrives=true&fields=id', { method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ name: 'BACKUP_' + Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyMMdd_HHmm') + '_APP_' + name, parents: [backupFolder_()] }) });
  if (cp.getResponseCode() !== 200) throw new Error('Không lưu được bản sao trước khi ghi (HTTP ' + cp.getResponseCode() + ')');
  P.setProperty('kh_lastbk', String(Date.now()));
  return JSON.parse(cp.getContentText()).id;
}

/**
 * Sửa 1 dòng trong sheet DANH SACH LUU TRU của file .xlsm.
 *  rowNum  : số dòng trong sheet
 *  expect  : { ma, ten } — phải khớp nội dung hiện tại của dòng, nếu không thì DỪNG (dữ liệu đã đổi)
 *  setV    : { 'E': 'Tên mới', 'L': '01/12/2026', ... }  (chuỗi rỗng = xoá ô)
 *  kinds   : { 'L': 'date' | 'num' | 'text' }
 *  opt     : { fill: 'yellow' — tô màu các ô A..AH của dòng, skipFormula: true — bỏ qua ô công thức thay vì báo lỗi }
 */
function xlsmOpen_() {
  const meta = JSON.parse(drive_('/drive/v3/files/' + KH_TARGET + '?fields=name,modifiedTime&supportsAllDrives=true').getContentText());
  const res = drive_('/drive/v3/files/' + KH_TARGET + '?alt=media&supportsAllDrives=true');
  if (res.getResponseCode() !== 200) throw new Error('Không tải được file trên Drive (HTTP ' + res.getResponseCode() + ')');
  const blob = res.getBlob(); blob.setContentType('application/zip');
  const parts = Utilities.unzip(blob), byName = {}, order = [];
  parts.forEach(b => { byName[b.getName()] = b; order.push(b.getName()); });
  const get = n => byName[n] ? byName[n].getDataAsString('UTF-8') : null;
  const wb = get('xl/workbook.xml'), rels = get('xl/_rels/workbook.xml.rels');
  let rid = null, tgt = null;
  wb.replace(/<sheet\b[^>]*>/g, t => { const nm = (t.match(/ name="([^"]*)"/) || [])[1]; if (nm != null && xUnesc_(nm) === QLKH_SHEET) rid = (t.match(/ r:id="([^"]*)"/) || [])[1]; return t; });
  if (!rid) throw new Error('Không thấy sheet ' + QLKH_SHEET);
  rels.replace(/<Relationship\b[^>]*>/g, t => { if ((t.match(/ Id="([^"]*)"/) || [])[1] === rid) tgt = (t.match(/ Target="([^"]*)"/) || [])[1]; return t; });
  const sheetPath = tgt.charAt(0) === '/' ? tgt.slice(1) : 'xl/' + tgt.replace(/^\.\//, '');
  const sst = [];
  (get('xl/sharedStrings.xml') || '').replace(/<si>([\s\S]*?)<\/si>/g, (m, si) => { let t = ''; si.replace(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g, (a, b) => { t += b; return a; }); sst.push(xUnesc_(t)); return m; });
  const cellVal = c => {
    if (!c) return '';
    if (c.t === 'inlineStr') { let t = ''; c.inner.replace(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g, (a, b) => { t += b; return a; }); return xUnesc_(t); }
    const v = (c.inner.match(/<v>([\s\S]*?)<\/v>/) || [])[1]; if (v == null) return '';
    return c.t === 's' ? (sst[+v] || '') : xUnesc_(v);
  };
  const parseRow = xml => {
    const out = {};
    xml.replace(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, (m, attrs, inner) => {
      const ref = attrs.match(/ r="([A-Z]+)(\d+)"/); if (!ref) return m;
      out[colNum_(ref[1])] = { xml: m, s: (attrs.match(/ s="(\d+)"/) || [])[1] || '0', t: (attrs.match(/ t="(\w+)"/) || [])[1] || '', inner: inner || '' };
      return m;
    });
    return out;
  };
  const rowRe = r => new RegExp('<row\\b[^>]*? r="' + r + '"[^>]*?(?:\\/>|>[\\s\\S]*?<\\/row>)');
  return { meta, byName, order, get, wb, sheetPath, s: get(sheetPath), cellVal, parseRow, rowRe };
}

/** Mở file 1 lần, sửa nhiều dòng trong bộ nhớ, rồi ghi lên Drive 1 lần (commit). */
function xlsmSession_() {
  const X = xlsmOpen_(), meta = X.meta, byName = X.byName, order = X.order, get = X.get, wb = X.wb, sheetPath = X.sheetPath;
  const cellVal = X.cellVal, parseRow = X.parseRow, rowRe = X.rowRe;
  let s = X.s, dirty = false;

  // --- kiểm tra cấu trúc (dòng tiêu đề)
  let hdr = null;
  for (let r = 1; r <= 20 && !hdr; r++) {
    const m = s.match(rowRe(r)); if (!m) continue;
    const cs = parseRow(m[0]);
    if (Object.keys(cs).some(k => isHdrTen_(cellVal(cs[k])))) hdr = cs;
  }
  if (!hdr) throw new Error('Không thấy dòng tiêu đề trong file');
  for (const L in K_CHECK) if (norm_(cellVal(hdr[colNum_(L)])) !== K_CHECK[L]) throw new Error('Cột ' + L + ' của file không đúng mẫu — đã dừng, không ghi.');

  // --- styles: định dạng ngày dd/mm/yyyy + màu vàng (thêm khi cần)
  let st = get('xl/styles.xml'), stDirty = false;
  const xfsM = st.match(/<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/);
  const xfs = xfsM[2].match(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g), added = [], xcache = {};
  let DATEFMT = 0, maxId = 163; const fmtCode = {}, FILLS = {};
  const nfM = st.match(/<numFmts count="(\d+)">([\s\S]*?)<\/numFmts>/);
  if (nfM) nfM[2].replace(/<numFmt numFmtId="(\d+)" formatCode="([^"]*)"\/>/g, (m, id, code) => { fmtCode[id] = xUnesc_(code); if (+id > maxId) maxId = +id; if (xUnesc_(code) === 'dd/mm/yyyy') DATEFMT = +id; return m; });
  const dateFmt = () => {
    if (DATEFMT) return DATEFMT;
    DATEFMT = maxId + 1; stDirty = true;
    const nf = '<numFmt numFmtId="' + DATEFMT + '" formatCode="dd/mm/yyyy"/>';
    st = nfM ? st.replace(nfM[0], '<numFmts count="' + (+nfM[1] + 1) + '">' + nfM[2] + nf + '</numFmts>') : st.replace('<fonts', '<numFmts count="1">' + nf + '</numFmts><fonts');
    return DATEFMT;
  };
  const RGB = { yellow: 'FFFFEB9C', red: 'FFFF0000' };
  const fillId = color => {
    if (color === 'none') return 0;
    if (FILLS[color] != null) return FILLS[color];
    const fm = st.match(/<fills count="(\d+)">([\s\S]*?)<\/fills>/);
    FILLS[color] = +fm[1]; stDirty = true;
    st = st.replace(fm[0], '<fills count="' + (FILLS[color] + 1) + '">' + fm[2] + '<fill><patternFill patternType="solid"><fgColor rgb="' + RGB[color] + '"/><bgColor indexed="64"/></patternFill></fill></fills>');
    return FILLS[color];
  };
  const isDateFmt = id => { id = +id; if ((id >= 14 && id <= 22) || (id >= 45 && id <= 47)) return true; const c = fmtCode[id]; return !!c && /[dy]/i.test(c.replace(/"[^"]*"|\[[^\]]*\]/g, '')) && !/[#0]/.test(c); };
  const xfFor = (base, kind, fill) => {
    const key = base + '|' + kind + '|' + (fill || '');
    if (xcache[key] != null) return xcache[key];
    let x = xfs[+base] || xfs[0];
    const nid = +(x.match(/ numFmtId="(\d+)"/) || [0, 0])[1];
    let setFmt = null;
    if (kind === 'date' && !isDateFmt(nid)) setFmt = dateFmt();
    if (kind === 'text' && isDateFmt(nid)) setFmt = 49;
    if (setFmt == null && !fill) { xcache[key] = +base; return +base; }
    if (setFmt != null) {
      x = / numFmtId="/.test(x) ? x.replace(/ numFmtId="\d+"/, ' numFmtId="' + setFmt + '"') : x.replace('<xf ', '<xf numFmtId="' + setFmt + '" ');
      x = / applyNumberFormat="/.test(x) ? x.replace(/ applyNumberFormat="\d"/, ' applyNumberFormat="1"') : x.replace('<xf ', '<xf applyNumberFormat="1" ');
    }
    if (fill) {
      const Y = fillId(fill);
      x = / fillId="/.test(x) ? x.replace(/ fillId="\d+"/, ' fillId="' + Y + '"') : x.replace('<xf ', '<xf fillId="' + Y + '" ');
      x = / applyFill="/.test(x) ? x.replace(/ applyFill="\d"/, ' applyFill="1"') : x.replace('<xf ', '<xf applyFill="1" ');
    }
    xcache[key] = xfs.length + added.length; added.push(x); stDirty = true;
    return xcache[key];
  };
  const serial = d => { const m = String(d).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); return m ? Math.round((Date.UTC(+m[3], +m[2] - 1, +m[1]) - Date.UTC(1899, 11, 30)) / 86400000) : null; };

  /** Sửa 1 dòng trong bộ nhớ. expect {ma, ten} phải khớp nội dung hiện tại; opt {fill:'yellow', skipFormula} */
  function edit(rowNum, expect, setV, kinds, opt) {
    opt = opt || {}; expect = expect || {};
    let rm = s.match(rowRe(rowNum));
    const matches = cs => (expect.ma == null || cmp_(expect.ma) === cmp_(cellVal(cs[colNum_('D')]))) && (expect.ten == null || cmp_(expect.ten) === cmp_(cellVal(cs[colNum_('E')])));
    if (rm && !matches(parseRow(rm[0])) && expect.ma && expect.ten) {      // dòng đã dịch (do xoá dòng khách phụ) → tìm lại gần đó
      for (let d = 1; d <= 15; d++) for (const r of [rowNum - d, rowNum + d]) {
        const m2 = r > 8 && s.match(rowRe(r)); if (m2 && matches(parseRow(m2[0]))) { rowNum = r; rm = m2; d = 99; break; }
      }
    }
    if (!rm) return { ok: false, err: 'norow', msg: 'Không thấy dòng ' + rowNum + ' trong file' };
    const rowXml = rm[0], cells = parseRow(rowXml);
    const curMa = cmp_(cellVal(cells[colNum_('D')])), curTen = cmp_(cellVal(cells[colNum_('E')]));
    if ((expect.ma != null && cmp_(expect.ma) !== curMa) || (expect.ten != null && cmp_(expect.ten) !== curTen)) {
      return { ok: false, err: 'changed', msg: 'Dòng ' + rowNum + ' trong file đã thay đổi (hiện là: ' + (curMa || '—') + ' · ' + (curTen || 'trống') + '). Tải lại danh sách rồi sửa lại.' };
    }
    const rowFill = opt.fill || null, colFill = opt.colFill || {};   // colFill: { E: 'red' } — tô màu riêng từng ô
    const fillOf = col => colFill[colName_(col)] || rowFill;
    const newCell = (col, v, base, kind) => {
      const ref = colName_(col) + rowNum, sa = ' s="' + xfFor(base, kind, fillOf(col)) + '"';
      v = String(v == null ? '' : v).trim();
      if (v === '') return '<c r="' + ref + '"' + sa + '/>';
      if (kind === 'formula') return '<c r="' + ref + '"' + sa + '><f>' + xEsc_(v.replace(/^=/, '')) + '</f></c>';
      if (kind === 'date') { const se = serial(v); if (se != null) return '<c r="' + ref + '"' + sa + '><v>' + se + '</v></c>'; }
      if (kind === 'num') { let n = v.replace(/[,\s]/g, ''); if (/^\d{1,3}(\.\d{3})+$/.test(n)) n = n.replace(/\./g, ''); if (/^-?\d+(\.\d+)?$/.test(n)) return '<c r="' + ref + '"' + sa + '><v>' + n + '</v></c>'; }
      const t = xEsc_(v);
      return '<c r="' + ref + '"' + sa + ' t="inlineStr"><is><t' + (t !== t.trim() ? ' xml:space="preserve"' : '') + '>' + t + '</t></is></c>';
    };
    const byCol = {};
    Object.keys(cells).forEach(c => { byCol[+c] = cells[c].xml; });
    for (const L of Object.keys(setV)) {
      const c = colNum_(L);
      if (c >= colNum_('AI') && c <= colNum_('AN')) return { ok: false, err: 'formula', msg: 'Không được ghi vào cột công thức ' + L };
      const old = cells[c];
      if (old && /<f[ >]/.test(old.xml)) { if (opt.skipFormula) continue; return { ok: false, err: 'formula', msg: 'Ô ' + L + rowNum + ' là công thức — không ghi đè' }; }
      byCol[c] = newCell(c, setV[L], old ? old.s : '0', kinds[L] || 'text');
    }
    const paint = (c, color) => {
      if (setV[colName_(c)] != null && !/<f[ >]/.test((cells[c] || {}).xml || '')) return;   // ô vừa ghi đã có màu
      const old = cells[c];
      if (!old) { byCol[c] = '<c r="' + colName_(c) + rowNum + '" s="' + xfFor('0', 'keep', color) + '"/>'; return; }
      byCol[c] = / s="\d+"/.test(old.xml) ? old.xml.replace(/ s="\d+"/, ' s="' + xfFor(old.s, 'keep', color) + '"') : old.xml.replace('<c ', '<c s="' + xfFor('0', 'keep', color) + '" ');
    };
    if (rowFill) for (let c = 1; c <= colNum_('AH'); c++) paint(c, rowFill);
    Object.keys(colFill).forEach(L => paint(colNum_(L), colFill[L]));
    const tag = rowXml.match(/^<row\b[^>]*?(?=\/?>)/)[0].replace(/ spans="[^"]*"/, '');
    const newRow = tag + '>' + Object.keys(byCol).map(Number).sort((a, b) => a - b).map(c => byCol[c]).join('') + '</row>';
    s = s.replace(rowXml, () => newRow); dirty = true;
    return { ok: true, row: rowNum };
  }

  /** Đưa khách ở dòng fromRow lên dòng chính mainRow: chép các ô thông tin cá nhân (giữ định dạng dòng chính). */
  function promote(mainRow, fromRow, cols) {
    const rm = s.match(rowRe(mainRow)), fm = s.match(rowRe(fromRow));
    if (!rm || !fm) return { ok: false, msg: 'Không thấy dòng ' + (!rm ? mainRow : fromRow) };
    const mc = parseRow(rm[0]), fc = parseRow(fm[0]), byCol = {};
    Object.keys(mc).forEach(c => { byCol[+c] = mc[c].xml; });
    cols.forEach(L => {
      const c = colNum_(L), src = fc[c], old = mc[c];
      if (old && /<f[ >]/.test(old.xml)) return;
      const st = old && old.s != null && old.s !== '' ? old.s : (src && src.s != null && src.s !== '' ? src.s : null);
      if (src && !/<f[ >]/.test(src.xml) && cellVal(src) !== '') {
        let x = src.xml.replace(/ r="[A-Z]+\d+"/, ' r="' + L + mainRow + '"');
        if (st != null) x = / s="\d+"/.test(x) ? x.replace(/ s="\d+"/, ' s="' + st + '"') : x.replace('<c ', '<c s="' + st + '" ');
        byCol[c] = x;
      } else byCol[c] = '<c r="' + L + mainRow + '"' + (st != null ? ' s="' + st + '"' : '') + '/>';
    });
    const tag = rm[0].match(/^<row\b[^>]*?(?=\/?>)/)[0].replace(/ spans="[^"]*"/, '');
    s = s.replace(rm[0], () => tag + '>' + Object.keys(byCol).map(Number).sort((a, b) => a - b).map(c => byCol[c]).join('') + '</row>');
    dirty = true; return { ok: true };
  }

  /** Các dòng khách phụ còn trống trong cùng căn (cùng mã phòng cột D, liền nhau) — trừ dòng đầu (dòng khách chính). */
  function blockExtras(rowNum) {
    const at = r => { const m = r > 8 && s.match(rowRe(r)); if (!m) return null; const cs = parseRow(m[0]); return { ma: cmp_(cellVal(cs[colNum_('D')])), ten: cmp_(cellVal(cs[colNum_('E')])) }; };
    const me = at(rowNum); if (!me || !me.ma) return [];
    let a = rowNum, b = rowNum, x;
    while ((x = at(a - 1)) && x.ma === me.ma) a--;
    while ((x = at(b + 1)) && x.ma === me.ma) b++;
    const out = [];
    for (let r = a + 1; r <= b; r++) if (!at(r).ten) out.push(r);
    return out;
  }
  let DEL = {};
  const OPS = [];          // các lần chèn/xoá dòng (để sửa hình vẽ, ghi chú, tên vùng lúc ghi file)
  function delRows(list) { (list || []).forEach(r => { DEL[r] = 1; }); if ((list || []).length) dirty = true; }
  /** Chèn n dòng trống ngay sau dòng `after`, chép định dạng + công thức của dòng mẫu `tmpl` (áp ngay trong bộ nhớ). */
  function insRows(after, n, tmpl) {
    if (!(n > 0)) return [];
    const M = XD_.mkIns(after, n);
    s = xdSheetM_(s, M, QLKH_SHEET); OPS.push(M);
    const tm = s.match(rowRe(M.one(tmpl)));
    if (!tm) throw new Error('Không thấy dòng mẫu ' + tmpl);
    const tr = M.one(tmpl), cs = parseRow(tm[0]), tag = tm[0].match(/^<row\b[^>]*?(?=\/?>)/)[0].replace(/ spans="[^"]*"/, '');
    let add = '', rows = [];
    for (let k = 1; k <= n; k++) {
      const nr = after + k; rows.push(nr);
      add += tag.replace(/ r="\d+"/, ' r="' + nr + '"') + '>' + Object.keys(cs).map(Number).sort((a, b) => a - b).map(c => {
        const x = cs[c], ref = colName_(c) + nr, sa = x.s && x.s !== '0' ? ' s="' + x.s + '"' : '';
        const f = x.xml.match(/<f>([^<]*)<\/f>/);
        return f ? '<c r="' + ref + '"' + sa + '><f>' + xEsc_(XD_.shift(xUnesc_(f[1]), nr - tr, 0)) + '</f></c>' : '<c r="' + ref + '"' + sa + '/>';
      }).join('') + '</row>';
    }
    const am = s.match(rowRe(after));
    if (am) { const i = s.indexOf(am[0]) + am[0].length; s = s.slice(0, i) + add + s.slice(i); }
    else throw new Error('Không thấy dòng ' + after);
    dirty = true;
    return rows;
  }
  /** Các dòng liền nhau có mã phòng (cột D) = ma: [{row, ten}] */
  function findBlock(ma) {
    ma = cmp_(ma); const out = [];
    const a = s.indexOf('<sheetData'), b = s.indexOf('>', a) + 1, e = s.indexOf('</sheetData>');
    s.slice(b, e).replace(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g, rx => {
      const r = +(rx.match(/ r="(\d+)"/) || [])[1]; if (r <= 8) return rx;
      const cs = parseRow(rx), m = cmp_(cellVal(cs[colNum_('D')]));
      if (m === ma && (!out.length || out[out.length - 1].row === r - 1)) out.push({ row: r, ten: cmp_(cellVal(cs[colNum_('E')])), toa: String(cellVal(cs[colNum_('C')]) || '').trim() });
      return rx;
    });
    return out;
  }

  /** Các dòng khách có số liệu lệch vào cột kỳ tính tiền Y..AC (không phải ngày / AA không phải công thức). */
  function scanShifted() {
    const out = [];
    const sdA = s.indexOf('<sheetData'), sdB = s.indexOf('>', sdA) + 1, sdE = s.indexOf('</sheetData>');
    s.slice(sdB, sdE).replace(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g, rx => {
      const n = +(rx.match(/ r="(\d+)"/) || [])[1]; if (n <= 8) return rx;
      const cs = parseRow(rx);
      if (!String(cellVal(cs[colNum_('E')])).trim()) return rx;
      const info = {}; let bad = false, adText = false;
      ['Y', 'Z', 'AA', 'AB', 'AC', 'AD'].forEach(L => {
        const c = cs[colNum_(L)], v = c ? cellVal(c) : '';
        const isD = !!c && c.t === '' && /^-?\d+(\.\d+)?$/.test(v) && isDateFmt(+((xfs[+c.s] || '').match(/ numFmtId="(\d+)"/) || [0, 0])[1]);
        info[L] = v;
        if (['Y', 'Z', 'AB', 'AC'].indexOf(L) >= 0 && v !== '' && !isD) bad = true;          // kỳ tính tiền phải là ngày
        if (L === 'AD' && c && v !== '' && (c.t === 's' || c.t === 'inlineStr' || c.t === 'str')) adText = true;   // chỉ số điện cũ phải là số
      });
      info.adText = adText;
      if (bad) out.push({ row: n, ma: cellVal(cs[colNum_('D')]), ten: cellVal(cs[colNum_('E')]), old: info });
      return rx;
    });
    return out;
  }

  /** Sao lưu (nếu cần) rồi ghi file lên Drive 1 lần. */
  function commit() {
    if (!dirty) return { ok: true, nothing: true };
    const changed = {}, D = Object.keys(DEL).map(Number).sort((a, b) => a - b);
    let wbx = wb;
    if (D.length) { const M = XD_.mk(D); s = xdSheetM_(s, M, QLKH_SHEET); OPS.push(M); }
    if (OPS.length) {
      const rp = sheetPath.replace(/worksheets\/(sheet\d+\.xml)$/, 'worksheets/_rels/$1.rels');
      (get(rp) || '').replace(/<Relationship\b[^>]*>/g, t => {
        const ty = (t.match(/ Type="([^"]*)"/) || [])[1] || '', tg = (t.match(/ Target="([^"]*)"/) || [])[1] || '';
        if (/TargetMode="External"/.test(t)) return t;
        const path = tg.charAt(0) === '/' ? tg.slice(1) : ('xl/' + tg.replace(/^\.\.\//, '')).replace('xl/worksheets/', 'xl/worksheets/');
        const kind = /\/drawing$/.test(ty) ? 'drawing' : /vmlDrawing$/.test(ty) ? 'vml' : /\/comments$/.test(ty) ? 'comments' : '';
        if (kind && get(path) != null) { let x = get(path); OPS.forEach(M => { x = xdAnchorsM_(x, M, kind); }); changed[path] = x; }
        return t;
      });
      wbx = wb.replace(/(<definedName\b[^>]*>)([^<]*)(<\/definedName>)/g, (m, o, f, c) => { let x = xUnesc_(f); OPS.forEach(M => { x = XD_.fixFormula(x, M, QLKH_SHEET, false); }); return o + xEsc_(x) + c; });
    }
    changed[sheetPath] = s;
    if (stDirty) { changed['xl/styles.xml'] = st.replace(xfsM[0], '<cellXfs count="' + (xfs.length + added.length) + '">' + xfs.join('') + added.join('') + '</cellXfs>'); }
    changed['xl/workbook.xml'] = /<calcPr\b[^>]*\/>/.test(wbx)
      ? wbx.replace(/<calcPr\b([^>]*)\/>/, (m, a) => '<calcPr' + a.replace(/ fullCalcOnLoad="\w+"/, '') + ' fullCalcOnLoad="1"/>')
      : wbx.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');
    const backupId = KH_TARGET === QLKH_ID ? maybeBackup_(meta.name, OPS.length > 0) : null;
    const blobs = [], first = '[Content_Types].xml';
    [first].concat(order.filter(n => n !== first)).forEach(n => {
      if (!byName[n]) return;
      blobs.push(changed[n] != null ? Utilities.newBlob(Utilities.newBlob(changed[n]).getBytes(), 'application/xml', n) : byName[n]);
    });
    const zip = Utilities.zip(blobs, meta.name);
    const up = drive_('/upload/drive/v3/files/' + KH_TARGET + '?uploadType=media&supportsAllDrives=true', { method: 'patch', contentType: QLKH_MIME, payload: zip.getBytes() });
    if (up.getResponseCode() !== 200) throw new Error('Ghi file lên Drive không được (HTTP ' + up.getResponseCode() + ')');
    return { ok: true, backupId: backupId, deleted: D.length };
  }
  return { edit: edit, commit: commit, scanShifted: scanShifted, blockExtras: blockExtras, delRows: delRows, insRows: insRows, findBlock: findBlock, promote: promote };
}
/** Sửa 1 dòng và ghi ngay (dùng cho hàm thử). */
function xlsmEditRow_(rowNum, expect, setV, kinds, opt) {
  const S = xlsmSession_(), r = S.edit(rowNum, expect, setV, kinds, opt);
  if (!r.ok) return r;
  return S.commit();
}

/* ======================= LỆNH CHO APP (HÀNG CHỜ — LƯU NHANH) =======================
 * App bấm Lưu -> khqueue: chỉ ghi nhận thay đổi vào hàng chờ (~1 giây) rồi trả lời ngay.
 * Sau đó app gọi khapply (chạy nền): mở file 1 lần, áp TẤT CẢ thay đổi đang chờ, ghi file 1 lần.
 * Dù app bị tắt, lần gọi khapply / khlist / khqueue kế tiếp sẽ ghi nốt phần còn chờ.
 */
const KQ_PREFIX = 'khq_', KQ_ERR = 'khq_errors';
function editSpec_(fields) {
  const setV = {}, kinds = {};
  Object.keys(fields || {}).forEach(k => {
    if (K_EDIT.indexOf(k) < 0) return;
    const L = KF[k]; setV[L] = fields[k];
    kinds[L] = K_DATE.indexOf(k) >= 0 ? 'date' : K_NUM.indexOf(k) >= 0 ? 'num' : 'text';
  });
  return { setV: setV, kinds: kinds };
}
function outSpec_() {
  const setV = {}, kinds = {};
  for (let c = colNum_('E'); c <= colNum_('AC'); c++) { setV[colName_(c)] = ''; kinds[colName_(c)] = 'text'; }
  return { setV: setV, kinds: kinds };
}
const K_ADD = ['ten', 'ns', 'qt', 'gt', 'cccd', 'hc', 'dc', 'bd', 'kt', 'lt', 'coc', 'tp', 'sn', 'dck', 'nuoc', 'bt', 'ttnp', 'sale', 'lh', 'sdt'];
const K_SNAP = ['ma', 'ten', 'ns', 'qt', 'gt', 'cccd', 'hc', 'dc', 'bd', 'kt', 'lt', 'coc', 'tp', 'sn', 'sale', 'lh', 'sdt', 'gc'];
function khSnap_(g) { const o = {}; K_SNAP.forEach(k => { if (g && g[k] != null && String(g[k]) !== '') o[k] = String(g[k]).slice(0, 200); }); return o; }
const K_PERSON = ['E', 'F', 'G', 'H', 'I', 'J', 'K', 'N'];      // cột thông tin riêng từng khách (tên, NS, QT, GT, CCCD, HC, địa chỉ, hạn lưu trú)
function khQueue_(p) {
  if (p.type === 'unitout') {
    const ma = String(p.ma || '').toUpperCase().trim(); if (!ma) return { ok: false, err: 'nothing' };
    const item = { t: Date.now(), type: 'unitout', ma: ma, guests: (p.guests || []).map(khSnap_).slice(0, 12), note: String(p.note || '').slice(0, 300), openUnit: p.openUnit !== false, by: p.by || '' };
    const key = KQ_PREFIX + item.t + '_' + Math.random().toString(36).slice(2, 7);
    PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(item));
    return { ok: true, id: key, t: item.t, pending: khPending_().length };
  }
  if (p.type === 'add') {
    const guests = (p.guests || []).map(g => { const f = {}; K_ADD.forEach(k => { if (g[k] != null && String(g[k]).trim() !== '') f[k] = String(g[k]).trim(); }); return f; }).filter(f => f.ten);
    if (!p.ma || !guests.length) return { ok: false, err: 'nothing', msg: 'Chưa có mã phòng hoặc khách' };
    const item = { t: Date.now(), type: 'add', ma: String(p.ma).toUpperCase().trim(), guests: guests, force: !!p.force, file: String(p.file || '').slice(0, 200), by: p.by || 'Chủ' };
    const key = KQ_PREFIX + item.t + '_' + Math.random().toString(36).slice(2, 7);
    PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(item));
    return { ok: true, id: key, t: item.t, pending: khPending_().length };
  }
  const row = +p.row;
  if (!(row > 8)) return { ok: false, err: 'badrow' };
  const item = { t: Date.now(), type: p.type === 'out' ? 'out' : 'save', row: row, expect: p.expect || {} };
  if (item.type === 'save') {
    const f = {}; Object.keys(p.fields || {}).forEach(k => { if (K_EDIT.indexOf(k) >= 0) f[k] = p.fields[k]; });
    if (!Object.keys(f).length) return { ok: false, err: 'nothing' };
    item.fields = f;
  } else {
    item.guest = khSnap_(p.guest || {}); item.note = p.note || ''; item.openUnit = !!p.openUnit;
    item.expect = { ma: item.guest.ma, ten: item.guest.ten };
  }
  const key = KQ_PREFIX + item.t + '_' + Math.random().toString(36).slice(2, 7);
  PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(item));
  return { ok: true, id: key, t: item.t, pending: khPending_().length };
}
function khPending_() {
  const all = PropertiesService.getScriptProperties().getProperties();
  return Object.keys(all).filter(k => k.indexOf(KQ_PREFIX) === 0 && k !== KQ_ERR).sort().map(k => { const o = JSON.parse(all[k]); o.key = k; return o; });
}
function khErrors_() { try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(KQ_ERR) || '[]'); } catch (e) { return []; } }
/** Áp mọi thay đổi đang chờ vào file (1 lần tải + 1 lần ghi). Gọi bên trong khoá. */
function khApply_() {
  const P = PropertiesService.getScriptProperties();
  const all = khPending_();
  if (!all.length) return { ok: true, applied: 0, errors: khErrors_() };
  const items = all.filter(it => it.type !== 'add'), adds = all.filter(it => it.type === 'add');
  if (!items.length) return khApplyAdds_(adds);
  const S = xlsmSession_(), done = [], errs = khErrors_(), after = [], outRows = [];
  let delList = [];
  const bad = (it, msg) => errs.unshift({ t: it.t, row: it.row || 0, ma: it.ma || (it.expect || {}).ma || '', ten: (it.expect || {}).ten || '', type: it.type, msg: msg, fields: it.fields || null });
  items.forEach(it => {
    if (it.type === 'unitout') {                           // CHECK-OUT CẢ CĂN: xoá trắng dòng chính (tô đỏ ô tên), xoá hẳn mọi dòng khách phụ
      const blk = S.findBlock(it.ma);
      if (!blk.length) return bad(it, 'Không thấy căn ' + it.ma + ' trong file');
      const sp = outSpec_(), r = S.edit(blk[0].row, { ma: it.ma }, sp.setV, sp.kinds, { colFill: { E: 'red' }, skipFormula: true });
      if (!r.ok) return bad(it, r.msg || r.err);
      delList = delList.concat(blk.slice(1).map(x => x.row));
      done.push(it); after.push(it); return;
    }
    if (it.type === 'out') {
      // khách CHÍNH (dòng đầu của căn) ra mà còn khách phụ ở lại → đưa khách phụ kế tiếp lên dòng chính, xoá dòng khách phụ đó
      const chk = S.edit(it.row, it.expect, {}, {}, {});
      if (!chk.ok) return bad(it, chk.msg || chk.err);
      const blk = S.findBlock((it.expect || {}).ma || '');
      const nx = blk.length && blk[0].row === chk.row ? blk.slice(1).find(x => x.ten) : null;
      if (nx) {
        const pr = S.promote(chk.row, nx.row, K_PERSON);
        if (!pr.ok) return bad(it, pr.msg);
        delList.push(nx.row); done.push(it); after.push(it); return;
      }
      const sp = outSpec_(), r = S.edit(chk.row, it.expect, sp.setV, sp.kinds, { colFill: { E: 'red' }, skipFormula: true });
      if (r.ok) { done.push(it); after.push(it); outRows.push(r.row); } else bad(it, r.msg || r.err);
      return;
    }
    const sp = editSpec_(it.fields);
    const r = S.edit(it.row, it.expect, sp.setV, sp.kinds, {});
    if (r.ok) done.push(it); else bad(it, r.msg || r.err);
  });
  // căn đã OUT: chỉ giữ 1 dòng khách chính, xoá hẳn các dòng khách phụ đã trống
  outRows.forEach(r => { delList = delList.concat(S.blockExtras(r)); });
  delList = delList.filter((r, i, a) => a.indexOf(r) === i);
  S.delRows(delList);
  const c = S.commit();   // lỗi ghi file -> ném lỗi, hàng chờ GIỮ NGUYÊN để thử lại lần sau
  items.forEach(it => P.deleteProperty(it.key));
  P.setProperty(KQ_ERR, JSON.stringify(errs.slice(0, 30)));
  // việc phụ của OUT (sau khi file đã ghi xong)
  if (!KH_TEST) after.forEach(it => {
    const gs = it.type === 'unitout' ? (it.guests || []) : [it.guest || {}], g = Object.assign({ ma: it.ma }, gs[0] || {});
    gs.forEach(x => {
      try {
        outLog_().appendRow([Utilities.formatDate(new Date(it.t), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm'), it.by || 'Admin', x.ma || it.ma, x.ten, x.ns, x.qt, x.gt, x.cccd, x.hc, x.dc,
          x.bd, x.kt, x.lt, x.coc, x.tp, x.sn, x.sale, x.lh, x.sdt, x.gc, (it.type === 'unitout' ? 'CHECK-OUT CẢ CĂN' : 'Xoá khách') + (it.note ? ' · ' + it.note : '')].map(v => v == null ? '' : "'" + String(v)));
      } catch (e) {}
    });
    if (it.openUnit) {
      try { const t = unitTable_(), ma = String(g.ma || it.ma || '').replace(/-/g, '').toUpperCase(), i = findUnit_(t, ma);
        if (i >= 0 && t.C.tt >= 0) t.sh.getRange(i + 1, t.C.tt + 1).setValue('Đang có sẵn'); } catch (e) {}
    }
  });
  const res = { ok: true, applied: done.length, failed: items.length - done.length, deleted: delList.length, backupId: c.backupId || null, errors: errs.slice(0, 30) };
  if (adds.length) { const r2 = khApplyAdds_(adds); res.applied += r2.applied; res.failed += r2.failed; res.added = r2.added; res.errors = r2.errors; res.runErrors = r2.runErrors; }
  return res;
}
/** Bổ sung khách check-in vào file: khách chính vào dòng chính của căn (bỏ màu đỏ), khách thêm chèn dòng ngay dưới. */
function khApplyAdds_(adds) {
  const P = PropertiesService.getScriptProperties(), errs = khErrors_(), added = [];
  let done = 0, failed = 0; const runErr = [];
  const S = xlsmSession_();
  // khoá so tên: bỏ dấu, chỉ chữ/số, xếp từ theo ABC (tên đảo thứ tự vẫn khớp)
  const nk = t => String(t || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/gi, 'D').toUpperCase()
    .replace(/\(.*?\)/g, ' ').replace(/[^A-Z0-9]+/g, ' ').trim().split(' ').filter(Boolean).sort().join(' ');
  let reg = {}; try { reg = JSON.parse(P.getProperty('khadd_reg') || '{}'); } catch (e) {}
  adds.forEach(it => {
    const fail = (msg, dup) => { failed++; const e = { t: it.t, row: 0, ma: it.ma, ten: (it.guests[0] || {}).ten || '', type: 'add', msg: msg, dup: !!dup }; errs.unshift(e); runErr.push(e); };
    const ghide = it.mode !== 'bao' || !!it.force;       // mặc định: khách trùng tên → tự ghi đè dòng cũ
    try {
      const blk = S.findBlock(it.ma);
      if (!blk.length) return fail('Không thấy căn ' + it.ma + ' trong file Quản lý khách hàng — thêm căn vào file trước');
      const occ = blk.filter(x => x.ten), dung = {}, khop = it.guests.map(g => {
        const o = occ.find(x => !dung[x.row] && nk(x.ten) && nk(x.ten) === nk(g.ten)); if (o) dung[o.row] = 1; return o || null; });
      const nTrung = khop.filter(Boolean).length;
      if (nTrung && !ghide) return fail('TRÙNG KHÁCH: căn ' + it.ma + ' đã có ' + khop.filter(Boolean).map(o => o.ten + ' (dòng ' + o.row + ')').join(', ') + ' — bấm "Ghi đè" để cập nhật thông tin mới, hoặc bỏ qua', true);
      if (!nTrung && it.file && reg[it.file] && !ghide) return fail('Hợp đồng này đã được bổ sung vào file lúc ' + reg[it.file].at + ' (dòng ' + reg[it.file].rows.join(', ') + ') — bấm "Ghi đè" nếu muốn ghi lại', true);
      const conLai = occ.filter(x => !dung[x.row]);
      if (!nTrung && conLai.length && !it.force) return fail('Căn ' + it.ma + ' còn ' + conLai.length + ' khách chưa OUT: ' + conLai.map(x => x.ten).join(', ') + ' — OUT khách cũ trước, hoặc chọn "Vẫn bổ sung"');
      const toa = blk[0].toa || (it.ma.split('-')[1] || '');
      const moi = it.guests.filter((g, i) => !khop[i]);
      const slots = blk.filter(x => !x.ten).map(x => x.row), need = moi.length - slots.length;
      const ghi = (r, g) => {
        const sp = editSpec_(g);
        sp.setV.D = it.ma; sp.kinds.D = 'text'; sp.setV.C = toa; sp.kinds.C = 'text';
        const x = S.edit(r, {}, sp.setV, sp.kinds, { colFill: { E: 'none' }, skipFormula: true });
        if (!x.ok) throw new Error(x.msg || x.err);
      };
      const upd = [];
      it.guests.forEach((g, i) => { if (khop[i]) { ghi(khop[i].row, g); upd.push(khop[i].row); } });
      const newRows = moi.length ? slots.slice(0, moi.length).concat(need > 0 ? S.insRows(blk[blk.length - 1].row, need, blk[blk.length - 1].row) : []) : [];
      newRows.forEach((r, i) => ghi(r, moi[i]));
      const rows = upd.concat(newRows).sort((a, b) => a - b);
      done++; added.push({ t: it.t, ma: it.ma, rows: rows, upd: upd, file: it.file });
    } catch (e) { fail(String(e.message || e)); }
  });
  if (done) S.commit();
  adds.forEach(it => P.deleteProperty(it.key));
  P.setProperty(KQ_ERR, JSON.stringify(errs.slice(0, 30)));
  if (added.length) {
    const at = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm');
    added.forEach(a => { if (a.file) { reg[a.file] = { at: at, rows: a.rows }; tdMarkAdd_(a.file, 'Đã bổ sung ' + at + ' · dòng ' + a.rows.join(', ')); } });
    const ks = Object.keys(reg); if (ks.length > 600) ks.slice(0, ks.length - 600).forEach(k => delete reg[k]);
    P.setProperty('khadd_reg', JSON.stringify(reg));
  }
  return { ok: true, applied: done, failed: failed, added: added, errors: errs.slice(0, 30), runErrors: runErr };
}
/** App check-in: bổ sung khách của hợp đồng vào file. Chủ → ghi ngay; nhân viên → đề xuất chờ duyệt. */
function khAddRoute_(p) {
  const who = hdAuth_(p, ['checkin', 'checkinout']);
  if (!who) return { ok: false, err: 'noperm' };
  const add = { type: 'add', ma: p.ma, guests: p.guests, force: !!p.force, file: p.file, mode: p.mode === 'bao' ? 'bao' : 'ghide' };
  if (!who.owner) {
    if (!(p.guests || []).length || !p.ma) return { ok: false, err: 'nothing' };
    const o = { t: Date.now(), status: 'cho', who: { ma: who.ma, ten: who.ten }, type: 'add', add: add, kq: String(p.kq || '').slice(0, 300) };
    if (p.lo) { o.lo = String(p.lo).slice(0, 80); o.src = String(p.src || '').slice(0, 120); }     // lô nạp từ file DS Check-in → chủ duyệt cả lô
    o.id = KP_PREFIX + o.t + '_' + Math.random().toString(36).slice(2, 7);
    kpSave_(o);
    if (!p.lo || p.loFirst) { try { nkAppend_(who.ma, who.ten, p.lo ? 'Gửi lô DS Check-in chờ duyệt' : kpLabel_(o), p.lo ? (o.src + ' · ' + (p.loN || '') + ' căn') : kpDetail_(o), 'checkin.html', p.devname); } catch (e) {} }
    return { ok: true, proposed: true, id: o.id };
  }
  add.by = 'Chủ';
  const q = khQueue_(add); if (!q.ok) return q;
  if (p.queueOnly) return { ok: true, queued: true, t: q.t };     // nạp nhiều căn: xếp hàng hết rồi gọi khaddrun ghi 1 lần
  let r; try { r = withLock_(() => khApply_()); } catch (e) { return { ok: false, err: 'apply', msg: String(e.message || e) }; }
  const a = (r.added || []).find(x => x.t === q.t), e = (r.runErrors || r.errors || []).find(x => x.t === q.t);
  if (a) return { ok: true, rows: a.rows, upd: a.upd || [] };
  return { ok: false, err: 'add', dup: !!(e && e.dup), msg: e ? e.msg : 'Chưa ghi được — thử lại' };
}
/** Chủ: ghi tất cả khách đang xếp hàng vào file trong 1 lần (1 lần tải / ghi file) */
function khAddRun_(p) {
  const who = hdAuth_(p, []);
  if (!who || !who.owner) return { ok: false, err: 'noperm' };
  let r; try { r = withLock_(() => khApply_()); } catch (e) { return { ok: false, err: 'apply', msg: String(e.message || e) }; }
  return { ok: true, added: r.added || [], errors: r.runErrors || r.errors || [] };
}
function khAddStatus_(p) {
  const who = hdAuth_(p, ['checkin', 'checkinout']);
  if (!who) return { ok: false, err: 'noperm' };
  let reg = {}; try { reg = JSON.parse(PropertiesService.getScriptProperties().getProperty('khadd_reg') || '{}'); } catch (e) {}
  const pend = {}; kpAll_().forEach(o => { if (o.type === 'add' && o.status === 'cho' && o.add && o.add.file) pend[o.add.file] = 1; });
  const out = {}; [].concat(p.files || []).forEach(f => { if (reg[f]) out[f] = { done: reg[f] }; else if (pend[f]) out[f] = { pend: 1 }; });
  return { ok: true, files: out };
}
function khStatus_() { return { ok: true, pending: khPending_().length, errors: khErrors_() }; }
function khClearErr_() { PropertiesService.getScriptProperties().deleteProperty(KQ_ERR); return { ok: true }; }
function outLog_() {
  const P = PropertiesService.getScriptProperties();
  let id = P.getProperty('kh_outlog'), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) {} }
  if (!ss) {
    ss = SpreadsheetApp.create(OUT_LOG_NAME);
    ss.getSheets()[0].appendRow(['Ngày trả phòng', 'Người thao tác', 'Mã phòng', 'Tên khách', 'Năm sinh', 'Quốc tịch', 'Giới tính', 'CCCD', 'Hộ chiếu', 'Địa chỉ',
      'Bắt đầu HĐ', 'Kết thúc HĐ', 'Hạn lưu trú', 'Tiền cọc', 'Tiền phòng', 'Số người', 'SALE', 'Liên hệ', 'SĐT', 'Ghi chú', 'Lý do / ghi chú OUT']);
    ss.getSheets()[0].setFrozenRows(1);
    P.setProperty('kh_outlog', ss.getId());
  }
  return ss.getSheets()[0];
}
function khExport_() {
  const res = drive_('/drive/v3/files/' + QLKH_ID + '?alt=media&supportsAllDrives=true');
  if (res.getResponseCode() !== 200) return { ok: false, err: 'download' };
  return { ok: true, name: 'QUAN LY KHACH HANG_' + Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'ddMMyy_HHmm') + '.xlsm',
    mime: QLKH_MIME, data: Utilities.base64Encode(res.getBlob().getBytes()) };
}
function khOutList_() {
  const id = PropertiesService.getScriptProperties().getProperty('kh_outlog');
  if (!id) return { ok: true, rows: [] };
  const v = SpreadsheetApp.openById(id).getSheets()[0].getDataRange().getDisplayValues();
  return { ok: true, header: v[0], rows: v.slice(1).reverse().slice(0, 300), url: SpreadsheetApp.openById(id).getUrl() };
}

/* ======================= XOÁ HẲN DÒNG (khách phụ đã OUT) =======================
 * Xoá dòng trong file .xlsm như Excel "Delete row": các dòng dưới dịch lên, đồng thời sửa
 * vị trí tham chiếu trong công thức, định dạng có điều kiện, bộ lọc, ô gộp, liên kết, ghi chú, hình vẽ, tên vùng.
 */
const XD_ = {
  // ánh xạ dòng sau khi xoá; D = mảng dòng bị xoá (tăng dần)
  mk(D) {
    const del = {}; D.forEach(r => del[r] = 1);
    const below = r => { let n = 0; for (let i = 0; i < D.length && D[i] < r; i++) n++; return n; };
    const upTo = r => { let n = 0; for (let i = 0; i < D.length && D[i] <= r; i++) n++; return n; };
    return {
      del: del,
      one: r => del[r] ? null : r - below(r),
      range: (a, b) => { const s = a - below(a), e = b - upTo(b); return e < s ? null : [s, e]; }
    };
  },
  // ánh xạ khi CHÈN n dòng ngay sau dòng `after`
  mkIns(after, n) {
    const f = r => r > after ? r + n : r;
    return { del: {}, one: f, range: (a, b) => [f(a), f(b)] };
  },
  // tham chiếu ô/vùng trong công thức: (sheet!)?A1(:B2)? hoặc 9:12
  RE: /('(?:[^']|'')+'!|[A-Za-z_][\w.]*!)?(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?|('(?:[^']|'')+'!|[A-Za-z_][\w.]*!)?(\$?)(\d+):(\$?)(\d+)/g,
  /** Áp fn lên từng tham chiếu (bỏ qua chuỗi "..." ); fn(m, info) trả chuỗi thay thế */
  walk(f, fn) {
    let out = '', i = 0;
    const parts = f.split(/("(?:[^"]|"")*")/);
    return parts.map((p, k) => k % 2 ? p : p.replace(XD_.RE, function (m) {
      const a = arguments, off = a[a.length - 2], str = a[a.length - 1];
      const prev = off > 0 ? str.charAt(off - 1) : '', next = str.charAt(off + m.length);
      if (/[A-Za-z0-9_.$]/.test(prev) || /[(A-Za-z0-9_!]/.test(next)) return m;   // tên hàm (LOG10() / tên vùng / một phần chữ khác
      return fn(m, a);
    })).join('');
  },
  sheetOf(pref) { if (!pref) return null; pref = pref.slice(0, -1); return pref.charAt(0) === "'" ? pref.slice(1, -1).replace(/''/g, "'") : pref; },
  /** Sửa công thức sau khi xoá dòng. own = tên sheet bị xoá dòng; inOwn = công thức nằm trong sheet đó */
  fixFormula(f, M, own, inOwn) {
    return XD_.walk(f, (m, a) => {
      if (a[3] != null) {           // dạng ô
        const sh = XD_.sheetOf(a[1]); if (sh ? sh !== own : !inOwn) return m;
        const pre = a[1] || '';
        if (a[7] == null) { const r = M.one(+a[5]); return r == null ? pre + '#REF!' : pre + a[2] + a[3] + a[4] + r; }
        const rg = M.range(+a[5], +a[9]); if (!rg) return pre + '#REF!';
        return pre + a[2] + a[3] + a[4] + rg[0] + ':' + a[6] + a[7] + a[8] + rg[1];
      }
      const sh = XD_.sheetOf(a[10]); if (sh ? sh !== own : !inOwn) return m;   // dạng cả dòng 9:12
      const rg = M.range(+a[12], +a[14]); if (!rg) return (a[10] || '') + '#REF!';
      return (a[10] || '') + a[11] + rg[0] + ':' + a[13] + rg[1];
    });
  },
  /** Dịch công thức dùng chung sang ô khác (dr dòng, dc cột) — phần tương đối mới dịch */
  shift(f, dr, dc) {
    const cs = (c, abs) => abs ? c : colName_(colNum_(c) + dc), rs = (r, abs) => abs ? r : String(+r + dr);
    return XD_.walk(f, (m, a) => {
      if (a[3] != null) return (a[1] || '') + a[2] + cs(a[3], a[2]) + a[4] + rs(a[5], a[4]) + (a[7] != null ? ':' + a[6] + cs(a[7], a[6]) + a[8] + rs(a[9], a[8]) : '');
      return (a[10] || '') + a[11] + rs(a[12], a[11]) + ':' + a[13] + rs(a[14], a[13]);
    });
  },
  /** sqref / ref: "A9:X1594 L72" */
  fixSqref(sq, M) {
    return String(sq).split(/\s+/).map(t => {
      const g = t.match(/^(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?$/); if (!g) return t;
      if (g[6] == null) { const r = M.one(+g[4]); return r == null ? null : g[1] + g[2] + g[3] + r; }
      const rg = M.range(+g[4], +g[8]); return rg ? g[1] + g[2] + g[3] + rg[0] + ':' + g[5] + g[6] + g[7] + rg[1] : null;
    }).filter(Boolean).join(' ');
  },
  /** Bỏ công thức dùng chung trong sheet (mỗi ô có công thức riêng) để xoá dòng an toàn */
  unshare(s) {
    const master = {};
    s.replace(/<c r="([A-Z]+)(\d+)"[^>]*>\s*<f t="shared" ref="[^"]*" si="(\d+)"[^>]*>([^<]*)<\/f>/g, (m, c, r, si, f) => { master[si] = { c: colNum_(c), r: +r, f: xUnesc_(f) }; return m; });
    return s.replace(/(<c r="([A-Z]+)(\d+)"[^>]*>\s*)<f t="shared"(?: ref="[^"]*")? si="(\d+)"[^>]*?(?:\/>|>([^<]*)<\/f>)/g, (m, pre, c, r, si, f) => {
      const M0 = master[si]; if (!M0) return m;
      const txt = XD_.shift(M0.f, +r - M0.r, colNum_(c) - M0.c);
      return pre + '<f>' + xEsc_(txt) + '</f>';
    });
  }
};

/** Xoá dòng trong chuỗi XML của sheet (D tăng dần). Trả XML mới. */
function xdSheet_(s, D, own) { return xdSheetM_(s, XD_.mk(D), own); }
function xdSheetM_(s, M, own) {
  s = XD_.unshare(s);
  const a = s.indexOf('<sheetData'), b = s.indexOf('>', a) + 1, e = s.indexOf('</sheetData>');
  const body = s.slice(b, e).replace(/<row\b[^>]*?(?:\/>|>[\s\S]*?<\/row>)/g, rx => {
    const r = +(rx.match(/ r="(\d+)"/) || [])[1]; if (!r) return rx;
    const nr = M.one(r); if (nr == null) return '';
    rx = rx.replace(/^(<row\b[^>]*? r=")\d+"/, '$1' + nr + '"');
    rx = rx.replace(/(<c r="[A-Z]+)\d+"/g, '$1' + nr + '"');
    rx = rx.replace(/(<f\b[^>]*>)([^<]*)(<\/f>)/g, (m, o, f, c) => o.replace(/ ref="([^"]*)"/, (q, rf) => ' ref="' + XD_.fixSqref(rf, M) + '"') + xEsc_(XD_.fixFormula(xUnesc_(f), M, own, true)) + c);
    return rx;
  });
  let head = s.slice(0, b), tail = s.slice(e);
  head = head.replace(/<dimension ref="([^"]*)"/, (m, r) => '<dimension ref="' + (XD_.fixSqref(r, M) || 'A1') + '"')
             .replace(/(<selection\b[^>]*? sqref=")([^"]*)"/g, (m, p, r) => p + (XD_.fixSqref(r, M) || 'A1') + '"')
             .replace(/(<selection\b[^>]*? activeCell=")([^"]*)"/g, (m, p, r) => p + (XD_.fixSqref(r, M) || 'A1') + '"');
  tail = tail.replace(/<conditionalFormatting sqref="([^"]*)">([\s\S]*?)<\/conditionalFormatting>/g, (m, sq, inner) => {
    const n = XD_.fixSqref(sq, M); if (!n) return '';
    return '<conditionalFormatting sqref="' + n + '">' + inner.replace(/<formula>([^<]*)<\/formula>/g, (q, f) => '<formula>' + xEsc_(XD_.fixFormula(xUnesc_(f), M, own, true)) + '</formula>') + '</conditionalFormatting>';
  });
  tail = tail.replace(/<mergeCell ref="([^"]*)"\/>/g, (m, r) => { const n = XD_.fixSqref(r, M); return n && /:/.test(n) ? '<mergeCell ref="' + n + '"/>' : ''; })
             .replace(/<mergeCells count="\d+">([\s\S]*?)<\/mergeCells>/, (m, inner) => { const k = (inner.match(/<mergeCell /g) || []).length; return k ? '<mergeCells count="' + k + '">' + inner + '</mergeCells>' : ''; })
             .replace(/<hyperlink ref="([^"]*)"([^>]*)\/>/g, (m, r, rest) => { const n = XD_.fixSqref(r, M); return n ? '<hyperlink ref="' + n + '"' + rest + '/>' : ''; })
             .replace(/<hyperlinks>\s*<\/hyperlinks>/, '')
             .replace(/(<autoFilter ref=")([^"]*)"/, (m, p, r) => p + (XD_.fixSqref(r, M) || r) + '"')
             .replace(/(<dataValidation\b[^>]*? sqref=")([^"]*)"/g, (m, p, r) => p + XD_.fixSqref(r, M) + '"')
             .replace(/<brk id="(\d+)"/g, (m, r) => '<brk id="' + (M.one(+r) || M.one(+r + 1) || r) + '"');
  return head + body + tail;
}
/** Hình vẽ / ghi chú gắn theo dòng (0-based trong drawing & vml) */
function xdAnchors_(xml, D, kind) { return xdAnchorsM_(xml, XD_.mk(D), kind); }
function xdAnchorsM_(xml, M, kind) {
  const z = r => { const n = M.one(r + 1); return n == null ? (M.one(r + 2) || r + 1) - 1 : n - 1; };
  if (kind === 'drawing') return xml.replace(/<xdr:row>(\d+)<\/xdr:row>/g, (m, r) => '<xdr:row>' + z(+r) + '</xdr:row>');
  if (kind === 'vml') return xml.replace(/<x:Row>(\d+)<\/x:Row>/g, (m, r) => '<x:Row>' + z(+r) + '</x:Row>')
    .replace(/(<x:Anchor>\s*\d+\s*,\s*\d+\s*,\s*)(\d+)(\s*,\s*\d+\s*,\s*\d+\s*,\s*\d+\s*,\s*)(\d+)/g, (m, a, r1, b, r2) => a + z(+r1) + b + z(+r2));
  if (kind === 'comments') return xml.replace(/<comment ref="([^"]*)"([\s\S]*?)<\/comment>/g, (m, r, rest) => { const n = XD_.fixSqref(r, M); return n ? '<comment ref="' + n + '"' + rest + '</comment>' : ''; });
  return xml;
}

/* Chạy tay: thử OUT + XOÁ DÒNG trên BẢN SAO, rồi mở bản sao bằng Google Sheets để kiểm tra (file thật không bị đụng). */
function thuXoaDongBanSao() {
  const cp = JSON.parse(drive_('/drive/v3/files/' + QLKH_ID + '/copy?supportsAllDrives=true&fields=id', { method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ name: 'THU_XOA_DONG_XOA_DUOC.xlsm' }) }).getContentText());
  KH_TARGET = cp.id; let gs = null;
  try {
    const L0 = khList_().rows, n0 = L0.length;
    // căn có >= 2 khách: OUT tất cả
    const by = {}; L0.forEach(r => { if (r.ma && r.ten) (by[r.ma] = by[r.ma] || []).push(r); });
    const ma = Object.keys(by).find(k => by[k].length >= 2 && by[k][0].row > 300), g = by[ma];
    const below = L0.find(r => r.row === g[g.length - 1].row + 3) || {};
    Logger.log('Thử OUT căn ' + ma + ': ' + g.map(r => r.row + ' ' + r.ten).join(', ') + ' · dòng dưới để so: ' + below.row + ' ' + below.ma + ' ' + below.ten);
    const S = xlsmSession_(), os = outSpec_(), rows = [];
    g.forEach(r => { const x = S.edit(r.row, { ma: r.ma, ten: r.ten }, os.setV, os.kinds, { colFill: { E: 'red' }, skipFormula: true }); rows.push(x.row); });
    let del = []; rows.forEach(r => { del = del.concat(S.blockExtras(r)); }); del = del.filter((r, i, a) => a.indexOf(r) === i);
    S.delRows(del); Logger.log('Xoá dòng: ' + del.join(',') + ' -> ' + JSON.stringify(S.commit()));
    const L1 = khList_().rows;
    Logger.log('Số dòng: ' + n0 + ' -> ' + L1.length + ' · căn ' + ma + ' còn: ' + JSON.stringify(L1.filter(r => r.ma === ma).map(r => [r.row, r.ten])));
    const b2 = L1.find(r => r.ten === below.ten && r.ma === below.ma); Logger.log('Dòng dưới giờ ở: ' + (b2 ? b2.row + ' tiền phòng=' + b2.tp + ' điện=' + b2.dmoi : 'KHÔNG THẤY'));
    // mở bằng Google Sheets: kiểm tra file đọc được + công thức không lỗi
    gs = JSON.parse(drive_('/drive/v3/files/' + cp.id + '/copy?supportsAllDrives=true&fields=id', { method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ name: 'THU_XOA_DONG_GS', mimeType: 'application/vnd.google-apps.spreadsheet' }) }).getContentText()).id;
    const sh = SpreadsheetApp.openById(gs).getSheetByName(QLKH_SHEET), v = sh.getDataRange().getDisplayValues(), f = sh.getDataRange().getFormulas();
    let errN = 0, ex = []; v.forEach((row, i) => row.forEach((c, j) => { if (/^#(REF|VALUE|N\/A|NAME|DIV)/.test(c) && f[i][j]) { errN++; if (ex.length < 6) ex.push((i + 1) + ':' + (j + 1) + ' ' + f[i][j]); } }));
    const fb = b2 ? f[b2.row - 1].filter(Boolean).slice(0, 4) : [];
    Logger.log('Google Sheets mở được · ô công thức lỗi: ' + errN + ' ' + ex.join(' | ') + ' · công thức dòng dưới: ' + fb.join(' ; '));
    const vba = Utilities.unzip(DriveApp.getFileById(KH_TARGET).getBlob().setContentType('application/zip')).some(b => b.getName() === 'xl/vbaProject.bin');
    Logger.log('Macro VBA còn nguyên: ' + vba);
  } finally {
    KH_TARGET = QLKH_ID;
    [cp.id, gs].filter(Boolean).forEach(id => drive_('/drive/v3/files/' + id + '?supportsAllDrives=true', { method: 'patch', contentType: 'application/json', payload: JSON.stringify({ trashed: true }) }));
    Logger.log('Đã chuyển bản sao vào thùng rác.');
  }
}

/** Chạy tay: dọn các căn đã trống hết khách mà còn nhiều dòng — chỉ giữ dòng đầu. xem trước: donDongTrong(true) */
function donDongTrong(chiXem) {
  const lock = LockService.getScriptLock(); lock.waitLock(60000);
  try {
    if (khPending_().length) khApply_();
    const L = khList_().rows, by = {};
    L.forEach(r => { if (r.ma) (by[r.ma] = by[r.ma] || []).push(r); });
    let del = [];
    Object.keys(by).forEach(ma => {
      const g = by[ma].sort((a, b) => a.row - b.row);
      const lien = g.every((r, i) => !i || r.row === g[i - 1].row + 1);
      if (g.length > 1 && lien && g.every(r => !r.ten)) del = del.concat(g.slice(1).map(r => r.row));
    });
    Logger.log('Căn trống còn dòng thừa: ' + del.length + ' dòng · ' + del.map(r => r + ' ' + L.find(x => x.row === r).ma).join(', '));
    if (chiXem || !del.length) return;
    const S = xlsmSession_(); S.delRows(del); Logger.log('Đã xoá -> ' + JSON.stringify(S.commit()));
  } finally { lock.releaseLock(); }
}

/** Chạy tay: xoá các dòng KHÁCH TRÙNG (cùng căn, cùng tên) đang có trong file — giữ 1 dòng (dòng chính của căn, hoặc dòng nhiều dữ liệu nhất). chiXem=true: chỉ liệt kê. */
function donKhachTrung(chiXem) {
  const lock = LockService.getScriptLock(); lock.waitLock(60000);
  try {
    if (khPending_().length) khApply_();
    const nk = t => String(t || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/đ/gi, 'D').toUpperCase()
      .replace(/\(.*?\)/g, ' ').replace(/[^A-Z0-9]+/g, ' ').trim().split(' ').filter(Boolean).sort().join(' ');
    const L = khList_().rows.slice().sort((a, b) => a.row - b.row), del = [], log = [];
    const day = (r) => Object.keys(r).filter(k => ['row', 'ma', 'ten'].indexOf(k) < 0 && r[k] != null && String(r[k]).trim() !== '').length;
    // gom từng khối liền nhau cùng mã căn
    let blk = [];
    const xet = () => {
      if (blk.length < 2) return;
      const by = {};
      blk.forEach((r, i) => { const k = nk(r.ten); if (k) (by[k] = by[k] || []).push(Object.assign({ i: i }, r)); });
      Object.keys(by).forEach(k => {
        const g = by[k]; if (g.length < 2) return;
        const giu = g.find(r => r.i === 0) || g.slice().sort((a, b) => day(b) - day(a) || a.row - b.row)[0];
        g.forEach(r => { if (r.row !== giu.row) { del.push(r.row); log.push(r.ma + ' · ' + r.ten + ' · xoá dòng ' + r.row + ' (giữ dòng ' + giu.row + ')'); } });
      });
    };
    L.forEach(r => { if (blk.length && (r.ma !== blk[0].ma || r.row !== blk[blk.length - 1].row + 1)) { xet(); blk = []; } if (r.ma) blk.push(r); else { xet(); blk = []; } });
    xet();
    Logger.log('Khách trùng: ' + del.length + ' dòng\n' + log.join('\n'));
    if (chiXem || !del.length) return { ok: true, n: del.length, list: log };
    const S = xlsmSession_(); S.delRows(del); const c = S.commit(); Logger.log('Đã xoá -> ' + JSON.stringify(c));
    return { ok: true, n: del.length, list: log, backupId: c.backupId || null };
  } finally { lock.releaseLock(); }
}
function donKhachTrungXem() { return donKhachTrung(true); }

/* Chạy tay: thử BỔ SUNG KHÁCH CHECK-IN (2 khách vào 1 căn trống → chèn 1 dòng) trên BẢN SAO, kiểm tra bằng Google Sheets. */
function thuBoSungBanSao() {
  const cp = JSON.parse(drive_('/drive/v3/files/' + QLKH_ID + '/copy?supportsAllDrives=true&fields=id', { method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ name: 'THU_BO_SUNG_XOA_DUOC.xlsm' }) }).getContentText());
  KH_TARGET = cp.id; let gs = null;
  try {
    const L0 = khList_().rows, by = {};
    L0.forEach(r => { if (r.ma) (by[r.ma] = by[r.ma] || []).push(r); });
    const ma = Object.keys(by).find(k => by[k].length === 1 && !by[k][0].ten && by[k][0].row > 300 && by[k][0].row < 560);
    const r0 = by[ma][0].row, below = L0.find(r => r.row === r0 + 1) || {};
    Logger.log('Căn trống thử: ' + ma + ' dòng ' + r0 + ' · dòng dưới: ' + below.ma + ' ' + below.ten);
    const S = xlsmSession_(), blk = S.findBlock(ma);
    const rows = [r0].concat(S.insRows(blk[blk.length - 1].row, 1, blk[blk.length - 1].row));
    const G = [{ ten: 'THU KHACH MOT', ns: '01/01/1990', qt: 'NN', gt: 'Nam', hc: 'X1234567', bd: '27/11/2026', kt: '26/02/2027', lt: '20/02/2027', coc: '7500000', tp: '7500000', sn: '2', sdt: '0900000000' },
               { ten: 'THU KHACH HAI', ns: '02/02/1992', qt: 'NN', gt: 'Nữ', hc: 'Y7654321', lt: '20/02/2027' }];
    rows.forEach((r, i) => { const sp = editSpec_(G[i]); sp.setV.D = ma; sp.kinds.D = 'text'; sp.setV.C = blk[0].toa; sp.kinds.C = 'text';
      Logger.log('Ghi dòng ' + r + ' -> ' + JSON.stringify(S.edit(r, {}, sp.setV, sp.kinds, { colFill: { E: 'none' }, skipFormula: true }))); });
    Logger.log('Ghi file -> ' + JSON.stringify(S.commit()));
    const L1 = khList_().rows;
    Logger.log('Số dòng: ' + L0.length + ' -> ' + L1.length + ' · căn ' + ma + ': ' + JSON.stringify(L1.filter(r => r.ma === ma).map(r => [r.row, r.ten, r.hc, r.bd, r.lt, r.tp])));
    const b2 = L1.find(r => r.ma === below.ma && r.ten === below.ten); Logger.log('Dòng dưới giờ ở: ' + (b2 ? b2.row + ' (trước ' + below.row + ')' : 'KHÔNG THẤY'));
    gs = JSON.parse(drive_('/drive/v3/files/' + cp.id + '/copy?supportsAllDrives=true&fields=id', { method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ name: 'THU_BO_SUNG_GS', mimeType: 'application/vnd.google-apps.spreadsheet' }) }).getContentText()).id;
    const sh = SpreadsheetApp.openById(gs).getSheetByName(QLKH_SHEET), v = sh.getDataRange().getDisplayValues(), f = sh.getDataRange().getFormulas();
    let errN = 0; v.forEach((row, i) => row.forEach((c, j) => { if (/^#(REF|VALUE|N.A|NAME|DIV)/.test(c) && f[i][j]) errN++; }));
    Logger.log('Google Sheets mở được · ô công thức lỗi: ' + errN + ' (file gốc: 33) · AA dòng mới: ' + (f[rows[1] - 1][26] || '(trống)') + ' · màu E dòng chính: ' + sh.getRange(rows[0], 5).getBackground());
  } finally {
    KH_TARGET = QLKH_ID;
    [cp.id, gs].filter(Boolean).forEach(id => drive_('/drive/v3/files/' + id + '?supportsAllDrives=true', { method: 'patch', contentType: 'application/json', payload: JSON.stringify({ trashed: true }) }));
    Logger.log('Đã chuyển bản sao vào thùng rác.');
  }
}

/* Chạy tay: thử XOÁ KHÁCH CHÍNH (đưa khách phụ lên) + CHECK-OUT CẢ CĂN trên BẢN SAO, kiểm tra bằng Google Sheets. File thật không bị đụng. */
function thuCheckoutBanSao() {
  const lock = LockService.getScriptLock(); lock.waitLock(60000);
  const P = PropertiesService.getScriptProperties();
  if (khPending_().length) { lock.releaseLock(); Logger.log('Đang có thay đổi chờ ghi — chạy lại sau'); return; }
  const cp = JSON.parse(drive_('/drive/v3/files/' + QLKH_ID + '/copy?supportsAllDrives=true&fields=id', { method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ name: 'THU_CHECKOUT_XOA_DUOC.xlsm' }) }).getContentText());
  KH_TARGET = cp.id; KH_TEST = true; let gs = null; const keys = [];
  try {
    const L0 = khList_().rows, by = {};
    L0.forEach(r => { if (r.ma) (by[r.ma] = by[r.ma] || []).push(r); });
    const ok = k => by[k].length >= 2 && by[k].every(r => r.ten) && by[k].every((r, i) => !i || r.row === by[k][i - 1].row + 1);
    const ks = Object.keys(by).filter(ok), A = ks[0], B = ks[ks.length - 1];
    Logger.log('Căn A (xoá khách chính): ' + A + ' ' + JSON.stringify(by[A].map(r => [r.row, r.ten, r.lt, r.bd, r.tp])));
    Logger.log('Căn B (check-out cả căn): ' + B + ' ' + JSON.stringify(by[B].map(r => [r.row, r.ten])));
    const m = by[A][0];
    keys.push(khQueue_({ type: 'out', row: m.row, guest: m, note: 'THU' }).id);
    keys.push(khQueue_({ type: 'unitout', ma: B, guests: by[B], note: 'THU', openUnit: false }).id);
    const r = khApply_();
    Logger.log('Kết quả: ' + JSON.stringify({ applied: r.applied, failed: r.failed, deleted: r.deleted, errors: (r.errors || []).slice(0, 2) }));
    const L1 = khList_().rows;
    Logger.log('Căn A sau: ' + JSON.stringify(L1.filter(x => x.ma === A).map(x => [x.row, x.ten, x.lt, x.bd, x.tp])));
    Logger.log('Căn B sau: ' + JSON.stringify(L1.filter(x => x.ma === B).map(x => [x.row, x.ten, x.bd])));
    Logger.log('Số dòng: ' + L0.length + ' -> ' + L1.length);
    gs = JSON.parse(drive_('/drive/v3/files/' + cp.id + '/copy?supportsAllDrives=true&fields=id', { method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ name: 'THU_CHECKOUT_GS', mimeType: 'application/vnd.google-apps.spreadsheet' }) }).getContentText()).id;
    const sh = SpreadsheetApp.openById(gs).getSheetByName(QLKH_SHEET), v = sh.getDataRange().getDisplayValues(), f = sh.getDataRange().getFormulas();
    let errN = 0; v.forEach((row, i) => row.forEach((c, j) => { if (/^#(REF|VALUE|N.A|NAME|DIV)/.test(c) && f[i][j]) errN++; }));
    const bRow = (L1.find(x => x.ma === B) || {}).row;
    Logger.log('Google Sheets mở được · ô công thức lỗi: ' + errN + ' (file gốc: 33) · màu E dòng căn B: ' + (bRow ? sh.getRange('E' + bRow).getBackground() : '?'));
  } finally {
    keys.forEach(k => k && P.deleteProperty(k));
    KH_TARGET = QLKH_ID; KH_TEST = false; lock.releaseLock();
    [cp.id, gs].filter(Boolean).forEach(id => drive_('/drive/v3/files/' + id + '?supportsAllDrives=true', { method: 'patch', contentType: 'application/json', payload: JSON.stringify({ trashed: true }) }));
    Logger.log('Đã chuyển bản sao vào thùng rác.');
  }
}

/* Chạy tay trong trình soạn thảo: thử đọc danh sách (không ghi gì). */
function thuDocKhach() {
  const r = khList_();
  Logger.log(r.ok ? ('Đọc được ' + r.rows.length + ' dòng. Dòng đầu: ' + JSON.stringify(r.rows[0])) : JSON.stringify(r));
}

/* Chạy tay: thử GHI trên BẢN SAO của file (file thật không bị đụng), xong tự xoá bản sao. */
function thuGhiBanSao() {
  const cp = JSON.parse(drive_('/drive/v3/files/' + QLKH_ID + '/copy?supportsAllDrives=true&fields=id', { method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ name: 'THU_GHI_APP_XOA_DUOC.xlsm' }) }).getContentText());
  KH_TARGET = cp.id;
  try {
    const L = khList_(), g = L.rows.find(r => r.ten && r.row > 8), e = L.rows.find(r => !r.ten && r.row > 8);
    Logger.log('Thử sửa dòng ' + g.row + ' (' + g.ma + ' · ' + g.ten + ')');
    const sp = editSpec_({ gc: 'THU GHI APP', dmoi: '12345', lt: '31/12/2027', tp: '7.000.000' });
    Logger.log('Sai tên -> ' + JSON.stringify(xlsmEditRow_(g.row, { ma: g.ma, ten: 'SAI' }, sp.setV, sp.kinds, {})));
    Logger.log('Sửa -> ' + JSON.stringify(xlsmEditRow_(g.row, { ma: g.ma, ten: g.ten }, sp.setV, sp.kinds, {})));
    const g2 = khList_().rows.find(r => r.row === g.row);
    Logger.log('Đọc lại bản sao: ghi chú=' + g2.gc + ' điện mới=' + g2.dmoi + ' hạn lưu trú=' + g2.lt + ' tiền phòng=' + g2.tp);
    const os = outSpec_(), r2 = xlsmEditRow_(g.row, { ma: g.ma, ten: g.ten }, os.setV, os.kinds, { colFill: { E: 'red' }, skipFormula: true });
    const g3 = khList_().rows.find(r => r.row === g.row);
    Logger.log('OUT -> ' + JSON.stringify(r2) + ' · sau OUT: tên="' + (g3 ? g3.ten : '(dòng trống)') + '" mã phòng=' + (g3 ? g3.ma : '') + ' điện mới=' + (g3 ? g3.dmoi : ''));
    const vba = Utilities.unzip(Object.assign(DriveApp.getFileById(KH_TARGET).getBlob(), {}).setContentType('application/zip')).some(b => b.getName() === 'xl/vbaProject.bin');
    Logger.log('Macro VBA còn nguyên: ' + vba);
  } finally {
    KH_TARGET = QLKH_ID;
    drive_('/drive/v3/files/' + cp.id + '?supportsAllDrives=true', { method: 'patch', contentType: 'application/json', payload: JSON.stringify({ trashed: true }) });
    Logger.log('Đã chuyển bản sao vào thùng rác.');
  }
}

/**
 * SỬA HÀNG LOẠT (chạy tay 1 lần): các dòng khách bị dán lệch số liệu DS Check-in vào cột Y..AC.
 *  - chép giá trị cũ sang bảng "LUXHOME - SỬA CỘT TIỀN THUÊ" (không mất thông tin hoa hồng…)
 *  - xoá trắng Y, Z, AB, AC (+ AD nếu đang chứa chữ/tiền dán lệch); đặt lại công thức AA = (Z-Y+1)/30
 *  - luôn lưu bản sao file trước khi ghi
 *  xem trước: suaCotTienThue(true)
 */
function suaCotTienThue(chiXem) {
  const lock = LockService.getScriptLock(); lock.waitLock(60000);
  try {
    if (khPending_().length) khApply_();
    const S = xlsmSession_(), list = S.scanShifted();
    Logger.log('Tìm thấy ' + list.length + ' dòng lệch cột. Ví dụ: ' + JSON.stringify(list.slice(0, 3)));
    if (chiXem === true || !list.length) return list.length;
    const ss = SpreadsheetApp.create('LUXHOME - SỬA CỘT TIỀN THUÊ (Y–AC) ' + Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm'));
    const sh = ss.getSheets()[0];
    const rows = [['Dòng', 'Mã phòng', 'Tên khách', 'Y (cũ)', 'Z (cũ)', 'AA (cũ)', 'AB (cũ)', 'AC (cũ)', 'AD (cũ)', 'Đã xoá AD?']]
      .concat(list.map(x => [x.row, x.ma, x.ten, x.old.Y, x.old.Z, x.old.AA, x.old.AB, x.old.AC, x.old.AD, x.old.adText ? 'Có' : ''].map(v => "'" + String(v))));
    sh.getRange(1, 1, rows.length, 10).setValues(rows); sh.setFrozenRows(1);
    let ok = 0;
    list.forEach(x => {
      const setV = { Y: '', Z: '', AB: '', AC: '', AA: '=(Z' + x.row + '-Y' + x.row + '+1)/30' }, kinds = { Y: 'date', Z: 'date', AB: 'date', AC: 'date', AA: 'formula' };
      if (x.old.adText) { setV.AD = ''; kinds.AD = 'num'; }
      const r = S.edit(x.row, { ma: x.ma, ten: x.ten }, setV, kinds, { skipFormula: true });
      if (r.ok) ok++;
    });
    PropertiesService.getScriptProperties().setProperty('kh_lastbk', '0');   // bắt buộc sao lưu lần này
    const c = S.commit();
    Logger.log('Đã sửa ' + ok + '/' + list.length + ' dòng. Bản sao file: ' + c.backupId + ' · Bảng giá trị cũ: ' + ss.getUrl());
    return ok;
  } finally { lock.releaseLock(); }
}
