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
 *    rồi xoá trắng thông tin khách của dòng (E..AC), tô vàng — giống quy tắc nút Duyệt check-in.
 */

const QLKH_ID     = '1olOswDQsPC8CAsaUiitbAePTeZm83ugV';
const QLKH_SHEET  = 'DANH SACH LUU TRU';
const QLKH_BACKUP = 'LUXHOME BACKUP QLKH';
const QLKH_MIME   = 'application/vnd.ms-excel.sheet.macroEnabled.12';
const OUT_LOG_NAME = 'LUXHOME - KHÁCH ĐÃ TRẢ PHÒNG';
let KH_TARGET = QLKH_ID;   // file đang ghi (hàm thử đổi sang bản sao)

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
function khSheet_() {
  const sh = SpreadsheetApp.openById(QLKH_ID).getSheetByName(QLKH_SHEET);
  if (!sh) throw new Error('Không mở được sheet ' + QLKH_SHEET);
  return sh;
}
function khList_() {
  const rng = khSheet_().getDataRange();
  const v = rng.getDisplayValues(), raw = rng.getValues();
  // ô ngày thật -> luôn trả dd/MM/yyyy (tránh nhầm định dạng ngày kiểu Mỹ khi hiển thị)
  for (let i = 0; i < raw.length; i++) for (let j = 0; j < raw[i].length; j++)
    if (raw[i][j] instanceof Date) v[i][j] = Utilities.formatDate(raw[i][j], 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy');
  let hr = -1;
  for (let i = 0; i < Math.min(v.length, 20); i++) if (v[i].some(isHdrTen_)) { hr = i; break; }
  if (hr < 0) throw new Error('Không thấy dòng tiêu đề "Tên khách"');
  const H = v[hr].map(norm_);
  for (const L in K_CHECK) if (H[colNum_(L) - 1] !== K_CHECK[L]) return { ok: false, err: 'layout', msg: 'Cột ' + L + ' của file không đúng mẫu — không đọc để tránh sai.' };
  const rows = [];
  for (let r = hr + 1; r < v.length; r++) {
    const row = v[r], o = { row: r + 1 };
    for (const k in KF) o[k] = String(row[colNum_(KF[k]) - 1] || '').trim();
    if (!o.ma && !o.ten && !o.toa) continue;
    // cột công thức (giá trị đã tính) — chỉ để xem
    o.sodien = String(row[colNum_('AI') - 1] || ''); o.tiendien = String(row[colNum_('AJ') - 1] || '');
    o.tiennuoc = String(row[colNum_('AK') - 1] || ''); o.tiendv = String(row[colNum_('AL') - 1] || '');
    o.tienthue = String(row[colNum_('AM') - 1] || ''); o.tong = String(row[colNum_('AN') - 1] || '');
    rows.push(o);
  }
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
function maybeBackup_(name) {
  const P = PropertiesService.getScriptProperties(), last = +(P.getProperty('kh_lastbk') || 0);
  if (Date.now() - last < 30 * 60 * 1000) return null;
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
function xlsmEditRow_(rowNum, expect, setV, kinds, opt) {
  opt = opt || {}; const fill = opt.fill || null;
  const meta = JSON.parse(drive_('/drive/v3/files/' + KH_TARGET + '?fields=name,modifiedTime&supportsAllDrives=true').getContentText());
  const res = drive_('/drive/v3/files/' + KH_TARGET + '?alt=media&supportsAllDrives=true');
  if (res.getResponseCode() !== 200) throw new Error('Không tải được file trên Drive (HTTP ' + res.getResponseCode() + ')');
  const blob = res.getBlob(); blob.setContentType('application/zip');
  const parts = Utilities.unzip(blob), byName = {}, order = [];
  parts.forEach(b => { byName[b.getName()] = b; order.push(b.getName()); });
  const get = n => byName[n] ? byName[n].getDataAsString('UTF-8') : null;
  const changed = {};

  // --- tìm sheet
  const wb = get('xl/workbook.xml'), rels = get('xl/_rels/workbook.xml.rels');
  let rid = null, tgt = null;
  wb.replace(/<sheet\b[^>]*>/g, t => { const nm = (t.match(/ name="([^"]*)"/) || [])[1]; if (nm != null && xUnesc_(nm) === QLKH_SHEET) rid = (t.match(/ r:id="([^"]*)"/) || [])[1]; return t; });
  if (!rid) throw new Error('Không thấy sheet ' + QLKH_SHEET);
  rels.replace(/<Relationship\b[^>]*>/g, t => { if ((t.match(/ Id="([^"]*)"/) || [])[1] === rid) tgt = (t.match(/ Target="([^"]*)"/) || [])[1]; return t; });
  const sheetPath = tgt.charAt(0) === '/' ? tgt.slice(1) : 'xl/' + tgt.replace(/^\.\//, '');
  let s = get(sheetPath);

  // --- shared strings (để đọc giá trị kiểm tra)
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

  // --- kiểm tra cấu trúc (dòng tiêu đề)
  let hdr = null;
  for (let r = 1; r <= 20 && !hdr; r++) {
    const m = s.match(rowRe(r)); if (!m) continue;
    const cs = parseRow(m[0]);
    if (Object.keys(cs).some(k => isHdrTen_(cellVal(cs[k])))) hdr = cs;
  }
  if (!hdr) throw new Error('Không thấy dòng tiêu đề trong file');
  for (const L in K_CHECK) if (norm_(cellVal(hdr[colNum_(L)])) !== K_CHECK[L]) throw new Error('Cột ' + L + ' của file không đúng mẫu — đã dừng, không ghi.');

  // --- dòng cần sửa + kiểm tra khớp
  const rm = s.match(rowRe(rowNum));
  if (!rm) throw new Error('Không thấy dòng ' + rowNum + ' trong file');
  const rowXml = rm[0], cells = parseRow(rowXml);
  const curMa = cmp_(cellVal(cells[colNum_('D')])), curTen = cmp_(cellVal(cells[colNum_('E')]));
  if ((expect.ma != null && cmp_(expect.ma) !== curMa) || (expect.ten != null && cmp_(expect.ten) !== curTen)) {
    return { ok: false, err: 'changed', msg: 'Dòng ' + rowNum + ' trong file đã thay đổi (hiện là: ' + (curMa || '—') + ' · ' + (curTen || 'trống') + '). Tải lại danh sách rồi sửa lại.' };
  }

  // --- styles: thêm định dạng ngày dd/mm/yyyy + màu vàng khi cần
  let st = get('xl/styles.xml');
  const xfsM = st.match(/<cellXfs count="(\d+)">([\s\S]*?)<\/cellXfs>/);
  const xfs = xfsM[2].match(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g), added = [], xcache = {};
  let YELLOW = null;
  if (fill === 'yellow') {
    const fm = st.match(/<fills count="(\d+)">([\s\S]*?)<\/fills>/);
    YELLOW = +fm[1];
    st = st.replace(fm[0], '<fills count="' + (YELLOW + 1) + '">' + fm[2] + '<fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/><bgColor indexed="64"/></patternFill></fill></fills>');
  }
  let DATEFMT = 0, maxId = 163; const fmtCode = {};
  const nfM = st.match(/<numFmts count="(\d+)">([\s\S]*?)<\/numFmts>/);
  if (nfM) nfM[2].replace(/<numFmt numFmtId="(\d+)" formatCode="([^"]*)"\/>/g, (m, id, code) => { fmtCode[id] = xUnesc_(code); if (+id > maxId) maxId = +id; if (xUnesc_(code) === 'dd/mm/yyyy') DATEFMT = +id; return m; });
  if (!DATEFMT) {
    DATEFMT = maxId + 1;
    const nf = '<numFmt numFmtId="' + DATEFMT + '" formatCode="dd/mm/yyyy"/>';
    st = nfM ? st.replace(nfM[0], '<numFmts count="' + (+nfM[1] + 1) + '">' + nfM[2] + nf + '</numFmts>') : st.replace('<fonts', '<numFmts count="1">' + nf + '</numFmts><fonts');
  }
  const isDateFmt = id => { id = +id; if ((id >= 14 && id <= 22) || (id >= 45 && id <= 47)) return true; const c = fmtCode[id]; return !!c && /[dy]/i.test(c.replace(/"[^"]*"|\[[^\]]*\]/g, '')) && !/[#0]/.test(c); };
  const xfFor = (base, kind) => {
    const key = base + '|' + kind + '|' + (YELLOW == null ? '' : YELLOW);
    if (xcache[key] != null) return xcache[key];
    let x = xfs[+base] || xfs[0];
    const nid = +(x.match(/ numFmtId="(\d+)"/) || [0, 0])[1];
    let setFmt = null;
    if (kind === 'date' && !isDateFmt(nid)) setFmt = DATEFMT;
    if (kind === 'text' && isDateFmt(nid)) setFmt = 49;
    if (setFmt == null && YELLOW == null) { xcache[key] = +base; return +base; }
    if (setFmt != null) {
      x = / numFmtId="/.test(x) ? x.replace(/ numFmtId="\d+"/, ' numFmtId="' + setFmt + '"') : x.replace('<xf ', '<xf numFmtId="' + setFmt + '" ');
      x = / applyNumberFormat="/.test(x) ? x.replace(/ applyNumberFormat="\d"/, ' applyNumberFormat="1"') : x.replace('<xf ', '<xf applyNumberFormat="1" ');
    }
    if (YELLOW != null) {
      x = / fillId="/.test(x) ? x.replace(/ fillId="\d+"/, ' fillId="' + YELLOW + '"') : x.replace('<xf ', '<xf fillId="' + YELLOW + '" ');
      x = / applyFill="/.test(x) ? x.replace(/ applyFill="\d"/, ' applyFill="1"') : x.replace('<xf ', '<xf applyFill="1" ');
    }
    xcache[key] = xfs.length + added.length; added.push(x);
    return xcache[key];
  };
  const serial = d => { const m = String(d).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/); return m ? Math.round((Date.UTC(+m[3], +m[2] - 1, +m[1]) - Date.UTC(1899, 11, 30)) / 86400000) : null; };
  const newCell = (col, v, base, kind) => {
    const ref = colName_(col) + rowNum, sa = ' s="' + xfFor(base, kind) + '"';
    v = String(v == null ? '' : v).trim();
    if (v === '') return '<c r="' + ref + '"' + sa + '/>';
    if (kind === 'date') { const se = serial(v); if (se != null) return '<c r="' + ref + '"' + sa + '><v>' + se + '</v></c>'; }
    if (kind === 'num') { let n = v.replace(/[,\s]/g, ''); if (/^\d{1,3}(\.\d{3})+$/.test(n)) n = n.replace(/\./g, ''); if (/^-?\d+(\.\d+)?$/.test(n)) return '<c r="' + ref + '"' + sa + '><v>' + n + '</v></c>'; }
    const t = xEsc_(v);
    return '<c r="' + ref + '"' + sa + ' t="inlineStr"><is><t' + (t !== t.trim() ? ' xml:space="preserve"' : '') + '>' + t + '</t></is></c>';
  };

  // --- dựng lại dòng
  const byCol = {};
  Object.keys(cells).forEach(c => { byCol[+c] = cells[c].xml; });
  Object.keys(setV).forEach(L => {
    const c = colNum_(L);
    if (c >= colNum_('AI') && c <= colNum_('AN')) throw new Error('Không được ghi vào cột công thức ' + L);
    const old = cells[c];
    if (old && /<f[ >]/.test(old.xml)) { if (opt.skipFormula) return; throw new Error('Ô ' + L + rowNum + ' là công thức — không ghi đè'); }
    byCol[c] = newCell(c, setV[L], old ? old.s : '0', kinds[L] || 'text');
  });
  if (fill === 'yellow') {
    for (let c = 1; c <= colNum_('AH'); c++) {
      if (setV[colName_(c)] != null) continue;
      const old = cells[c];
      if (!old) { byCol[c] = '<c r="' + colName_(c) + rowNum + '" s="' + xfFor('0', 'keep') + '"/>'; continue; }
      byCol[c] = / s="\d+"/.test(old.xml) ? old.xml.replace(/ s="\d+"/, ' s="' + xfFor(old.s, 'keep') + '"') : old.xml.replace('<c ', '<c s="' + xfFor('0', 'keep') + '" ');
    }
  }
  const tag = rowXml.match(/^<row\b[^>]*?(?=\/?>)/)[0].replace(/ spans="[^"]*"/, '');
  const newRow = tag + '>' + Object.keys(byCol).map(Number).sort((a, b) => a - b).map(c => byCol[c]).join('') + '</row>';
  s = s.replace(rowXml, () => newRow);
  changed[sheetPath] = s;
  if (added.length || fill === 'yellow' || !nfM || DATEFMT > maxId) {
    st = st.replace(xfsM[0], '<cellXfs count="' + (xfs.length + added.length) + '">' + xfs.join('') + added.join('') + '</cellXfs>');
    changed['xl/styles.xml'] = st;
  }
  // tính lại công thức khi mở file
  changed['xl/workbook.xml'] = /<calcPr\b[^>]*\/>/.test(wb)
    ? wb.replace(/<calcPr\b([^>]*)\/>/, (m, a) => '<calcPr' + a.replace(/ fullCalcOnLoad="\w+"/, '') + ' fullCalcOnLoad="1"/>')
    : wb.replace('</workbook>', '<calcPr fullCalcOnLoad="1"/></workbook>');

  // --- sao lưu rồi ghi
  const backupId = KH_TARGET === QLKH_ID ? maybeBackup_(meta.name) : null;
  const blobs = [], first = '[Content_Types].xml';
  [first].concat(order.filter(n => n !== first)).forEach(n => {
    if (!byName[n]) return;
    blobs.push(changed[n] != null ? Utilities.newBlob(Utilities.newBlob(changed[n]).getBytes(), 'application/xml', n) : byName[n]);
  });
  const zip = Utilities.zip(blobs, meta.name);
  const up = drive_('/upload/drive/v3/files/' + KH_TARGET + '?uploadType=media&supportsAllDrives=true', { method: 'patch', contentType: QLKH_MIME, payload: zip.getBytes() });
  if (up.getResponseCode() !== 200) throw new Error('Ghi file lên Drive không được (HTTP ' + up.getResponseCode() + ')');
  return { ok: true, backupId: backupId };
}

/* ======================= LỆNH CHO APP ======================= */
function khSave_(p) {
  const row = +p.row, f = p.fields || {};
  if (!(row > 8)) return { ok: false, err: 'badrow' };
  const setV = {}, kinds = {};
  Object.keys(f).forEach(k => {
    if (K_EDIT.indexOf(k) < 0) return;
    const L = KF[k]; setV[L] = f[k];
    kinds[L] = K_DATE.indexOf(k) >= 0 ? 'date' : K_NUM.indexOf(k) >= 0 ? 'num' : 'text';
  });
  if (!Object.keys(setV).length) return { ok: false, err: 'nothing' };
  return xlsmEditRow_(row, { ma: p.expect && p.expect.ma, ten: p.expect && p.expect.ten }, setV, kinds, {});
}
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
function khOut_(p) {
  const g = p.guest || {}, row = +p.row;
  if (!(row > 8)) return { ok: false, err: 'badrow' };
  // xoá trắng thông tin khách (E..AC) — giữ STT, phòng, tòa, mã phòng, chỉ số điện & đơn giá (AD..AH) cho khách sau
  const setV = {}, kinds = {};
  for (let c = colNum_('E'); c <= colNum_('AC'); c++) { setV[colName_(c)] = ''; kinds[colName_(c)] = 'text'; }
  const r = xlsmEditRow_(row, { ma: g.ma, ten: g.ten }, setV, kinds, { fill: 'yellow', skipFormula: true });
  if (!r.ok) return r;
  try {
    outLog_().appendRow([Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm'), 'Admin', g.ma, g.ten, g.ns, g.qt, g.gt, g.cccd, g.hc, g.dc,
      g.bd, g.kt, g.lt, g.coc, g.tp, g.sn, g.sale, g.lh, g.sdt, g.gc, p.note || ''].map(x => x == null ? '' : "'" + String(x)));
  } catch (e) { r.logErr = String(e.message || e); }
  // mở căn trên catalog nếu được yêu cầu (phòng trống)
  if (p.openUnit) {
    try {
      const t = unitTable_(), ma = String(g.ma || '').replace(/-/g, '').toUpperCase(), i = findUnit_(t, ma);
      if (i >= 0 && t.C.tt >= 0) { t.sh.getRange(i + 1, t.C.tt + 1).setValue('Đang có sẵn'); r.unitOpened = ma; }
      else r.unitOpened = '';
    } catch (e) { r.unitErr = String(e.message || e); }
  }
  return r;
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
    Logger.log('Sai tên -> ' + JSON.stringify(khSave_({ row: g.row, expect: { ma: g.ma, ten: 'SAI' }, fields: { gc: 'x' } })));
    Logger.log('Sửa -> ' + JSON.stringify(khSave_({ row: g.row, expect: { ma: g.ma, ten: g.ten }, fields: { gc: 'THU GHI APP', dmoi: '12345', lt: '31/12/2027', tp: '7.000.000' } })));
    const sh = SpreadsheetApp.openById(KH_TARGET).getSheetByName(QLKH_SHEET);
    Logger.log('Đọc lại bản sao: AG=' + sh.getRange('AG' + g.row).getDisplayValue() + ' AE=' + sh.getRange('AE' + g.row).getDisplayValue() +
      ' N=' + sh.getRange('N' + g.row).getDisplayValue() + ' P=' + sh.getRange('P' + g.row).getValue() + ' AN=' + sh.getRange('AN' + g.row).getFormula());
    const r2 = xlsmEditRow_(g.row, { ma: g.ma, ten: g.ten }, (() => { const o = {}; for (let c = 5; c <= 29; c++) o[colName_(c)] = ''; return o; })(), {}, { fill: 'yellow', skipFormula: true });
    Logger.log('OUT -> ' + JSON.stringify(r2) + ' · tên sau OUT="' + SpreadsheetApp.openById(KH_TARGET).getSheetByName(QLKH_SHEET).getRange('E' + g.row).getDisplayValue() + '"');
    const vba = Utilities.unzip(Object.assign(DriveApp.getFileById(KH_TARGET).getBlob(), {}).setContentType('application/zip')).some(b => b.getName() === 'xl/vbaProject.bin');
    Logger.log('Macro VBA còn nguyên: ' + vba);
  } finally {
    KH_TARGET = QLKH_ID;
    drive_('/drive/v3/files/' + cp.id + '?supportsAllDrives=true', { method: 'patch', contentType: 'application/json', payload: JSON.stringify({ trashed: true }) });
    Logger.log('Đã chuyển bản sao vào thùng rác.');
  }
}
