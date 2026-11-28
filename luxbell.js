/* LUXHOME · CHUÔNG THÔNG BÁO việc chờ duyệt (góc trên bên phải) — chỉ hiện trên máy của CHỦ
   Nguồn: đề xuất của nhân viên (máy chủ Quản lý căn hộ), đợt check-in chờ duyệt (máy chủ Check-in),
   máy nhân viên chờ duyệt (máy chủ Lưu trú). */
(function () {
  var CK = "https://script.google.com/macros/s/AKfycbylsYx2qOuQgp92bb5o_Q0Pg7K3I6FYfjvhlyBxUvBYgcE5yOOBth3PyniM3wgGtT4y/exec";
  var NV = "https://script.google.com/macros/s/AKfycbyFhO_xjflBHiofIY_XuQCcUNKd1VCwTUKhW5Jvaya51WBD5IId5owYfZM_Q9DY0aZoWQ/exec";
  function ls(k) { try { return localStorage.getItem(k) || ''; } catch (e) { return ''; } }
  var PW = ls('luxAdminPw'), CKPW = ls('ck_pw'), MG = ls('mg_pin');
  var TOK = new URLSearchParams(location.search).get('tok') || ((!PW && !CKPW && !MG) ? ls('nv_tok') : '');
  var NV_MODE = !!TOK;                                                   // máy / trang nhân viên → chuông "lệnh của tôi đã được xử lý"
  if (!NV_MODE && !PW && !CKPW && !MG) return;
  var esc = function (s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); };
  function jsonp(url) { return fetch(url + (url.indexOf('?') < 0 ? '?' : '&') + 'callback=cb&t=' + Date.now(), { cache: 'no-store' }).then(function (r) { return r.text(); }).then(function (t) { return JSON.parse(t.slice(t.indexOf('{'), t.lastIndexOf('}') + 1)); }); }

  var css = document.createElement('style');
  css.textContent =
    '#lxBell{position:fixed;top:calc(env(safe-area-inset-top,0px) + 12px);right:14px;z-index:9000;width:46px;height:46px;border-radius:50%;border:0;cursor:pointer;' +
    'background:rgba(255,255,255,.95);box-shadow:0 4px 16px rgba(0,40,60,.25);font-size:1.35rem;display:flex;align-items:center;justify-content:center}' +
    '#lxBell.rung{animation:lxRung 1s ease 2}@keyframes lxRung{0%,100%{transform:rotate(0)}20%{transform:rotate(14deg)}40%{transform:rotate(-12deg)}60%{transform:rotate(8deg)}80%{transform:rotate(-5deg)}}' +
    '#lxBell .c{position:absolute;top:-3px;right:-3px;min-width:21px;height:21px;padding:0 6px;border-radius:999px;background:#dc2626;color:#fff;font:800 .74rem/21px system-ui,sans-serif;box-shadow:0 0 0 2px #fff}' +
    '#lxPanel{position:fixed;top:calc(env(safe-area-inset-top,0px) + 66px);right:12px;z-index:9001;width:min(380px,calc(100vw - 24px));max-height:75vh;overflow:auto;' +
    'background:#fff;color:#13242a;border-radius:18px;box-shadow:0 18px 50px rgba(0,30,45,.30);display:none;font-family:-apple-system,"Segoe UI",Roboto,system-ui,sans-serif}' +
    '#lxPanel.on{display:block}#lxPanel h4{margin:0;padding:14px 16px;font-size:1rem;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #e8eef0;position:sticky;top:0;background:#fff;border-radius:18px 18px 0 0}' +
    '#lxPanel h4 button{border:0;background:#eef5f7;border-radius:999px;padding:5px 10px;font-weight:700;cursor:pointer;color:#0e7ba8}' +
    '.lxS{display:block;text-decoration:none;color:inherit;padding:12px 16px;border-bottom:1px solid #f0f4f5}.lxS:hover{background:#f6fbfc}' +
    '.lxS .t{display:flex;align-items:center;gap:10px;font-weight:800;font-size:.92rem}.lxS .t i{font-style:normal;width:34px;height:34px;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:1.1rem;flex:none}' +
    '.lxS .t b{margin-left:auto;min-width:24px;height:24px;border-radius:999px;color:#fff;font-size:.78rem;display:flex;align-items:center;justify-content:center;padding:0 7px}' +
    '.lxS ul{margin:6px 0 0 44px;padding:0;list-style:none;font-size:.8rem;color:#5f727a}.lxS li{padding:2px 0}.lxS li b{color:#13242a}' +
    '.lxS .go{margin:6px 0 0 44px;font-size:.78rem;font-weight:800;color:#0e7ba8}.lxE{padding:26px 16px;text-align:center;color:#5f727a;font-size:.9rem}' +
    '.lxF{padding:10px 16px;font-size:.72rem;color:#90a4ac;text-align:center}' +
    '@media (prefers-color-scheme:dark){#lxPanel,#lxPanel h4{background:#13222a;color:#e8f1f3}#lxPanel h4,.lxS{border-color:#20333c}.lxS:hover{background:#18303a}.lxS li b{color:#e8f1f3}#lxBell{background:#1b2f38}}';
  document.head.appendChild(css);
  var bell = document.createElement('button'); bell.id = 'lxBell'; bell.title = 'Việc chờ duyệt'; bell.innerHTML = '🔔<span class="c" style="display:none"></span>';
  var panel = document.createElement('div'); panel.id = 'lxPanel';
  document.body.appendChild(bell); document.body.appendChild(panel);
  bell.onclick = function (e) { e.stopPropagation(); panel.classList.toggle('on'); if (panel.classList.contains('on')) { ve(); tai(); } };
  document.addEventListener('click', function (e) { if (!panel.contains(e.target) && e.target !== bell) panel.classList.remove('on'); });

  var D = { de: null, hd: null, dots: null, dev: null, at: 0 };
  function tf(t) { var d = new Date(t); return isNaN(d) ? '' : ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() + 1)).slice(-2) + ' ' + ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2); }
  function sec(ic, bg, title, n, items, href, go) {
    if (!n) return '';
    return '<a class="lxS" href="' + href + '"><div class="t"><i style="background:' + bg + '22">' + ic + '</i>' + title + '<b style="background:' + bg + '">' + n + '</b></div>'
      + (items.length ? '<ul>' + items.join('') + '</ul>' : '') + '<div class="go">' + go + ' ›</div></a>';
  }
  function tong() {
    return (D.de ? D.de.length : 0) + (D.dots ? D.dots.length : 0) + (D.dev ? D.dev.length : 0);
  }
  function ve() {
    var n = tong(), c = bell.querySelector('.c');
    c.textContent = n > 99 ? '99+' : n; c.style.display = n ? 'block' : 'none';
    var h = '<h4>🔔 Yêu cầu chờ duyệt<button onclick="event.stopPropagation();window.lxBellTai(1)">↻ Làm mới</button></h4>';
    if (D.de && D.de.length) {
      var lo = {}, le = [];
      D.de.forEach(function (o) { if (o.lo) { (lo[o.lo] = lo[o.lo] || { who: o.who, src: o.src, t: o.t, n: 0 }).n++; } else le.push(o); });
      var it = Object.keys(lo).map(function (k) { var g = lo[k]; return '<li>📁 <b>Lô DS Check-in</b> · ' + g.n + ' căn · ' + esc(g.who) + '</li>'; })
        .concat(le.slice(0, 6).map(function (o) { return '<li><b>' + esc(o.lb) + '</b> · ' + esc(o.who) + ' · ' + tf(o.t) + '<br>' + esc(o.dt) + '</li>'; }));
      if (le.length > 6) it.push('<li>… và ' + (le.length - 6) + ' đề xuất khác</li>');
      h += sec('📋', '#dc2626', 'Đề xuất của nhân viên', D.de.length, it, 'khach.html?view=props', 'Mở để duyệt');
    }
    if (D.dots && D.dots.length) h += sec('🛎️', '#ea580c', 'Đợt check-in chờ duyệt', D.dots.length,
      D.dots.slice(0, 5).map(function (x) { return '<li><b>' + esc(x.ten || 'Nhân viên') + '</b>' + (x.so ? ' · ' + esc(x.so) + ' khách' : '') + (x.ngay ? ' · ' + esc(x.ngay) : '') + '</li>'; }),
      'checkin.html?tab=cho', 'Mở Chờ duyệt');
    if (D.dev && D.dev.length) h += sec('📱', '#7c3aed', 'Máy nhân viên chờ duyệt', D.dev.length,
      D.dev.slice(0, 5).map(function (x) { return '<li><b>' + esc(x.nv) + '</b> · ' + esc(x.name || 'máy mới') + '</li>'; }), 'nvadmin.html', 'Mở Phân quyền trợ lý');
    if (!tong()) h += '<div class="lxE">' + (D.at ? '✅ Không có yêu cầu nào của nhân viên đang chờ' : '⏳ Đang tải…') + '</div>';
    h += '<div class="lxF">' + (D.at ? 'Cập nhật lúc ' + tf(D.at) + ' · tự làm mới mỗi phút' : '') + '</div>';
    panel.innerHTML = h;
  }
  var dangTai = false, truoc = -1;
  function tai(force) {
    if (dangTai) return; dangTai = true;
    var jobs = [];
    if (PW && window.luxApi) jobs.push(window.luxApi('thongbao', { pw: PW }).then(function (d) { if (d && d.ok) { D.de = d.de || []; } }).catch(function () {}));
    if (CKPW) jobs.push(jsonp(CK + '?list=1&pw=' + encodeURIComponent(CKPW)).then(function (d) { if (d && d.ok) D.dots = (d.dots || []).filter(function (x) { return x.st === 'CHỜ DUYỆT'; }); }).catch(function () {}));
    if (MG) jobs.push(jsonp(NV + '?nvadmin=list&pw=' + encodeURIComponent(MG)).then(function (d) {
      if (d && d.ok) { var a = []; (d.list || []).forEach(function (nv) { (nv.devices || []).forEach(function (x) { if (x.status !== 'ok') a.push({ nv: nv.ten || nv.ma || '', name: x.name }); }); }); D.dev = a; } }).catch(function () {}));
    Promise.all(jobs).then(function () {
      dangTai = false; D.at = Date.now(); var n = tong();
      if (truoc >= 0 && n > truoc) { bell.classList.remove('rung'); void bell.offsetWidth; bell.classList.add('rung'); }
      truoc = n; ve();
    });
  }
  if (NV_MODE) {
    var dev = ls('nv_devid'), ua = navigator.userAgent || '', dn = /iPhone/.test(ua) ? 'iPhone' : /iPad/.test(ua) ? 'iPad' : /Android/.test(ua) ? 'Android' : (navigator.platform || 'Máy');
    var q = '&tok=' + encodeURIComponent(TOK) + '&dev=' + encodeURIComponent(dev) + '&devname=' + encodeURIComponent(dn), tq = '?tok=' + encodeURIComponent(TOK);
    var SEEN = {}; try { SEEN = JSON.parse(ls('lx_seen') || '{}'); } catch (e) {}
    var NV = { items: [], at: 0, moi: {} };
    var LAB = { save: '✏️ Sửa thông tin khách', out: '🗑 Xoá khách khỏi căn', unitout: '🏁 Check-out cả căn', add: '🛎️ Bổ sung khách', export: '⬇️ Xuất Excel' };
    var STL = { duyet: ['✅ Đã duyệt', '#15803d'], tuchoi: ['✕ Từ chối', '#b91c1c'], loi: ['⚠️ Lỗi ghi file', '#b45309'], huy: ['↩️ Chủ đã huỷ lệnh', '#b45309'],
      'ĐÃ DUYỆT': ['✅ Đã duyệt', '#15803d'], 'ĐÃ CẬP NHẬT FILE': ['✅ Đã duyệt · đã ghi file', '#15803d'], 'TRẢ LẠI': ['↩️ Chủ trả lại', '#b91c1c'], 'ĐÃ NHẬP EXCEL': ['✅ Đã nhập Excel', '#15803d'] };
    var chuaXem = function () { return NV.items.filter(function (x) { return !SEEN[x.k]; }).length; };
    tong = chuaXem;
    ve = function () {
      var n = chuaXem(), c = bell.querySelector('.c');
      c.textContent = n > 99 ? '99+' : n; c.style.display = n ? 'block' : 'none';
      var h = '<h4>🔔 Yêu cầu của tôi đã được xử lý<button onclick="event.stopPropagation();window.lxBellTai(1)">↻ Làm mới</button></h4>';
      if (!NV.items.length) h += '<div class="lxE">' + (NV.at ? 'Chưa có yêu cầu nào được chủ xử lý' : '⏳ Đang tải…') + '</div>';
      NV.items.slice(0, 25).forEach(function (x) {
        var st = STL[x.st] || [x.st, '#5f727a'];
        h += '<a class="lxS" href="' + x.href + '"><div class="t"><i style="background:' + st[1] + '22">' + x.ic + '</i><span>' + esc(x.lb) + '</span>'
          + (SEEN[x.k] && !NV.moi[x.k] ? '' : '<b style="background:#dc2626">Mới</b>') + '</div><ul><li><b style="color:' + st[1] + '">' + st[0] + '</b>' + (x.at ? ' · ' + esc(x.at) : '') + '</li>'
          + (x.dt ? '<li>' + esc(x.dt) + '</li>' : '') + (x.ly ? '<li>📝 ' + esc(x.ly) + '</li>' : '') + '</ul></a>';
      });
      h += '<div class="lxF">' + (NV.at ? 'Cập nhật lúc ' + tf(NV.at) + ' · tự làm mới mỗi phút' : '') + '</div>';
      panel.innerHTML = h;
    };
    tai = function () {
      if (dangTai) return; dangTai = true;
      var a = [], jobs = [];
      if (window.luxApi) jobs.push(window.luxApi('khmyreqs', { tok: TOK, dev: dev, devname: dn }).then(function (d) {
        if (d && d.ok) (d.items || []).forEach(function (o) {
          if (o.status === 'cho') return;
          var g = o.expect || {};
          a.push({ k: 'K' + o.id + o.status, t: o.huyAt || o.decidedAt || o.t, st: o.status, ic: (LAB[o.type] || '📋').split(' ')[0], lb: (LAB[o.type] || o.type).replace(/^\S+\s/, '') + (o.lo ? ' (lô DS Check-in)' : ''),
            dt: [o.ma || g.ma, g.ten].filter(Boolean).join(' · '), ly: o.reason || o.err || '', at: tf(o.huyAt || o.decidedAt || o.t), href: 'khach.html' + tq + '&view=mine' });
        });
      }).catch(function () {}));
      jobs.push(jsonp(CK + '?list=1' + q).then(function (d) {
        if (d && d.ok) (d.dots || []).forEach(function (x) {
          if (!x.st || x.st === 'CHỜ DUYỆT') return;
          a.push({ k: 'D' + x.id + x.st, t: 0, st: x.st, ic: '🛎️', lb: 'Đợt check-in' + (x.so ? ' · ' + x.so + ' khách' : ''), dt: x.ngay ? 'Gửi ' + x.ngay : '', ly: x.dnote || '', at: '', href: 'checkin.html' + tq + '&tab=cua' });
        });
      }).catch(function () {}));
      Promise.all(jobs).then(function () {
        dangTai = false; a.sort(function (x, y) { return (y.t || 0) - (x.t || 0); }); NV.items = a; NV.at = Date.now();
        if (!ls('lx_seen_init')) { a.forEach(function (x) { SEEN[x.k] = 1; }); try { localStorage.setItem('lx_seen', JSON.stringify(SEEN)); localStorage.setItem('lx_seen_init', '1'); } catch (e) {} }
        var n = chuaXem(); if (truoc >= 0 && n > truoc) { bell.classList.remove('rung'); void bell.offsetWidth; bell.classList.add('rung'); }
        truoc = n; ve();
      });
    };
    bell.title = 'Yêu cầu của tôi đã được xử lý';
    bell.onclick = function (e) {
      e.stopPropagation(); panel.classList.toggle('on');
      if (panel.classList.contains('on')) { NV.moi = {}; NV.items.forEach(function (x) { if (!SEEN[x.k]) NV.moi[x.k] = 1; SEEN[x.k] = 1; }); ve(); try { localStorage.setItem('lx_seen', JSON.stringify(SEEN)); } catch (er) {} var c = bell.querySelector('.c'); c.style.display = 'none'; tai(); }
    };
  }
  window.lxBellTai = tai;
  ve(); tai();
  setInterval(function () { if (!document.hidden) tai(); }, 60000);
  document.addEventListener('visibilitychange', function () { if (!document.hidden) tai(); });
})();
