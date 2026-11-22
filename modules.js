// ===== DANH SÁCH MODULE DÙNG CHUNG (LUXHOME) =====
// Chỉ cần thêm/sửa 1 dòng ở đây là module tự xuất hiện đồng bộ:
//   - Màn hình chính (index.html)  -> hiện tất cả (kể cả mục quản trị)
//   - Phân quyền app phụ (nvadmin) -> hiện các mục staff:true để admin tick chọn
//   - App phụ (nv.html)            -> nhân viên thấy đúng mục được cấp quyền
// staff:false = mục chỉ dành cho admin, KHÔNG hiện trong phân quyền.
// Mục chưa xây xong thì KHÔNG đưa vào đây (để ở danh sách "Sắp có" trong index.html).
window.LUX_MODULES = [
  { k:"quanlycanho", ic:"🏢", nm:"Quản lý căn hộ",     en:"Apartment management",  ds:"Danh mục · Khóa/Mở căn", url:"quanlycanho.html", staff:true },
  { k:"luutru",   ic:"🛂", nm:"Lưu trú & hạn visa",    en:"Residence & visa",      ds:"Nhắc gia hạn lưu trú",   url:"luutru.html",   staff:true },
  { k:"quanlythanhtoan", ic:"💳", nm:"Quản lý thanh toán", en:"Payment management", ds:"Gửi hóa đơn · Công nợ", url:"quanlythanhtoan.html", staff:true },
  { k:"thongbao", ic:"📢", nm:"Thông báo cư dân",      en:"Resident notices",      ds:"Báo cư dân theo toà",    url:"thongbao.html", staff:true },
  { k:"nvadmin",  ic:"🧑‍💼", nm:"Quản lý quyền app phụ", en:"Staff app permissions", ds:"Phân quyền app phụ",     url:"nvadmin.html",  staff:false },
  { k:"noibo",    ic:"🔁", nm:"Hoạt động định kỳ",     en:"Periodic activities",   ds:"Vệ sinh, PCCC…",         url:"noibo.html",    staff:true },
  { k:"hopdong",  ic:"📄", nm:"Hợp đồng PDF",          en:"Contracts & annexes",   ds:"Tạo HĐ/PLHĐ",            url:"hopdong.html",  staff:true },
  { k:"checkin",  ic:"🛎️", nm:"Cập nhật khách check-in", en:"Check-in updates",  ds:"Gửi DS check-in chờ duyệt", url:"checkin.html", staff:true },
];
// Mã quyền cũ (trước khi gộp) -> mã mới, để nhân viên đã được cấp quyền cũ vẫn thấy mục mới
window.LUX_PERM_ALIAS = { catalog:"quanlycanho", canho:"quanlycanho", bill:"quanlythanhtoan", congno:"quanlythanhtoan" };
// Khi LƯU quyền: mục gộp được lưu kèm mã cũ để máy chủ (Apps Script) vẫn nhận quyền từng chức năng con
window.LUX_PERM_EXPAND = { quanlycanho:["catalog","canho"], quanlythanhtoan:["bill","congno"] };
window.luxPermsSave = function(ps){ var E=window.LUX_PERM_EXPAND, o=[]; (ps||[]).forEach(function(k){ [k].concat(E[k]||[]).forEach(function(x){ if(o.indexOf(x)<0) o.push(x); }); }); return o; };
window.luxPerms = function(ps){ var A=window.LUX_PERM_ALIAS, o=[]; (ps||[]).forEach(function(k){ k=A[k]||k; if(o.indexOf(k)<0) o.push(k); }); return o; };
