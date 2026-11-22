// ===== DANH SÁCH MODULE DÙNG CHUNG (LUXHOME) =====
// Chỉ cần thêm/sửa 1 dòng ở đây là module tự xuất hiện đồng bộ:
//   - Màn hình chính (index.html)  -> hiện tất cả (kể cả mục quản trị)
//   - Phân quyền app phụ (nvadmin) -> hiện các mục staff:true để admin tick chọn
//   - App phụ (nv.html)            -> nhân viên thấy đúng mục được cấp quyền
// staff:false = mục chỉ dành cho admin, KHÔNG hiện trong phân quyền.
// Mục chưa xây xong thì KHÔNG đưa vào đây (để ở danh sách "Sắp có" trong index.html).
window.LUX_MODULES = [
  { k:"catalog",  ic:"🏢", nm:"Danh mục căn hộ",      en:"Apartment catalog",     ds:"Xem danh mục căn hộ",    url:"catalog.html#k=LUXADMIN2026", staff:true },
  { k:"luutru",   ic:"🛂", nm:"Lưu trú & hạn visa",    en:"Residence & visa",      ds:"Nhắc gia hạn lưu trú",   url:"luutru.html",   staff:true },
  { k:"bill",     ic:"🧾", nm:"Gửi hóa đơn",           en:"Send invoices",         ds:"Gửi bill cho khách",     url:"bill.html",     staff:true },
  { k:"thongbao", ic:"📢", nm:"Thông báo cư dân",      en:"Resident notices",      ds:"Báo cư dân theo toà",    url:"thongbao.html", staff:true },
  { k:"nvadmin",  ic:"🧑‍💼", nm:"Quản lý quyền app phụ", en:"Staff app permissions", ds:"Phân quyền app phụ",     url:"nvadmin.html",  staff:false },
  { k:"noibo",    ic:"🔁", nm:"Hoạt động định kỳ",     en:"Periodic activities",   ds:"Vệ sinh, PCCC…",         url:"noibo.html",    staff:true },
  { k:"hopdong",  ic:"📄", nm:"Hợp đồng PDF",          en:"Contracts & annexes",   ds:"Tạo HĐ/PLHĐ",            url:"hopdong.html",  staff:true },
  { k:"congno",   ic:"💰", nm:"Theo dõi công nợ",      en:"Debt & payments",       ds:"Công nợ khách",          url:"congno.html",   staff:true },
  { k:"checkin",  ic:"🛎️", nm:"Cập nhật khách check-in", en:"Check-in updates",  ds:"Gửi DS check-in chờ duyệt", url:"checkin.html", staff:true },
  { k:"canho",    ic:"🔑", nm:"Khóa / Mở căn hộ",      en:"Lock / open units",     ds:"Khoá/mở căn khi in-out", url:"canho.html",    staff:true },
];
