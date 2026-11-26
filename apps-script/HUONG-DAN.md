# Cài máy chủ "Quản lý căn hộ" (làm 1 lần, khoảng 5 phút)

Đăng nhập Google bằng **tài khoản đang sở hữu Sheet "KHO ẢNH CĂN HỘ - CSKH" và thư mục "KHO ẢNH CĂN HỘ"**.

1. Mở https://script.google.com → **Dự án mới**. Đặt tên: `LUXHOME Căn hộ API`.
2. Xóa hết mã có sẵn trong `Code.gs`, dán nội dung 4 file nối tiếp nhau (cùng 1 file): `canho-api.gs`, `khach-api.gs`, `nhatky-api.gs`, `dexuat-api.gs` → bấm 💾 Lưu.
3. Ở thanh trên, chọn hàm **capQuyen** → bấm **Chạy** → **Xem xét quyền** → chọn tài khoản → *Nâng cao* → *Đi tới LUXHOME Căn hộ API* → **Cho phép**.
4. **Triển khai → Tùy chọn triển khai mới** → bánh răng ⚙ → **Ứng dụng web**:
   - Thực thi với tư cách: **Tôi**
   - Người có quyền truy cập: **Bất kỳ ai**
   → **Triển khai** → copy **URL ứng dụng web** (kết thúc bằng `/exec`).
5. Gửi URL đó cho Claude để gắn vào app (`canho-api.js`).
6. Sau khi app chạy ổn: mở Sheet "KHO ẢNH CĂN HỘ - CSKH" → **Chia sẻ** → *Quyền truy cập chung* → **Bị hạn chế**.
   Từ lúc này Sheet (kể cả trang MẬT KHẨU) không còn ai ngoài bạn đọc được; app vẫn chạy bình thường qua máy chủ.

## Khi sửa mã máy chủ sau này
**Triển khai → Quản lý các lượt triển khai** → ✏️ → Phiên bản: **Phiên bản mới** → Triển khai. (Giữ nguyên URL, không cần sửa app.)
