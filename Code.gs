# QUANLYLUANVAN — Hệ thống quản lý luận văn tốt nghiệp
### Khoa Kỹ thuật Công trình — Trường Đại học Lạc Hồng

Web app hoàn chỉnh, chạy với **tên miền riêng**, chi phí hạ tầng ~0đ:

| Thành phần | Công nghệ | Chi phí |
|---|---|---|
| Giao diện (frontend) | HTML/JS tĩnh — host trên **GitHub Pages** | Miễn phí |
| API + CSDL (backend) | **Google Apps Script** + **Google Sheets** | Miễn phí |
| File luận văn, PDF phiếu chấm | **Google Drive** của Khoa | Miễn phí |
| Tên miền | Bạn mua (vd `luanvan-ktct.vn`) | ~200–400k/năm |

## 5 vai trò (khóa theo tài khoản — không thể vào nhầm)

- **Sinh viên** — tự tạo tài khoản; đăng ký đề tài; xem thời gian biểu; ghi tiến độ; nộp file luận văn; xem phân công GVHD/GVPB/hội đồng.
- **Người quản lý / Thư ký** — duyệt hồ sơ; phân công GVHD & GVPB; lập hội đồng; nạp danh sách giảng viên (CSV); đổi rubric; xem tổng hợp điểm; xuất PDF tổng hợp.
- **GV hướng dẫn** — thấy SV mình hướng dẫn; theo dõi tiến độ + ghi chú; điểm danh buổi duyệt bài; chấm điểm + nhận xét; ký số; xuất phiếu PDF.
- **GV phản biện** — thấy SV được phân công; mở file luận văn; nhận xét, chấm điểm; ký số; xuất phiếu PDF.
- **Hội đồng** (chủ tịch + 2 ủy viên) — thấy SV thuộc hội đồng của mình (kèm số hội đồng); chấm theo rubric; chủ tịch xem điểm của cả 5 người chấm.

**Cách hệ thống chống vào nhầm vai trò:** không có ô "chọn vai trò". Đăng nhập bằng email → hệ thống tra vai trò và các phân công gắn với email đó → chỉ hiện đúng khu làm việc và đúng sinh viên thuộc quyền. Một giảng viên kiêm nhiều vai (vd vừa GVHD vừa ủy viên hội đồng) sẽ thấy các tab tách biệt, mỗi tab ghi rõ vai trò.

## Cấu trúc repo

```
index.html        ← giao diện web app (mở là chạy, chưa có API thì tự vào chế độ DEMO)
config.js         ← DÁN API_URL của Apps Script vào đây
mock.js           ← dữ liệu demo (chỉ dùng khi API_URL trống)
backend/
  Code.gs         ← backend API — dán vào Google Apps Script
  appsscript.json ← manifest Apps Script
README.md
```

## Triển khai (3 phần, ~20 phút)

### Phần 1 — Backend (Google Apps Script)
1. Tài khoản Google của Khoa → tạo **Google Trang tính** mới tên `CSDL Luận văn KTCT`.
2. **Tiện ích mở rộng → Apps Script** → xóa code mặc định, dán toàn bộ `backend/Code.gs`.
3. (Nên làm) Project Settings → tick *Show appsscript.json* → dán nội dung `backend/appsscript.json`.
4. Trong editor chọn hàm **`setup`** → **Run** → cấp quyền (Review permissions → Allow). Các sheet dữ liệu + tài khoản mẫu được tạo.
5. **Deploy → New deployment → Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone** *(bắt buộc — frontend ở tên miền khác gọi API vào đây; dữ liệu vẫn an toàn vì mọi thao tác đều yêu cầu đăng nhập email + mật khẩu)*
6. Copy **Web app URL** (`https://script.google.com/macros/s/…/exec`).

### Phần 2 — Frontend (GitHub Pages)
1. Mở `config.js`, dán URL vừa copy vào `API_URL`.
2. Đưa toàn bộ file lên repo GitHub `nguyendinhdu-maker/QUANLYLUANVAN` (kéo-thả trên github.com: **Add file → Upload files**).
3. Repo → **Settings → Pages** → Source: **Deploy from a branch** → Branch `main` / folder `/ (root)` → Save.
4. Vài phút sau web chạy tại `https://nguyendinhdu-maker.github.io/QUANLYLUANVAN/`.

### Phần 3 — Gắn tên miền riêng
1. Mua tên miền (Mắt Bão, PA, Tenten, iNET…).
2. Repo → Settings → Pages → **Custom domain** → nhập `luanvan.tenmiencuaban.vn` → Save (GitHub tự tạo file CNAME).
3. Ở trang quản trị tên miền, thêm bản ghi DNS:
   - **CNAME** | `luanvan` (hoặc `www`) | `nguyendinhdu-maker.github.io`
   - (Nếu dùng tên miền gốc: 4 bản ghi **A** → `185.199.108.153`, `185.199.109.153`, `185.199.110.153`, `185.199.111.153`)
4. Chờ DNS (5–60 phút) → bật **Enforce HTTPS** trong Settings → Pages.

## Tài khoản mặc định sau khi chạy `setup()`

- Quản lý/Thư ký: `quanly@lhu.edu.vn` / `ktct@2026` *(đổi ngay sau khi triển khai)*
- Giảng viên mẫu: `hoang.tran@lhu.edu.vn`, `hong.le@lhu.edu.vn`, `son.do@lhu.edu.vn`… / `123456`
- Sinh viên tự bấm **"Tạo tài khoản"** trên trang đăng nhập.

Nạp danh sách giảng viên thật: đăng nhập Quản lý → tab **Danh sách giảng viên** → tải CSV
(`email ; họ tên ; chức danh ; mật khẩu`).

## Quy trình nghiệp vụ

1. SV tạo tài khoản → đăng ký đề tài → Thư ký duyệt.
2. Thư ký phân công GVHD, GVPB, lập hội đồng (chủ tịch + 2 ủy viên), xếp SV vào hội đồng.
3. SV thực hiện: ghi tiến độ, GVHD ghi chú + điểm danh buổi duyệt bài.
4. SV nộp file luận văn (PDF → Google Drive Khoa).
5. GVHD, GVPB, 3 thành viên hội đồng chấm theo rubric + nhận xét → **ký số** (khóa phiếu).
6. Điểm cuối = trung bình 5 người chấm. Chủ tịch thấy điểm mọi thành viên; Thư ký xem tổng hợp toàn khoa.
7. Xuất **PDF** phiếu GVHD / phiếu GVPB / phiếu hội đồng / biên bản tổng hợp — đầy đủ thông tin SV, GV, chữ ký — tự lưu vào Drive.

## Ghi chú kỹ thuật & bảo mật

- Mật khẩu băm SHA-256 + salt, lưu trong sheet `NguoiDung`. Phiên đăng nhập hết hạn sau 6 giờ.
- "Chữ ký số" là xác thực nội bộ (người ký + thời gian + số serial). Cần chữ ký pháp lý → tích hợp CA (VNPT-CA, FPT-CA).
- Giới hạn file nộp ~20MB (giới hạn Apps Script).
- Đổi rubric sẽ xóa điểm & chữ ký hiện có — chỉ đổi trước khi chấm.
- Mở `index.html` khi `API_URL` trống = **chế độ DEMO** (dữ liệu mô phỏng, có sẵn tài khoản demo hiển thị ở trang đăng nhập).
