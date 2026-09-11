# Deploy frontend lên GitHub Pages

## 1. Chuẩn bị repository

1. Tạo một GitHub repository (có thể public hoặc private nếu dùng GitHub
   Pro/Team/Enterprise — Pages với repo private cần gói trả phí).
2. Đưa toàn bộ nội dung thư mục `web/` vào **thư mục gốc** của repo, hoặc
   vào thư mục `docs/` nếu bạn muốn dùng branch chính cho code khác.
3. Đảm bảo đã sửa `GAS_URL` trong `auth.js` trỏ đúng Web App URL đã deploy.

## 2. Bật GitHub Pages

1. Vào **Settings > Pages** của repository.
2. **Source**: chọn branch (`main`) và thư mục (`/root` hoặc `/docs` tuỳ
   bước 1).
3. Bấm **Save**. Sau 1–2 phút, GitHub cung cấp URL dạng:
   `https://<username>.github.io/<repo>/`

## 3. Giới hạn CORS phía Apps Script (tuỳ chọn, khuyến nghị)

Mặc định `CONFIG.ALLOWED_ORIGIN = '*'` trong `Config.gs` không thực sự giới
hạn gì vì Apps Script Web App không set header `Access-Control-Allow-Origin`
theo cách kiểm soát được từ code — trình duyệt vẫn cho phép gọi cross-origin
tới `script.google.com`. Lớp bảo vệ thực sự nằm ở:

- Đăng nhập bằng CCCD + mật khẩu cho endpoint đọc dữ liệu lương.
- `SYNC_API_KEY` cho endpoint `/sync`.

Nếu muốn hạn chế thêm, có thể kiểm tra `e.parameter` hoặc header `Referer`
trong `doPost()`, nhưng lưu ý header này có thể bị giả mạo và không nên là
lớp bảo mật duy nhất.

## 4. Cập nhật sau khi sửa frontend

Chỉ cần `git push` các thay đổi trong thư mục đã cấu hình làm Pages source;
GitHub Pages tự build lại (thường trong vài chục giây tới vài phút).

## 5. Tên miền riêng (tuỳ chọn)

Nếu công ty có tên miền riêng, có thể trỏ CNAME về
`<username>.github.io` và thêm file `CNAME` chứa tên miền vào thư mục Pages,
theo hướng dẫn chính thức của GitHub.
