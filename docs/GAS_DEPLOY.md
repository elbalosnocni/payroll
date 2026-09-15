# Deploy Google Apps Script làm Web App (API)

## 1. Deploy lần đầu

1. Trong Apps Script Editor (đã dán đủ code ở `gas/`), bấm **Deploy > New deployment**.
2. Chọn loại: **Web app**.
3. Cấu hình:
   - **Execute as**: `Me` (tài khoản của bạn — script cần quyền đọc/ghi Sheet).
   - **Who has access**: `Anyone` (bắt buộc, để GitHub Pages/VBA gọi được mà
     không cần đăng nhập Google — bảo mật thực sự nằm ở lớp CCCD+mật khẩu và
     `SYNC_API_KEY`, không phải ở quyền truy cập Apps Script).
4. Bấm **Deploy**, cấp quyền nếu được hỏi.
5. Copy **Web app URL** dạng:
   `https://script.google.com/macros/s/XXXXXXXXXXXXXXXX/exec`
6. Dán URL này vào:
   - `web/auth.js` → biến `GAS_URL`
   - `vba/Config.bas` → hằng số `GAS_URL`

## 2. Mỗi lần sửa code Apps Script

Apps Script Web App **không tự cập nhật** deployment cũ khi bạn sửa code.
Sau khi sửa, phải:

1. **Deploy > Manage deployments**.
2. Chọn deployment hiện tại > biểu tượng bút chì (Edit).
3. Ở mục **Version**, chọn **New version**.
4. Bấm **Deploy**.

URL Web App giữ nguyên, không cần cập nhật lại phía frontend/VBA.

## 3. Vì sao dùng `Content-Type: text/plain` khi gọi API?

Trình duyệt sẽ gửi preflight request (`OPTIONS`) trước mọi request có
`Content-Type: application/json`. Apps Script Web App **không hỗ trợ**
`doOptions`, nên preflight sẽ thất bại và gây lỗi CORS. Giải pháp phổ biến:
gửi với header `Content-Type: text/plain;charset=utf-8` (không kích hoạt
preflight) nhưng vẫn để **nội dung body là JSON hợp lệ** — phía
`Code.gs > doPost()` vẫn `JSON.parse()` bình thường vì Apps Script đọc từ
`e.postData.contents`, không quan tâm header khai báo là gì.

## 4. Kiểm tra nhanh

Mở terminal / Postman, gọi:

```bash
curl -X POST "<GAS_URL>" -H "Content-Type: text/plain" -d "{\"action\":\"ping\"}"
```

Kết quả mong đợi: `{"ok":true,"data":{"pong":true}}`

## 5. Bảo mật cần lưu ý

- `SYNC_API_KEY` trong Script Properties phải là chuỗi dài, ngẫu nhiên, và
  **khớp chính xác** với giá trị trong `vba/Config.bas`. Không commit giá
  trị thật lên Git repository công khai — nếu repo GitHub Pages là public,
  hãy tách `web/auth.js` phần `GAS_URL` ra biến môi trường/khác repo private
  nếu cần giữ bí mật đường dẫn (thực tế URL Apps Script khó đoán nên rủi ro
  thấp, nhưng vẫn nên hạn chế commit các secret khác).
- Đừng đặt `Who has access` là "Anyone" cho một API xử lý dữ liệu nhạy cảm
  mà không có lớp xác thực riêng — project này đã có lớp xác thực CCCD +
  mật khẩu (nhân viên) và `SYNC_API_KEY` (VBA), nên việc mở public endpoint
  là chấp nhận được, nhưng không nên bỏ các lớp xác thực đó đi.
