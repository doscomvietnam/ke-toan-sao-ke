# Kế toán DOSCOM — Xử lý sao kê ngân hàng

Tự động tạo Phiếu thu / Phiếu chi từ file sao kê ngân hàng.

## Cấu trúc

```
.
├── kt.py                       # Logic xử lý chính (vừa CLI vừa web)
├── api/process.py              # Vercel serverless function (Flask)
├── index.html                  # Trang upload
├── vercel.json                 # Cấu hình deploy Vercel
├── requirements.txt            # Python dependencies
├── Template-dau-ra.xlsx        # Template phiếu thu / chi
├── ngan-hang.xlsx              # Bảng tra số TK → tên ngân hàng
├── ma-thong-ke.xlsx            # Bảng tra TK Nợ + Mã KMCP → Mã thống kê
└── lich-su-giao-dich-tai-khoan (17).xlsx   # File sao kê mẫu để test
```

## Chạy local (test trước khi deploy)

```bash
pip install -r requirements.txt
python api/process.py
```

Mở trình duyệt vào http://localhost:5000 — chỉ thấy API. Để xem giao diện upload, mở trực tiếp `index.html` (hoặc dùng `python -m http.server` ở thư mục gốc rồi sửa fetch URL nếu khác origin).

Cách đơn giản nhất khi dev local: dùng Vercel CLI.

```bash
npm i -g vercel
vercel dev
```

Lệnh này khởi động cả static (index.html) và serverless function (api/process.py) trên cùng port — y hệt môi trường production.

## Deploy Vercel

1. Push toàn bộ thư mục này lên GitHub.
2. Vào https://vercel.com → New Project → import repo.
3. Vercel tự nhận `vercel.json`, `requirements.txt`, `api/`. Bấm Deploy.
4. Sau khi deploy xong, gửi URL cho kế toán dùng.

### Cập nhật bảng tra ngân hàng / mã thống kê

Khi cần thêm dòng vào `ngan-hang.xlsx` hoặc `ma-thong-ke.xlsx`:

1. Sửa file `.xlsx` trong repo.
2. Commit + push → Vercel tự deploy lại.

## CLI (chạy bằng dòng lệnh)

```bash
python kt.py
```

Đường dẫn cố định ở phần `CẤU HÌNH ĐƯỜNG DẪN` đầu file `kt.py`. Kết quả lưu thành `Ket-qua-ke-toan.xlsx`.
