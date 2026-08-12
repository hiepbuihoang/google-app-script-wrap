# google-app-script-wrap

Bọc web app Google Apps Script **"Quản lý Công việc"** vào domain riêng `os.smiletrip.vn`.

Toàn bộ site là 1 file tĩnh: [`public/index.html`](public/index.html).

## Cách hoạt động

`os.smiletrip.vn` phục vụ một trang bọc, nhúng thẳng `/exec` bằng iframe chiếm đúng
100% viewport. Người dùng vào domain là thấy app y hệt vào link Apps Script gốc.

Đã kiểm tra: response của `/exec` **không set `X-Frame-Options`**, nên nhúng được
ngay, không cần sửa gì trong Apps Script.

Ưu điểm so với reverse proxy: `google.script.run` vẫn chạy bình thường vì iframe
load đúng origin gốc của Google.

Đánh đổi: URL không đổi khi điều hướng bên trong app (không deep-link / bookmark
trang con), và không có SEO.

## Deploy — Cloudflare Pages (khuyên dùng, free)

1. Push repo này lên GitHub.
2. Cloudflare Dashboard → **Workers & Pages** → **Create** → **Pages** → **Connect to Git** → chọn repo.
3. Build settings:
   - Framework preset: **None**
   - Build command: *(để trống)*
   - Build output directory: **`public`**
4. Deploy xong → tab **Custom domains** → **Set up a custom domain** → nhập `os.smiletrip.vn`.
5. DNS:
   - Nếu `smiletrip.vn` đã dùng nameserver Cloudflare → Cloudflare tự thêm record, xong.
   - Nếu DNS đang ở nhà cung cấp khác (PA Vietnam, Mắt Bão, Nhân Hòa…) → thêm thủ công:

     | Type  | Name | Value                    |
     |-------|------|--------------------------|
     | CNAME | `os` | `<ten-project>.pages.dev` |

SSL cấp tự động, chờ vài phút.

### Hosting khác

- **Vercel / Netlify**: trỏ thư mục publish vào `public`, rồi add domain `os.smiletrip.vn` (CNAME theo hướng dẫn của họ).
- **VPS sẵn có (nginx)**: copy `public/index.html` vào webroot của vhost `os.smiletrip.vn`, chạy `certbot` cho SSL. Thêm record A trỏ về IP VPS.

## Chạy thử local

```bash
npx serve public -l 4173
```

## Đổi link Apps Script

Khi tạo deployment mới trên Apps Script, sửa URL ở **2 chỗ** trong
`public/index.html`: thuộc tính `src` của `<iframe>` và `href` của link
"Mở trực tiếp" trong khối `#fallback`.
