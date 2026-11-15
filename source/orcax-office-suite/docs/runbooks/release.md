# Quy trình Phát hành và Vận hành (Release Runbook)

Tài liệu này xác định các cổng chất lượng (release gates), quy trình đóng gói và triển khai bản phát hành cho OrcaX Office Suite.

## 1. Cổng kiểm tra trước khi phát hành (Pre-flight Release Gates)

Trước khi kích hoạt bất kỳ quy trình phát hành hoặc gắn tag phiên bản:

1. **Kiểm tra cam kết Non-AI & Cấu hình:**

   ```powershell
   npm run check:foundation
   ```

   _Yêu cầu:_ Không có bất kỳ dependency AI bị cấm nào (`openai`, `langchain`, `generative-ai`, ...); `synchronize` trong TypeORM phải luôn là `false`.

2. **Kiểm tra định dạng và quy ước mã nguồn:**

   ```powershell
   npm run format:check
   npm run lint
   ```

3. **Chạy toàn bộ bộ kiểm thử tự động:**

   ```powershell
   npm test
   ```

   _Yêu cầu:_ Toàn bộ unit test và cross-application test (compatibility, e2e smoke, integration) đều phải đạt kết quả `PASS` 100%.

4. **Đóng gói toàn bộ ứng dụng:**
   ```powershell
   npm run build
   ```
   _Yêu cầu:_ `apps/api`, `apps/web-portal`, và `apps/office-js` đều phải biên dịch thành công không có cảnh báo nghiêm trọng.

## 2. Thứ tự Triển khai (Deployment Sequence)

1. **Chuẩn bị môi trường & Cơ sở dữ liệu:**
   - Đảm bảo các biến môi trường bắt buộc đã được cấu hình qua `.env` (bí mật tối thiểu 32 ký tự cho `AUTH_TOKEN_SECRET`).
   - Thực hiện sao lưu cơ sở dữ liệu hiện tại trước khi can thiệp:
     ```powershell
     .\infrastructure\deployment\backup-restore.ps1 -Action backup
     ```

2. **Chạy di chuyển cơ sở dữ liệu (Migrations):**

   ```powershell
   npm run migration:run
   ```

   Kiểm tra danh sách migration đã chạy:

   ```powershell
   npm run migration:show
   ```

3. **Triển khai ứng dụng API & Web Portal:**
   - Cập nhật phiên bản container hoặc khởi động process dịch vụ.
   - Kiểm tra liveness và readiness probes:
     - `GET http://localhost:3001/health` -> HTTP 200 `{ status: "ok" }`
     - `GET http://localhost:3001/health/ready` -> HTTP 200 (database ping thành công)

4. **Triển khai Add-in Manifests:**
   - Cung cấp manifest `manifests/word-manifest.xml` và `manifests/excel-manifest.xml` cho người dùng quản trị tổ chức để phân phối tập trung qua Microsoft 365 Admin Center hoặc tải trực tiếp.

## 3. Kế hoạch Hoàn tác (Rollback Plan)

Khi phát hiện sự cố nghiêm trọng sau triển khai:

1. Chuyển hướng lưu lượng truy cập về phiên bản trước (traffic rollback).
2. Hoàn tác migration cơ sở dữ liệu nếu có migration mới bị lỗi:
   ```powershell
   npm run migration:revert
   ```
3. Trong trường hợp hỏng hóc dữ liệu, thực hiện khôi phục từ bản sao lưu tạo trước triển khai:
   ```powershell
   .\infrastructure\deployment\backup-restore.ps1 -Action restore -File <backup_file>
   ```
4. Ghi nhận sự cố và kích hoạt phân tích nguyên nhân gốc rễ (RCA).
