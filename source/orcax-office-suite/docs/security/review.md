# Đánh giá An ninh & Kiến trúc Bảo mật (Security Review & Baseline)

Tài liệu này tổng hợp kết quả đánh giá các biện pháp kiểm soát an ninh thông tin, quyền riêng tư dữ liệu và cách ly người dùng trong OrcaX Office Suite theo baseline MVP và Beta Readiness (E10).

## 1. Cam kết Quyền riêng tư & Xử lý Local-first

- **Non-AI Policy:** Tuyệt đối không sử dụng các mô hình ngôn ngữ lớn (LLM), API generative AI hoặc các thư viện suy luận xác suất. Toàn bộ tính năng chuẩn hóa và kiểm tra tuân theo quy tắc toán học, biểu thức chính quy và định dạng quy chuẩn tất định (deterministic).
- **Không tự động tải tài liệu:** Nội dung các tài liệu Word và bảng tính Excel được xử lý hoàn toàn cục bộ trên client (thông qua Office.js và taskpane). API máy chủ chỉ tiếp nhận metadata kết quả kiểm tra (tổng số lỗi, cảnh báo, mã quy tắc) và thông tin cấu hình in ấn. Không có endpoint upload nội dung tệp Office mặc định.

## 2. Cách ly Đa người thuê (Multi-Tenancy Isolation)

- **Biên giới Tenant:** Ranh giới cách ly người thuê dựa trên `workspace_id`. Mỗi người dùng sở hữu một không gian cá nhân (`personal`) và có thể tham gia vào các không gian tổ chức (`organization`).
- **Phạm vi dữ liệu:** Mọi bảng chứa dữ liệu nghiệp vụ đều mang cột `workspace_id` và chỉ mục liên kết. Tầng dịch vụ backend luôn trích xuất `workspace_id` từ `AccessContext` của token đã được xác thực, ngăn chặn hoàn toàn nguy cơ truy cập chéo dữ liệu giữa các workspace (IDOR).
- **Chuyển đổi Không gian làm việc:** Luôn yêu cầu cờ `confirmed: true`. Việc chuyển đổi workspace sẽ cấp phát token mới có `wid` khớp với phiên làm việc, token cũ của workspace trước đó ngay lập tức bị từ chối.

## 3. Tính Bất biến của Nhật ký Kiểm toán (Audit Immutability)

- Bảng `audit_events` được bảo vệ bằng trigger cấp cơ sở dữ liệu (`audit_events_immutable`), ngăn chặn mọi thao tác `UPDATE`, `DELETE` và `TRUNCATE`.
- Trình ghi nhật ký `AuditRecorder` thực hiện thanh lọc dữ liệu trước khi ghi: tự động từ chối metadata chứa mật khẩu, token phiên, cookie, thông tin thẻ tín dụng hoặc nội dung văn bản.
- Dung lượng mỗi bản ghi nhật ký kiểm toán bị giới hạn dưới 16 KB.

## 4. Xác thực và Phân quyền (Authentication & Authorization)

- **Xác thực:** Sử dụng Bearer Access Token thời hạn ngắn kết hợp với Refresh Cookie (`orcax_session`) có cờ `HttpOnly`, `SameSite=Lax`.
- **Phân quyền kết hợp RBAC & ABAC:**
  - Mô hình vai trò: `personal_owner`, `organization_owner`, `organization_administrator`, `member`, `auditor`.
  - Quyền hạn được kiểm soát thông qua `@RequirePermission(...)` và đối soát thuộc tính động (trạng thái workspace hoạt động, vai trò thành viên).
- **CORS:** Chỉ chấp nhận duy nhất origin được cấu hình tại biến môi trường `WEB_ORIGIN`.

## 5. Quản lý Bí mật và Khóa

- Không có giá trị bí mật mặc định trong mã nguồn.
- File cấu hình `.env` được loại trừ khỏi kho lưu trữ mã nguồn qua `.gitignore`.
- Khóa `AUTH_TOKEN_SECRET` bắt buộc có độ dài tối thiểu 32 ký tự, được kiểm tra nghiêm ngặt khi khởi động ứng dụng.
- Webhook bí mật `PAYMENT_WEBHOOK_SECRET` được xác thực qua header chuyên dụng cho mọi thông điệp từ cổng thanh toán.
