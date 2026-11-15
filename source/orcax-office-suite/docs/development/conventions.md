# Quy ước phát triển

Các quy ước này giữ baseline E01/E02. Epic sau mở rộng module, không phá các nguyên tắc dưới đây.

## Nơi đặt code

- Hành vi HTTP nằm trong `apps/api`. Controller mỏng. Rule nghiệp vụ nằm trong service của đúng module.
- Giao diện cổng thông tin nằm trong `apps/web-portal`. Copy người dùng nhìn thấy viết bằng tiếng Việt.
- DTO, mã lỗi, envelope và kết quả rule Office nằm trong `packages/api-contracts`.
- Adapter VSTO và Office.js không gọi chéo nội bộ của nhau. Chúng chỉ dùng contract đã xuất bản.
- Không thêm dịch vụ deploy riêng cho một module Nest. Monolith tiếp tục là một API cho đến khi có quyết định kiến trúc mới.

## API

Prefix nghiệp vụ là `/api/v1`, lấy từ `API_VERSION`.

Handler trả dữ liệu thô. Interceptor bọc thành:

```json
{
  "data": {},
  "error": null,
  "meta": { "correlationId": "", "apiVersion": "v1" }
}
```

Lỗi có `error.code`, `error.message` và `error.details`. Thông điệp tổng quát là tiếng Việt. Khi thêm DTO, `message` của `class-validator` cũng viết tiếng Việt.

Mã lỗi dùng `ErrorCodes` trong contract. Không bịa mã mới ở controller.

Mọi DTO bật `whitelist` và `forbidNonWhitelisted`. Tenant là workspace đang hoạt động trên phiên đã kiểm tra. Không lấy `workspace_id` từ một header hay body do client tự khai, trừ thao tác chuyển workspace đã xác nhận. Bản ghi thuộc tenant phải mang `workspace_id` và được lọc bằng `filterByWorkspace` hoặc `assertRecordInWorkspace`. Chi tiết: [tenancy.md](../architecture/tenancy.md).

## Xác thực và phiên

- Public (không Bearer): `POST /auth/register`, `/auth/login`, `/auth/refresh`, `/auth/email/verify`, `/auth/email/resend`, `/auth/password/forgot`, `/auth/password/reset`, và `GET /dev/mailbox` khi `EMAIL_TRANSPORT=log`.
- Access token Bearer sống ngắn, chứa `sub` (user), `sid` (session), `wid` (workspace). Cookie `orcax_session` là refresh HttpOnly.
- `wid` trên token phải khớp `sessions.active_workspace_id`. Sau khi chuyển workspace, token cũ bị từ chối.
- Chuyển workspace cần `confirmed: true`. Đăng nhập luôn vào personal; `last_workspace_id` chỉ là gợi ý.
- Tạo tổ chức chỉ từ personal workspace active. Invite/đổi role chỉ trên organization workspace.
- Register/forgot/resend trả thông điệp generic, không tiết lộ email đã tồn tại.
- Quyền: RBAC trong catalog; ABAC qua `assertPolicy` / `@RequirePermission` (loại workspace, membership active, ownership). Chi tiết: [authorization.md](../architecture/authorization.md).
- `GET /audit-events` chỉ trong `workspace_id` đang hoạt động; cần `workspace.audit.read`; việc tìm kiếm cũng được audit.
- License gắn `workspace_id`. Phát hành/kích hoạt/thu hồi qua module licensing; entitlement check không đọc workspace khác. Chi tiết: [licensing.md](../architecture/licensing.md).

## Database

- Thêm file migration mới trong `apps/api/src/database/migrations`. Không sửa migration đã chạy ở môi trường dùng chung.
- Không bật `synchronize`.
- Đổi schema theo hướng mở rộng trước, thu hẹp sau, để `migration:revert` còn đường lui.
- Giao dịch bao các ghi nhiều bản ghi. Side effect ra ngoài process dùng idempotency khi epic đó bắt đầu.
- Tên cột snake_case. Tên property TypeScript camelCase.

## Audit và log

Chỉ `AuditRecorder.append` được ghi `audit_events`. Không thêm method sửa hoặc xóa.

Metadata audit là object nhỏ, JSON được, không vòng. Cấm key mật khẩu, token, cookie, số thẻ và nội dung tài liệu. Không ghi những thứ đó vào log.

Correlation id đi theo response header `x-correlation-id`.

## Office và quyền riêng tư

Không thêm route nhận body tài liệu trong E01. Test và fixture cũng không upload file Office.

Lệnh COM/VSTO chạy trên UI thread của Office. Office.js và VSTO cùng một hình dạng `RuleResult`, không dùng chung binary.

Host không hỗ trợ một lệnh thì adapter nói rõ giới hạn. Không giả bộ đã áp dụng thay đổi.

## Cổng thông tin

Layout gồm sidebar, thanh trên, vùng làm việc và truy cập nhanh. Khung tối thiểu 1100px vì bản đầu là desktop-first.

Điều khiển chưa có hành vi thì để disabled hoặc để chữ "chưa", không dẫn tới màn hình giả lập nghiệp vụ.

## Kiểm thử

- Logic thuần có unit test và chạy bằng `npm test` mà không cần database. API dùng test runner của Node vì NestJS 12 là ESM; Jest trên Node 22 hiện không tải được các package đó.
- Schema và trigger có integration test. CI chạy test đó trên PostgreSQL 16, gồm `migration:run`, `migration:revert`, rồi `migration:run` lại.
- `npm run check:foundation` chặn dependency AI đã biết và mọi chỗ bật `synchronize`.
- Đổi hành vi thì cập nhật test cùng lúc. Vỏ module được phép nhận controller khi epic của nó bắt đầu, và test biên giới phải được sửa có chủ đích.

## Bí mật

`.env` không được commit. Mẫu nằm ở `infrastructure/environments/env.example` và để trống mật khẩu. Workflow CI chỉ chứa mật khẩu của database tạm, sống cùng job.
