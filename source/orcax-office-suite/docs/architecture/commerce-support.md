# Commerce & Support (E09)

Cung cấp khả năng thương mại (đơn hàng, giá theo gói, thanh toán mock có webhook idempotent, hoàn tiền, xuất hóa đơn và đối soát) cùng hệ thống hỗ trợ kỹ thuật (yêu cầu hỗ trợ, thảo luận và trung tâm trợ giúp có phân loại, phiên bản).

Tất cả xử lý tuân thủ chính sách deterministic, non-AI. Không gửi nội dung tài liệu Office trong các luồng thương mại và hỗ trợ.

## Thương mại (Commerce)

### Danh mục giá

- Gói cá nhân miễn phí: `personal-free` (0 VND).
- Gói tổ chức theo năm: `organization-starter` (2.990.000 VND / năm).
- Gói tổ chức vĩnh viễn: `organization-perpetual` (29.900.000 VND một lần).

### Vòng đời đơn hàng

- `draft` → `pending_payment` → `paid` (hoặc `cancelled`).
- Hoàn tiền: `refunded` (toàn phần) hoặc `partially_refunded`.
- Khi đơn hàng hoàn tất thanh toán (`paid`), hệ thống tự động:
  1. Xuất bản ghi hóa đơn `invoices` với số hóa đơn duy nhất dạng `INV-YYYY-XXXXXX`.
  2. Kích hoạt phát hành giấy phép tương ứng qua `licensing.issueFromPaidOrder`.
  3. Ghi nhật ký kiểm toán `commerce.payment.succeeded`.

### Webhook & Idempotency

- Endpoint `POST /api/v1/payments/webhooks/mock` yêu cầu header xác thực `x-orcax-webhook-secret`.
- Khóa `idempotencyKey` đảm bảo mỗi sự kiện thanh toán từ cổng thanh toán chỉ được xử lý đúng một lần. Sự kiện gửi lại (replay) được phát hiện và bỏ qua an toàn.
- Nếu thanh toán lặp lại cho đơn đã thanh toán hoặc đã hoàn tiền, giao dịch được ghi nhận và đẩy vào hàng đợi đối soát `reconciliation_items` với trạng thái `unmatched`.

### Đối soát & Hoàn tiền

- Nhân viên tài chính hoặc quản trị viên tổ chức có quyền `commerce.reconcile` để xem và giải quyết các mục đối soát (`matched` hoặc `ignored`).
- Quyền `commerce.refund` cho phép hoàn tiền theo thanh toán, không vượt quá tổng số tiền đã thanh toán.

## Hỗ trợ & Trợ giúp (Support & Help Center)

### Phiếu yêu cầu hỗ trợ (Support Tickets)

- Người dùng tạo phiếu hỗ trợ gắn với workspace hiện tại (`POST /api/v1/support/tickets`).
- Danh mục hỗ trợ: `general`, `account`, `billing`, `licensing`, `word_addin`, `excel_addin`.
- Mức độ ưu tiên: `low`, `medium`, `high`, `urgent`.
- Trạng thái: `open` → `in_progress` → `resolved` → `closed`.
- Thảo luận hai chiều giữa người dùng và quản trị viên/nhân viên (`POST /api/v1/support/tickets/:id/comments`).
- Phân quyền: Người dùng thông thường có `support.ticket.read`, `support.ticket.create`; Quản trị viên có thêm `support.ticket.manage` để đổi trạng thái xử lý.

### Trung tâm trợ giúp (Help Center)

- Hướng dẫn và giải đáp thắc mắc theo phiên bản sản phẩm (`product_version`).
- Truy cập công khai đối với người dùng đăng nhập qua `GET /api/v1/support/help/articles`.
- Quản trị viên cập nhật tài liệu qua `POST /api/v1/support/help/articles` với quyền `support.help.manage`.

## API chính

| Method | Path                                   | Quyền                           |
| ------ | -------------------------------------- | ------------------------------- |
| GET    | `/api/v1/plan-prices`                  | `commerce.order.read`           |
| GET    | `/api/v1/orders`                       | `commerce.order.read`           |
| POST   | `/api/v1/orders`                       | `commerce.order.create`         |
| POST   | `/api/v1/orders/:id/finalize`          | `commerce.order.create`         |
| POST   | `/api/v1/payments/checkout`            | `commerce.order.create`         |
| POST   | `/api/v1/payments/webhooks/mock`       | Header `x-orcax-webhook-secret` |
| POST   | `/api/v1/refunds`                      | `commerce.refund`               |
| GET    | `/api/v1/reconciliation-items`         | `commerce.reconcile`            |
| POST   | `/api/v1/reconciliation-items/:id`     | `commerce.reconcile`            |
| GET    | `/api/v1/support/tickets`              | `support.ticket.read`           |
| POST   | `/api/v1/support/tickets`              | `support.ticket.create`         |
| GET    | `/api/v1/support/tickets/:id`          | `support.ticket.read`           |
| POST   | `/api/v1/support/tickets/:id/comments` | `support.ticket.create`         |
| PATCH  | `/api/v1/support/tickets/:id/status`   | `support.ticket.manage`         |
| GET    | `/api/v1/support/help/articles`        | `support.help.read`             |
| GET    | `/api/v1/support/help/articles/:slug`  | `support.help.read`             |
| POST   | `/api/v1/support/help/articles`        | `support.help.manage`           |
