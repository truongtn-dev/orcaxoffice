# OrcaX Office Suite

OrcaX Office Suite: cổng thông tin Next.js, API NestJS dạng modular monolith, PostgreSQL với TypeORM, và hợp đồng dùng chung cho add-in Word/Excel. Đã hoàn thiện toàn bộ các epic MVP P0/Beta: identity, workspace, phân quyền/audit, licensing, profile & rule engine, Word/Excel Office.js MVP, admin portal, commerce (đơn hàng, bảng giá, thanh toán mock, đối soát), support (phiếu hỗ trợ, trung tâm trợ giúp), và bộ kiểm thử/runbook sẵn sàng cho Beta.

Sản phẩm tuân thủ nghiêm ngặt chính sách không dùng AI (Non-AI). Nội dung Word và Excel được xử lý cục bộ trên máy người dùng, không tải lên máy chủ theo mặc định.

## Yêu cầu

- Node.js 22 trở lên
- PostgreSQL 16 cho migration và API
- Docker là một cách chạy PostgreSQL cục bộ, không bắt buộc nếu đã có máy chủ cơ sở dữ liệu

## Cài đặt

Từ thư mục này:

```powershell
Copy-Item infrastructure\environments\env.example .env
npm install
```

Mở `.env` và đặt `DATABASE_PASSWORD` cùng `AUTH_TOKEN_SECRET` (ít nhất 32 ký tự, mỗi môi trường một giá trị). Ứng dụng không khởi động khi thiếu các giá trị này. `EMAIL_TRANSPORT=log` chỉ dùng khi phát triển; production phải dùng `smtp`. Không commit file `.env`.

Nếu dùng Docker:

```powershell
docker compose --env-file .env -f infrastructure\deployment\docker-compose.yml up -d
npm run migration:run
```

Nếu PostgreSQL cài trực tiếp trên máy, trỏ các biến `DATABASE_*` tới instance đó rồi chạy cùng lệnh migration.

## Chạy

Hai terminal, cùng thư mục repo:

```powershell
npm run dev:api
npm run dev:web
```

- API: `http://localhost:3001`
- Cổng thông tin: `http://localhost:3000`
- Sống: `GET /health`
- Sẵn sàng, có kiểm tra database: `GET /health/ready`
- API nghiệp vụ sau này nằm dưới `/api/v1`

## Kiểm tra

```powershell
npm run check:foundation
npm run format:check
npm run lint
npm test
npm run build
```

`npm test` không cần PostgreSQL. `npm run test:integration` chạy migration trên database đã cấu hình, kiểm tra audit chỉ được thêm bản ghi, ranh giới workspace cá nhân/tổ chức, và luồng đăng ký → xác minh → chuyển workspace → membership.

Khi phát triển email, `EMAIL_TRANSPORT=log` và `GET /api/v1/dev/mailbox?email=` đọc mã xác minh trong process. Production bắt buộc `EMAIL_TRANSPORT=smtp` cùng `SMTP_HOST` / `SMTP_PORT`.

## Cấu trúc

```text
apps/api            NestJS modular monolith
apps/web-portal     Next.js, JavaScript, tiếng Việt
apps/office-vsto    Contract C# VSTO (scaffold)
apps/office-js      Office.js Word/Excel MVP + manifest
packages/api-contracts
packages/rule-definitions
packages/profile-specifications
infrastructure      Mẫu môi trường, PostgreSQL, ghi chú log
docs                Kiến trúc và quy ước
```

`rule-definitions` và `profile-specifications` phục vụ E05 (rule packs + gói hồ sơ `orcax.profile.v1`). `test-fixtures` vẫn giữ chỗ.

## Phạm vi mở rộng sau Beta

Các tính năng nâng cao như trình dựng trực quan quy trình nghiệp vụ (visual workflow builder), biểu mẫu động phức tạp (advanced dynamic forms), quản lý kho nhị phân tài liệu (DMS binary repository) và tích hợp nhà cung cấp chữ ký số được hoãn sang giai đoạn sau Beta theo lộ trình SRS. Toàn bộ các module MVP P0 (E01 đến E10) đã hoàn thiện. `WEB_ORIGIN` phải trùng origin của cổng thông tin thì cookie phiên mới được trình duyệt gửi.
