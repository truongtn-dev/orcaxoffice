# Authorization & Audit (E03)

E03 hoàn thiện lớp quyền trên nền tenancy E02.

## RBAC

Vai trò và quyền nằm trong `apps/api/src/modules/authorization/catalog.ts` và được seed bằng migration. Runtime kiểm tra qua `can(role, permission)` và `AuthorizationService.assert`.

## ABAC

`evaluatePolicy` / `assertPolicy` thêm thuộc tính:

- `workspaceKinds` — personal hoặc organization
- `membershipStatus` / `workspaceStatus` — mặc định phải `active`
- `ownership: 'self'` — actor phải trùng `resourceOwnerId`

Decorator `@RequirePermission(...)` gắn lên handler. `PermissionsGuard` (APP_GUARD) đọc metadata sau khi `AccessGuard` đã gắn `AccessContext`.

## Audit đọc

`GET /api/v1/audit-events` chỉ trả sự kiện có `workspace_id` bằng workspace đang hoạt động. Cần `workspace.audit.read`. Mỗi lần tìm kiếm ghi `audit.searched`.

Không có API sửa/xóa audit. Bảng vẫn append-only.

## Ngoài phạm vi E03

- Department / group hierarchy
- ABAC theo sensitivity hoặc workflow state (chưa có module đó)
- License entitlement (E04)
