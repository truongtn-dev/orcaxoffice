# Hợp đồng tenancy (E02+)

Biên giới tenant của OrcaX Office là `workspace_id`. Không dùng `organization_id` nullable làm khóa phạm vi cho bảng nghiệp vụ.

## Hai loại workspace

| Thuộc tính        | Personal                    | Organization                      |
| ----------------- | --------------------------- | --------------------------------- |
| `kind`            | `personal`                  | `organization`                    |
| `owner_user_id`   | bắt buộc, duy nhất mỗi user | null                              |
| `organization_id` | null                        | bắt buộc                          |
| `status`          | luôn `active`               | `active`, `suspended`, `archived` |

Constraint database giữ hai dạng này tách nhau. Không được “đổi kind” của một hàng workspace.

## Phiên và token

- Đăng nhập luôn kích hoạt workspace cá nhân.
- `users.last_workspace_id` chỉ là gợi ý (`suggestedWorkspace`), không tự chuyển.
- Đổi workspace: `POST /api/v1/workspaces/switch` với `{ workspaceId, confirmed: true }`.
- Access token mang `wid`. Guard từ chối nếu `wid` khác `sessions.active_workspace_id`.
- Cookie `orcax_session` là refresh HttpOnly. Bearer access token sống ngắn.

## Quy tắc cho module sau (E03+)

1. Mọi bảng tenant-owned có cột `workspace_id uuid NOT NULL` và index theo `workspace_id`.
2. Service lấy `workspaceId` từ `AccessContext` (`requireWorkspaceId(access)`), không từ body/query do client tự khai.
3. Đọc/ghi kiểm tra bằng `assertRecordInWorkspace` hoặc lọc bằng `filterByWorkspace`.
4. Dữ liệu personal và organization là hai tập hàng khác nhau. Không join “cùng user nên thấy hết”.
5. Tạo tổ chức chỉ khi actor đang ở personal workspace active.
6. Invite, đổi vai trò, tạm ngưng thành viên chỉ trên organization workspace.

Import helper từ `apps/api/src/tenancy`.
