# Admin portal (E08)

Cổng quản trị gom tổ chức, thành viên/vai trò, phòng ban-nhóm, hồ sơ, giấy phép và cấu hình workspace. Widget trên `/quan-tri` chỉ hiện theo quyền RBAC hiện tại (deterministic, không AI).

## Tổ chức & thành viên

- Đổi tên / tạm ngưng / lưu trữ tổ chức: chủ sở hữu (`workspace.update` + role owner).
- Thành viên: mời, tạm ngưng, gỡ, đổi vai trò (`workspace.member.*`).
- Phòng ban (cây `parent_id`) và nhóm; gán user vào đích để phân phối hồ sơ.

## Hồ sơ

`POST /profiles/:id/distribute` gắn hồ sơ **đã duyệt** tới department/group. Hồ sơ có distribution chỉ hiện với thành viên thuộc đích (admin/publisher vẫn thấy tất cả).

## Cấu hình

`workspace_settings` + `workspace_setting_history`. Khóa cho phép: `portal.locale`, `portal.timezone`, `notifications.email.enabled`, `security.session.hint`. Mọi thay đổi audit.

## API chính

| Path                                 | Quyền                          |
| ------------------------------------ | ------------------------------ |
| `GET /admin/overview`                | `workspace.read`               |
| `GET/POST /departments`, `/groups`   | `workspace.department.manage`  |
| `POST /organizations/current/status` | `workspace.update` (owner)     |
| `GET/POST /workspace-settings`       | `workspace.config.read/manage` |

Support tickets / help center thuộc E09.
