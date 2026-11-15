# Licensing (E04)

Giấy phép là bản ghi tenant-owned: luôn có `workspace_id`. Personal và organization không chia sẻ license.

## Catalogue

Migration seed:

- Product `orcax-office`
- Plans: `personal-free`, `organization-starter` (subscription), `organization-perpetual`
- Entitlements: `word.core`, `excel.core`, `portal.workspace`

Thanh toán chưa có (E09). Phát hành qua `POST /licenses` với quyền `license.issue`.

## Lifecycle

Trạng thái runtime: `pending`, `active`, `grace`, `expired`, `revoked`. Dùng được khi `active` hoặc `grace`. Thu hồi license cũng thu hồi mọi activation còn sống. Gia hạn chỉ áp dụng subscription.

## Activation

`POST /licenses/activate` gắn `deviceKey` (hash SHA-256) trong giới hạn `device_limit` của plan. Add-in sau này gọi cùng API; E04 chưa ship VSTO/Office.js.

## Kiểm tra quyền lợi

`GET /entitlements/check?code=word.core` trả `allowed` theo license còn dùng được của workspace hiện tại.

## Quyền

`license.read`, `license.issue`, `license.manage`, `license.activate` — gán theo role trong catalog.
