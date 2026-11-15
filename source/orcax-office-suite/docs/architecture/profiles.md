# Profiles & rules (E05)

Hồ sơ (profile) và mẫu (template) là nền tảng deterministic cho Word/Excel. Thực thi Office Object Model nằm ở add-in; API chỉ giữ metadata, vòng đời, trích dẫn pháp lý và hợp đồng kết quả.

## Phạm vi hồ sơ

- `system`: catalogue dùng chung, seed migration, chỉ đọc từ API.
- `personal` / `organization`: gắn `workspace_id` của không gian hiện tại.
- Host: `word` | `excel`.

## Vòng đời hồ sơ

`draft` → `in_review` → `approved` → `deprecated`.

Phiên bản đã duyệt là immutable. Tạo phiên bản mới copy rule + legal refs (refs trở lại unverified).

## Cổng xác minh pháp lý

Không duyệt (`approved`) khi còn `legal_references.verified = false`. Xác minh ghi `verification_evidence`.

## Rule & báo cáo

Package `@orcax/rule-definitions` cung cấp curated rules và `buildValidationReport`. `POST /profiles/:id/evaluate` đóng dấu `profileVersion` lên kết quả; add-in gửi observation/result, không upload nội dung tài liệu mặc định.

Import/export dùng format `orcax.profile.v1` (`@orcax/profile-specifications`).

## Template

Vòng đời: `draft` → `review` → `approved` → `published` → `deprecated`. Bản ghi `field_definitions` và `data_mappings` deterministic; `GET /templates/:id/mapping-preview` cho xem nguồn → đích → transform. Chỉ `approved`/`published` dùng cho thao tác governed.

## Quyền

`profile.read|write|review|publish`, `template.read|write|publish`.

## Phân phối (E08)

`profile_distributions` gắn hồ sơ đã duyệt tới department/group. Thành viên chỉ thấy hồ sơ được phân phối cho đích mình thuộc về (trừ `profile.publish`).

Thực thi rule Office nằm ở add-in (E06/E07).
