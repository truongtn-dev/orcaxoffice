# Word & Excel add-ins (E06 / E07)

Add-in thực thi trên host Office. Nest chỉ cung cấp bootstrap (hồ sơ đã duyệt, entitlement, capability matrix) và ghi metadata lần chạy.

## Local-first

- Không có route nhận body tài liệu/workbook.
- `POST /add-in/word/validation-runs` chỉ lưu summary pass/fail/warning/corrected + profile version.
- `POST /add-in/excel/print-runs` chỉ lưu paper/orientation + cờ applied/pdfExported.

## API

| Method | Path                           | Quyền             |
| ------ | ------------------------------ | ----------------- |
| GET    | `/add-in/word/bootstrap`       | `addin.word.use`  |
| POST   | `/add-in/word/validation-runs` | `addin.word.use`  |
| GET    | `/add-in/excel/bootstrap`      | `addin.excel.use` |
| POST   | `/add-in/excel/print-runs`     | `addin.excel.use` |

Entitlement: `word.core` / `excel.core`.

## Office.js

`apps/office-js` — manifest Ribbon → task pane tiếng Việt. Validate/print planner chạy cục bộ; xác nhận trước thay đổi high-risk.

## VSTO

`apps/office-vsto` giữ contract C# cùng hình dạng kết quả. Project designer đầy đủ chờ Windows build; MVP chạy được là Office.js.
