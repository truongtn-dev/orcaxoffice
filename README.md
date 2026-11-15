# OrcaX Office Suite 2026

> Hệ thống Chuẩn hóa Thể thức Văn bản & Tự động hóa Văn phòng Việt Nam (Nghị định 30/2020/NĐ-CP & Hướng dẫn 05-HD/VPTW)

## 📌 Tổng quan dự án

**OrcaX Office Suite** là bộ giải pháp toàn diện cho tự động hóa và chuẩn hóa văn bản tại Việt Nam:
- **Chuẩn hóa tự động 1-Click**: Áp dụng triệt để thể thức văn bản hành chính theo Nghị định 30/2020/NĐ-CP và văn bản Đảng theo Hướng dẫn 05-HD/VPTW.
- **Thư viện 59 Biểu mẫu chuẩn Quốc gia**: Đầy đủ các mẫu văn bản cấp Bộ, cấp Tỉnh, cấp Sở, cấp Xã/Phường, văn bản Đảng các cấp, văn bản hợp nhất (Nghị định 250) và giáo án (Công văn 5512).
- **Thanh công cụ Word Add-in (Tab ORCAX-FORMAT)**: Tích hợp trực tiếp vào Microsoft Word với giao diện Fluent Ribbon hiện đại, không sử dụng emoji, icon chuẩn Microsoft Office chính hãng (`imageMso`).
- **Hệ thống kiến trúc Monorepo**: Kiến trúc hoàn chỉnh với Frontend (Next.js/React), Backend API (FastAPI / Node.js), Engine xử lý văn bản, và hệ thống kiểm định quy chuẩn tự động.

---

## 📂 Cấu trúc thư mục

```text
OrcaXOffice/
├── distribution/
│   ├── orcax-commercial-pack/       # Gói thương mại tối giản giao khách hàng
│   │   ├── Cai_Dat_OrcaX_1Click.bat # Cài đặt tự động vào Word & kích hoạt bản quyền
│   │   ├── Go_Cai_Dat_OrcaX.bat     # Gỡ cài đặt an toàn
│   │   ├── KEY_BAN_QUYEN.reg        # File kích hoạt bản quyền vĩnh viễn
│   │   ├── Huong_Dan_Cai_Dat.txt    # Tài liệu hướng dẫn sử dụng
│   │   ├── OrcaX_Addin.dotm         # Add-in Word với Tab ORCAX-FORMAT
│   │   └── templates/               # 59 biểu mẫu văn bản chuẩn Quốc gia
│   └── orcax-office-suite/          # Add-in Word & Excel đầy đủ
├── source/
│   └── orcax-office-suite/          # Monorepo mã nguồn dự án
│       ├── apps/                    # Web App, API Server, Desktop Wrapper
│       ├── packages/                # Core engines, Rulebooks, DTOs
│       ├── scripts/                 # Build scripts & automation tools
│       ├── tests/                   # Automated test suites
│       └── docs/                    # Tài liệu đặc tả kỹ thuật
├── templates/                       # 59 file biểu mẫu chuẩn (.docx)
├── OrcaX_Setup_v1.0.zip             # File zip thương mại sẵn sàng phân phối
└── OrcaX_Office_Suite_Software_Requirements_Specification_v1.0.docx # SRS Specification
```

---

## 🚀 Hướng dẫn Cài đặt & Sử dụng Add-in cho Khách hàng

### Cách 1: Cài đặt tự động (Khuyên dùng)
1. Giải nén file `OrcaX_Setup_v1.0.zip`.
2. Chạy file `Cai_Dat_OrcaX_1Click.bat`.
3. Mở Microsoft Word, tab **ORCAX-FORMAT** sẽ tự động hiển thị trên thanh Ribbon.

### Cách 2: Cài đặt thủ công
1. Nhấn `Windows + R`, gõ `%appdata%\Microsoft\Word\STARTUP` rồi nhấn `Enter`.
2. Copy file `OrcaX_Addin.dotm` vào thư mục vừa mở.
3. Nhấp đúp file `KEY_BAN_QUYEN.reg` để kích hoạt bản quyền vĩnh viễn.
4. Mở Word và sử dụng.

---

## ⚖️ Bản quyền & Hỗ trợ
- **Đơn vị phát triển**: OrcaX Office Suite Development Team
- **Website**: https://orcax.vn
- **Hỗ trợ kỹ thuật**: support@orcax.vn
