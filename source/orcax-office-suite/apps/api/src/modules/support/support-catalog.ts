export const SUPPORT_CATEGORIES = [
  { code: 'general', name: 'Chung' },
  { code: 'account', name: 'Tài khoản & Workspace' },
  { code: 'billing', name: 'Thanh toán & Đơn hàng' },
  { code: 'licensing', name: 'Giấy phép & Kích hoạt' },
  { code: 'word_addin', name: 'Add-in Word' },
  { code: 'excel_addin', name: 'Add-in Excel' },
] as const;

export const INITIAL_HELP_ARTICLES = [
  {
    id: '34000000-0000-4000-8000-000000000001',
    slug: 'bat-dau-voi-orcax-office',
    title: 'Bắt đầu với OrcaX Office Suite',
    category: 'general',
    content:
      'OrcaX Office Suite là hệ sinh thái tự động hóa quy trình nghiệp vụ và năng suất văn phòng tiếng Việt. Sản phẩm gồm các add-in cho Microsoft Word và Excel kết hợp với cổng thông tin quản trị tập trung. Mọi quy tắc xử lý tuân thủ chính sách non-AI: chuẩn hóa tài liệu và bảng tính theo hồ sơ quy định, không gửi nội dung tệp lên máy chủ theo mặc định.',
    productVersion: '1.0',
    status: 'published' as const,
  },
  {
    id: '34000000-0000-4000-8000-000000000002',
    slug: 'huong-dan-cai-dat-add-in-word-excel',
    title: 'Hướng dẫn cài đặt Add-in Word và Excel',
    category: 'word_addin',
    content:
      'Để sử dụng Add-in Office.js trên Microsoft Word và Excel, tải tệp manifest từ mục Add-in trên cổng thông tin, sau đó thực hiện sideload vào ứng dụng Word/Excel trên máy tính. Đăng nhập bằng tài khoản OrcaX để kích hoạt và đồng bộ hồ sơ định dạng.',
    productVersion: '1.0',
    status: 'published' as const,
  },
  {
    id: '34000000-0000-4000-8000-000000000003',
    slug: 'quan-ly-khong-gian-lam-viec-va-giay-phep',
    title: 'Quản lý Không gian làm việc và Giấy phép',
    category: 'licensing',
    content:
      'Mỗi người dùng có một Không gian cá nhân (Personal workspace). Khi tạo tổ chức, một Không gian tổ chức (Organization workspace) được tạo riêng biệt. Giấy phép gắn theo từng workspace. Để kích hoạt thiết bị, sử dụng mã kích hoạt hoặc đăng nhập trực tiếp từ taskpane của Add-in.',
    productVersion: '1.0',
    status: 'published' as const,
  },
  {
    id: '34000000-0000-4000-8000-000000000004',
    slug: 'chuan-in-an-a4-tren-excel',
    title: 'Tối ưu chuẩn in ấn A4 trên Excel',
    category: 'excel_addin',
    content:
      'Add-in Excel của OrcaX cung cấp tính năng chuẩn hóa trang in A4: tự động căn lề, điều chỉnh vùng in, xem trước trang in và xuất PDF trực tiếp theo quy chuẩn văn bản hành chính Việt Nam.',
    productVersion: '1.0',
    status: 'published' as const,
  },
] as const;
