export const WORD_CAPABILITIES = Object.freeze({
  host: 'word' as const,
  entitlement: 'word.core',
  requirementSets: Object.freeze(['WordApi 1.3']),
  actions: Object.freeze([
    Object.freeze({
      id: 'select-profile',
      label: 'Chọn hồ sơ',
      risk: 'low',
      supported: true,
    }),
    Object.freeze({
      id: 'validate',
      label: 'Kiểm tra theo hồ sơ',
      risk: 'low',
      supported: true,
    }),
    Object.freeze({
      id: 'preview-apply',
      label: 'Xem trước và áp dụng',
      risk: 'high',
      supported: true,
      requiresConfirmation: true,
    }),
    Object.freeze({
      id: 'validation-report',
      label: 'Báo cáo kết quả',
      risk: 'low',
      supported: true,
    }),
  ]),
  notes: Object.freeze([
    'Nội dung tài liệu xử lý cục bộ trong host Office.',
    'API chỉ nhận metadata báo cáo (số pass/fail), không nhận body tài liệu.',
  ]),
});

export const EXCEL_CAPABILITIES = Object.freeze({
  host: 'excel' as const,
  entitlement: 'excel.core',
  requirementSets: Object.freeze(['ExcelApi 1.4']),
  actions: Object.freeze([
    Object.freeze({
      id: 'select-print-profile',
      label: 'Chọn hồ sơ in',
      risk: 'low',
      supported: true,
    }),
    Object.freeze({
      id: 'apply-a4',
      label: 'Áp dụng khổ A4',
      risk: 'high',
      supported: true,
      requiresConfirmation: true,
    }),
    Object.freeze({
      id: 'print-preview',
      label: 'Xem trước in',
      risk: 'low',
      supported: true,
    }),
    Object.freeze({
      id: 'export-pdf',
      label: 'Xuất PDF',
      risk: 'low',
      supported: true,
      hostDependent: true,
    }),
  ]),
  notes: Object.freeze([
    'Workbook xử lý cục bộ; hủy xem trước không đổi workbook.',
    'Xuất PDF phụ thuộc host (Desktop hỗ trợ tốt hơn Web).',
  ]),
});
