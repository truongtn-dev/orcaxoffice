function readPrintRule(profile) {
  const rule = (profile.rules || []).find((item) => item.checkKind === 'print.page');
  if (!rule) {
    return {
      paperSize: 'A4',
      orientation: 'portrait',
      marginMm: { top: 15, bottom: 15, left: 12, right: 12 },
      ruleCode: null,
      ruleVersion: null,
    };
  }
  const config = rule.checkConfig || {};
  return {
    paperSize: config.paperSize || 'A4',
    orientation: config.orientation || 'portrait',
    marginMm: {
      top: config.marginMm?.top ?? 15,
      bottom: config.marginMm?.bottom ?? 15,
      left: config.marginMm?.left ?? 12,
      right: config.marginMm?.right ?? 12,
    },
    ruleCode: rule.code,
    ruleVersion: rule.version,
  };
}

function inspectExcelPrint(profile, snapshot) {
  const expected = readPrintRule(profile);
  const actual = snapshot.pageSetup || {};
  const mismatches = [];
  if (String(actual.paperSize || '').toUpperCase() !== String(expected.paperSize).toUpperCase()) {
    mismatches.push({
      field: 'paperSize',
      expected: expected.paperSize,
      actual: actual.paperSize || null,
    });
  }
  if (
    String(actual.orientation || '').toLowerCase() !== String(expected.orientation).toLowerCase()
  ) {
    mismatches.push({
      field: 'orientation',
      expected: expected.orientation,
      actual: actual.orientation || null,
    });
  }
  for (const edge of ['top', 'bottom', 'left', 'right']) {
    const wanted = expected.marginMm[edge];
    const got = actual.marginMm ? actual.marginMm[edge] : null;
    if (got == null || Number(got) < Number(wanted)) {
      mismatches.push({
        field: `margin.${edge}`,
        expected: wanted,
        actual: got,
      });
    }
  }
  return {
    profileId: profile.id,
    profileVersion: profile.currentVersion || profile.version?.version || 1,
    expected,
    mismatches,
    fitsA4: mismatches.length === 0,
  };
}

function planExcelPrint(profile, snapshot) {
  const inspection = inspectExcelPrint(profile, snapshot);
  if (inspection.fitsA4) {
    return {
      profileId: profile.id,
      profileVersion: inspection.profileVersion,
      changes: [],
      preview: inspection.expected,
      canAutoApply: true,
    };
  }
  return {
    profileId: profile.id,
    profileVersion: inspection.profileVersion,
    changes: [
      {
        id: 'apply-a4-print',
        risk: 'high',
        summary: `Áp dụng ${inspection.expected.paperSize} ${inspection.expected.orientation}.`,
        payload: inspection.expected,
      },
    ],
    preview: inspection.expected,
    canAutoApply: false,
  };
}

function applyExcelPrintPlan(plan, { confirmed, hostApply }) {
  if (!confirmed && plan.changes.some((item) => item.risk === 'high')) {
    return {
      applied: false,
      cancelled: true,
      settings: null,
      message: 'Đã hủy. Workbook không đổi.',
    };
  }
  const settings = plan.preview;
  if (typeof hostApply === 'function' && plan.changes.length > 0) {
    hostApply(settings);
  }
  return {
    applied: plan.changes.length > 0,
    cancelled: false,
    settings,
    message:
      plan.changes.length === 0 ? 'Trang in đã khớp hồ sơ.' : 'Đã áp dụng cấu hình in theo hồ sơ.',
  };
}

function canExportPdf(capabilities) {
  const action = (capabilities?.actions || []).find((item) => item.id === 'export-pdf');
  return Boolean(action && action.supported);
}

module.exports = {
  readPrintRule,
  inspectExcelPrint,
  planExcelPrint,
  applyExcelPrintPlan,
  canExportPdf,
};
