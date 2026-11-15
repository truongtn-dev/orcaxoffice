(function (root) {
  function readPrintRule(profile) {
    const rule = (profile.rules || []).find((item) => item.checkKind === 'print.page');
    if (!rule) {
      return {
        paperSize: 'A4',
        orientation: 'portrait',
        marginMm: { top: 15, bottom: 15, left: 12, right: 12 },
      };
    }
    const config = rule.checkConfig || {};
    return {
      paperSize: config.paperSize || 'A4',
      orientation: config.orientation || 'portrait',
      marginMm: {
        top: (config.marginMm && config.marginMm.top) || 15,
        bottom: (config.marginMm && config.marginMm.bottom) || 15,
        left: (config.marginMm && config.marginMm.left) || 12,
        right: (config.marginMm && config.marginMm.right) || 12,
      },
    };
  }

  function inspectExcelPrint(profile, snapshot) {
    const expected = readPrintRule(profile);
    const actual = (snapshot && snapshot.pageSetup) || {};
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
    ['top', 'bottom', 'left', 'right'].forEach((edge) => {
      const wanted = expected.marginMm[edge];
      const got = actual.marginMm ? actual.marginMm[edge] : null;
      if (got == null || Number(got) < Number(wanted)) {
        mismatches.push({ field: 'margin.' + edge, expected: wanted, actual: got });
      }
    });
    return {
      expected,
      mismatches,
      fitsA4: mismatches.length === 0,
      profileVersion: profile.currentVersion || 1,
    };
  }

  function planExcelPrint(profile, snapshot) {
    const inspection = inspectExcelPrint(profile, snapshot);
    if (inspection.fitsA4) {
      return {
        changes: [],
        preview: inspection.expected,
        profileVersion: inspection.profileVersion,
      };
    }
    return {
      profileVersion: inspection.profileVersion,
      preview: inspection.expected,
      changes: [
        {
          id: 'apply-a4-print',
          risk: 'high',
          summary: `Áp dụng ${inspection.expected.paperSize} ${inspection.expected.orientation}.`,
          payload: inspection.expected,
        },
      ],
    };
  }

  function applyExcelPrintPlan(plan, confirmed, hostApply) {
    if (!confirmed && plan.changes.some((item) => item.risk === 'high')) {
      return { applied: false, cancelled: true, message: 'Đã hủy. Workbook không đổi.' };
    }
    if (plan.changes.length && typeof hostApply === 'function') {
      hostApply(plan.preview);
    }
    return {
      applied: plan.changes.length > 0,
      cancelled: false,
      message: plan.changes.length
        ? 'Đã áp dụng cấu hình in theo hồ sơ.'
        : 'Trang in đã khớp hồ sơ.',
    };
  }

  root.OrcaXExcel = {
    inspectExcelPrint,
    planExcelPrint,
    applyExcelPrintPlan,
  };
})(typeof window !== 'undefined' ? window : globalThis);
