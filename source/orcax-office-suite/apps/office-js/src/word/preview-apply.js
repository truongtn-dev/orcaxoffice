function planWordFormatting(profile, snapshot) {
  const changes = [];
  for (const rule of profile.rules || []) {
    if (rule.checkKind !== 'style.font') {
      continue;
    }
    const config = rule.checkConfig || {};
    const fontName = (config.fontNames || [])[0];
    const fontSizePt = config.fontSizePt;
    if (!fontName && fontSizePt == null) {
      continue;
    }
    const mismatched = (snapshot.fontRuns || []).filter((run) => {
      const fontOk =
        !fontName || String(run.fontName || '').toLowerCase() === String(fontName).toLowerCase();
      const sizeOk = fontSizePt == null || Number(run.fontSizePt) === Number(fontSizePt);
      return !(fontOk && sizeOk);
    });
    if (mismatched.length === 0) {
      continue;
    }
    changes.push({
      id: `apply-${rule.code}`,
      ruleCode: rule.code,
      risk: 'high',
      summary: `Áp dụng ${fontName || ''} ${fontSizePt || ''}pt cho ${mismatched.length} đoạn.`,
      payload: {
        fontName: fontName || null,
        fontSizePt: fontSizePt ?? null,
        locations: mismatched.map((item) => item.location).filter(Boolean),
      },
    });
  }
  return {
    profileId: profile.id,
    profileVersion: profile.currentVersion || profile.version?.version || 1,
    changes,
    canAutoApply: changes.every((item) => item.risk === 'low'),
  };
}

function applyWordPlan(plan, { confirmed, hostApply }) {
  if (!plan) {
    throw new TypeError('plan is required');
  }
  if (!confirmed && plan.changes.some((item) => item.risk === 'high')) {
    return {
      applied: false,
      cancelled: true,
      appliedChangeIds: [],
      message: 'Đã hủy. Tài liệu không đổi.',
    };
  }
  const appliedChangeIds = [];
  for (const change of plan.changes) {
    if (typeof hostApply === 'function') {
      hostApply(change);
    }
    appliedChangeIds.push(change.id);
  }
  return {
    applied: true,
    cancelled: false,
    appliedChangeIds,
    message: `Đã áp dụng ${appliedChangeIds.length} thay đổi.`,
  };
}

module.exports = { planWordFormatting, applyWordPlan };
