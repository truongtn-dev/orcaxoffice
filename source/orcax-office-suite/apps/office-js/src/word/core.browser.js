(function (root) {
  function createRuleResult(input) {
    return {
      ruleId: input.ruleId,
      severity: input.severity,
      location: input.location == null ? null : String(input.location),
      explanation: input.explanation == null ? null : String(input.explanation),
      correction: input.correction == null ? null : String(input.correction),
    };
  }

  function buildValidationReport(input) {
    const results = input.results || [];
    const summary = { pass: 0, fail: 0, warning: 0, corrected: 0 };
    for (const result of results) {
      if (result.severity === 'pass') summary.pass += 1;
      if (result.severity === 'fail') summary.fail += 1;
      if (result.severity === 'warning') summary.warning += 1;
      if (result.correction) summary.corrected += 1;
    }
    return {
      profileId: input.profileId,
      profileVersionId: input.profileVersionId || null,
      profileVersion: input.profileVersion,
      profileCode: input.profileCode || null,
      generatedAt: new Date().toISOString(),
      summary,
      results,
    };
  }

  function normalizeFont(name) {
    return String(name || '')
      .trim()
      .toLowerCase();
  }

  function observation(rule, passed, extras) {
    const severity = passed ? 'pass' : rule.defaultSeverity || 'fail';
    return createRuleResult({
      ruleId: `${rule.code}@${rule.version}`,
      severity,
      location: extras.location || null,
      explanation: extras.explanation || null,
      correction: extras.correction || null,
    });
  }

  function evaluateRule(rule, snapshot) {
    if (rule.checkKind === 'style.font') {
      const config = rule.checkConfig || {};
      const allowed = (config.fontNames || []).map(normalizeFont);
      const expectedSize = config.fontSizePt;
      const results = [];
      for (const run of snapshot.fontRuns || []) {
        const fontOk = !allowed.length || allowed.includes(normalizeFont(run.fontName));
        const sizeOk = expectedSize == null || Number(run.fontSizePt) === Number(expectedSize);
        if (fontOk && sizeOk) continue;
        results.push(
          observation(rule, false, {
            location: run.location || null,
            explanation: `Phông/cỡ chữ không khớp hồ sơ (${run.fontName || '?'} ${run.fontSizePt || '?'}pt).`,
            correction: `Đặt ${(config.fontNames || []).join(' / ')}${expectedSize ? ` ${expectedSize}pt` : ''}.`,
          }),
        );
      }
      if (!results.length) {
        results.push(
          observation(rule, true, { location: 'document', explanation: 'Phông chữ khớp hồ sơ.' }),
        );
      }
      return results;
    }
    if (rule.checkKind === 'structure.heading') {
      const maxSkip = Number((rule.checkConfig || {}).maxSkip ?? 0);
      let previous = 0;
      const headings = snapshot.headings || [];
      for (let i = 0; i < headings.length; i += 1) {
        const level = Number(headings[i].level);
        if (previous > 0 && level > previous + 1 + maxSkip) {
          return [
            observation(rule, false, {
              location: headings[i].location || null,
              explanation: `Bỏ cấp đề mục từ ${previous} sang ${level}.`,
              correction: 'Chèn đủ cấp đề mục trung gian.',
            }),
          ];
        }
        previous = level;
      }
      return [
        observation(rule, true, { location: 'document', explanation: 'Cấp đề mục liên tục.' }),
      ];
    }
    return [
      observation(rule, false, {
        explanation: `Loại kiểm tra ${rule.checkKind} chưa được hỗ trợ trên host này.`,
      }),
    ];
  }

  function validateWordDocument(profile, snapshot) {
    const results = [];
    for (const rule of profile.rules || []) {
      results.push.apply(results, evaluateRule(rule, snapshot || {}));
    }
    return buildValidationReport({
      profileId: profile.id,
      profileVersionId: profile.version && profile.version.id,
      profileVersion: profile.currentVersion || (profile.version && profile.version.version) || 1,
      profileCode: profile.code,
      results,
    });
  }

  function planWordFormatting(profile, snapshot) {
    const changes = [];
    for (const rule of profile.rules || []) {
      if (rule.checkKind !== 'style.font') continue;
      const config = rule.checkConfig || {};
      const fontName = (config.fontNames || [])[0];
      const fontSizePt = config.fontSizePt;
      const mismatched = (snapshot.fontRuns || []).filter((run) => {
        const fontOk =
          !fontName || String(run.fontName || '').toLowerCase() === String(fontName).toLowerCase();
        const sizeOk = fontSizePt == null || Number(run.fontSizePt) === Number(fontSizePt);
        return !(fontOk && sizeOk);
      });
      if (!mismatched.length) continue;
      changes.push({
        id: 'apply-' + rule.code,
        ruleCode: rule.code,
        risk: 'high',
        summary: `Áp dụng ${fontName || ''} ${fontSizePt || ''}pt cho ${mismatched.length} đoạn.`,
        payload: { fontName: fontName || null, fontSizePt: fontSizePt ?? null },
      });
    }
    return {
      profileId: profile.id,
      profileVersion: profile.currentVersion || 1,
      changes,
    };
  }

  function applyWordPlan(plan, confirmed, hostApply) {
    if (!confirmed && plan.changes.some((item) => item.risk === 'high')) {
      return { applied: false, cancelled: true, message: 'Đã hủy. Tài liệu không đổi.' };
    }
    for (const change of plan.changes) {
      if (typeof hostApply === 'function') hostApply(change);
    }
    return {
      applied: true,
      cancelled: false,
      message: `Đã áp dụng ${plan.changes.length} thay đổi.`,
    };
  }

  root.OrcaXWord = {
    validateWordDocument,
    planWordFormatting,
    applyWordPlan,
  };
})(typeof window !== 'undefined' ? window : globalThis);
