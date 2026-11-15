const { buildValidationReport, observationToResult } = require('@orcax/rule-definitions');

function normalizeFont(name) {
  return String(name || '')
    .trim()
    .toLowerCase();
}

function evaluateStyleFont(rule, snapshot) {
  const config = rule.checkConfig || {};
  const allowed = (config.fontNames || []).map(normalizeFont);
  const expectedSize = config.fontSizePt;
  const observations = [];
  for (const run of snapshot.fontRuns || []) {
    const fontOk = allowed.length === 0 || allowed.includes(normalizeFont(run.fontName));
    const sizeOk = expectedSize == null || Number(run.fontSizePt) === Number(expectedSize);
    if (fontOk && sizeOk) {
      continue;
    }
    observations.push({
      ruleCode: rule.code,
      ruleVersion: rule.version,
      passed: false,
      defaultSeverity: rule.defaultSeverity || 'fail',
      location: run.location || null,
      explanation: `Phông/cỡ chữ không khớp hồ sơ (${run.fontName || '?'} ${run.fontSizePt || '?'}pt).`,
      correction: `Đặt ${(config.fontNames || []).join(' / ')}${expectedSize ? ` ${expectedSize}pt` : ''}.`,
    });
  }
  if (observations.length === 0) {
    observations.push({
      ruleCode: rule.code,
      ruleVersion: rule.version,
      passed: true,
      location: 'document',
      explanation: 'Phông chữ nội dung khớp hồ sơ.',
    });
  }
  return observations;
}

function evaluateHeadingLevels(rule, snapshot) {
  const maxSkip = Number((rule.checkConfig || {}).maxSkip ?? 0);
  const levels = (snapshot.headings || []).map((item) => Number(item.level));
  let previous = 0;
  for (let index = 0; index < levels.length; index += 1) {
    const level = levels[index];
    if (previous > 0 && level > previous + 1 + maxSkip) {
      return [
        {
          ruleCode: rule.code,
          ruleVersion: rule.version,
          passed: false,
          defaultSeverity: rule.defaultSeverity || 'warning',
          location: (snapshot.headings[index] && snapshot.headings[index].location) || null,
          explanation: `Bỏ cấp đề mục từ ${previous} sang ${level}.`,
          correction: 'Chèn đủ cấp đề mục trung gian.',
        },
      ];
    }
    previous = level;
  }
  return [
    {
      ruleCode: rule.code,
      ruleVersion: rule.version,
      passed: true,
      location: 'document',
      explanation: 'Cấp đề mục liên tục.',
    },
  ];
}

function evaluateRule(rule, snapshot) {
  if (rule.checkKind === 'style.font') {
    return evaluateStyleFont(rule, snapshot);
  }
  if (rule.checkKind === 'structure.heading') {
    return evaluateHeadingLevels(rule, snapshot);
  }
  return [
    {
      ruleCode: rule.code,
      ruleVersion: rule.version,
      passed: false,
      defaultSeverity: 'warning',
      location: null,
      explanation: `Loại kiểm tra ${rule.checkKind} chưa được hỗ trợ trên host này.`,
      correction: null,
    },
  ];
}

function validateWordDocument({ profile, snapshot }) {
  if (!profile || !Array.isArray(profile.rules)) {
    throw new TypeError('profile.rules is required');
  }
  const observations = [];
  for (const rule of profile.rules) {
    observations.push(...evaluateRule(rule, snapshot || {}));
  }
  return buildValidationReport({
    profileId: profile.id,
    profileVersionId: profile.version?.id || null,
    profileVersion: profile.currentVersion || profile.version?.version || 1,
    profileCode: profile.code,
    results: observations.map((item) => observationToResult(item)),
  });
}

module.exports = {
  validateWordDocument,
  evaluateRule,
};
