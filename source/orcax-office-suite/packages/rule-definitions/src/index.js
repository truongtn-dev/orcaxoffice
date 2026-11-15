const { createRuleResult, RuleSeverities } = require('@orcax/api-contracts');

const RULE_PACK_VERSION = '1.0.0';

const CURATED_RULES = Object.freeze([
  Object.freeze({
    code: 'word.style.font-body',
    version: '1.0.0',
    host: 'word',
    title: 'Phông chữ nội dung',
    description: 'Nội dung chính dùng Times New Roman 13pt hoặc tương đương đã cấu hình.',
    checkKind: 'style.font',
    defaultSeverity: 'fail',
    checkConfig: Object.freeze({
      fontNames: Object.freeze(['Times New Roman', 'Times']),
      fontSizePt: 13,
    }),
  }),
  Object.freeze({
    code: 'word.structure.heading-levels',
    version: '1.0.0',
    host: 'word',
    title: 'Cấp đề mục liên tục',
    description: 'Không bỏ cấp đề mục khi đi từ Heading 1 xuống Heading 3.',
    checkKind: 'structure.heading',
    defaultSeverity: 'warning',
    checkConfig: Object.freeze({ maxSkip: 0 }),
  }),
  Object.freeze({
    code: 'excel.print.a4-portrait',
    version: '1.0.0',
    host: 'excel',
    title: 'Khổ in A4 dọc',
    description: 'Trang in đặt khổ A4, hướng dọc, lề tối thiểu theo hồ sơ.',
    checkKind: 'print.page',
    defaultSeverity: 'fail',
    checkConfig: Object.freeze({
      paperSize: 'A4',
      orientation: 'portrait',
      marginMm: Object.freeze({ top: 15, bottom: 15, left: 12, right: 12 }),
    }),
  }),
]);

function assertRuleDefinition(input) {
  if (!input || typeof input.code !== 'string' || input.code.trim() === '') {
    throw new TypeError('code is required');
  }
  if (typeof input.version !== 'string' || input.version.trim() === '') {
    throw new TypeError('version is required');
  }
  if (input.host !== 'word' && input.host !== 'excel') {
    throw new TypeError('host must be word or excel');
  }
  if (typeof input.checkKind !== 'string' || input.checkKind.trim() === '') {
    throw new TypeError('checkKind is required');
  }
  if (
    input.defaultSeverity !== RuleSeverities.FAIL &&
    input.defaultSeverity !== RuleSeverities.WARNING &&
    input.defaultSeverity !== 'fail' &&
    input.defaultSeverity !== 'warning'
  ) {
    throw new TypeError('defaultSeverity must be fail or warning');
  }
  return Object.freeze({
    code: input.code,
    version: input.version,
    host: input.host,
    title: String(input.title || input.code),
    description: String(input.description || ''),
    checkKind: input.checkKind,
    defaultSeverity: input.defaultSeverity,
    checkConfig:
      input.checkConfig && typeof input.checkConfig === 'object' ? input.checkConfig : {},
  });
}

function ruleId(code, version) {
  return `${code}@${version}`;
}

function observationToResult(observation) {
  if (!observation || typeof observation.ruleCode !== 'string') {
    throw new TypeError('ruleCode is required');
  }
  const version = typeof observation.ruleVersion === 'string' ? observation.ruleVersion : '1.0.0';
  let severity = observation.severity;
  if (!severity) {
    if (observation.passed === true) {
      severity = RuleSeverities.PASS;
    } else if (observation.passed === false) {
      severity = observation.defaultSeverity || RuleSeverities.FAIL;
    } else {
      throw new TypeError('severity or passed is required');
    }
  }
  return createRuleResult({
    ruleId: ruleId(observation.ruleCode, version),
    severity,
    location: observation.location,
    explanation: observation.explanation,
    correction: observation.correction,
  });
}

function buildValidationReport(input) {
  if (!input || typeof input.profileId !== 'string' || typeof input.profileVersion !== 'number') {
    throw new TypeError('profileId and profileVersion are required');
  }
  const results = Array.isArray(input.results)
    ? input.results.map((item) =>
        item.ruleId ? createRuleResult(item) : observationToResult(item),
      )
    : [];
  const summary = {
    pass: 0,
    fail: 0,
    warning: 0,
    corrected: 0,
  };
  for (const result of results) {
    if (result.severity === RuleSeverities.PASS) {
      summary.pass += 1;
    } else if (result.severity === RuleSeverities.FAIL) {
      summary.fail += 1;
    } else if (result.severity === RuleSeverities.WARNING) {
      summary.warning += 1;
    }
    if (result.correction) {
      summary.corrected += 1;
    }
  }
  return Object.freeze({
    profileId: input.profileId,
    profileVersionId: input.profileVersionId || null,
    profileVersion: input.profileVersion,
    profileCode: input.profileCode || null,
    generatedAt: input.generatedAt || new Date().toISOString(),
    summary: Object.freeze(summary),
    results: Object.freeze(results),
  });
}

module.exports = {
  RULE_PACK_VERSION,
  CURATED_RULES,
  assertRuleDefinition,
  ruleId,
  observationToResult,
  buildValidationReport,
};
