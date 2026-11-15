const assert = require('node:assert/strict');
const { describe, it } = require('node:test');
const {
  CURATED_RULES,
  assertRuleDefinition,
  buildValidationReport,
  observationToResult,
  ruleId,
} = require('./index');

describe('rule definitions', () => {
  it('freezes curated deterministic rules', () => {
    assert.ok(CURATED_RULES.length >= 3);
    assert.equal(CURATED_RULES[0].code, 'word.style.font-body');
    assert.throws(() => assertRuleDefinition({ code: 'x' }), /version/);
  });

  it('builds a stamped validation report', () => {
    const report = buildValidationReport({
      profileId: 'p1',
      profileVersion: 2,
      profileVersionId: 'pv1',
      profileCode: 'word-standard-vi',
      results: [
        observationToResult({
          ruleCode: 'word.style.font-body',
          ruleVersion: '1.0.0',
          passed: false,
          location: 'paragraph:3',
          explanation: 'Sai phông',
          correction: 'Đặt Times New Roman 13pt',
        }),
        {
          ruleId: ruleId('word.structure.heading-levels', '1.0.0'),
          severity: 'pass',
        },
      ],
    });
    assert.equal(report.summary.fail, 1);
    assert.equal(report.summary.pass, 1);
    assert.equal(report.summary.corrected, 1);
    assert.equal(report.profileVersion, 2);
  });
});
