const assert = require('node:assert/strict');
const { describe, it } = require('node:test');
const { validateWordDocument } = require('../src/word/validate');
const { planWordFormatting, applyWordPlan } = require('../src/word/preview-apply');

const profile = {
  id: 'p1',
  code: 'word-standard-vi',
  currentVersion: 1,
  version: { id: 'pv1', version: 1 },
  rules: [
    {
      code: 'word.style.font-body',
      version: '1.0.0',
      checkKind: 'style.font',
      defaultSeverity: 'fail',
      checkConfig: { fontNames: ['Times New Roman'], fontSizePt: 13 },
    },
    {
      code: 'word.structure.heading-levels',
      version: '1.0.0',
      checkKind: 'structure.heading',
      defaultSeverity: 'warning',
      checkConfig: { maxSkip: 0 },
    },
  ],
};

describe('word validation', () => {
  it('flags wrong fonts and skipped headings', () => {
    const report = validateWordDocument({
      profile,
      snapshot: {
        fontRuns: [{ location: 'paragraph:1', fontName: 'Arial', fontSizePt: 12 }],
        headings: [
          { location: 'paragraph:2', level: 1 },
          { location: 'paragraph:3', level: 3 },
        ],
      },
    });
    assert.equal(report.profileVersion, 1);
    assert.ok(report.summary.fail >= 1);
    assert.ok(report.summary.warning >= 1);
  });

  it('requires confirmation before high-risk apply', () => {
    const plan = planWordFormatting(profile, {
      fontRuns: [{ location: 'paragraph:1', fontName: 'Arial', fontSizePt: 12 }],
    });
    assert.ok(plan.changes.length >= 1);
    const cancelled = applyWordPlan(plan, { confirmed: false });
    assert.equal(cancelled.cancelled, true);
    const applied = applyWordPlan(plan, { confirmed: true, hostApply: () => undefined });
    assert.equal(applied.applied, true);
  });
});
