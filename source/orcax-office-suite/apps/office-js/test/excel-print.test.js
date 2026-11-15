const assert = require('node:assert/strict');
const { describe, it } = require('node:test');
const { inspectExcelPrint, planExcelPrint, applyExcelPrintPlan } = require('../src/excel/print');

const profile = {
  id: 'p2',
  code: 'excel-print-a4',
  currentVersion: 1,
  rules: [
    {
      code: 'excel.print.a4-portrait',
      version: '1.0.0',
      checkKind: 'print.page',
      defaultSeverity: 'fail',
      checkConfig: {
        paperSize: 'A4',
        orientation: 'portrait',
        marginMm: { top: 15, bottom: 15, left: 12, right: 12 },
      },
    },
  ],
};

describe('excel print', () => {
  it('detects A4 mismatches and cancels without confirmation', () => {
    const snapshot = {
      pageSetup: {
        paperSize: 'Letter',
        orientation: 'landscape',
        marginMm: { top: 5, bottom: 5, left: 5, right: 5 },
      },
    };
    const inspection = inspectExcelPrint(profile, snapshot);
    assert.equal(inspection.fitsA4, false);
    const plan = planExcelPrint(profile, snapshot);
    assert.equal(plan.preview.paperSize, 'A4');
    const cancelled = applyExcelPrintPlan(plan, { confirmed: false });
    assert.equal(cancelled.cancelled, true);
    const applied = applyExcelPrintPlan(plan, {
      confirmed: true,
      hostApply: () => undefined,
    });
    assert.equal(applied.applied, true);
  });
});
