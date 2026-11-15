const assert = require('node:assert/strict');
const { describe, it } = require('node:test');

const {
  API_VERSION,
  CORRELATION_ID_HEADER,
  success,
  failure,
  createRuleResult,
  RuleSeverities,
} = require('../../packages/api-contracts/src');

const { buildValidationReport, CURATED_RULES } = require('../../packages/rule-definitions/src');

const {
  PLAN_PRICES,
  calcLineTotal,
  calcOrderTotal,
} = require('../../apps/api/dist/modules/commerce/commerce-catalog');

const {
  SUPPORT_CATEGORIES,
  INITIAL_HELP_ARTICLES,
} = require('../../apps/api/dist/modules/support/support-catalog');

describe('e2e product smoke suite', () => {
  it('enforces API response contracts and envelope conventions', () => {
    const correlationId = '11111111-2222-3333-4444-555555555555';
    const payload = success(
      { status: 'ok', active: true },
      { correlationId, apiVersion: API_VERSION },
    );

    assert.equal(payload.error, null);
    assert.deepEqual(payload.data, { status: 'ok', active: true });
    assert.equal(payload.meta.correlationId, correlationId);
    assert.equal(payload.meta.apiVersion, 'v1');

    const failPayload = failure(
      { code: 'BAD_REQUEST', message: 'Dữ liệu không hợp lệ.' },
      { correlationId, apiVersion: API_VERSION },
    );
    assert.equal(failPayload.data, null);
    assert.equal(failPayload.error.code, 'BAD_REQUEST');
    assert.equal(failPayload.error.message, 'Dữ liệu không hợp lệ.');
  });

  it('generates stamped deterministic validation reports for Office add-ins', () => {
    const results = [
      createRuleResult({
        ruleId: 'word.style.font-body@1.0.0',
        severity: RuleSeverities.FAIL,
        location: 'Paragraph 1',
        explanation: 'Font Arial không đúng chuẩn Times New Roman',
        correction: 'Chuyển sang Times New Roman',
      }),
      createRuleResult({
        ruleId: 'word.structure.heading-levels@1.0.0',
        severity: RuleSeverities.PASS,
      }),
    ];

    const report = buildValidationReport({
      profileId: 'prof-001',
      profileVersion: 1,
      host: 'word',
      results,
    });

    assert.equal(report.profileId, 'prof-001');
    assert.equal(report.profileVersion, 1);
    assert.equal(report.summary.fail, 1);
    assert.equal(report.summary.pass, 1);
    assert.equal(report.summary.warning, 0);
    assert.equal(report.results.length, 2);
    assert.ok(report.generatedAt);
  });

  it('simulates commerce order creation and price calculation end-to-end', () => {
    const starterPlan = PLAN_PRICES.find((p) => p.planCode === 'organization-starter');
    assert.ok(starterPlan);

    const lineTotal = calcLineTotal(starterPlan.amountMinor, 2);
    assert.equal(lineTotal, 5_980_000);

    const orderTotal = calcOrderTotal([{ lineTotalMinor: lineTotal }]);
    assert.equal(orderTotal, 5_980_000);
  });

  it('provides support catalog categories and versioned help articles', () => {
    assert.ok(SUPPORT_CATEGORIES.length >= 6);
    assert.ok(INITIAL_HELP_ARTICLES.length >= 4);

    const generalArticle = INITIAL_HELP_ARTICLES.find((a) => a.slug === 'bat-dau-voi-orcax-office');
    assert.ok(generalArticle);
    assert.equal(generalArticle.status, 'published');
    assert.equal(generalArticle.productVersion, '1.0');
  });
});
