const assert = require('node:assert/strict');
const { describe, it } = require('node:test');

const {
  PLAN_PRICES,
  calcLineTotal,
  calcOrderTotal,
} = require('../../apps/api/dist/modules/commerce/commerce-catalog');

const { SUPPORT_CATEGORIES } = require('../../apps/api/dist/modules/support/support-catalog');

describe('cross-module flow invariants', () => {
  it('ensures each plan price maps to an approved billing plan', () => {
    for (const price of PLAN_PRICES) {
      assert.ok(price.id);
      assert.ok(price.planId);
      assert.ok(price.planCode);
      assert.equal(price.currency, 'VND');
      assert.ok(typeof price.amountMinor === 'number');
      assert.ok(['once', 'month', 'year'].includes(price.billingPeriod));
    }
  });

  it('verifies support categories align with functional domains', () => {
    const requiredDomains = [
      'general',
      'account',
      'billing',
      'licensing',
      'word_addin',
      'excel_addin',
    ];
    const activeDomains = SUPPORT_CATEGORIES.map((c) => c.code);

    for (const domain of requiredDomains) {
      assert.equal(
        activeDomains.includes(domain),
        true,
        `Support category domain ${domain} must be registered`,
      );
    }
  });

  it('guarantees non-negative order calculations under boundary inputs', () => {
    assert.equal(calcLineTotal(0, 1), 0);
    assert.equal(calcOrderTotal([{ lineTotalMinor: 0 }, { lineTotalMinor: 100 }]), 100);
  });
});
