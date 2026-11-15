import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import {
  calcLineTotal,
  calcOrderTotal,
  PLAN_PRICES,
} from '../src/modules/commerce/commerce-catalog';

describe('commerce rules', () => {
  it('calculates line totals with non-negative integer amounts and positive quantity', () => {
    assert.equal(calcLineTotal(2_990_000, 1), 2_990_000);
    assert.equal(calcLineTotal(2_990_000, 3), 8_970_000);
    assert.equal(calcLineTotal(0, 5), 0);

    assert.throws(() => calcLineTotal(-1, 1), /unitAmountMinor must be a non-negative integer/);
    assert.throws(() => calcLineTotal(100, 0), /quantity must be an integer >= 1/);
    assert.throws(() => calcLineTotal(100.5, 1), /unitAmountMinor must be a non-negative integer/);
  });

  it('calculates total order amount across line items', () => {
    const items = [
      { lineTotalMinor: 2_990_000 },
      { lineTotalMinor: 5_980_000 },
      { lineTotalMinor: 0 },
    ];
    assert.equal(calcOrderTotal(items), 8_970_000);
    assert.equal(calcOrderTotal([]), 0);
  });

  it('declares prices matching active product plans', () => {
    const personalFree = PLAN_PRICES.find((p) => p.planCode === 'personal-free');
    const orgStarter = PLAN_PRICES.find((p) => p.planCode === 'organization-starter');
    const orgPerpetual = PLAN_PRICES.find((p) => p.planCode === 'organization-perpetual');

    assert.ok(personalFree);
    assert.equal(personalFree.amountMinor, 0);
    assert.equal(personalFree.currency, 'VND');

    assert.ok(orgStarter);
    assert.equal(orgStarter.amountMinor, 2_990_000);
    assert.equal(orgStarter.billingPeriod, 'year');

    assert.ok(orgPerpetual);
    assert.equal(orgPerpetual.amountMinor, 29_900_000);
    assert.equal(orgPerpetual.billingPeriod, 'once');
  });
});
