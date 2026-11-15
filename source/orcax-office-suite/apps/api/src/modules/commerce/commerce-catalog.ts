import { PLANS } from '../licensing/license-catalog';

/** Amounts in VND minor units (đồng). */
export const PLAN_PRICES = [
  {
    id: '32000000-0000-4000-8000-000000000001',
    planId: PLANS[0].id,
    planCode: 'personal-free',
    currency: 'VND',
    amountMinor: 0,
    billingPeriod: 'once' as const,
  },
  {
    id: '32000000-0000-4000-8000-000000000002',
    planId: PLANS[1].id,
    planCode: 'organization-starter',
    currency: 'VND',
    amountMinor: 2_990_000,
    billingPeriod: 'year' as const,
  },
  {
    id: '32000000-0000-4000-8000-000000000003',
    planId: PLANS[2].id,
    planCode: 'organization-perpetual',
    currency: 'VND',
    amountMinor: 29_900_000,
    billingPeriod: 'once' as const,
  },
] as const;

export function calcLineTotal(unitAmountMinor: number, quantity: number): number {
  if (!Number.isInteger(unitAmountMinor) || unitAmountMinor < 0) {
    throw new TypeError('unitAmountMinor must be a non-negative integer');
  }
  if (!Number.isInteger(quantity) || quantity < 1) {
    throw new TypeError('quantity must be an integer >= 1');
  }
  return unitAmountMinor * quantity;
}

export function calcOrderTotal(items: readonly { lineTotalMinor: number }[]): number {
  return items.reduce((sum, item) => sum + item.lineTotalMinor, 0);
}
