import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import {
  EXCEL_CAPABILITIES,
  WORD_CAPABILITIES,
} from '../src/modules/word-automation/add-in-capabilities';

describe('add-in capabilities', () => {
  it('declares word and excel local-first matrices', () => {
    assert.equal(WORD_CAPABILITIES.host, 'word');
    assert.equal(WORD_CAPABILITIES.entitlement, 'word.core');
    assert.ok(WORD_CAPABILITIES.actions.some((item) => item.id === 'preview-apply'));
    assert.equal(EXCEL_CAPABILITIES.host, 'excel');
    assert.equal(EXCEL_CAPABILITIES.entitlement, 'excel.core');
    assert.ok(EXCEL_CAPABILITIES.actions.some((item) => item.id === 'export-pdf'));
  });
});
