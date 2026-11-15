import assert from 'node:assert/strict';
import { describe, it } from 'node:test';
import { INITIAL_HELP_ARTICLES, SUPPORT_CATEGORIES } from '../src/modules/support/support-catalog';

describe('support catalog and rules', () => {
  it('defines curated support categories for core products', () => {
    const codes = new Set(SUPPORT_CATEGORIES.map((c) => c.code));
    assert.equal(codes.has('general'), true);
    assert.equal(codes.has('licensing'), true);
    assert.equal(codes.has('word_addin'), true);
    assert.equal(codes.has('excel_addin'), true);
    assert.equal(codes.has('billing'), true);
    assert.equal(codes.has('account'), true);
  });

  it('contains published initial help articles with valid categories and versioning', () => {
    assert.ok(INITIAL_HELP_ARTICLES.length >= 4);
    const categoryCodes = new Set(SUPPORT_CATEGORIES.map((c) => c.code as string));

    for (const article of INITIAL_HELP_ARTICLES) {
      assert.ok(article.id);
      assert.ok(article.slug);
      assert.ok(article.title);
      assert.ok(article.content);
      assert.equal(article.status, 'published');
      assert.equal(article.productVersion, '1.0');
      assert.equal(
        categoryCodes.has(article.category),
        true,
        `Article ${article.slug} has invalid category ${article.category}`,
      );
    }
  });

  it('ensures slugs of initial help articles are unique', () => {
    const slugs = INITIAL_HELP_ARTICLES.map((a) => a.slug);
    const uniqueSlugs = new Set(slugs);
    assert.equal(slugs.length, uniqueSlugs.size);
  });
});
