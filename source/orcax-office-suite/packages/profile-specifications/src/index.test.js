const assert = require('node:assert/strict');
const { describe, it } = require('node:test');
const { PROFILE_PACKAGE_FORMAT, createProfilePackage, validateProfilePackage } = require('./index');

describe('profile package schema', () => {
  it('accepts a documented package', () => {
    const pkg = createProfilePackage({
      profile: {
        code: 'word-standard-vi',
        name: 'Chuẩn văn bản Việt',
        host: 'word',
        scope: 'system',
        source: 'Nội bộ OrcaX đã rà soát',
      },
      rules: [
        {
          code: 'word.style.font-body',
          version: '1.0.0',
          checkKind: 'style.font',
          defaultSeverity: 'fail',
        },
      ],
      legalReferences: [{ citation: 'QCVN nội bộ mẫu', verified: true }],
    });
    assert.equal(pkg.format, PROFILE_PACKAGE_FORMAT);
    assert.equal(validateProfilePackage(pkg).ok, true);
  });

  it('rejects invalid packages', () => {
    const result = validateProfilePackage({ format: 'other', profile: {}, rules: [] });
    assert.equal(result.ok, false);
    assert.ok(result.errors.length > 0);
  });
});
