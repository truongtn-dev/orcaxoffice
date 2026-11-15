const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { describe, it } = require('node:test');

const { CURATED_RULES } = require('../../packages/rule-definitions/src');
const { validateProfilePackage } = require('../../packages/profile-specifications/src');

describe('office add-in compatibility baseline', () => {
  const manifestsDir = path.resolve(__dirname, '../../apps/office-js/manifests');

  it('declares valid Word Office.js manifest with Taskpane and Desktop FormFactor', () => {
    const wordManifestPath = path.join(manifestsDir, 'word-manifest.xml');
    assert.equal(fs.existsSync(wordManifestPath), true, 'word-manifest.xml must exist');

    const content = fs.readFileSync(wordManifestPath, 'utf8');
    assert.match(content, /<Host Name="Document"\s*\/>/, 'Must declare Document host');
    assert.match(content, /<DesktopFormFactor>/, 'Must declare DesktopFormFactor');
    assert.match(content, /xsi:type="ShowTaskpane"/, 'Must support ShowTaskpane action');
    assert.match(content, /<Id>[0-9a-fA-F-]+<\/Id>/, 'Must have valid UUID add-in ID');
  });

  it('declares valid Excel Office.js manifest with Workbook host and Desktop FormFactor', () => {
    const excelManifestPath = path.join(manifestsDir, 'excel-manifest.xml');
    assert.equal(fs.existsSync(excelManifestPath), true, 'excel-manifest.xml must exist');

    const content = fs.readFileSync(excelManifestPath, 'utf8');
    assert.match(content, /<Host Name="Workbook"\s*\/>/, 'Must declare Workbook host');
    assert.match(content, /<DesktopFormFactor>/, 'Must declare DesktopFormFactor');
    assert.match(content, /<Id>[0-9a-fA-F-]+<\/Id>/, 'Must have valid UUID add-in ID');
  });

  it('ensures rule definitions cover both Word and Excel hosts deterministically', () => {
    assert.ok(Array.isArray(CURATED_RULES), 'CURATED_RULES must be an array');
    assert.ok(CURATED_RULES.length > 0, 'CURATED_RULES must not be empty');

    const wordRules = CURATED_RULES.filter((r) => r.host === 'word');
    const excelRules = CURATED_RULES.filter((r) => r.host === 'excel');

    assert.ok(wordRules.length > 0, 'Word rules must exist');
    assert.ok(excelRules.length > 0, 'Excel rules must exist');

    for (const rule of CURATED_RULES) {
      assert.ok(rule.code, 'Every rule must have a code');
      assert.ok(rule.title, 'Every rule must have a title');
      assert.ok(['word', 'excel'].includes(rule.host), 'Host must be word or excel');
      assert.ok(rule.description, 'Every rule must have a description');
    }
  });

  it('validates profile package compatibility according to orcax.profile.v1 schema', () => {
    const validWordPackage = {
      format: 'orcax.profile.v1',
      profile: {
        code: 'qd-1822-word',
        name: 'Chuẩn thể thức 1822',
        host: 'word',
      },
      rules: [
        {
          code: 'word.style.font-body',
          version: '1.0.0',
          checkKind: 'style.font',
          defaultSeverity: 'fail',
        },
      ],
      legalReferences: [],
    };

    const result = validateProfilePackage(validWordPackage);
    assert.equal(result.ok, true, `Package should be valid: ${JSON.stringify(result.errors)}`);
  });
});
