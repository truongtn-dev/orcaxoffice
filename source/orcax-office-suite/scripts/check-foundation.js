const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const bannedDependencies = new Set([
  'openai',
  '@anthropic-ai/sdk',
  'langchain',
  '@langchain/core',
  '@langchain/openai',
  '@google/generative-ai',
  'ai',
]);

const failures = [];

function readJson(file) {
  return JSON.parse(fs.readFileSync(file, 'utf8'));
}

function dependencyNames(manifest) {
  return Object.keys({
    ...manifest.dependencies,
    ...manifest.devDependencies,
    ...manifest.optionalDependencies,
  });
}

function isBanned(specifier) {
  if (bannedDependencies.has(specifier)) {
    return true;
  }
  return [...bannedDependencies].some((name) => specifier.startsWith(`${name}/`));
}

function checkManifest(file) {
  const manifest = readJson(file);
  for (const name of dependencyNames(manifest)) {
    if (isBanned(name)) {
      failures.push(`${path.relative(root, file)} depends on ${name}`);
    }
  }
}

function walk(dir, visit) {
  if (!fs.existsSync(dir)) {
    return;
  }
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    if (entry.name === 'node_modules' || entry.name === 'dist' || entry.name === '.next') {
      continue;
    }
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      walk(fullPath, visit);
      continue;
    }
    visit(fullPath);
  }
}

function checkSource(file) {
  if (!/\.(js|ts|jsx|tsx)$/.test(file)) {
    return;
  }
  const source = fs.readFileSync(file, 'utf8');
  const importPattern = /(?:require\(\s*|from\s+)['"]([^'"]+)['"]/g;
  for (const match of source.matchAll(importPattern)) {
    if (isBanned(match[1])) {
      failures.push(`${path.relative(root, file)} imports ${match[1]}`);
    }
  }
  if (/synchronize\s*:\s*true/.test(source)) {
    failures.push(`${path.relative(root, file)} enables schema synchronize`);
  }
}

const gitignore = fs.readFileSync(path.join(root, '.gitignore'), 'utf8');
if (!/(^|\n)\.env(\r?\n|$)/.test(gitignore)) {
  failures.push('.gitignore must ignore .env');
}

for (const relativePath of [
  'package.json',
  path.join('apps', 'api', 'package.json'),
  path.join('apps', 'web-portal', 'package.json'),
  path.join('packages', 'api-contracts', 'package.json'),
]) {
  checkManifest(path.join(root, relativePath));
}

walk(path.join(root, 'apps'), checkSource);
walk(path.join(root, 'packages'), checkSource);

if (failures.length > 0) {
  process.stderr.write(`${failures.join('\n')}\n`);
  process.exit(1);
}
