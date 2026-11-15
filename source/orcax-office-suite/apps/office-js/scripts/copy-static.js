const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const dist = path.join(root, 'dist');

function copyDir(from, to) {
  fs.mkdirSync(to, { recursive: true });
  for (const entry of fs.readdirSync(from, { withFileTypes: true })) {
    const source = path.join(from, entry.name);
    const target = path.join(to, entry.name);
    if (entry.isDirectory()) {
      copyDir(source, target);
    } else {
      fs.copyFileSync(source, target);
    }
  }
}

fs.rmSync(dist, { recursive: true, force: true });
copyDir(path.join(root, 'public'), dist);
fs.mkdirSync(path.join(dist, 'src', 'word'), { recursive: true });
fs.mkdirSync(path.join(dist, 'src', 'excel'), { recursive: true });
fs.copyFileSync(
  path.join(root, 'src', 'word', 'core.browser.js'),
  path.join(dist, 'src', 'word', 'core.browser.js'),
);
fs.copyFileSync(
  path.join(root, 'src', 'excel', 'core.browser.js'),
  path.join(dist, 'src', 'excel', 'core.browser.js'),
);
fs.mkdirSync(path.join(dist, 'manifests'), { recursive: true });
fs.copyFileSync(
  path.join(root, 'manifests', 'word-manifest.xml'),
  path.join(dist, 'manifests', 'word-manifest.xml'),
);
fs.copyFileSync(
  path.join(root, 'manifests', 'excel-manifest.xml'),
  path.join(dist, 'manifests', 'excel-manifest.xml'),
);

// Minimal opaque PNG placeholders for sideload manifests.
const png = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  'base64',
);
const assets = path.join(dist, 'assets');
fs.mkdirSync(assets, { recursive: true });
for (const name of ['icon-16.png', 'icon-32.png', 'icon-64.png', 'icon-80.png']) {
  fs.writeFileSync(path.join(assets, name), png);
}

console.log('office-js dist ready');
