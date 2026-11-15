const { spawnSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const integration = process.argv.includes('--integration');
const directory = __dirname;
const files = fs
  .readdirSync(directory)
  .filter((name) =>
    integration ? name.endsWith('.integration-spec.ts') : name.endsWith('.spec.ts'),
  )
  .map((name) => path.join(directory, name));

if (files.length === 0) {
  process.stderr.write('No test files found.\n');
  process.exit(1);
}

const args = ['--require', 'ts-node/register', '--test', '--test-reporter', 'spec'];
if (integration) {
  args.push('--test-concurrency=1');
}

const result = spawnSync(process.execPath, [...args, ...files], { stdio: 'inherit' });

process.exit(result.status ?? 1);
