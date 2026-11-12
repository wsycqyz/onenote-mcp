import assert from 'node:assert/strict';
import fs from 'node:fs';
import { test } from 'node:test';

test('package-lock.json only references the public npm registry', () => {
  const lock = JSON.parse(fs.readFileSync(new URL('../package-lock.json', import.meta.url), 'utf8'));
  const elsewhere = Object.entries(lock.packages)
    .filter(([, entry]) => entry.resolved && !entry.resolved.startsWith('https://registry.npmjs.org/'))
    .map(([key]) => key);
  assert.deepEqual(elsewhere, [], 'Run `npm run lockfile:public` before committing.');
});
