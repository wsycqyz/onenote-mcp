#!/usr/bin/env node
// Points every "resolved" URL in package-lock.json at the public npm registry, for
// installs made through a registry mirror. Integrity hashes stay valid: the tarballs are identical.
import fs from 'node:fs';

const file = new URL('../package-lock.json', import.meta.url);
const lock = JSON.parse(fs.readFileSync(file, 'utf8'));
let changed = 0;
for (const [key, entry] of Object.entries(lock.packages ?? {})) {
  if (!entry.resolved || !key.includes('node_modules/')) continue;
  const name = key.slice(key.lastIndexOf('node_modules/') + 'node_modules/'.length);
  const url = `https://registry.npmjs.org/${name}/-/${name.split('/').pop()}-${entry.version}.tgz`;
  if (entry.resolved !== url) {
    entry.resolved = url;
    changed++;
  }
}
fs.writeFileSync(file, `${JSON.stringify(lock, null, 2)}\n`);
console.log(`package-lock.json: ${changed} resolved URL(s) rewritten to registry.npmjs.org`);
