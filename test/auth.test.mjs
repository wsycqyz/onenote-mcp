import assert from 'node:assert/strict';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { test } from 'node:test';
import { Auth, fileCachePlugin } from '../lib/auth.mjs';
import { CONSUMERS_TENANT_ID } from '../lib/config.mjs';

test('concurrent cache saves leave the latest complete file and no temp files', async (t) => {
  const dir = await fs.promises.mkdtemp(path.join(os.tmpdir(), 'onenote-mcp-test-'));
  t.after(() => fs.promises.rm(dir, { recursive: true, force: true }));
  const file = path.join(dir, 'msal-cache.json');
  const plugin = fileCachePlugin(file);
  const save = (n) =>
    plugin.afterCacheAccess({
      cacheHasChanged: true,
      tokenCache: { serialize: () => JSON.stringify({ n, padding: 'x'.repeat(100_000) }) },
    });

  await Promise.all(Array.from({ length: 25 }, (_, n) => save(n)));

  assert.equal(JSON.parse(await fs.promises.readFile(file, 'utf8')).n, 24);
  assert.deepEqual((await fs.promises.readdir(dir)).filter((name) => name.endsWith('.tmp')), []);
  let loaded;
  await plugin.beforeCacheAccess({ tokenCache: { deserialize: (data) => (loaded = JSON.parse(data).n) } });
  assert.equal(loaded, 24);
});

test('concurrent token requests share one MSAL call', async () => {
  let calls = 0;
  const account = { username: 'me@outlook.com', tenantId: CONSUMERS_TENANT_ID };
  const pca = {
    getAllAccounts: async () => [account],
    acquireTokenSilent: async () => {
      calls++;
      await new Promise((resolve) => setTimeout(resolve, 20));
      return { accessToken: 'token', account };
    },
  };
  const auth = new Auth({}, { pca });
  assert.deepEqual(await Promise.all([auth.getAccessToken(), auth.getAccessToken(), auth.getAccessToken()]), ['token', 'token', 'token']);
  assert.equal(calls, 1);
  await auth.getAccessToken();
  assert.equal(calls, 2);
});
