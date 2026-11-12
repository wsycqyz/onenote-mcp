import assert from 'node:assert/strict';
import { test } from 'node:test';
import { loadConfig } from '../lib/config.mjs';
import { createGraph } from '../lib/graph.mjs';
import { createSectionGuard, WriteDeniedError } from '../lib/guard.mjs';
import { createFakeGraph } from './fake-graph.mjs';

const getToken = async () => 'test-token';

test('config is read-only unless write-section names a section', () => {
  assert.equal(loadConfig({}).mode, 'read-only');
  assert.equal(loadConfig({ ONENOTE_WRITE_SECTION: 'AI Manage' }).writeSection, null);

  const missing = loadConfig({ ONENOTE_ACCESS: 'write-section' });
  assert.equal(missing.mode, 'read-only');
  assert.equal(missing.warnings.length, 1);

  const typo = loadConfig({ ONENOTE_ACCESS: 'write', ONENOTE_WRITE_SECTION: 'AI Manage' });
  assert.equal(typo.mode, 'read-only');

  const write = loadConfig({ ONENOTE_ACCESS: 'write-section', ONENOTE_WRITE_SECTION: ' AI Manage ' });
  assert.equal(write.mode, 'write-section');
  assert.equal(write.writeSection, 'AI Manage');
});

test('read-only graph has no writer and refuses non-OneNote URLs', async () => {
  const fake = createFakeGraph();
  const { reader, writer } = createGraph({ getToken, fetchImpl: fake.fetchImpl });
  assert.equal(writer, null);
  await assert.rejects(reader.get('https://example.com/me/onenote/pages'), /only Microsoft Graph/);
  await assert.rejects(reader.get('/me/drive/root/children'), /only \/me and OneNote paths/);
  await assert.rejects(reader.get('/me/onenote/../../users/someone'), /only \/me and OneNote paths/);
  assert.equal(fake.requests.length, 0);
});

test('writer rejects IDs that could escape the page or section path', async () => {
  const fake = createFakeGraph();
  const { writer } = createGraph({ getToken, fetchImpl: fake.fetchImpl, allowWrites: true });
  await assert.rejects(writer.deletePage('../sections/sec-journal'), /Invalid OneNote pageId/);
  await assert.rejects(writer.createPage('sec-ai/pages?x=1', '<p/>'), /Invalid OneNote sectionId/);
  assert.equal(fake.requests.length, 0);
});

function writableGuard(fake, options = {}) {
  const { reader, writer } = createGraph({ getToken, fetchImpl: fake.fetchImpl, allowWrites: true });
  return createSectionGuard({ reader, writer, sectionName: 'AI Manage', ...options });
}

test('guard writes only to pages of the configured section', async () => {
  const fake = createFakeGraph();
  const guard = writableGuard(fake);

  const created = await guard.createPage({ title: 'Hi', bodyXhtml: '<p>x</p>' });
  assert.equal(created.section.id, 'sec-ai');
  await assert.rejects(guard.createPage({ title: 'Hi', bodyXhtml: '', sectionId: 'sec-journal' }), WriteDeniedError);
  await assert.rejects(guard.patchPage('page-journal', []), WriteDeniedError);
  await assert.rejects(guard.deletePage('page-journal'), WriteDeniedError);
  await guard.patchPage('page-ai', [{ target: 'body', action: 'append', content: '<p>y</p>' }]);
  await guard.deletePage('page-ai');

  assert.deepEqual(
    fake.writes().map((r) => `${r.method} ${r.path}`),
    [
      'POST /v1.0/me/onenote/sections/sec-ai/pages',
      'PATCH /v1.0/me/onenote/pages/page-ai/content',
      'DELETE /v1.0/me/onenote/pages/page-ai',
    ],
  );
});

test('guard denies writes when the section is missing or ambiguous', async () => {
  const missing = createFakeGraph({ sections: [{ id: 's1', displayName: 'Journal', parentNotebook: { displayName: 'A' } }] });
  await assert.rejects(writableGuard(missing).createPage({ title: 't', bodyXhtml: '' }), /No OneNote section named "AI Manage"/);

  const twoNotebooks = [
    { id: 's1', displayName: 'AI Manage', parentNotebook: { id: 'a', displayName: 'Personal' } },
    { id: 's2', displayName: 'ai manage', parentNotebook: { id: 'b', displayName: 'Work stuff' } },
  ];
  const ambiguous = createFakeGraph({ sections: twoNotebooks });
  await assert.rejects(writableGuard(ambiguous).createPage({ title: 't', bodyXhtml: '' }), /2 sections are named/);
  assert.equal(ambiguous.writes().length, 0);

  const pinned = createFakeGraph({ sections: twoNotebooks });
  const { section } = await writableGuard(pinned, { notebookName: 'Work stuff' }).createPage({ title: 't', bodyXhtml: '' });
  assert.equal(section.id, 's2');
});

test('guard re-checks the section before every write, ignoring the display cache', async () => {
  const fake = createFakeGraph();
  const guard = writableGuard(fake);
  await guard.createPage({ title: 't', bodyXhtml: '' });

  fake.state.sections[0].displayName = 'Renamed';
  await assert.rejects(guard.createPage({ title: 't', bodyXhtml: '' }), WriteDeniedError);
  await assert.rejects(guard.patchPage('page-ai', []), WriteDeniedError);

  fake.state.sections[0].displayName = 'AI Manage';
  fake.state.sections.push({ id: 'sec-dup', displayName: 'AI Manage', parentNotebook: { id: 'nb-1', displayName: 'Personal' } });
  await assert.rejects(guard.deletePage('page-ai'), /2 sections are named/);
  assert.equal(fake.writes().length, 1);
});
