import assert from 'node:assert/strict';
import { test } from 'node:test';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { loadConfig } from '../lib/config.mjs';
import { createOneNoteServer, READ_TOOLS, WRITE_TOOLS } from '../lib/server.mjs';
import { createFakeGraph, fakeAuth, SOURCE_HTML } from './fake-graph.mjs';

async function connect(env, fakeOptions) {
  const fake = createFakeGraph(fakeOptions);
  const server = createOneNoteServer({ config: loadConfig(env), auth: fakeAuth, fetchImpl: fake.fetchImpl });
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  const client = new Client({ name: 'test', version: '1.0.0' });
  await Promise.all([server.connect(serverTransport), client.connect(clientTransport)]);
  const call = async (name, args = {}) => {
    try {
      const result = await client.callTool({ name, arguments: args });
      return { error: Boolean(result.isError), text: result.content.map((c) => c.text).join('\n') };
    } catch (error) {
      return { error: true, text: error.message };
    }
  };
  return { client, call, fake };
}

const WRITE_ENV = { ONENOTE_ACCESS: 'write-section', ONENOTE_WRITE_SECTION: 'AI Manage' };

test('read-only server exposes no write tools and sends no writes', async () => {
  const { client, call, fake } = await connect({});
  const names = (await client.listTools()).tools.map((t) => t.name).sort();
  assert.deepEqual(names, [...READ_TOOLS].sort());

  const page = await call('getPage', { pageId: 'page-journal' });
  assert.equal(page.error, false);
  assert.match(page.text, /body of page-journal/);
  assert.equal((await call('createPage', { title: 't', content: 'c' })).error, true);
  assert.equal(fake.writes().length, 0);
});

test('write-section server writes only inside the configured section', async () => {
  const { client, call, fake } = await connect(WRITE_ENV);
  const names = (await client.listTools()).tools.map((t) => t.name).sort();
  assert.deepEqual(names, [...READ_TOOLS, ...WRITE_TOOLS].sort());

  assert.equal((await call('createPage', { title: 'Plan', content: '- [ ] task' })).error, false);
  const outside = await call('createPage', { title: 'x', content: 'y', sectionId: 'sec-journal' });
  assert.equal(outside.error, true);
  assert.match(outside.text, /Write denied/);
  assert.equal((await call('appendToPage', { pageId: 'page-journal', content: 'x' })).error, true);
  assert.equal((await call('updatePage', { pageId: 'page-journal', commands: [{ target: 'title', action: 'replace', content: 'x' }] })).error, true);
  assert.equal((await call('deletePage', { pageId: 'page-journal', expectedTitle: 'Diary' })).error, true);
  assert.equal((await call('deletePage', { pageId: 'page-ai', expectedTitle: 'wrong title' })).error, true);
  assert.equal((await call('appendToPage', { pageId: 'page-ai', content: 'more' })).error, false);
  assert.equal((await call('deletePage', { pageId: 'page-ai', expectedTitle: 'AI notes' })).error, false);

  const writes = fake.writes();
  assert.deepEqual(
    writes.map((r) => `${r.method} ${r.path}`),
    [
      'POST /v1.0/me/onenote/sections/sec-ai/pages',
      'PATCH /v1.0/me/onenote/pages/page-ai/content',
      'DELETE /v1.0/me/onenote/pages/page-ai',
    ],
  );
  assert.match(writes[0].body, /<p data-tag="to-do">task<\/p>/);
});

test('searchPages skips password-protected sections', async () => {
  const { call } = await connect({}, { encrypted: ['sec-journal'] });
  const result = await call('searchPages', { query: 'notes' });
  assert.equal(result.error, false);
  const body = JSON.parse(result.text);
  assert.deepEqual(body.results.map((p) => p.id), ['page-ai']);
  assert.match(body.skippedSections[0], /Journal \(password protected\)/);
});

test('composePage copies blocks and images verbatim into the writable section only', async () => {
  const imageBytes = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x00, 0xff, 0x0d, 0x0a]);
  const fileBytes = Buffer.from('%PDF-1.7 \x00\xff binary');
  const pages = {
    'page-src': {
      id: 'page-src',
      title: 'Notes',
      lastModifiedDateTime: '2020-01-01T00:00:00Z',
      parentSection: { id: 'sec-journal', displayName: 'Journal' },
      parentNotebook: { id: 'nb-1', displayName: 'Personal' },
      html: SOURCE_HTML,
    },
  };
  const resources = {
    'res-1': { bytes: imageBytes, contentType: 'image/png' },
    'res-2': { bytes: fileBytes, contentType: 'application/pdf' },
  };
  const { call, fake } = await connect(WRITE_ENV, { pages, resources });

  const listing = await call('getPage', { pageId: 'page-src', format: 'blocks' });
  const header = JSON.parse(listing.text.split('\n\n---\n\n')[0]);
  assert.equal(header.blockCount, 7);
  assert.match(listing.text, /^\[4\] object: \[attachment: doc\.pdf\]/m);
  assert.match(listing.text, /^\[5\] p: after attachment$/m);
  assert.match(listing.text, /^\[6\] img: \(image\) screenshot text$/m);

  const base = { title: 'Notes (reformatted)', sourcePageId: 'page-src', sourceVersion: header.blocksVersion };
  const sections = [{ heading: 'Keys', refs: [1, 6] }, { heading: 'Rest', refs: ['2-5'] }];
  assert.match((await call('composePage', { ...base, sections })).text, /Not placed: 7/);
  assert.match((await call('composePage', { ...base, sourceVersion: 'stale', sections, omit: [7] })).text, /source page changed/);
  assert.equal(fake.writes().length, 0);

  const result = await call('composePage', { ...base, sections, omit: [7] });
  assert.equal(result.error, false, result.text);
  const body = JSON.parse(result.text);
  assert.deepEqual([body.blocksCopied, body.omitted, body.filesCopied], [6, [7], 2]);
  assert.deepEqual(body.verification, {
    ok: true,
    mismatchedBlocks: [],
    images: { expected: 1, found: 1, identical: 1 },
    attachments: { expected: 1, found: 1, identical: 1 },
  });

  assert.deepEqual(fake.writes().map((r) => `${r.method} ${r.path}`), ['POST /v1.0/me/onenote/sections/sec-ai/pages']);
  const [created] = fake.state.created;
  assert.match(created.posted, /<p style="color:#7f7f7f">Source: Personal \/ Journal \/ Notes \(last modified 2020-01-01T00:00:00Z\)<\/p><h2>Keys<\/h2>/);
  assert.match(created.posted, /beta p\*\*\*w0rd &lt;host IP&gt;/);
  assert.match(created.posted, /<img [^>]*src="name:part1"/);
  assert.match(created.posted, /<object data-attachment="doc\.pdf" type="application\/pdf" data="name:part2"><\/object><p>after attachment<\/p>/);
  assert.doesNotMatch(created.posted, /position:absolute|gamma/);
  assert.deepEqual(created.parts.find((p) => p.name === 'part1').bytes, imageBytes);
  assert.deepEqual(created.parts.find((p) => p.name === 'part2').bytes, fileBytes);
});
