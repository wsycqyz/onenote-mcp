import assert from 'node:assert/strict';
import { test } from 'node:test';
import { attachResources, composeBody, extractBlocks, verifyFiles, verifyText } from '../lib/compose.mjs';
import { FILE_URL, IMAGE_URL, SOURCE_HTML } from './fake-graph.mjs';

test('extractBlocks numbers top-level blocks, keeps content after a self-closing attachment, and drops positioning', () => {
  const { blocks, version } = extractBlocks(SOURCE_HTML);
  assert.deepEqual(blocks.map((b) => b.type), ['p', 'table', 'p', 'object', 'p', 'img', 'p']);
  assert.match(blocks[0].html, /^<p style="margin-top:0pt">alpha <span style="font-weight:bold">key-123<\/span><\/p>$/);
  assert.match(blocks[2].html, /beta p\*\*\*w0rd &lt;host IP&gt;/);
  assert.equal(blocks[3].html, `<object data-attachment="doc.pdf" type="application/pdf" data="${FILE_URL}"></object>`);
  assert.equal(blocks[4].html, '<p>after attachment</p>');
  assert.doesNotMatch(blocks[5].html, /position/);
  assert.equal(blocks[6].html, '<p>gamma text node</p>');
  assert.equal(extractBlocks(SOURCE_HTML).version, version);
  assert.notEqual(extractBlocks(SOURCE_HTML.replace('key-123', 'key-124')).version, version);
});

test('composeBody requires every block exactly once unless omitted', () => {
  const { blocks } = extractBlocks(SOURCE_HTML);
  const sections = [{ heading: 'Keys & codes', refs: [1, 6] }, { heading: 'Rest', refs: ['2-5'] }];
  const { bodyHtml, used, omitted } = composeBody({ blocks, sections, omit: [7] });
  assert.match(bodyHtml, /^<h2>Keys &amp; codes<\/h2><p style/);
  assert.deepEqual(used.map((b) => b.ref), [1, 6, 2, 3, 4, 5]);
  assert.deepEqual(omitted, [7]);

  assert.throws(() => composeBody({ blocks, sections }), /Not placed: 7/);
  assert.throws(() => composeBody({ blocks, sections: [{ refs: ['1-7', 2] }] }), /Block 2 is used more than once/);
  assert.throws(() => composeBody({ blocks, sections: [{ refs: ['1-7'] }], omit: [7] }), /both used and listed in omit/);
  assert.throws(() => composeBody({ blocks, sections: [{ refs: ['1-8'] }] }), /Block 8 does not exist/);
});

test('verifyText catches a changed character', () => {
  const { blocks } = extractBlocks(SOURCE_HTML);
  const copy = blocks.map((b) => b.html).join('');
  assert.equal(verifyText(blocks, `<html><body><h2>x</h2>${copy}</body></html>`).ok, true);
  assert.deepEqual(verifyText(blocks, `<html><body>${copy.replace('key-123', 'key-I23')}</body></html>`).mismatched, [1]);
});

test('attachResources turns images and attachments into multipart parts with real content types', async () => {
  const { blocks } = extractBlocks(SOURCE_HTML);
  const png = Buffer.from('89504e470d0a1a0a', 'hex');
  const downloads = [];
  const { html, parts } = await attachResources(blocks[3].html + blocks[5].html, async (url) => {
    downloads.push(url);
    // Graph serves every resource as application/octet-stream.
    return { bytes: url === IMAGE_URL ? png : Buffer.from('%PDF-1.7'), contentType: 'application/octet-stream' };
  });
  assert.deepEqual(downloads, [IMAGE_URL, FILE_URL]);
  assert.deepEqual(parts.map((p) => [p.name, p.kind, p.contentType]), [['part1', 'image', 'image/png'], ['part2', 'attachment', 'application/pdf']]);
  assert.match(html, /<object [^>]*data="name:part2"><\/object><img [^>]*src="name:part1"/);
  assert.doesNotMatch(html, /data-fullres-src|graph\.microsoft\.com/);
});

test('verifyFiles compares every image and attachment on the new page by hash', async () => {
  const { parts } = await attachResources(`<img src="${IMAGE_URL}" /><object data="${FILE_URL}" />`, async (url) => ({
    bytes: Buffer.from(url),
    contentType: 'image/png',
  }));
  const page = `<html><body><img src="${IMAGE_URL}" /><object data="${FILE_URL}" /></body></html>`;
  const same = await verifyFiles(page, parts, async (url) => ({ bytes: Buffer.from(url) }));
  assert.deepEqual(same, { ok: true, images: { expected: 1, found: 1, identical: 1 }, attachments: { expected: 1, found: 1, identical: 1 } });

  const changed = await verifyFiles(page, parts, async (url) => ({ bytes: Buffer.from(url === FILE_URL ? 'other' : url) }));
  assert.deepEqual([changed.ok, changed.attachments.identical], [false, 0]);
  const missing = await verifyFiles(`<html><body><img src="${IMAGE_URL}" /></body></html>`, parts, async (url) => ({ bytes: Buffer.from(url) }));
  assert.deepEqual([missing.ok, missing.attachments.found], [false, 0]);
});
