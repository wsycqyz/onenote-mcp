import assert from 'node:assert/strict';
import { test } from 'node:test';
import { attachResources, composeBody, extractBlocks, verifyCopy } from '../lib/compose.mjs';
import { IMAGE_URL, SOURCE_HTML } from './fake-graph.mjs';

test('extractBlocks numbers top-level blocks and drops positioning', () => {
  const { blocks, version } = extractBlocks(SOURCE_HTML);
  assert.deepEqual(blocks.map((b) => b.type), ['p', 'table', 'p', 'img', 'p']);
  assert.match(blocks[0].html, /^<p style="margin-top:0pt">alpha <span style="font-weight:bold">key-123<\/span><\/p>$/);
  assert.match(blocks[2].html, /beta p\*\*\*w0rd &lt;host IP&gt;/);
  assert.doesNotMatch(blocks[3].html, /position/);
  assert.equal(blocks[4].html, '<p>gamma text node</p>');
  assert.equal(extractBlocks(SOURCE_HTML).version, version);
  assert.notEqual(extractBlocks(SOURCE_HTML.replace('key-123', 'key-124')).version, version);
});

test('composeBody requires every block exactly once unless omitted', () => {
  const { blocks } = extractBlocks(SOURCE_HTML);
  const sections = [{ heading: 'Keys & codes', refs: [1, 4] }, { heading: 'Rest', refs: ['2-3'] }];
  const { bodyHtml, used, omitted } = composeBody({ blocks, sections, omit: [5] });
  assert.match(bodyHtml, /^<h2>Keys &amp; codes<\/h2><p style/);
  assert.deepEqual(used.map((b) => b.ref), [1, 4, 2, 3]);
  assert.deepEqual(omitted, [5]);

  assert.throws(() => composeBody({ blocks, sections }), /Not placed: 5/);
  assert.throws(() => composeBody({ blocks, sections: [{ refs: ['1-5', 2] }] }), /Block 2 is used more than once/);
  assert.throws(() => composeBody({ blocks, sections: [{ refs: ['1-5'] }], omit: [5] }), /both used and listed in omit/);
  assert.throws(() => composeBody({ blocks, sections: [{ refs: ['1-6'] }] }), /Block 6 does not exist/);
});

test('verifyCopy catches a changed character or a missing image', () => {
  const { blocks } = extractBlocks(SOURCE_HTML);
  const copy = blocks.map((b) => b.html).join('');
  assert.equal(verifyCopy(blocks, `<html><body><h2>x</h2>${copy}</body></html>`).ok, true);
  const altered = verifyCopy(blocks, `<html><body>${copy.replace('key-123', 'key-I23')}</body></html>`);
  assert.deepEqual(altered.mismatched, [1]);
  assert.equal(verifyCopy(blocks, `<html><body>${copy.replace(/<img[^>]*>/, '')}</body></html>`).ok, false);
});

test('attachResources turns OneNote images into multipart parts', async () => {
  const { blocks } = extractBlocks(SOURCE_HTML);
  const downloads = [];
  const { html, parts } = await attachResources(blocks[3].html, async (url) => {
    downloads.push(url);
    return { bytes: Buffer.from([1, 2, 3]), contentType: 'image/png; charset=binary' };
  });
  assert.deepEqual(downloads, [IMAGE_URL]);
  assert.deepEqual(parts.map((p) => [p.name, p.contentType]), [['part1', 'image/png']]);
  assert.match(html, /src="name:part1"/);
  assert.doesNotMatch(html, /data-fullres-src|graph\.microsoft\.com/);
});
