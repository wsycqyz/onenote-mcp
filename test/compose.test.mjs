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
  assert.throws(() => composeBody({ blocks, sections: [{ refs: [99, 'x'] }] }), /Invalid block reference "x"/);
});

test('composeBody theme developer restyles blocks without touching their text, one row per refs item', () => {
  const { blocks } = extractBlocks(
    '<html><body><div style="position:absolute;left:48px;top:120px;width:624px">' +
      '<p lang="zh-CN" data-tag="to-do" style="margin-top:5.5pt;font-family:Aptos">' +
      '<span style="font-family:Microsoft YaHei;font-size:14pt;color:red;background-color:yellow;font-weight:bold" lang="en-US">key-1</span> ' +
      '<span style="text-decoration:line-through;font-weight:bold">old</span> <a href="https://example.com/x" style="color:#333;font-size:14pt">link</a> <b>bold</b></p>' +
      '<table style="border:1px solid"><tr><td style="border:1px solid;background-color:#eee;font-family:Aptos;width:120px">cell</td></tr></table>' +
      '<cite style="font-size:9pt;color:#595959">From <a href="https://example.com/">site</a></cite>' +
      '<p>ls -la /tmp | grep log</p>' +
      '<table style="border:0px"><tr><td style="border:0px">outer<table style="border:1px solid"><tr><td style="border:1px solid">inner</td></tr></table></td></tr></table>' +
      '</div></body></html>',
  );
  const sections = [{ heading: 'Keys & <codes>', refs: ['1-2', 3] }, { heading: 'Shell', panel: 'dark', refs: [4, 5] }];
  const { bodyHtml, used } = composeBody({ blocks, sections, theme: 'developer' });

  assert.equal(verifyText(used, `<html><body>${bodyHtml}</body></html>`).ok, true);
  assert.doesNotMatch(bodyHtml, /Aptos|YaHei|14pt|color:red|yellow|#eee|<b>/);
  assert.match(bodyHtml, /^<table border="0" style="width:900px"><tr><td style="background-color:#24292f;[^"]*"><p [^>]*>.*Keys &amp; &lt;codes&gt;<\/span><\/p><\/td><\/tr>/);
  assert.match(
    bodyHtml,
    /<tr><td style="background-color:#f6f8fa;[^"]*"><p lang="zh-CN" data-tag="to-do" style="font-family:Consolas;font-size:10\.5pt;color:#1f2328;margin-top:0pt;margin-bottom:0pt"><span lang="en-US">key-1<\/span> <span style="text-decoration:line-through">old<\/span> <a href="https:\/\/example\.com\/x" style="color:#0969da">link<\/a> bold<\/p><table border="1"><tbody><tr><td style="font-family:Consolas;[^"]*;width:120px">cell<\/td><\/tr><\/tbody><\/table><\/td><\/tr><tr><td style="background-color:#eaeef2;[^"]*"><cite style="font-family:Consolas;font-size:9pt;color:#59636e;[^"]*">From <a href="https:\/\/example\.com\/" style="color:#0969da">site<\/a><\/cite><\/td><\/tr><\/table>/,
  );
  assert.match(bodyHtml, /<td style="background-color:#0d1117;[^"]*"><p style="font-family:Consolas;font-size:10\.5pt;color:#e6edf3;[^"]*">ls -la \/tmp \| grep log<\/p>/);
  // A bordered table nested in a borderless one keeps its border without spreading it to the outer table.
  assert.match(bodyHtml, /<td style="background-color:#161b22;[^"]*"><table border="0"><tbody><tr><td style="[^"]*">outer<table border="1">/);
  assert.throws(() => composeBody({ blocks, sections, theme: 'fancy' }), /Unknown theme "fancy"/);
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
