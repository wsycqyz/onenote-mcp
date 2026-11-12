import assert from 'node:assert/strict';
import { test } from 'node:test';
import { JSDOM } from 'jsdom';
import { buildPageXhtml, htmlToMarkdown, toXhtmlFragment } from '../lib/convert.mjs';

test('markdown becomes well-formed XHTML in the shapes OneNote keeps', () => {
  const markdown = [
    '# T',
    'A & B ~~old~~ <script>x()</script>',
    '- [ ] open\n- [x] done\n- plain',
    '```\nif x:\n    y()\n```',
    'line<br>\n\u0001',
  ].join('\n\n');
  const body = toXhtmlFragment(markdown, 'markdown');
  // OneNote drops data-tag on <li>, so task items become tagged paragraphs; plain items stay a list.
  assert.match(body, /<p data-tag="to-do">open<\/p>\s*<p data-tag="to-do:completed">done<\/p>\s*<ul>\s*<li>plain<\/li>\s*<\/ul>/);
  // OneNote mangles <pre>, so code is written as one monospace paragraph per line.
  assert.match(body, /<p style="font-family:Consolas[^"]*">if\u00A0x:<\/p><p style="font-family:Consolas[^"]*">\u00A0{4}y\(\)<\/p>/);
  assert.match(body, /<span style="text-decoration:line-through">old<\/span>/);
  assert.match(body, /<br \/>/);
  assert.doesNotMatch(body, /script|<input|<pre|<li data-tag/);
  // Throws if the page is not well-formed XML.
  new JSDOM(buildPageXhtml('Q & "A"', body).replace('<!DOCTYPE html>', ''), { contentType: 'application/xhtml+xml' });
});

test('OneNote HTML becomes readable Markdown', () => {
  const mono = 'style="font-family:consolas"';
  const html =
    '<html><head><title>T</title></head><body><div>' +
    '<p>Plain <span style="font-weight:bold">bold</span> <span style="font-style:italic">it</span> ' +
    `<span style="text-decoration:line-through">gone</span> <span ${mono}>inline</span> text</p>` +
    '<p data-tag="to-do">milk</p><p data-tag="to-do:completed">bread</p><p data-tag="important">key_point</p>' +
    `<p ${mono}>if\u00A0x:</p><p><span ${mono}>\u00A0\u00A0y()</span><span ${mono}>\uFFFC</span><span ${mono}>z()</span></p>` +
    '<table><tr><td>a|b</td><td>c</td></tr><tr><td>1</td><td>2</td></tr></table>' +
    '<object data-attachment="doc.pdf" data="https://graph.microsoft.com/v1.0/x/$value"></object></div></body></html>';
  const markdown = htmlToMarkdown(html);
  assert.match(markdown, /^Plain \*\*bold\*\* \*it\* ~~gone~~ `inline` text$/m);
  assert.match(markdown, /^\[ \] milk$/m);
  assert.match(markdown, /^\[x\] bread$/m);
  assert.match(markdown, /^\[important\] key_point$/m);
  assert.match(markdown, /^```\nif x:\n {2}y\(\)\nz\(\)\n```$/m);
  assert.match(markdown, /^\| a\\\|b \| c \|\n\| --- \| --- \|\n\| 1 \| 2 \|$/m);
  assert.match(markdown, /\[attachment: doc\.pdf\]\(https:\/\/graph\.microsoft\.com\/v1\.0\/x\/\$value\)/);
});
