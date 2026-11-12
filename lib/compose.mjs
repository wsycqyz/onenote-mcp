import { createHash } from 'node:crypto';
import { JSDOM } from 'jsdom';
import { escapeHtml, expandSelfClosingTags, htmlToMarkdown } from './convert.mjs';

// Building blocks for copying a page's content verbatim into a new layout.

const SKIPPED = new Set(['BR', 'SCRIPT', 'STYLE', 'META', 'TITLE']);
const POSITIONING = /(^|;)\s*(position|left|top|z-index)\s*:[^;]*/gi;
const XHTML_NS = / xmlns="http:\/\/www\.w3\.org\/1999\/xhtml"/g;
const RANGE = /^(\d+)\s*-\s*(\d+)$/;

// Visible characters only: whitespace layout differs between OneNote pages, characters must not.
export function compactText(text) {
  return String(text ?? '').replace(/[\s\u00A0\uFFFC]+/g, '');
}

function normalizeText(text) {
  return String(text ?? '')
    .replace(/[\s\u00A0\uFFFC]+/g, ' ')
    .trim();
}

function collectBlocks(container, out) {
  for (const node of [...container.childNodes]) {
    if (node.nodeType === 3) {
      if (node.nodeValue.trim()) {
        const paragraph = node.ownerDocument.createElement('p');
        paragraph.append(node.cloneNode());
        out.push(paragraph);
      }
    } else if (node.nodeType === 1 && !SKIPPED.has(node.nodeName)) {
      // OneNote wraps content in absolutely positioned outline <div>s; their children are the blocks.
      if (node.nodeName === 'DIV') collectBlocks(node, out);
      else out.push(node);
    }
  }
}

/**
 * Splits OneNote page HTML into numbered top-level blocks (paragraphs, tables,
 * lists, images, ...). `version` identifies this exact content, so a later copy
 * can detect that the page changed in between.
 */
export function extractBlocks(html) {
  const dom = new JSDOM(expandSelfClosingTags(html));
  const { document, XMLSerializer } = dom.window;
  const serializer = new XMLSerializer();
  const nodes = [];
  collectBlocks(document.body, nodes);

  const blocks = nodes.map((node, index) => {
    const clone = node.cloneNode(true);
    for (const element of [clone, ...clone.querySelectorAll('*')]) {
      element.removeAttribute('id');
      element.removeAttribute('data-id');
    }
    const style = (clone.getAttribute('style') ?? '').replace(POSITIONING, '').replace(/^;+|;+$/g, '').trim();
    if (style) clone.setAttribute('style', style);
    else clone.removeAttribute('style');
    return {
      ref: index + 1,
      type: node.nodeName.toLowerCase(),
      html: serializer.serializeToString(clone).replace(XHTML_NS, ''),
      text: compactText(node.textContent),
      alt: node.nodeName === 'IMG' ? normalizeText(node.getAttribute('alt')) : '',
    };
  });
  const version = createHash('sha256')
    .update(blocks.map((block) => block.html).join('\n'))
    .digest('hex')
    .slice(0, 16);
  return { blocks, version };
}

/** One-line preview of a block for choosing a layout. */
export function blockPreview(block, maxChars = 160) {
  const text =
    block.type === 'img'
      ? `(image) ${block.alt}`
      : htmlToMarkdown(block.html).replace(/!\[[^\]]*\]\([^)]*\)/g, '(image)').replace(/\s*\n+\s*/g, ' / ');
  return text.length > maxChars ? `${text.slice(0, maxChars)}…` : text;
}

function expandRefs(items) {
  return items.flatMap((item) => {
    if (typeof item === 'number') return [item];
    const match = RANGE.exec(String(item).trim());
    if (!match) throw new Error(`Invalid block reference ${JSON.stringify(item)}; use a number or a range like "3-7".`);
    const [from, to] = [Number(match[1]), Number(match[2])];
    if (to < from) throw new Error(`Invalid block range "${item}".`);
    return Array.from({ length: to - from + 1 }, (_, i) => from + i);
  });
}

function compactRanges(refs) {
  const sorted = [...refs].sort((a, b) => a - b);
  const ranges = [];
  for (const ref of sorted) {
    const last = ranges.at(-1);
    if (last && ref === last[1] + 1) last[1] = ref;
    else ranges.push([ref, ref]);
  }
  return ranges.map(([from, to]) => (from === to ? `${from}` : `${from}-${to}`)).join(', ');
}

/**
 * Lays out source blocks under optional headings. Every block must be used
 * exactly once or listed in `omit`, so nothing is lost or duplicated silently.
 */
export function composeBody({ blocks, sections, omit = [] }) {
  const byRef = new Map(blocks.map((block) => [block.ref, block]));
  const check = (ref) => {
    if (!byRef.has(ref)) throw new Error(`Block ${ref} does not exist; the page has blocks 1-${blocks.length}.`);
  };
  const omitted = new Set(omit);
  omitted.forEach(check);

  const used = [];
  const seen = new Set();
  const html = [];
  for (const section of sections) {
    if (section.heading) {
      const level = section.headingLevel ?? 2;
      html.push(`<h${level}>${escapeHtml(section.heading)}</h${level}>`);
    }
    for (const ref of expandRefs(section.refs ?? [])) {
      check(ref);
      if (seen.has(ref)) throw new Error(`Block ${ref} is used more than once.`);
      if (omitted.has(ref)) throw new Error(`Block ${ref} is both used and listed in omit.`);
      seen.add(ref);
      used.push(byRef.get(ref));
      html.push(byRef.get(ref).html);
    }
  }
  const missing = blocks.filter((block) => !seen.has(block.ref) && !omitted.has(block.ref)).map((block) => block.ref);
  if (missing.length) {
    throw new Error(`Every block must be copied or listed in omit. Not placed: ${compactRanges(missing)}.`);
  }
  return { bodyHtml: html.join(''), used, omitted: [...omitted].sort((a, b) => a - b) };
}

const GRAPH_RESOURCE = /^https:\/\/graph\.microsoft\.com\/v1\.0\/(me|users\/[^/]+|users\('[^/]+'\))\/onenote\/resources\/[^/]+\/(\$value|content)$/;

const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex');
const MIME_TYPE = /^[\w.+-]+\/[\w.+-]+$/;

function sniffImageType(bytes) {
  const head = bytes.subarray(0, 12).toString('hex');
  if (head.startsWith('ffd8ff')) return 'image/jpeg';
  if (head.startsWith('89504e47')) return 'image/png';
  if (head.startsWith('47494638')) return 'image/gif';
  if (head.startsWith('424d')) return 'image/bmp';
  if (head.startsWith('49492a00') || head.startsWith('4d4d002a')) return 'image/tiff';
  if (head.startsWith('52494646') && head.endsWith('57454250')) return 'image/webp';
  return null;
}

// Graph serves resources as application/octet-stream, and OneNote drops image parts sent with that
// type, so prefer the type the source page declares, then the file's magic bytes.
function partContentType(kind, declared, bytes, served) {
  const candidates = kind === 'image' ? [declared, sniffImageType(bytes), served] : [declared, served];
  const valid = candidates
    .map((type) => (type ?? '').split(';')[0].trim().toLowerCase())
    .filter((type) => MIME_TYPE.test(type) && type !== 'application/octet-stream');
  return (kind === 'image' ? valid.find((type) => type.startsWith('image/')) : valid[0]) ?? 'application/octet-stream';
}

/**
 * Replaces OneNote resource URLs (images, attachments) with multipart part
 * references and downloads their bytes, so the copy doesn't depend on the source page.
 */
export async function attachResources(bodyHtml, download) {
  const dom = new JSDOM(`<!DOCTYPE html><html><body>${expandSelfClosingTags(bodyHtml)}</body></html>`);
  const { document, XMLSerializer } = dom.window;
  const parts = [];
  const take = async (url, kind, declaredType) => {
    if (!url || !GRAPH_RESOURCE.test(url)) return null;
    const { bytes, contentType } = await download(url);
    const name = `part${parts.length + 1}`;
    parts.push({ name, kind, contentType: partContentType(kind, declaredType, bytes, contentType), bytes, sha256: sha256(bytes) });
    return `name:${name}`;
  };
  for (const image of document.querySelectorAll('img')) {
    const fullres = image.getAttribute('data-fullres-src');
    const declared = fullres ? image.getAttribute('data-fullres-src-type') : image.getAttribute('data-src-type');
    const reference = await take(fullres || image.getAttribute('src'), 'image', declared);
    if (!reference) continue;
    image.setAttribute('src', reference);
    for (const attribute of ['data-fullres-src', 'data-src-type', 'data-fullres-src-type']) image.removeAttribute(attribute);
  }
  for (const attachment of document.querySelectorAll('object')) {
    const reference = await take(attachment.getAttribute('data'), 'attachment', attachment.getAttribute('type'));
    if (reference) attachment.setAttribute('data', reference);
  }
  const serializer = new XMLSerializer();
  const html = [...document.body.childNodes].map((node) => serializer.serializeToString(node)).join('').replace(XHTML_NS, '');
  return { html, parts };
}

/** Checks that the new page contains every copied block's characters, in order. */
export function verifyText(used, newPageHtml) {
  const { document } = new JSDOM(expandSelfClosingTags(newPageHtml)).window;
  const pageText = compactText(document.body.textContent);
  const mismatched = [];
  let cursor = 0;
  for (const block of used) {
    if (!block.text) continue;
    const at = pageText.indexOf(block.text, cursor);
    if (at === -1) mismatched.push(block.ref);
    else cursor = at + block.text.length;
  }
  return { ok: mismatched.length === 0, mismatched };
}

/**
 * Downloads the new page's images and attachments and compares them, in page
 * order, with the uploaded parts by SHA-256.
 */
export async function verifyFiles(newPageHtml, parts, download) {
  const { document } = new JSDOM(expandSelfClosingTags(newPageHtml)).window;
  const found = {
    image: [...document.querySelectorAll('img')].map((img) => img.getAttribute('data-fullres-src') || img.getAttribute('src')),
    attachment: [...document.querySelectorAll('object')].map((object) => object.getAttribute('data')),
  };
  const result = {};
  for (const kind of ['image', 'attachment']) {
    const expected = parts.filter((part) => part.kind === kind);
    let identical = 0;
    if (found[kind].length === expected.length) {
      for (const [index, url] of found[kind].entries()) {
        if (GRAPH_RESOURCE.test(url ?? '') && sha256((await download(url)).bytes) === expected[index].sha256) identical++;
      }
    }
    result[`${kind}s`] = { expected: expected.length, found: found[kind].length, identical };
  }
  const ok = Object.values(result).every((r) => r.found === r.expected && r.identical === r.expected);
  return { ok, ...result };
}
