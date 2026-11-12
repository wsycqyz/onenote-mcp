import { createHash } from 'node:crypto';
import { JSDOM } from 'jsdom';
import { escapeHtml, htmlToMarkdown } from './convert.mjs';

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
  const dom = new JSDOM(String(html ?? ''));
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

/**
 * Replaces OneNote resource URLs (images, attachments) with multipart part
 * references and downloads their bytes, so the copy doesn't depend on the source page.
 */
export async function attachResources(bodyHtml, download) {
  const dom = new JSDOM(`<!DOCTYPE html><html><body>${bodyHtml}</body></html>`);
  const { document, XMLSerializer } = dom.window;
  const parts = [];
  const take = async (url) => {
    if (!url || !GRAPH_RESOURCE.test(url)) return null;
    const { bytes, contentType } = await download(url);
    const name = `part${parts.length + 1}`;
    parts.push({ name, contentType: contentType.split(';')[0].trim(), bytes });
    return `name:${name}`;
  };
  for (const image of document.querySelectorAll('img')) {
    const reference = await take(image.getAttribute('data-fullres-src') || image.getAttribute('src'));
    if (!reference) continue;
    image.setAttribute('src', reference);
    for (const attribute of ['data-fullres-src', 'data-src-type', 'data-fullres-src-type']) image.removeAttribute(attribute);
  }
  for (const attachment of document.querySelectorAll('object')) {
    const reference = await take(attachment.getAttribute('data'));
    if (reference) attachment.setAttribute('data', reference);
  }
  const serializer = new XMLSerializer();
  const html = [...document.body.childNodes].map((node) => serializer.serializeToString(node)).join('').replace(XHTML_NS, '');
  return { html, parts };
}

/** Checks that the new page contains every copied block's characters, in order, and the same number of images. */
export function verifyCopy(used, newPageHtml) {
  const { document } = new JSDOM(String(newPageHtml ?? '')).window;
  const pageText = compactText(document.body.textContent);
  const mismatched = [];
  let cursor = 0;
  for (const block of used) {
    if (!block.text) continue;
    const at = pageText.indexOf(block.text, cursor);
    if (at === -1) mismatched.push(block.ref);
    else cursor = at + block.text.length;
  }
  const expectedImages = used.reduce((count, block) => count + (block.html.match(/<img\b/g)?.length ?? 0), 0);
  const images = document.querySelectorAll('img').length;
  return { ok: mismatched.length === 0 && images === expectedImages, mismatched, images, expectedImages };
}
