import { createHash } from 'node:crypto';
import { JSDOM } from 'jsdom';
import { escapeHtml, expandSelfClosingTags, htmlToMarkdown } from './convert.mjs';

// Building blocks for copying a page's content verbatim into a new layout.

const SKIPPED = new Set(['BR', 'SCRIPT', 'STYLE', 'META', 'TITLE']);
const POSITIONING = /(^|;)\s*(position|left|top|z-index)\s*:[^;]*/gi;
const XHTML_NS = / xmlns="http:\/\/www\.w3\.org\/1999\/xhtml"/g;
// A block reference: "5" or "3-7", optionally prefixed with its source page, as in "2:5" or "2:3-7".
const REF = /^(?:(\d+)\s*:\s*)?(\d+)(?:\s*-\s*(\d+))?$/;

export const THEMES = Object.freeze(['original', 'developer']);
export const PANELS = Object.freeze(['light', 'dark']);

// "developer" theme: one monospace font and GitHub-style colours; each section is a window-style table.
const DEVELOPER = Object.freeze({
  font: 'Consolas',
  size: '10.5pt',
  smallSize: '9pt',
  headingSize: '12pt',
  titleSizes: Object.freeze({ 1: '13pt', 2: '11.5pt', 3: '10.5pt' }),
  width: '900px',
  chrome: '#24292f',
  chromeText: '#f6f8fa',
  dots: Object.freeze(['#ff5f56', '#ffbd2e', '#27c93f']),
  panels: Object.freeze({
    light: Object.freeze({ rows: ['#f6f8fa', '#eaeef2'], text: '#1f2328', muted: '#59636e', link: '#0969da' }),
    dark: Object.freeze({ rows: ['#0d1117', '#161b22'], text: '#e6edf3', muted: '#8b949e', link: '#79c0ff' }),
  }),
});
const SPACER = `<p style="font-size:${DEVELOPER.smallSize};margin-top:0pt;margin-bottom:0pt">&#160;</p>`;
const CHARACTER_TAGS = new Set(['B', 'STRONG', 'I', 'EM', 'U', 'INS', 'FONT', 'MARK', 'BIG', 'SMALL']);
const TEXT_CONTAINERS = new Set(['P', 'LI', 'TD', 'TH', 'BLOCKQUOTE', 'PRE', 'DIV']);
const HEADINGS = new Set(['H1', 'H2', 'H3', 'H4', 'H5', 'H6']);
const BORDER = /^border(-(top|right|bottom|left))?(-width)?$/;

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

/** Parses one refs or omit item. Without a page prefix it names a block of page 1. */
function parseRef(item, multi) {
  const match = typeof item === 'number' ? [null, undefined, item, item] : REF.exec(String(item).trim());
  if (!match) {
    const syntax = multi ? '"page:block" or a range like "2:3-7"' : 'a number or a range like "3-7"';
    throw new Error(`Invalid block reference ${JSON.stringify(item)}; use ${syntax}.`);
  }
  if (multi && match[1] === undefined) {
    throw new Error(`Block reference ${JSON.stringify(item)} needs its page: with several source pages use "page:block", like "2:5" or "2:3-7".`);
  }
  const [page, from, to] = [Number(match[1] ?? 1), Number(match[2]), Number(match[3] ?? match[2])];
  if (to < from) throw new Error(`Invalid block range "${item}".`);
  return { page, from, to };
}

const byPosition = (a, b) => a.page - b.page || a.ref - b.ref;

function compactRanges(blocks, multi) {
  const ranges = [];
  for (const { page, ref } of [...blocks].sort(byPosition)) {
    const last = ranges.at(-1);
    if (last && last.page === page && ref === last.to + 1) last.to = ref;
    else ranges.push({ page, from: ref, to: ref });
  }
  return ranges.map(({ page, from, to }) => `${multi ? `${page}:` : ''}${from === to ? from : `${from}-${to}`}`).join(', ');
}

function declarations(style) {
  return String(style ?? '')
    .split(';')
    .map((declaration) => declaration.trim())
    .filter(Boolean)
    .map((declaration) => {
      const at = declaration.indexOf(':');
      return at === -1 ? [declaration.toLowerCase(), ''] : [declaration.slice(0, at).trim().toLowerCase(), declaration.slice(at + 1).trim()];
    });
}

const hasBorder = (element) =>
  declarations(element.getAttribute('style')).some(([name, value]) => BORDER.test(name) && !/^(0(\.0+)?(px|pt)?|none|hidden)\b/i.test(value));

/**
 * Returns a function that replaces a block's fonts, sizes, colours, highlights, margins and
 * bold/italic/underline with the developer theme. Text nodes are never touched; links, note tags,
 * language, strike-through, alignment, table structure and images stay. One DOM is reused for every
 * block, because a JSDOM window per block costs about 1 MB of heap until the call returns.
 */
function developerRenderer() {
  const { document, XMLSerializer } = new JSDOM('<!DOCTYPE html><html><body></body></html>').window;
  const serializer = new XMLSerializer();
  return (html, palette) => {
    document.body.innerHTML = html;
    const text = `font-family:${DEVELOPER.font};font-size:${DEVELOPER.size};color:${palette.text}`;
    for (const element of [...document.body.querySelectorAll('*')]) {
      const tag = element.nodeName;
      if (tag === 'IMG' || tag === 'OBJECT') continue;
      if (CHARACTER_TAGS.has(tag)) {
        element.replaceWith(...element.childNodes);
        continue;
      }
      const declared = declarations(element.getAttribute('style'));
      const styles = declared
        .filter(([name]) => name === 'text-align' || name === 'list-style-type' || (name === 'width' && /^(TABLE|TD|TH)$/.test(tag)))
        .map(([name, value]) => `${name}:${value}`);
      if (declared.some(([name, value]) => name === 'text-decoration' && /line-through/i.test(value))) {
        styles.push('text-decoration:line-through');
      }
      if (tag === 'TABLE') {
        // OneNote ignores CSS borders on input tables; only the border attribute draws them.
        // Cells are still unprocessed here; rows/cells exclude those of nested tables.
        const cells = [...element.rows].flatMap((row) => [...row.cells]);
        const bordered = Number(element.getAttribute('border')) > 0 || [element, ...cells].some(hasBorder);
        element.setAttribute('border', bordered ? '1' : '0');
        element.removeAttribute('bgcolor');
      } else if (TEXT_CONTAINERS.has(tag)) {
        styles.unshift(`${text};margin-top:0pt;margin-bottom:0pt`);
      } else if (HEADINGS.has(tag)) {
        styles.unshift(`${text};font-size:${DEVELOPER.headingSize};font-weight:bold;margin-top:0pt;margin-bottom:0pt`);
      } else if (tag === 'CITE') {
        styles.unshift(`font-family:${DEVELOPER.font};font-size:${DEVELOPER.smallSize};color:${palette.muted};margin-top:0pt;margin-bottom:0pt`);
      } else if (tag === 'A') {
        styles.unshift(`color:${palette.link}`);
      }
      if (styles.length) element.setAttribute('style', styles.join(';'));
      else element.removeAttribute('style');
      if (tag === 'SPAN' && element.attributes.length === 0) element.replaceWith(...element.childNodes);
    }
    return [...document.body.childNodes].map((node) => serializer.serializeToString(node)).join('').replace(XHTML_NS, '');
  };
}

function developerTitle(heading, level) {
  const size = DEVELOPER.titleSizes[level] ?? DEVELOPER.titleSizes[2];
  const dots = DEVELOPER.dots.map((color) => `<span style="color:${color}">&#9679;</span>`).join(' ');
  return (
    `<tr><td style="background-color:${DEVELOPER.chrome};width:${DEVELOPER.width}">` +
    `<p style="font-family:${DEVELOPER.font};font-size:${size};color:${DEVELOPER.chromeText};margin-top:0pt;margin-bottom:0pt">` +
    `${dots}&#160;&#160;&#160;<span style="font-weight:bold">${escapeHtml(heading)}</span></p></td></tr>`
  );
}

/** A section as a window-style table: the heading in a title bar, then one row per refs item. */
function developerSection(section, rows, render) {
  const palette = DEVELOPER.panels[section.panel ?? 'light'];
  if (!palette) throw new Error(`Unknown panel "${section.panel}"; use ${PANELS.join(' or ')}.`);
  const html = [];
  if (section.heading) html.push(developerTitle(section.heading, section.headingLevel ?? 2));
  rows.forEach((row, index) => {
    const cell = `background-color:${palette.rows[index % 2]};width:${DEVELOPER.width};font-family:${DEVELOPER.font};font-size:${DEVELOPER.size};color:${palette.text}`;
    html.push(`<tr><td style="${cell}">${row.map((block) => render(block.html, palette)).join('')}</td></tr>`);
  });
  return html.length ? `<table border="0" style="width:${DEVELOPER.width}">${html.join('')}</table>${SPACER}` : '';
}

/** The lines that start a composed page and name its source pages. */
export function sourceLinesHtml(lines, theme = 'original') {
  if (theme !== 'developer') return lines.map((text) => `<p style="color:#7f7f7f">${escapeHtml(text)}</p>`).join('');
  const style = `font-family:${DEVELOPER.font};font-size:${DEVELOPER.smallSize};color:${DEVELOPER.panels.light.muted};margin-top:0pt;margin-bottom:0pt`;
  return `${lines.map((text) => `<p style="${style}">${escapeHtml(text)}</p>`).join('')}${SPACER}`;
}

/**
 * Lays out source blocks under optional headings. Every block must be used
 * exactly once or listed in `omit`, so nothing is lost or duplicated silently.
 * `blocks` are the blocks of one source page. To merge pages, pass `sources`,
 * one block list per page; refs and omit then name blocks as "page:block"
 * (1-based page), and blocks are labelled that way in errors and results.
 * Theme "original" copies each block's HTML as is. Theme "developer" restyles
 * the blocks and renders each section as a table with one row per refs item.
 */
export function composeBody({ blocks, sources, sections, omit = [], theme = 'original' }) {
  if (!THEMES.includes(theme)) throw new Error(`Unknown theme "${theme}"; use ${THEMES.join(' or ')}.`);
  const pages = sources ?? [blocks];
  const multi = pages.length > 1;
  const all = new Map();
  pages.forEach((pageBlocks, index) => {
    for (const block of pageBlocks) {
      const page = index + 1;
      all.set(`${page}:${block.ref}`, { ...block, page, label: multi ? `${page}:${block.ref}` : block.ref });
    }
  });
  const find = (page, ref) => {
    const block = all.get(`${page}:${ref}`);
    if (block) return block;
    const pageBlocks = pages[page - 1];
    if (!pageBlocks) {
      throw new Error(`Page ${page} does not exist; there ${pages.length === 1 ? 'is 1 source page' : `are ${pages.length} source pages`}.`);
    }
    const has = pageBlocks.length ? `blocks 1-${pageBlocks.length}` : 'no blocks';
    throw new Error(multi ? `Block ${page}:${ref} does not exist; page ${page} has ${has}.` : `Block ${ref} does not exist; the page has ${has}.`);
  };
  // Both ends are checked first, so a huge range fails instead of allocating.
  const expand = ({ page, from, to }) => {
    find(page, from);
    find(page, to);
    return Array.from({ length: to - from + 1 }, (_, i) => find(page, from + i));
  };
  const omitted = new Set(omit.flatMap((item) => expand(parseRef(item, multi))));

  const used = [];
  const seen = new Set();
  const take = (block) => {
    if (seen.has(block)) throw new Error(`Block ${block.label} is used more than once.`);
    if (omitted.has(block)) throw new Error(`Block ${block.label} is both used and listed in omit.`);
    seen.add(block);
    used.push(block);
    return block;
  };
  // Lay out and validate everything before rendering. Each section's refs are parsed before any is
  // taken, so a malformed ref is reported first, as before themes existed.
  const laidOut = sections.map((section) => {
    const items = (section.refs ?? []).map((item) => parseRef(item, multi));
    return { section, rows: items.map((ref) => expand(ref).map(take)) };
  });
  const missing = [...all.values()].filter((block) => !seen.has(block) && !omitted.has(block));
  if (missing.length) {
    throw new Error(`Every block must be copied or listed in omit. Not placed: ${compactRanges(missing, multi)}.`);
  }

  const html = [];
  const render = theme === 'developer' ? developerRenderer() : null;
  for (const { section, rows } of laidOut) {
    if (render) {
      html.push(developerSection(section, rows, render));
      continue;
    }
    if (section.heading) {
      const level = section.headingLevel ?? 2;
      html.push(`<h${level}>${escapeHtml(section.heading)}</h${level}>`);
    }
    for (const block of rows.flat()) html.push(block.html);
  }
  return { bodyHtml: html.join(''), used, omitted: [...omitted].sort(byPosition).map((block) => block.label) };
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
    if (at === -1) mismatched.push(block.label ?? block.ref);
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
