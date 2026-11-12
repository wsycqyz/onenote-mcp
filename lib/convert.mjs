import { JSDOM } from 'jsdom';
import { marked } from 'marked';
import TurndownService from 'turndown';

export const CONTENT_FORMATS = Object.freeze(['markdown', 'html', 'text']);

const VOID_ELEMENTS = new Set(['area', 'base', 'br', 'col', 'embed', 'hr', 'img', 'input', 'link', 'meta', 'param', 'source', 'track', 'wbr']);
const SELF_CLOSING_TAG = /<([A-Za-z][\w:-]*)((?:\s+[^\s<>/=]+(?:\s*=\s*(?:"[^"]*"|'[^']*'|[^\s"'<>]+))?)*)\s*\/>/g;

/**
 * OneNote returns XHTML-style self-closing tags such as <object ... />. HTML parsers ignore
 * the slash on non-void elements, so the element would swallow everything after it.
 */
export function expandSelfClosingTags(html) {
  return String(html ?? '').replace(SELF_CLOSING_TAG, (tag, name, attributes) =>
    VOID_ELEMENTS.has(name.toLowerCase()) ? tag : `<${name}${attributes}></${name}>`,
  );
}

const INVALID_XML_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g;
const UNSAFE_ELEMENTS = 'script, style, iframe, frame, frameset, form, button, select, textarea, link, meta, base, embed, applet';
const BLOCK_ELEMENTS = new Set(['UL', 'OL', 'TABLE', 'PRE', 'DIV', 'P', 'BLOCKQUOTE', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6']);
const MONOSPACE = /font-family:\s*['"]?(consolas|courier new|courier|monospace|lucida console|cascadia)/i;
const CODE_LINE_STYLE = 'font-family:Consolas;margin-top:0pt;margin-bottom:0pt';
// OneNote turns line breaks inside <pre> into U+FFFC.
const OBJECT_REPLACEMENT = /\uFFFC/g;

export function escapeHtml(text) {
  return String(text)
    .replace(INVALID_XML_CHARS, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function textToHtml(text) {
  return String(text)
    .split(/\r?\n[ \t]*\r?\n/)
    .filter((paragraph) => paragraph.length > 0)
    .map((paragraph) => `<p>${escapeHtml(paragraph).replace(/\r?\n/g, '<br />')}</p>`)
    .join('');
}

// OneNote keeps data-tag only on paragraphs, so tagged list items become tagged <p> elements.
function liftTaggedListItems(document) {
  for (const list of [...document.querySelectorAll('ul, ol')].reverse()) {
    const items = [...list.children].filter((child) => child.nodeName === 'LI');
    if (!items.some((li) => li.hasAttribute('data-tag'))) continue;
    const replacement = [];
    let run = null;
    for (const li of items) {
      if (!li.hasAttribute('data-tag')) {
        if (!run) replacement.push((run = document.createElement(list.nodeName.toLowerCase())));
        run.append(li);
        continue;
      }
      run = null;
      const paragraph = document.createElement('p');
      paragraph.setAttribute('data-tag', li.getAttribute('data-tag'));
      replacement.push(paragraph);
      for (const child of [...li.childNodes]) {
        const empty = paragraph.childNodes.length === 0;
        if (empty && child.nodeType === 3 && !child.nodeValue.trim()) continue;
        if (child.nodeName === 'P' && empty) paragraph.append(...child.childNodes);
        else if (BLOCK_ELEMENTS.has(child.nodeName)) replacement.push(child);
        else paragraph.append(child);
      }
    }
    list.replaceWith(...replacement);
  }
}

// OneNote mangles <pre>; write one monospace paragraph per line instead.
function preToCodeLines(document) {
  for (const pre of [...document.querySelectorAll('pre')]) {
    const lines = pre.textContent.replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n');
    pre.replaceWith(
      ...lines.map((line) => {
        const paragraph = document.createElement('p');
        paragraph.setAttribute('style', CODE_LINE_STYLE);
        paragraph.textContent = line.replace(/\t/g, '    ').replace(/ /g, '\u00A0') || '\u00A0';
        return paragraph;
      }),
    );
  }
}

/** Converts Markdown, HTML or plain text into a well-formed XHTML fragment OneNote accepts. */
export function toXhtmlFragment(content, format = 'markdown') {
  if (!CONTENT_FORMATS.includes(format)) throw new Error(`Unsupported content format "${format}".`);
  const source = String(content ?? '').replace(INVALID_XML_CHARS, '');
  let html;
  if (format === 'html') html = expandSelfClosingTags(source);
  else if (format === 'text') html = textToHtml(source);
  else html = marked.parse(source, { gfm: true, async: false });

  const dom = new JSDOM(`<!DOCTYPE html><html><body>${html}</body></html>`);
  const { document, XMLSerializer } = dom.window;
  document.querySelectorAll(UNSAFE_ELEMENTS).forEach((node) => node.remove());

  // Markdown task lists ("- [ ] item") become OneNote to-do tags.
  for (const box of document.querySelectorAll('input[type="checkbox"]')) {
    const holder = box.closest('li') ?? box.closest('p');
    if (holder && !holder.hasAttribute('data-tag')) {
      holder.setAttribute('data-tag', box.hasAttribute('checked') ? 'to-do:completed' : 'to-do');
    }
    const next = box.nextSibling;
    if (next && next.nodeType === 3) next.textContent = next.textContent.replace(/^\s+/, '');
    box.remove();
  }
  document.querySelectorAll('input').forEach((node) => node.remove());
  for (const element of document.body.querySelectorAll('*')) {
    for (const attribute of [...element.attributes]) {
      if (/^on/i.test(attribute.name)) element.removeAttribute(attribute.name);
    }
  }
  liftTaggedListItems(document);
  preToCodeLines(document);
  for (const strike of [...document.querySelectorAll('del, strike')]) {
    const span = document.createElement('span');
    span.setAttribute('style', 'text-decoration:line-through');
    span.append(...strike.childNodes);
    strike.replaceWith(span);
  }

  const serializer = new XMLSerializer();
  return [...document.body.childNodes]
    .map((node) => serializer.serializeToString(node))
    .join('')
    .replace(/ xmlns="http:\/\/www\.w3\.org\/1999\/xhtml"/g, '');
}

export function buildPageXhtml(title, bodyXhtml) {
  return `<!DOCTYPE html><html><head><title>${escapeHtml(title)}</title></head><body>${bodyXhtml}</body></html>`;
}

function tagLabel(value) {
  return value
    .split(',')
    .map((tag) => tag.trim())
    .filter(Boolean)
    .map((tag) => {
      const done = tag.endsWith(':completed');
      const name = done ? tag.slice(0, -':completed'.length) : tag;
      if (name.startsWith('to-do')) return done ? '[x]' : '[ ]';
      return `[${name}${done ? ': done' : ''}]`;
    })
    .join(' ');
}

function isMonospace(element) {
  return MONOSPACE.test(element.getAttribute('style') || '');
}

// A link written as its own URL loses nothing in a code block.
const urlForm = (value) => String(value ?? '').trim().replace(/^(https?:\/\/|mailto:)/i, '').replace(/\/$/, '');

// A paragraph is a code line when all of its visible text is monospace (or it is a blank monospace line).
// Paragraphs with strike-through, or with a link whose text isn't its URL, stay prose: a code block would drop them.
function isCodeParagraph(paragraph, window) {
  if (paragraph.hasAttribute('data-tag')) return false;
  if (paragraph.querySelector('del, s, strike, [style*="line-through"]')) return false;
  if ([...paragraph.querySelectorAll('a[href]')].some((a) => urlForm(a.textContent) !== urlForm(a.getAttribute('href')))) return false;
  const walker = paragraph.ownerDocument.createTreeWalker(paragraph, window.NodeFilter.SHOW_TEXT);
  let visible = false;
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (!node.nodeValue.replace(/[\s\u00A0\uFFFC]/g, '')) continue;
    visible = true;
    let mono = false;
    for (let element = node.parentElement; element && !mono; element = element === paragraph ? null : element.parentElement) {
      mono = isMonospace(element);
    }
    if (!mono) return false;
  }
  return visible || isMonospace(paragraph) || [...paragraph.querySelectorAll('[style]')].some(isMonospace);
}

function mergeCodeParagraphs(document, window) {
  const code = new Set([...document.querySelectorAll('p')].filter((p) => isCodeParagraph(p, window)));
  for (const first of code) {
    if (!first.isConnected || code.has(first.previousElementSibling)) continue;
    const run = [first];
    while (code.has(run.at(-1).nextElementSibling)) run.push(run.at(-1).nextElementSibling);
    const pre = document.createElement('pre');
    const block = document.createElement('code');
    block.textContent = run
      .map((p) => {
        p.querySelectorAll('br').forEach((br) => br.replaceWith('\n'));
        return p.textContent.replace(OBJECT_REPLACEMENT, '\n').replace(/\u00A0/g, ' ').replace(/\s+$/, '');
      })
      .join('\n');
    pre.append(block);
    first.before(pre);
    run.forEach((p) => p.remove());
  }
}

// OneNote stores inline formatting as styled spans; wrap them in semantic elements Turndown understands.
function unwrapStyledSpans(document) {
  for (const span of [...document.querySelectorAll('span[style]')]) {
    if (!span.textContent.trim() || span.closest('pre')) continue;
    const style = span.getAttribute('style');
    const tags = [];
    if (MONOSPACE.test(style)) tags.push('code');
    if (/text-decoration:[^;]*line-through/i.test(style)) tags.push('del');
    if (/font-style:\s*italic/i.test(style)) tags.push('em');
    if (/font-weight:\s*(bold|[6-9]00)/i.test(style)) tags.push('strong');
    let inner = [...span.childNodes];
    for (const tag of tags) {
      const wrapper = document.createElement(tag);
      wrapper.append(...inner);
      inner = [wrapper];
    }
    span.append(...inner);
  }
}

function createTurndown() {
  const service = new TurndownService({
    headingStyle: 'atx',
    codeBlockStyle: 'fenced',
    bulletListMarker: '-',
    emDelimiter: '*',
  });
  // Output is for reading, so keep text literal instead of Markdown-escaping it.
  service.escape = (text) => text;

  const cellText = (cell) =>
    service.turndown(cell).replace(/\s*\n+\s*/g, ' <br> ').replace(/\|/g, '\\|').trim();

  service.addRule('table', {
    filter: 'table',
    replacement(_content, table) {
      const rows = [...table.rows].map((row) => [...row.cells].map(cellText));
      if (rows.length === 0) return '';
      const width = Math.max(...rows.map((row) => row.length));
      const line = (cells) => `| ${Array.from({ length: width }, (_, i) => cells[i] ?? '').join(' | ')} |`;
      return `\n\n${[line(rows[0]), line(Array(width).fill('---')), ...rows.slice(1).map(line)].join('\n')}\n\n`;
    },
  });
  service.addRule('strikethrough', {
    filter: ['del', 's', 'strike'],
    replacement: (content) => (content.trim() ? `~~${content}~~` : content),
  });
  return service;
}

let turndown;

/** Converts OneNote page HTML into readable Markdown. */
export function htmlToMarkdown(html) {
  turndown ??= createTurndown();
  const dom = new JSDOM(expandSelfClosingTags(html));
  const { document } = dom.window;
  document.querySelectorAll('head, script, style').forEach((node) => node.remove());

  mergeCodeParagraphs(document, dom.window);
  unwrapStyledSpans(document);

  for (const element of document.querySelectorAll('[data-tag]')) {
    const label = tagLabel(element.getAttribute('data-tag'));
    if (!label) continue;
    const marker = document.createTextNode(`${label} `);
    if (['IMG', 'OBJECT', 'BR', 'HR'].includes(element.nodeName)) element.before(marker);
    else element.insertBefore(marker, element.firstChild);
  }

  // Empty <object> attachments would be dropped as blank, so turn them into links.
  for (const attachment of document.querySelectorAll('object')) {
    const link = document.createElement('a');
    link.setAttribute('href', attachment.getAttribute('data') || '');
    link.textContent = `attachment: ${attachment.getAttribute('data-attachment') || 'file'}`;
    attachment.replaceWith(link);
  }

  return turndown
    .turndown(document.body)
    .replace(OBJECT_REPLACEMENT, '')
    .replace(/^[ \t\u00A0]+$/gm, '')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}
