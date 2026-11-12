import { randomUUID } from 'node:crypto';

export const GRAPH_ORIGIN = 'https://graph.microsoft.com';
export const GRAPH_ROOT = `${GRAPH_ORIGIN}/v1.0`;

const ID_PATTERN = /^[A-Za-z0-9!._-]{1,512}$/;

// The only request shapes the write capability may send.
const WRITE_ROUTES = Object.freeze({
  POST: /^\/v1\.0\/me\/onenote\/sections\/[A-Za-z0-9!._-]+\/pages$/,
  PATCH: /^\/v1\.0\/me\/onenote\/pages\/[A-Za-z0-9!._-]+\/content$/,
  DELETE: /^\/v1\.0\/me\/onenote\/pages\/[A-Za-z0-9!._-]+$/,
});

// /me, or OneNote paths. @odata.nextLink URLs may use users('...')/onenote/... instead of /me/onenote/...
const READ_PATH = /^\/v1\.0\/(me|(me|users\/[^/]+|users\('[^/]+'\))\/onenote\/.+)$/;

export class GraphError extends Error {
  constructor(status, code, message) {
    super(message);
    this.name = 'GraphError';
    this.status = status;
    this.code = code;
  }
}

export function assertValidId(id, label = 'id') {
  if (typeof id !== 'string' || !ID_PATTERN.test(id)) {
    throw new Error(`Invalid OneNote ${label}: ${JSON.stringify(id)}`);
  }
  return id;
}

export function encodeId(id, label) {
  return encodeURIComponent(assertValidId(id, label));
}

const PART_NAME = /^[A-Za-z0-9]{1,64}$/;
const MIME_TYPE = /^[\w.+-]+\/[\w.+-]+$/;

/** Builds the multipart body OneNote expects: a "Presentation" HTML part plus named file parts. */
export function buildMultipart(html, parts) {
  const boundary = `onenote-mcp-${randomUUID()}`;
  const chunks = [];
  const add = (name, contentType, data) => {
    chunks.push(Buffer.from(`--${boundary}\r\nContent-Disposition: form-data; name="${name}"\r\nContent-Type: ${contentType}\r\n\r\n`));
    chunks.push(Buffer.isBuffer(data) ? data : Buffer.from(data));
    chunks.push(Buffer.from('\r\n'));
  };
  add('Presentation', 'application/xhtml+xml', html);
  for (const part of parts) {
    if (!PART_NAME.test(part.name)) throw new Error(`Invalid multipart part name: ${JSON.stringify(part.name)}`);
    add(part.name, MIME_TYPE.test(part.contentType ?? '') ? part.contentType : 'application/octet-stream', part.bytes);
  }
  chunks.push(Buffer.from(`--${boundary}--\r\n`));
  return { body: Buffer.concat(chunks), contentType: `multipart/form-data; boundary=${boundary}` };
}

function resolveUrl(method, pathOrUrl) {
  const url = new URL(pathOrUrl.startsWith('https://') ? pathOrUrl : GRAPH_ROOT + pathOrUrl);
  if (url.origin !== GRAPH_ORIGIN) throw new Error(`Refusing request to ${url.origin}: only Microsoft Graph is allowed.`);
  if (method === 'GET') {
    if (!READ_PATH.test(url.pathname)) throw new Error(`Refusing to read ${url.pathname}: only /me and OneNote paths are allowed.`);
  } else {
    const route = WRITE_ROUTES[method];
    if (!route || !route.test(url.pathname) || url.search) {
      throw new Error(`Refusing ${method} ${url.pathname}: not an allowed OneNote write.`);
    }
  }
  return url;
}

function retryDelayMs(response, attempt) {
  const header = Number(response.headers.get('retry-after'));
  const seconds = Number.isFinite(header) && header > 0 ? header : 2 ** attempt;
  return Math.min(seconds, 30) * 1000;
}

async function toGraphError(response, method, url) {
  let code = '';
  let detail = response.statusText;
  try {
    const body = JSON.parse(await response.text());
    code = body?.error?.code ?? '';
    detail = body?.error?.message ?? detail;
  } catch {
    // Non-JSON error body.
  }
  return new GraphError(
    response.status,
    code,
    `Graph ${method} ${url.pathname} failed: HTTP ${response.status}${code ? ` (${code})` : ''}: ${detail}`,
  );
}

/**
 * Creates a Microsoft Graph client split into two capabilities. `reader` can only
 * send GET requests. `writer` exists only when `allowWrites` is true, and can only
 * create pages in a section, patch page content, or delete a page; callers must
 * check that the target is allowed (see guard.mjs).
 */
export function createGraph({ getToken, fetchImpl = globalThis.fetch, allowWrites = false, maxRetries = 3, sleep } = {}) {
  const wait = sleep ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms)));

  async function send(method, pathOrUrl, { body, contentType, responseType = 'json' } = {}) {
    if (method !== 'GET' && !allowWrites) {
      throw new Error(`Blocked ${method} request: this OneNote MCP server is read-only.`);
    }
    const url = resolveUrl(method, pathOrUrl);
    for (let attempt = 0; ; attempt++) {
      const headers = { Authorization: `Bearer ${await getToken()}` };
      if (contentType) headers['Content-Type'] = contentType;
      const response = await fetchImpl(url.toString(), { method, headers, body });
      // Only GETs are retried on 503/504; a write may already have been applied.
      const retryable = response.status === 429 || (method === 'GET' && (response.status === 503 || response.status === 504));
      if (retryable && attempt < maxRetries) {
        await wait(retryDelayMs(response, attempt));
        continue;
      }
      if (!response.ok) throw await toGraphError(response, method, url);
      if (response.status === 204) return null;
      if (responseType === 'binary') {
        return {
          bytes: Buffer.from(await response.arrayBuffer()),
          contentType: response.headers.get('content-type') ?? 'application/octet-stream',
        };
      }
      const text = await response.text();
      if (responseType === 'text') return text;
      return text ? JSON.parse(text) : null;
    }
  }

  const reader = Object.freeze({
    get: (path) => send('GET', path),
    getText: (path) => send('GET', path, { responseType: 'text' }),
    getBinary: (path) => send('GET', path, { responseType: 'binary' }),
    async getAll(path, { max = 5000 } = {}) {
      const items = [];
      let next = path;
      while (next && items.length < max) {
        const page = await send('GET', next);
        items.push(...(page?.value ?? []));
        next = page?.['@odata.nextLink'];
      }
      return items.slice(0, max);
    },
  });

  if (!allowWrites) return { reader, writer: null };

  const writer = Object.freeze({
    // parts: files referenced from the page as "name:<part name>" (images, attachments).
    createPage: async (sectionId, html, parts = []) =>
      send('POST', `/me/onenote/sections/${encodeId(sectionId, 'sectionId')}/pages`, {
        ...(parts.length ? buildMultipart(html, parts) : { body: html, contentType: 'application/xhtml+xml' }),
      }),
    patchPageContent: async (pageId, commands) =>
      send('PATCH', `/me/onenote/pages/${encodeId(pageId, 'pageId')}/content`, {
        body: JSON.stringify(commands),
        contentType: 'application/json',
      }),
    deletePage: async (pageId) => send('DELETE', `/me/onenote/pages/${encodeId(pageId, 'pageId')}`),
  });
  return { reader, writer };
}
