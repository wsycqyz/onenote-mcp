// Fake Microsoft Graph for tests: a tiny OneNote model plus a request log.

// A OneNote-style source page: an outline <div>, a self-closing attachment followed by more
// content, a positioned image and a bare text node.
export const IMAGE_URL = "https://graph.microsoft.com/v1.0/users('me@outlook.com')/onenote/resources/res-1/$value";
export const FILE_URL = "https://graph.microsoft.com/v1.0/users('me@outlook.com')/onenote/resources/res-2/$value";
export const SOURCE_HTML =
  '<html><head><title>Notes</title></head><body data-absolute-enabled="true">' +
  '<div data-id="_default" style="position:absolute;left:48px;top:120px;width:624px">' +
  '<p style="margin-top:0pt">alpha <span style="font-weight:bold">key-123</span></p><br />' +
  '<table><tr><td>a</td><td>b</td></tr></table><p>beta p***w0rd &lt;host IP&gt;</p>' +
  `<object data-attachment="doc.pdf" type="application/pdf" data="${FILE_URL}" /><p>after attachment</p></div>` +
  `<img style="position:absolute;left:10px;top:400px" src="${IMAGE_URL}" data-fullres-src="${IMAGE_URL}" ` +
  'data-src-type="image/png" alt="screenshot text" width="100" height="50" />' +
  '<div style="position:absolute;left:48px;top:600px">gamma text node</div></body></html>';

// Splits a multipart page-creation body into its parts (binary-safe).
export function parseMultipart(body) {
  const raw = Buffer.from(body).toString('latin1');
  const boundary = raw.slice(2, raw.indexOf('\r\n'));
  return raw
    .split(`--${boundary}`)
    .slice(1, -1)
    .map((chunk) => {
      const split = chunk.indexOf('\r\n\r\n');
      const head = chunk.slice(0, split);
      return {
        name: /name="([^"]+)"/.exec(head)[1],
        contentType: /Content-Type: ([^\r\n]+)/.exec(head)[1],
        bytes: Buffer.from(chunk.slice(split + 4, -2), 'latin1'),
      };
    });
}

export function createFakeGraph({ sections, pages, resources = {}, encrypted = [] } = {}) {
  const state = {
    encrypted: new Set(encrypted),
    sections: sections ?? [
      { id: 'sec-ai', displayName: 'AI Manage', parentNotebook: { id: 'nb-1', displayName: 'Personal' } },
      { id: 'sec-journal', displayName: 'Journal', parentNotebook: { id: 'nb-1', displayName: 'Personal' } },
    ],
    pages: pages ?? {
      'page-ai': { id: 'page-ai', title: 'AI notes', parentSection: { id: 'sec-ai', displayName: 'AI Manage' } },
      'page-journal': { id: 'page-journal', title: 'Diary', parentSection: { id: 'sec-journal', displayName: 'Journal' } },
    },
    resources,
    created: [],
  };
  const requests = [];

  const json = (status, body) =>
    new Response(JSON.stringify(body), { status, headers: { 'Content-Type': 'application/json' } });
  const encryptedError = () => json(403, { error: { code: '20185', message: 'Encrypted sections are not accessible.' } });

  // Like OneNote: uploaded parts become page resources, and name:<part> references become their URLs.
  function createPage(sectionId, body, contentType) {
    const parts = contentType.startsWith('multipart/') ? parseMultipart(body) : [];
    const posted = parts.length ? parts.find((p) => p.name === 'Presentation').bytes.toString('utf8') : String(body);
    const id = `page-new-${state.created.length + 1}`;
    let html = posted;
    for (const part of parts.filter((p) => p.name !== 'Presentation')) {
      const resourceId = `${id}-${part.name}`;
      state.resources[resourceId] = { bytes: part.bytes, contentType: part.contentType };
      html = html.replaceAll(`name:${part.name}`, `https://graph.microsoft.com/v1.0/users('me@outlook.com')/onenote/resources/${resourceId}/$value`);
    }
    const title = /<title>([^<]*)<\/title>/.exec(html)?.[1] ?? '';
    const page = { id, title, html, parentSection: { id: sectionId, displayName: 'AI Manage' } };
    state.pages[id] = page;
    state.created.push({ ...page, posted, parts });
    return json(201, { id, title, parentSection: page.parentSection });
  }

  async function fetchImpl(href, { method = 'GET', headers = {}, body } = {}) {
    const url = new URL(href);
    requests.push({ method, path: url.pathname, search: url.search, body });
    const path = url.pathname.replace(/^\/v1\.0/, '');
    let match;
    if (method === 'GET' && path === '/me') return json(200, { displayName: 'Me', userPrincipalName: 'me@outlook.com' });
    if (method === 'GET' && path === '/me/onenote/notebooks') return json(200, { value: [{ id: 'nb-1', displayName: 'Personal' }] });
    if (method === 'GET' && path === '/me/onenote/sections') return json(200, { value: state.sections });
    if (method === 'GET' && path === '/me/onenote/pages') {
      return state.encrypted.size ? encryptedError() : json(200, { value: Object.values(state.pages) });
    }
    if (method === 'GET' && (match = path.match(/^\/users\('[^']+'\)\/onenote\/resources\/([^/]+)\/\$value$/))) {
      const resource = state.resources[decodeURIComponent(match[1])];
      if (!resource) return json(404, { error: { code: '20102', message: 'No such resource.' } });
      return new Response(resource.bytes, { status: 200, headers: { 'Content-Type': resource.contentType } });
    }
    if ((match = path.match(/^\/me\/onenote\/sections\/([^/]+)\/pages$/))) {
      const sectionId = decodeURIComponent(match[1]);
      if (method === 'POST') return createPage(sectionId, body, headers['Content-Type'] ?? '');
      if (state.encrypted.has(sectionId)) return encryptedError();
      return json(200, { value: Object.values(state.pages).filter((p) => p.parentSection.id === sectionId) });
    }
    if ((match = path.match(/^\/me\/onenote\/pages\/([^/]+)(\/content)?$/))) {
      const page = state.pages[decodeURIComponent(match[1])];
      if (!page) return json(404, { error: { code: '20102', message: 'The requested resource does not exist.' } });
      if (match[2] && method === 'GET') {
        const html = page.html ?? `<html><head><title>${page.title}</title></head><body><p>body of ${page.id}</p></body></html>`;
        return new Response(html, { status: 200 });
      }
      if (match[2] && method === 'PATCH') return new Response(null, { status: 204 });
      if (!match[2] && method === 'DELETE') return new Response(null, { status: 204 });
      if (!match[2] && method === 'GET') {
        const { html: _html, ...meta } = page;
        return json(200, meta);
      }
    }
    return json(400, { error: { code: 'unexpected', message: `${method} ${path}` } });
  }

  return { state, requests, fetchImpl, writes: () => requests.filter((r) => r.method !== 'GET') };
}

export const fakeAuth = Object.freeze({
  getAccessToken: async () => 'test-token',
  getAccount: async () => ({ username: 'me@outlook.com' }),
  status: async () => ({ signedIn: true, account: 'me@outlook.com' }),
  startSignIn: async () => {
    throw new Error('not used in tests');
  },
});
