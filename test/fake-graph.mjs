// Fake Microsoft Graph for tests: a tiny OneNote model plus a request log.
export function createFakeGraph({ sections, pages, encrypted = [] } = {}) {
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
  };
  const requests = [];

  const json = (status, body) =>
    new Response(JSON.stringify(body), { status, headers: { 'Content-Type': 'application/json' } });
  const encryptedError = () => json(403, { error: { code: '20185', message: 'Encrypted sections are not accessible.' } });

  async function fetchImpl(href, { method = 'GET', body } = {}) {
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
    if ((match = path.match(/^\/me\/onenote\/sections\/([^/]+)\/pages$/))) {
      const sectionId = decodeURIComponent(match[1]);
      if (method === 'POST') return json(201, { id: 'page-new', title: 'created', parentSection: { id: sectionId } });
      if (state.encrypted.has(sectionId)) return encryptedError();
      return json(200, { value: Object.values(state.pages).filter((p) => p.parentSection.id === sectionId) });
    }
    if ((match = path.match(/^\/me\/onenote\/pages\/([^/]+)(\/content)?$/))) {
      const page = state.pages[decodeURIComponent(match[1])];
      if (!page) return json(404, { error: { code: '20102', message: 'The requested resource does not exist.' } });
      if (match[2] && method === 'GET') return new Response(`<html><head><title>${page.title}</title></head><body><p>body of ${page.id}</p></body></html>`, { status: 200 });
      if (match[2] && method === 'PATCH') return new Response(null, { status: 204 });
      if (!match[2] && method === 'DELETE') return new Response(null, { status: 204 });
      if (!match[2] && method === 'GET') return json(200, page);
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
