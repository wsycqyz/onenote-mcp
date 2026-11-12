import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { attachResources, blockPreview, composeBody, extractBlocks, verifyCopy } from './compose.mjs';
import { ACCESS_WRITE_SECTION } from './config.mjs';
import { CONTENT_FORMATS, escapeHtml, htmlToMarkdown, toXhtmlFragment } from './convert.mjs';
import { createGraph, encodeId } from './graph.mjs';
import { createSectionGuard, sameName, SECTIONS_QUERY, WriteDeniedError } from './guard.mjs';

export const SERVER_VERSION = '0.3.0';
export const READ_TOOLS = Object.freeze(['authStatus', 'authenticate', 'listNotebooks', 'listSections', 'listPages', 'getPage', 'searchPages']);
export const WRITE_TOOLS = Object.freeze(['createPage', 'composePage', 'appendToPage', 'updatePage', 'deletePage']);

const PAGE_FIELDS = 'id,title,createdDateTime,lastModifiedDateTime,links';
const PARENTS = 'parentSection($select=id,displayName),parentNotebook($select=id,displayName)';
const UPDATE_TARGET = /^(body|title|#?[A-Za-z0-9:{}._-]{1,200})$/;
// OneNote API error codes.
const ENCRYPTED_SECTION = '20185';
const TOO_MANY_SECTIONS = '20266';

const READ = Object.freeze({ readOnlyHint: true, openWorldHint: true });

function ok(data) {
  return { content: [{ type: 'text', text: typeof data === 'string' ? data : JSON.stringify(data, null, 2) }] };
}

function safe(handler) {
  return async (...args) => {
    try {
      return await handler(...args);
    } catch (error) {
      return { isError: true, content: [{ type: 'text', text: error?.message ?? String(error) }] };
    }
  };
}

function odataString(value) {
  return `'${String(value).replace(/'/g, "''")}'`;
}

function summarizePage(page) {
  return {
    id: page.id,
    title: page.title,
    notebook: page.parentNotebook?.displayName,
    section: page.parentSection?.displayName,
    sectionId: page.parentSection?.id,
    level: page.level,
    order: page.order,
    created: page.createdDateTime,
    lastModified: page.lastModifiedDateTime,
  };
}

/** Builds the OneNote MCP server. Write tools are registered only in write-section mode. */
export function createOneNoteServer({ config, auth, fetchImpl, sleep }) {
  const pause = sleep ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms)));
  const writable = config.mode === ACCESS_WRITE_SECTION;
  const { reader, writer } = createGraph({ getToken: () => auth.getAccessToken(), fetchImpl, allowWrites: writable });
  const guard = writable
    ? createSectionGuard({ reader, writer, sectionName: config.writeSection, notebookName: config.writeNotebook })
    : null;
  const writeScope = writable
    ? `the "${config.writeSection}" section${config.writeNotebook ? ` of notebook "${config.writeNotebook}"` : ''}`
    : null;
  const accessText = writable
    ? `Read access to all notebooks; write access only to ${writeScope}.`
    : 'Read-only: this server cannot create, change or delete anything in OneNote.';

  const server = new McpServer(
    { name: 'onenote', version: SERVER_VERSION },
    { instructions: `Personal OneNote (Microsoft account). ${accessText}` },
  );

  async function writableSectionId() {
    if (!guard) return null;
    try {
      return (await guard.resolveSection()).id;
    } catch {
      return null;
    }
  }

  async function findSections({ sectionId, sectionName, notebookName }) {
    const sections = await reader.getAll(SECTIONS_QUERY);
    return sections.filter(
      (s) =>
        (!sectionId || s.id === sectionId) &&
        (!sectionName || sameName(s.displayName, sectionName)) &&
        (!notebookName || sameName(s.parentNotebook?.displayName, notebookName)),
    );
  }

  // A new page can take a few seconds before Graph returns it.
  async function verifyNewPage(pageId, used) {
    if (!pageId) return { ok: false, error: 'Graph returned no page ID.' };
    for (let attempt = 1; ; attempt++) {
      try {
        return verifyCopy(used, await reader.getText(`/me/onenote/pages/${encodeId(pageId, 'pageId')}/content`));
      } catch (error) {
        if (error.status !== 404 || attempt >= 8) return { ok: false, error: error.message };
        await pause(2500);
      }
    }
  }

  server.registerTool(
    'authStatus',
    {
      title: 'OneNote sign-in and access status',
      description: `Show the signed-in personal Microsoft account and what this server may change. ${accessText}`,
      annotations: READ,
    },
    safe(async () => {
      const status = { access: config.mode, writableScope: writeScope, warnings: config.warnings, ...(await auth.status()) };
      if (status.signedIn) {
        try {
          const me = await reader.get('/me?$select=displayName,userPrincipalName');
          status.profile = { displayName: me.displayName, userPrincipalName: me.userPrincipalName };
        } catch (error) {
          status.profileError = error.message;
        }
        if (guard) {
          try {
            status.writeTarget = await guard.resolveSection();
          } catch (error) {
            status.writeTargetError = error.message;
          }
        }
      }
      return ok(status);
    }),
  );

  server.registerTool(
    'authenticate',
    {
      title: 'Sign in to OneNote',
      description:
        'Start a device code sign-in with the PERSONAL Microsoft account (work or school accounts are rejected). ' +
        'Returns a URL and a code that the user must enter in a browser; call authStatus afterwards to confirm.',
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: true },
    },
    safe(async () => {
      const account = await auth.getAccount();
      if (account) {
        try {
          await auth.getAccessToken();
          return ok(`Already signed in as ${account.username} (personal Microsoft account).`);
        } catch {
          // Expired: start a new sign-in below.
        }
      }
      const prompt = await auth.startSignIn();
      return ok(
        `${prompt.message}\n\nAsk the user to open ${prompt.verificationUri}, enter the code ${prompt.userCode} and sign in ` +
          `with their PERSONAL Microsoft account before ${prompt.expiresAt}. Then call authStatus to confirm.`,
      );
    }),
  );

  server.registerTool(
    'listNotebooks',
    { title: 'List notebooks', description: 'List all OneNote notebooks.', annotations: READ },
    safe(async () => {
      const notebooks = await reader.getAll(
        '/me/onenote/notebooks?$select=id,displayName,isDefault,isShared,createdDateTime,lastModifiedDateTime',
      );
      return ok(
        notebooks.map((n) => ({
          id: n.id,
          name: n.displayName,
          isDefault: n.isDefault,
          isShared: n.isShared,
          lastModified: n.lastModifiedDateTime,
        })),
      );
    }),
  );

  server.registerTool(
    'listSections',
    {
      title: 'List sections',
      description: 'List OneNote sections (including those inside section groups), optionally for one notebook.',
      inputSchema: {
        notebookId: z.string().optional().describe('Only sections of this notebook ID.'),
        notebookName: z.string().optional().describe('Only sections of the notebook with this name.'),
      },
      annotations: READ,
    },
    safe(async ({ notebookId, notebookName }) => {
      const [sections, writableId] = await Promise.all([reader.getAll(SECTIONS_QUERY), writableSectionId()]);
      const result = sections
        .filter(
          (s) =>
            (!notebookId || s.parentNotebook?.id === notebookId) &&
            (!notebookName || sameName(s.parentNotebook?.displayName, notebookName)),
        )
        .map((s) => ({
          id: s.id,
          name: s.displayName,
          notebook: s.parentNotebook?.displayName,
          notebookId: s.parentNotebook?.id,
          sectionGroup: s.parentSectionGroup?.displayName,
          isDefault: s.isDefault,
          lastModified: s.lastModifiedDateTime,
          ...(writable ? { writable: s.id === writableId } : {}),
        }));
      return ok(result);
    }),
  );

  server.registerTool(
    'listPages',
    {
      title: 'List pages in a section',
      description: 'List the pages of one section, by section ID or by section name (plus notebook name if the name is not unique).',
      inputSchema: {
        sectionId: z.string().optional().describe('Section ID (preferred).'),
        sectionName: z.string().optional().describe('Section name, used when sectionId is not given.'),
        notebookName: z.string().optional().describe('Notebook name, to disambiguate sectionName.'),
        top: z.number().int().min(1).max(1000).optional().describe('Maximum pages to return (default 200).'),
      },
      annotations: READ,
    },
    safe(async ({ sectionId, sectionName, notebookName, top = 200 }) => {
      if (!sectionId && !sectionName) throw new Error('Provide sectionId or sectionName.');
      const matches = await findSections({ sectionId, sectionName, notebookName });
      if (matches.length === 0) throw new Error('No matching section found. Use listSections to see section names and IDs.');
      if (matches.length > 1) {
        const options = matches.map((s) => `${s.parentNotebook?.displayName} / ${s.displayName}: ${s.id}`).join('\n');
        throw new Error(`Several sections match; pass sectionId or notebookName:\n${options}`);
      }
      const [section] = matches;
      const pages = await reader.getAll(
        `/me/onenote/sections/${encodeId(section.id, 'sectionId')}/pages?$select=id,title,createdDateTime,lastModifiedDateTime,level,order&pagelevel=true&$top=100`,
        { max: top },
      );
      pages.sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
      return ok({
        notebook: section.parentNotebook?.displayName,
        section: section.displayName,
        sectionId: section.id,
        pages: pages.map(summarizePage),
      });
    }),
  );

  server.registerTool(
    'getPage',
    {
      title: 'Read a page',
      description:
        'Read one page by ID. format "markdown" (default) is best for reading; "html" returns the raw OneNote HTML ' +
        'with element IDs, which updatePage can target; "blocks" lists numbered blocks and a blocksVersion for composePage.',
      inputSchema: {
        pageId: z.string().describe('Page ID from listPages or searchPages.'),
        format: z.enum(['markdown', 'html', 'blocks']).optional().describe('Output format (default markdown).'),
        maxChars: z.number().int().min(1000).max(500000).optional().describe('Truncate the body after this many characters (default 60000).'),
      },
      annotations: READ,
    },
    safe(async ({ pageId, format = 'markdown', maxChars = 60000 }) => {
      const id = encodeId(pageId, 'pageId');
      const [meta, html] = await Promise.all([
        reader.get(`/me/onenote/pages/${id}?$select=${PAGE_FIELDS}&$expand=${PARENTS}`),
        reader.getText(`/me/onenote/pages/${id}/content${format === 'html' ? '?includeIDs=true' : ''}`),
      ]);
      let body;
      let blockInfo = {};
      if (format === 'blocks') {
        const { blocks, version } = extractBlocks(html);
        blockInfo = { blocksVersion: version, blockCount: blocks.length };
        body = blocks.map((block) => `[${block.ref}] ${block.type}: ${blockPreview(block)}`).join('\n');
      } else {
        body = format === 'html' ? html : htmlToMarkdown(html);
      }
      if (body.length > maxChars) {
        body = `${body.slice(0, maxChars)}\n\n[truncated: ${body.length - maxChars} more characters; raise maxChars to read more]`;
      }
      const header = {
        ...summarizePage(meta),
        webUrl: meta.links?.oneNoteWebUrl?.href,
        ...blockInfo,
        ...(writable ? { writable: meta.parentSection?.id === (await writableSectionId()) } : {}),
      };
      return ok(`${JSON.stringify(header, null, 2)}\n\n---\n\n${body}`);
    }),
  );

  server.registerTool(
    'searchPages',
    {
      title: 'Search pages by title',
      description: 'Find pages whose title contains the query (case-insensitive), optionally within a notebook or section.',
      inputSchema: {
        query: z.string().min(1).describe('Text to look for in page titles.'),
        notebookName: z.string().optional().describe('Only search this notebook.'),
        sectionName: z.string().optional().describe('Only search sections with this name.'),
        top: z.number().int().min(1).max(500).optional().describe('Maximum results (default 50).'),
      },
      annotations: READ,
    },
    safe(async ({ query, notebookName, sectionName, top = 50 }) => {
      const filter = encodeURIComponent(`contains(tolower(title),${odataString(query.toLowerCase())})`);
      const options = `$filter=${filter}&$select=id,title,createdDateTime,lastModifiedDateTime&$expand=${PARENTS}&$top=100`;
      const skippedSections = [];
      const searchSections = async (sections) => {
        const found = [];
        for (const section of sections) {
          if (found.length >= top) break;
          try {
            found.push(
              ...(await reader.getAll(`/me/onenote/sections/${encodeId(section.id, 'sectionId')}/pages?${options}`, {
                max: top - found.length,
              })),
            );
          } catch (error) {
            if (error.code !== ENCRYPTED_SECTION) throw error;
            skippedSections.push(`${section.parentNotebook?.displayName} / ${section.displayName} (password protected)`);
          }
        }
        return found;
      };

      let pages;
      if (notebookName || sectionName) {
        pages = await searchSections(await findSections({ sectionName, notebookName }));
      } else {
        try {
          pages = await reader.getAll(`/me/onenote/pages?${options}`, { max: top });
        } catch (error) {
          // One account-wide query fails if any section is password protected or there are too many sections.
          if (error.code !== ENCRYPTED_SECTION && error.code !== TOO_MANY_SECTIONS) throw error;
          pages = await searchSections(await findSections({}));
        }
      }
      const result = { results: pages.slice(0, top).map(summarizePage) };
      if (skippedSections.length) result.skippedSections = skippedSections;
      return ok(result);
    }),
  );

  if (!guard) return server;

  const contentFormat = z.enum(CONTENT_FORMATS).optional().describe('Format of content: markdown (default), html or text.');

  server.registerTool(
    'createPage',
    {
      title: `Create a page in ${writeScope}`,
      description: `Create a new page. Pages can only be created in ${writeScope}; every other section is read-only.`,
      inputSchema: {
        title: z.string().min(1).max(500).describe('Page title.'),
        content: z.string().max(1_000_000).describe('Page body.'),
        contentFormat,
        sectionId: z.string().optional().describe(`Optional; must be the ID of ${writeScope}.`),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
    },
    safe(async ({ title, content, contentFormat: format = 'markdown', sectionId }) => {
      const { page, section } = await guard.createPage({ title, bodyXhtml: toXhtmlFragment(content, format), sectionId });
      return ok({
        created: true,
        id: page?.id,
        title: page?.title ?? title,
        notebook: section.notebookName,
        section: section.name,
        webUrl: page?.links?.oneNoteWebUrl?.href,
      });
    }),
  );

  server.registerTool(
    'composePage',
    {
      title: `Compose a page in ${writeScope} from another page's blocks`,
      description:
        `Create a page in ${writeScope} that copies the blocks of a source page verbatim into a new layout, for ` +
        'reformatting without retyping. First call getPage with format "blocks" for block numbers and blocksVersion. ' +
        'Every block must be used exactly once or listed in omit. Images and attachments are copied. The page starts ' +
        'with a line naming the source, and the result reports whether the new page text matches the source.',
      inputSchema: {
        title: z.string().min(1).max(500).describe('Title of the new page.'),
        sourcePageId: z.string().describe('ID of the page to copy blocks from.'),
        sourceVersion: z.string().describe('blocksVersion from getPage with format "blocks".'),
        sections: z
          .array(
            z.object({
              heading: z.string().max(300).optional().describe('Heading added above these blocks.'),
              headingLevel: z.number().int().min(1).max(3).optional().describe('Heading level (default 2).'),
              refs: z
                .array(z.union([z.number().int().min(1), z.string()]))
                .optional()
                .describe('Block numbers or ranges like "3-7", in the order to copy them.'),
            }),
          )
          .min(1)
          .max(300),
        omit: z.array(z.number().int().min(1)).optional().describe('Blocks deliberately left out; report them to the user.'),
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
    },
    safe(async ({ title, sourcePageId, sourceVersion, sections, omit = [] }) => {
      const id = encodeId(sourcePageId, 'sourcePageId');
      const [meta, html] = await Promise.all([
        reader.get(`/me/onenote/pages/${id}?$select=${PAGE_FIELDS}&$expand=${PARENTS}`),
        reader.getText(`/me/onenote/pages/${id}/content`),
      ]);
      const { blocks, version } = extractBlocks(html);
      if (version !== sourceVersion) {
        throw new Error('The source page changed since it was read (blocksVersion differs). Call getPage with format "blocks" again.');
      }
      const { bodyHtml, used, omitted } = composeBody({ blocks, sections, omit });
      const { html: body, parts } = await attachResources(bodyHtml, (url) => reader.getBinary(url));
      const source = [meta.parentNotebook?.displayName, meta.parentSection?.displayName, meta.title].join(' / ');
      const sourceLine = `<p style="color:#7f7f7f">Source: ${escapeHtml(source)} (last modified ${escapeHtml(meta.lastModifiedDateTime ?? 'unknown')})</p>`;
      const { page, section } = await guard.createPage({ title, bodyXhtml: sourceLine + body, parts });
      return ok({
        created: true,
        id: page?.id,
        title,
        notebook: section.notebookName,
        section: section.name,
        webUrl: page?.links?.oneNoteWebUrl?.href,
        blocksCopied: used.length,
        omitted,
        filesCopied: parts.length,
        verification: await verifyNewPage(page?.id, used),
      });
    }),
  );

  server.registerTool(
    'appendToPage',
    {
      title: `Append to a page in ${writeScope}`,
      description: `Append content to the end of a page. Only pages in ${writeScope} can be changed.`,
      inputSchema: {
        pageId: z.string().describe('Page ID.'),
        content: z.string().min(1).max(1_000_000).describe('Content to append.'),
        contentFormat,
      },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: true },
    },
    safe(async ({ pageId, content, contentFormat: format = 'markdown' }) => {
      const { page } = await guard.patchPage(pageId, [
        { target: 'body', action: 'append', content: toXhtmlFragment(content, format) },
      ]);
      return ok({ updated: true, id: page.id, title: page.title });
    }),
  );

  server.registerTool(
    'updatePage',
    {
      title: `Update a page in ${writeScope}`,
      description:
        `Apply OneNote PATCH commands to a page. Only pages in ${writeScope} can be changed. ` +
        'target is "body", "title" (action must be "replace", content is plain text), "#<data-id>", or a generated ' +
        'element id from getPage with format "html" (for example "p:{...}{12}"). action is append, insert ' +
        '(with position before or after), prepend or replace.',
      inputSchema: {
        pageId: z.string().describe('Page ID.'),
        commands: z
          .array(
            z.object({
              target: z.string().regex(UPDATE_TARGET, 'Invalid target.'),
              action: z.enum(['append', 'insert', 'prepend', 'replace']),
              position: z.enum(['before', 'after']).optional(),
              content: z.string().max(1_000_000),
              contentFormat,
            }),
          )
          .min(1)
          .max(50),
      },
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: false, openWorldHint: true },
    },
    safe(async ({ pageId, commands }) => {
      const patch = commands.map((command) => {
        if (command.target === 'title') {
          if (command.action !== 'replace') throw new Error('The title target only supports action "replace".');
          return { target: 'title', action: 'replace', content: command.content };
        }
        const result = {
          target: command.target,
          action: command.action,
          content: toXhtmlFragment(command.content, command.contentFormat ?? 'markdown'),
        };
        if (command.position) result.position = command.position;
        return result;
      });
      const { page } = await guard.patchPage(pageId, patch);
      return ok({ updated: true, id: page.id, title: page.title, commands: patch.length });
    }),
  );

  server.registerTool(
    'deletePage',
    {
      title: `Delete a page in ${writeScope}`,
      description: `Delete a page. Only pages in ${writeScope} can be deleted. expectedTitle must match the page title.`,
      inputSchema: {
        pageId: z.string().describe('Page ID.'),
        expectedTitle: z.string().describe('Current title of the page, as a safety check.'),
      },
      annotations: { readOnlyHint: false, destructiveHint: true, idempotentHint: true, openWorldHint: true },
    },
    safe(async ({ pageId, expectedTitle }) => {
      const { page } = await guard.assertPageWritable(pageId);
      if ((page.title ?? '').trim() !== expectedTitle.trim()) {
        throw new WriteDeniedError(`Not deleted: the page title is "${page.title}", not "${expectedTitle}".`);
      }
      await guard.deletePage(pageId);
      return ok({ deleted: true, id: page.id, title: page.title });
    }),
  );

  return server;
}
