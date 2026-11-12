import { buildPageXhtml } from './convert.mjs';
import { assertValidId, encodeId } from './graph.mjs';

export class WriteDeniedError extends Error {
  constructor(message) {
    super(message);
    this.name = 'WriteDeniedError';
  }
}

export const SECTIONS_QUERY =
  '/me/onenote/sections?$select=id,displayName,isDefault,lastModifiedDateTime' +
  '&$expand=parentNotebook($select=id,displayName),parentSectionGroup($select=id,displayName)';

export function sameName(a, b) {
  return typeof a === 'string' && typeof b === 'string' && a.trim().toLocaleLowerCase() === b.trim().toLocaleLowerCase();
}

/**
 * Restricts every write to the pages of one section, identified by name (and
 * optionally by notebook name). The section must match exactly once. The
 * resolved section is cached briefly for display only; writes resolve it again.
 */
export function createSectionGuard({ reader, writer, sectionName, notebookName = null, cacheMs = 60_000, now = Date.now }) {
  if (!writer) throw new Error('createSectionGuard requires a writer.');
  if (!sectionName) throw new Error('createSectionGuard requires a section name.');
  const where = notebookName ? ` in notebook "${notebookName}"` : '';
  let cached = null;

  async function resolveSection({ fresh = false } = {}) {
    if (!fresh && cached && now() - cached.at < cacheMs) return cached.section;
    const sections = await reader.getAll(SECTIONS_QUERY);
    const matches = sections.filter(
      (s) => sameName(s.displayName, sectionName) && (!notebookName || sameName(s.parentNotebook?.displayName, notebookName)),
    );
    if (matches.length === 0) {
      throw new WriteDeniedError(
        `No OneNote section named "${sectionName}"${where} exists, so nothing can be written. Ask the user to create it.`,
      );
    }
    if (matches.length > 1) {
      const notebooks = matches.map((s) => s.parentNotebook?.displayName ?? '?').join(', ');
      throw new WriteDeniedError(
        `${matches.length} sections are named "${sectionName}" (notebooks: ${notebooks}); writes stay disabled until ONENOTE_WRITE_NOTEBOOK selects one.`,
      );
    }
    const [match] = matches;
    const section = Object.freeze({
      id: match.id,
      name: match.displayName,
      notebookId: match.parentNotebook?.id ?? null,
      notebookName: match.parentNotebook?.displayName ?? null,
    });
    cached = { at: now(), section };
    return section;
  }

  function describe(section) {
    return `"${section.name}"${section.notebookName ? ` (notebook "${section.notebookName}")` : ''}`;
  }

  async function assertPageWritable(pageId) {
    assertValidId(pageId, 'pageId');
    // Writes never use the cached section, so a rename or a duplicate name takes effect at once.
    const section = await resolveSection({ fresh: true });
    const page = await reader.get(
      `/me/onenote/pages/${encodeId(pageId, 'pageId')}?$select=id,title&$expand=parentSection($select=id,displayName)`,
    );
    if (!page?.id || page.parentSection?.id !== section.id) {
      throw new WriteDeniedError(
        `Write denied: page "${page?.title ?? pageId}" is in section "${page?.parentSection?.displayName ?? 'unknown'}". ` +
          `Only pages in the ${describe(section)} section can be changed.`,
      );
    }
    return { page, section };
  }

  return Object.freeze({
    resolveSection,
    assertPageWritable,
    async createPage({ title, bodyXhtml, sectionId, parts = [] }) {
      const section = await resolveSection({ fresh: true });
      if (sectionId && sectionId !== section.id) {
        throw new WriteDeniedError(`Write denied: pages can only be created in the ${describe(section)} section.`);
      }
      const page = await writer.createPage(section.id, buildPageXhtml(title, bodyXhtml), parts);
      return { page, section };
    },
    // Writes target the page ID Graph returned during the check, not the caller's input.
    async patchPage(pageId, commands) {
      const { page, section } = await assertPageWritable(pageId);
      await writer.patchPageContent(page.id, commands);
      return { page, section };
    },
    async deletePage(pageId) {
      const { page, section } = await assertPageWritable(pageId);
      await writer.deletePage(page.id);
      return { page, section };
    },
  });
}
