# OneNote MCP Server (personal account, guarded fork)

A Model Context Protocol (MCP) server that lets AI assistants read Microsoft OneNote, and optionally write to **one** section, for a **personal Microsoft account**.

> Fork of [danosb/onenote-mcp](https://github.com/danosb/onenote-mcp), which builds on [azure-onenote-mcp-server](https://github.com/ZubeidHendricks/azure-onenote-mcp-server) by Zubeid Hendricks. The server was rewritten; see [Changes from upstream](#changes-from-upstream).

## Access model

The server is **read-only by default**. Write tools exist only when it is started with `ONENOTE_ACCESS=write-section`, and then every write is limited to one section:

| Mode | Started with | Tools |
| --- | --- | --- |
| `read-only` (default) | no `ONENOTE_*` settings, or `ONENOTE_ACCESS=read-only` | read tools only; the Graph client has no write capability |
| `write-section` | `ONENOTE_ACCESS=write-section` and `ONENOTE_WRITE_SECTION=<section name>` | read tools plus `createPage`, `composePage`, `appendToPage`, `updatePage`, `deletePage` |

In `write-section` mode, the guard (`lib/guard.mjs`) resolves the section by name, case-insensitive. Set `ONENOTE_WRITE_NOTEBOOK` to limit the match to one notebook. Writes are refused when no section matches or more than one does. The guard resolves the section again before every write, so renaming the section or adding a second section with the same name revokes write access immediately. Before each page write, the guard also reads the target page from Graph and checks its parent section, then writes to the page ID Graph returned. The Graph layer (`lib/graph.mjs`) sends only three write shapes: create a page in a section, PATCH a page's content, and delete a page. It refuses every other method, path and host.

Sign-in uses the `consumers` authority, so only personal Microsoft accounts (Outlook.com, Live, Hotmail) can sign in. Work or school accounts are rejected, and so is a cached account that isn't personal.

Password-protected sections can't be read through Microsoft Graph. `searchPages` skips them and lists them in `skippedSections`.

## Install

Requires Node.js 20 or later.

```
cd onenote-mcp
npm install
npm test
npm run auth
```

`npm run auth` prints a URL (`https://www.microsoft.com/link`) and a code. Sign in with your **personal** Microsoft account and accept the consent prompt for OneNote and basic profile. MSAL keeps the token cache, including the refresh token, in `%LOCALAPPDATA%\onenote-mcp\msal-cache.json` (`~/.local/state/onenote-mcp/` on Linux and macOS). Sessions renew silently from that cache, so you don't sign in again every hour. Other commands:

```
npm run auth -- --status     # show the signed-in account
npm run auth -- --sign-out   # remove the cached account
```

Within one process, token-cache saves run one at a time and atomically replace the file. Processes don't lock the cache against each other, so close Copilot sessions that use the server before you sign out. A session that refreshes its token at that moment could save the account back.

An assistant can also sign in through the `authenticate` tool. It returns the URL and code, and the user finishes in a browser.

## Configuration

| Variable | Default | Meaning |
| --- | --- | --- |
| `ONENOTE_ACCESS` | `read-only` | `read-only` or `write-section`. Any other value means read-only. |
| `ONENOTE_WRITE_SECTION` | (none) | Name of the only writable section (`write-section` mode). |
| `ONENOTE_WRITE_NOTEBOOK` | (none) | Notebook that must contain the writable section. |
| `ONENOTE_CACHE_DIR` | `%LOCALAPPDATA%\onenote-mcp` | Folder for the MSAL token cache. |
| `ONENOTE_CLIENT_ID` | `14d82eec-204b-4c2f-b7e8-296a70dab67e` | Public client ID ("Microsoft Graph Command Line Tools"). It must allow personal accounts and the device code flow. |

### GitHub Copilot CLI

Read-only for every session, in `~/.copilot/mcp-config.json`:

```json
{
  "mcpServers": {
    "onenote": {
      "type": "local",
      "command": "node",
      "args": ["C:\\path\\to\\onenote-mcp\\onenote-mcp.mjs"],
      "env": { "ONENOTE_ACCESS": "read-only" },
      "tools": ["authStatus", "authenticate", "listNotebooks", "listSections", "listPages", "getPage", "searchPages"]
    }
  }
}
```

Writable in one folder: put a `.mcp.json` with a server of the **same name** in that folder. Copilot CLI loads workspace servers from trusted folders, and they replace a user-level server of the same name:

```json
{
  "mcpServers": {
    "onenote": {
      "type": "local",
      "command": "node",
      "args": ["C:\\path\\to\\onenote-mcp\\onenote-mcp.mjs"],
      "env": {
        "ONENOTE_ACCESS": "write-section",
        "ONENOTE_WRITE_SECTION": "AI Manage",
        "ONENOTE_WRITE_NOTEBOOK": "My Notebook"
      }
    }
  }
}
```

Other MCP clients (Claude Desktop, Cursor, VS Code) use the same `command`, `args` and `env`.

## Tools

| Tool | Mode | Description |
| --- | --- | --- |
| `authStatus` | all | Signed-in account, access mode and the resolved writable section |
| `authenticate` | all | Start device code sign-in (personal accounts only) |
| `listNotebooks` | all | All notebooks |
| `listSections` | all | Sections, including ones in section groups; filter by `notebookId` or `notebookName` |
| `listPages` | all | Pages of a section (`sectionId`, or `sectionName` plus optional `notebookName`) |
| `getPage` | all | Page content as Markdown (default), raw HTML with element IDs (`format: "html"`), or numbered blocks plus a `blocksVersion` for `composePage` (`format: "blocks"`) |
| `searchPages` | all | Pages whose title contains `query`, optionally within a notebook or section |
| `createPage` | write-section | New page in the writable section; `content` is Markdown, HTML or text |
| `composePage` | write-section | New page in the writable section built from a source page's blocks, copied verbatim under added headings (see below) |
| `appendToPage` | write-section | Append content to a page in the writable section |
| `updatePage` | write-section | OneNote PATCH commands (`target`: `body`, `title`, `#data-id` or a generated element ID) |
| `deletePage` | write-section | Delete a page in the writable section; `expectedTitle` must match the page title |

Content conversion (`lib/convert.mjs`):

- Writes: Markdown becomes XHTML. `- [ ]` and `- [x]` items become OneNote to-do tags, and code blocks become one Consolas paragraph per line. Strikethrough becomes a styled span. Scripts, forms and event handlers are removed.
- Reads: OneNote HTML becomes Markdown. Styled spans become bold, italic, strikethrough or code. Runs of monospace paragraphs become fenced code blocks. Note tags show as `[ ]`, `[x]` or `[tag]`, and tables, images and attachments become Markdown.

### Reformatting without retyping (`composePage`)

`getPage` with `format: "blocks"` splits a page into numbered top-level blocks (paragraphs, tables, lists, images) and returns a `blocksVersion`. `composePage` takes a layout of sections, each with an optional heading and a list of block numbers or ranges such as `"3-7"`. It copies the blocks' original HTML into a new page in the writable section, so text is never retyped (`lib/compose.mjs`). Specifically, it:

- Refuses if the source page changed since it was read, if a block is used twice, or if a block is neither used nor listed in `omit`.
- Downloads images and attachments from the source page and uploads them with the new page (multipart request). Each part gets the content type the source page declares, or else the one its magic bytes indicate. Graph serves every resource as `application/octet-stream`, and OneNote drops image parts sent with that type.
- Starts the page with a line naming the source page.
- Re-reads the new page and reports in `verification` whether every copied block's characters appear in order and every image and attachment is byte-identical to the source (SHA-256).

OneNote returns attachments as XHTML self-closing tags (`<object ... />`). HTML parsers ignore that slash, so the server expands such tags before parsing; otherwise an attachment would swallow the content after it.

## Development

`npm test` runs the unit tests against a fake Microsoft Graph, so they never contact OneNote. Put throwaway scripts in `.scratch/`, which git ignores. If you install packages through an npm registry mirror, run `npm run lockfile:public` before committing, so `package-lock.json` only references registry.npmjs.org. A test enforces this.

## Changes from upstream

- The tools take real arguments. Upstream tools ignored their input, apart from a Cursor-only `random_string`. `createPage` always wrote a fixed page into the first section, and `listPages` only listed the first section.
- MSAL with a persistent token cache replaces a raw one-hour access token in `.access-token.txt`. Sign-in is limited to personal accounts.
- The write guard is new: read-only by default, and writes are limited to one section.
- Content conversion between Markdown and OneNote HTML is new.
- Search handles password-protected sections, and Graph throttling (HTTP 429) is retried.
- The server uses the published `@modelcontextprotocol/sdk` package instead of a locally built SDK clone, and the standalone helper scripts were removed.
- Tests: `npm test` uses `node:test` with a fake Graph.

## License

MIT
