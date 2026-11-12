#!/usr/bin/env node
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { Auth } from './lib/auth.mjs';
import { loadConfig } from './lib/config.mjs';
import { createOneNoteServer } from './lib/server.mjs';

const config = loadConfig();
for (const warning of config.warnings) console.error(`[onenote-mcp] WARNING: ${warning}`);

const server = createOneNoteServer({ config, auth: new Auth(config) });
await server.connect(new StdioServerTransport());

const scope = config.writeSection
  ? `writes limited to section "${config.writeSection}"${config.writeNotebook ? ` in notebook "${config.writeNotebook}"` : ''}`
  : 'read-only';
console.error(`[onenote-mcp] started: ${scope}; token cache ${config.cacheFile}`);
