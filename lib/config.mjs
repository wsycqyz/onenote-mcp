import os from 'node:os';
import path from 'node:path';

// "Microsoft Graph Command Line Tools": a Microsoft public client that allows
// device code sign-in with personal Microsoft accounts.
export const DEFAULT_CLIENT_ID = '14d82eec-204b-4c2f-b7e8-296a70dab67e';

// Tenant ID that Microsoft identity platform uses for every personal (MSA) account.
export const CONSUMERS_TENANT_ID = '9188040d-6c67-4c5b-b112-36a304b66dad';

export const ACCESS_READ_ONLY = 'read-only';
export const ACCESS_WRITE_SECTION = 'write-section';

function defaultCacheDir(env) {
  if (env.ONENOTE_CACHE_DIR) return env.ONENOTE_CACHE_DIR;
  if (process.platform === 'win32') {
    return path.join(env.LOCALAPPDATA || path.join(os.homedir(), 'AppData', 'Local'), 'onenote-mcp');
  }
  return path.join(env.XDG_STATE_HOME || path.join(os.homedir(), '.local', 'state'), 'onenote-mcp');
}

/**
 * Reads the server configuration from environment variables. Anything that is
 * missing or invalid falls back to read-only mode.
 */
export function loadConfig(env = process.env) {
  const warnings = [];
  const access = (env.ONENOTE_ACCESS || ACCESS_READ_ONLY).trim().toLowerCase();
  const writeSection = (env.ONENOTE_WRITE_SECTION || '').trim();
  const writeNotebook = (env.ONENOTE_WRITE_NOTEBOOK || '').trim();

  let mode = ACCESS_READ_ONLY;
  if (access === ACCESS_WRITE_SECTION) {
    if (writeSection) {
      mode = ACCESS_WRITE_SECTION;
    } else {
      warnings.push('ONENOTE_ACCESS=write-section but ONENOTE_WRITE_SECTION is empty; running read-only.');
    }
  } else if (access !== ACCESS_READ_ONLY) {
    warnings.push(`Unknown ONENOTE_ACCESS "${access}"; running read-only.`);
  }

  const writable = mode === ACCESS_WRITE_SECTION;
  return Object.freeze({
    mode,
    writeSection: writable ? writeSection : null,
    writeNotebook: writable && writeNotebook ? writeNotebook : null,
    clientId: (env.ONENOTE_CLIENT_ID || DEFAULT_CLIENT_ID).trim(),
    cacheFile: path.join(defaultCacheDir(env), 'msal-cache.json'),
    warnings: Object.freeze(warnings),
  });
}
