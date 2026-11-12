import { randomUUID } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { InteractionRequiredAuthError, LogLevel, PublicClientApplication } from '@azure/msal-node';
import { CONSUMERS_TENANT_ID } from './config.mjs';

// Personal Microsoft accounts only: the consumers authority rejects work or school accounts.
export const AUTHORITY = 'https://login.microsoftonline.com/consumers';
export const SCOPES = Object.freeze(['Notes.ReadWrite', 'User.Read']);

export class AuthRequiredError extends Error {
  constructor(message) {
    super(message);
    this.name = 'AuthRequiredError';
  }
}

const SIGN_IN_HINT =
  'Call the "authenticate" tool (or run `npm run auth` in the onenote-mcp folder) and sign in with your PERSONAL Microsoft account.';

export function isPersonalAccount(account) {
  return Boolean(account) && account.tenantId === CONSUMERS_TENANT_ID;
}

const RENAME_RETRY_CODES = new Set(['EPERM', 'EACCES', 'EBUSY']);

// Windows refuses to replace a file that another process has open, so retry the atomic rename.
async function replaceFile(tmp, file) {
  for (let attempt = 1; ; attempt++) {
    try {
      await fs.promises.rename(tmp, file);
      return;
    } catch (error) {
      if (!RENAME_RETRY_CODES.has(error.code) || attempt >= 10) throw error;
      await new Promise((resolve) => setTimeout(resolve, 25 * attempt));
    }
  }
}

/**
 * MSAL cache plugin that persists the token cache (including the refresh token) to a local file.
 * Within one process, reads and saves run one at a time in call order, and each save
 * atomically replaces the file.
 */
export function fileCachePlugin(file) {
  let queue = Promise.resolve();
  const enqueue = (task) => {
    const run = queue.then(task);
    queue = run.catch(() => {});
    return run;
  };

  return {
    beforeCacheAccess(context) {
      return enqueue(async () => {
        let data;
        try {
          data = await fs.promises.readFile(file, 'utf8');
        } catch (error) {
          if (error.code === 'ENOENT') return;
          throw error;
        }
        try {
          context.tokenCache.deserialize(data);
        } catch (error) {
          console.error(`[onenote-mcp] Ignoring unreadable token cache ${file}: ${error.message}`);
        }
      });
    },
    afterCacheAccess(context) {
      if (!context.cacheHasChanged) return Promise.resolve();
      const data = context.tokenCache.serialize();
      return enqueue(async () => {
        await fs.promises.mkdir(path.dirname(file), { recursive: true });
        const tmp = `${file}.${process.pid}.${randomUUID()}.tmp`;
        await fs.promises.writeFile(tmp, data, { mode: 0o600 });
        try {
          await replaceFile(tmp, file);
        } catch (error) {
          await fs.promises.rm(tmp, { force: true });
          throw new Error(`Could not save the OneNote token cache ${file}: ${error.message}`);
        }
      });
    },
  };
}

export class Auth {
  constructor(config, { pca } = {}) {
    this.config = config;
    this.pca =
      pca ??
      new PublicClientApplication({
        auth: { clientId: config.clientId, authority: AUTHORITY },
        cache: { cachePlugin: fileCachePlugin(config.cacheFile) },
        system: {
          loggerOptions: {
            logLevel: LogLevel.Warning,
            piiLoggingEnabled: false,
            // stdout carries the MCP protocol, so MSAL logs must go to stderr.
            loggerCallback: (_level, message) => console.error(`[msal] ${message}`),
          },
        },
      });
    this.pending = null;
    this.tokenRequest = null;
  }

  async getAccount() {
    const accounts = await this.pca.getAllAccounts();
    return accounts.find(isPersonalAccount) ?? null;
  }

  /** Concurrent callers share one in-flight request, so a token is refreshed at most once at a time. */
  getAccessToken() {
    this.tokenRequest ??= this.#acquireAccessToken().finally(() => {
      this.tokenRequest = null;
    });
    return this.tokenRequest;
  }

  async #acquireAccessToken() {
    const account = await this.getAccount();
    if (!account) throw new AuthRequiredError(`Not signed in to OneNote. ${SIGN_IN_HINT}`);
    let result;
    try {
      result = await this.pca.acquireTokenSilent({ account, scopes: [...SCOPES] });
    } catch (error) {
      if (error instanceof InteractionRequiredAuthError) {
        throw new AuthRequiredError(`The OneNote sign-in has expired. ${SIGN_IN_HINT}`);
      }
      throw error;
    }
    if (!isPersonalAccount(result.account)) {
      throw new AuthRequiredError(`The cached account is not a personal Microsoft account. ${SIGN_IN_HINT}`);
    }
    return result.accessToken;
  }

  /**
   * Starts (or reuses) a device code sign-in. Resolves as soon as the code is
   * available; token polling continues in the background.
   */
  async startSignIn() {
    if (this.pending && !this.pending.done) return this.pending.prompt;

    const pending = { done: false, error: null, username: null, prompt: null };
    pending.prompt = new Promise((resolvePrompt, rejectPrompt) => {
      this.pca
        .acquireTokenByDeviceCode({
          scopes: [...SCOPES],
          deviceCodeCallback: (response) =>
            resolvePrompt({
              message: response.message,
              userCode: response.userCode,
              verificationUri: response.verificationUri,
              expiresAt: new Date(Date.now() + response.expiresIn * 1000).toISOString(),
            }),
        })
        .then(async (result) => {
          if (!result) throw new Error('Sign-in was cancelled.');
          const cache = this.pca.getTokenCache();
          if (!isPersonalAccount(result.account)) {
            if (result.account) await cache.removeAccount(result.account);
            throw new Error('The signed-in account is not a personal Microsoft account; it was removed.');
          }
          // Keep exactly one account in the cache.
          for (const other of await cache.getAllAccounts()) {
            if (other.homeAccountId !== result.account.homeAccountId) await cache.removeAccount(other);
          }
          pending.username = result.account.username;
        })
        .catch((error) => {
          pending.error = error;
          rejectPrompt(error);
        })
        .finally(() => {
          pending.done = true;
        });
    });
    this.pending = pending;
    return pending.prompt;
  }

  async status() {
    const account = await this.getAccount();
    const status = { signedIn: Boolean(account), account: account?.username ?? null };
    const pending = this.pending;
    if (pending) {
      if (!pending.done) {
        status.signIn = 'in-progress';
        status.prompt = await pending.prompt.catch(() => null);
      } else if (pending.error) {
        status.signIn = 'failed';
        status.signInError = pending.error.message;
      } else {
        status.signIn = 'completed';
      }
    }
    return status;
  }

  async signOut() {
    const cache = this.pca.getTokenCache();
    const accounts = await cache.getAllAccounts();
    for (const account of accounts) await cache.removeAccount(account);
    return accounts.map((a) => a.username);
  }
}
