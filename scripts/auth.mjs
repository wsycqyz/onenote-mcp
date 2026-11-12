#!/usr/bin/env node
// Terminal sign-in helper: node scripts/auth.mjs [--status | --sign-out]
import { Auth } from '../lib/auth.mjs';
import { loadConfig } from '../lib/config.mjs';

const config = loadConfig();
const auth = new Auth(config);
const arg = process.argv[2] ?? '';

async function verify() {
  await auth.getAccessToken();
  const account = await auth.getAccount();
  console.log(`Signed in as ${account.username} (personal Microsoft account).`);
  console.log(`Token cache: ${config.cacheFile}`);
}

try {
  if (arg === '--sign-out') {
    const removed = await auth.signOut();
    console.log(removed.length ? `Signed out: ${removed.join(', ')}` : 'No cached account.');
  } else if (arg === '--status') {
    if (await auth.getAccount()) await verify();
    else console.log('Not signed in.');
  } else if (arg) {
    console.error('Usage: node scripts/auth.mjs [--status | --sign-out]');
    process.exitCode = 2;
  } else {
    let signedIn = false;
    if (await auth.getAccount()) {
      try {
        await verify();
        signedIn = true;
      } catch {
        // Expired; sign in again.
      }
    }
    if (!signedIn) {
      const prompt = await auth.startSignIn();
      console.log(prompt.message);
      console.log('Sign in with your PERSONAL Microsoft account (work or school accounts are rejected).');
      while (!auth.pending.done) await new Promise((resolve) => setTimeout(resolve, 1000));
      if (auth.pending.error) throw auth.pending.error;
      await verify();
    }
  }
} catch (error) {
  console.error(`Sign-in failed: ${error.message}`);
  process.exitCode = 1;
}
