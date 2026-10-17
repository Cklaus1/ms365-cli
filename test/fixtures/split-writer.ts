// Child process for the IDT-2 concurrency gate: one process per account, each round a fresh plugin (as a fresh
// CLI process would be) that loads the shared store, adds/refreshes its own account and saves.
// argv: <account n> <rounds> <keychain|file-items>; env: MS365_CLI_CONFIG_DIR (+ the stub keychain env).
import fs from 'fs';
import path from 'path';
import { configDir, FileTokenCachePlugin, fileCredentialStore, storageResolver, withFileLock, writeFileAtomic } from '../../src/auth.js';
import { createFileStore } from '../../src/credential-store.js';

const [n, rounds, mode, barrier] = process.argv.slice(2);

// Start barrier: report ready, then wait for the parent's "go" so all writers contend at once.
if (barrier) {
  fs.writeFileSync(path.join(barrier, `ready-${n}`), '');
  const deadline = Date.now() + 20_000;
  while (!fs.existsSync(path.join(barrier, 'go')) && Date.now() < deadline) Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 2);
}
const CID = '11111111-1111-1111-1111-111111111111';
const env = 'login.microsoftonline.com';

for (let r = 0; r < Number(rounds); r++) {
  let items;
  if (mode === 'file-items') {
    fs.mkdirSync(path.join(configDir(), 'items'), { recursive: true });
    items = createFileStore({ dir: () => path.join(configDir(), 'items'), writeFileAtomic, withFileLock });
  } else {
    items = storageResolver();
  }
  const plugin = new FileTokenCachePlugin(fileCredentialStore(), items);
  let loaded = '';
  await plugin.beforeCacheAccess({ tokenCache: { deserialize: (d: string) => { loaded = d; } } } as never);
  const cache = loaded ? JSON.parse(loaded) : { Account: {}, IdToken: {}, AccessToken: {}, RefreshToken: {}, AppMetadata: {} };
  // Round r signs in a new account uid<n>x<r> (an account lost by an unlocked write stays lost) and refreshes the
  // tokens of this writer's first account uid<n>x0 (a changed-only write).
  for (const k of new Set([0, r])) addAccount(cache, `uid${n}x${k}.utid`, k === 0 ? r : k);
  await plugin.afterCacheAccess({ cacheHasChanged: true, tokenCache: { serialize: () => JSON.stringify(cache) } } as never);
}

function addAccount(cache: Record<string, Record<string, unknown>>, hid: string, version: number) {
  cache.Account[`${hid}-${env}-utid`] = { home_account_id: hid, environment: env, realm: 'utid', local_account_id: hid.split('.')[0], username: `${hid.split('.')[0]}@contoso.com`, authority_type: 'MSSTS' };
  cache.RefreshToken[`${hid}-${env}-refreshtoken-${CID}----`] = { home_account_id: hid, environment: env, credential_type: 'RefreshToken', client_id: CID, secret: `rt-${hid}-${version}` };
  cache.AccessToken[`${hid}-${env}-accesstoken-${CID}-utid-.default--`] = { home_account_id: hid, environment: env, credential_type: 'AccessToken', client_id: CID, secret: `at-${hid}-${version}`, realm: 'utid', target: '.default', cached_at: '1', expires_on: '2', extended_expires_on: '3', token_type: 'Bearer' };
  cache.AppMetadata[`appmetadata-${env}-${CID}`] = { client_id: CID, environment: env, family_id: '1' };
}
