import { describe, it, expect } from 'vitest';
import crypto from 'crypto';
import fs from 'fs';
import os from 'os';
import path from 'path';
// @ts-expect-error -- plain ESM script without type declarations
import { CASES, FIXTURE_DIR, check } from '../scripts/contract.mjs';

// AHT-15: the shared contract fixture (Chief MSC-020 vendors test/fixtures/contract/). The check regenerates every
// case from the current source (src/ through tsx: npm test runs before any build) and fails on any byte difference,
// so a contract change must regenerate the fixture (`npm run contract`) in the same commit. `npm run contract:check`
// runs the same comparison against the built dist/.

interface Manifest {
  v: number;
  ms365: { commit: string | null; dirty: boolean | null; distSha256: string | null };
  clock: string;
  files: Record<string, { argv: string[]; exitCode: number; capture: string; sha256: string }>;
}
interface ErrorLine { error: { code: string; exitCode: number } }
interface ListPage { incompleteReason?: string }
interface WatchLine { type: string; reason?: string; removedReason?: string }

const read = (f: string) => fs.readFileSync(path.join(FIXTURE_DIR, f), 'utf8');

describe('contract fixture (AHT-15)', () => {
  it('matches the current source byte for byte', async () => {
    expect(await check({ entry: 'src' })).toEqual([]);
  }, 60_000);

  it('fails on a changed byte, a missing file and a stray file', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'contract-tamper-'));
    try {
      fs.cpSync(FIXTURE_DIR, dir, { recursive: true });
      fs.appendFileSync(path.join(dir, 'auth-status.json'), ' ');
      fs.rmSync(path.join(dir, 'health-stale.json'));
      fs.writeFileSync(path.join(dir, 'extra.json'), '{}');
      expect(await check({ entry: 'src', fixtureDir: dir })).toEqual([
        'extra.json: not generated any more',
        'auth-status.json: bytes differ',
        'health-stale.json: missing from the fixture',
      ]);
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  }, 60_000);

  it('the manifest lists every file with its sha256, argv and exit code, and the dist digest it came from', () => {
    const m: Manifest = JSON.parse(read('manifest.json'));
    const onDisk = fs.readdirSync(FIXTURE_DIR).filter(f => f !== 'manifest.json').sort();
    expect(Object.keys(m.files).sort()).toEqual(onDisk);
    for (const [f, e] of Object.entries(m.files)) {
      expect(e.sha256, f).toBe(crypto.createHash('sha256').update(fs.readFileSync(path.join(FIXTURE_DIR, f))).digest('hex'));
    }
    expect(m.ms365.commit).toMatch(/^[0-9a-f]{40}$/);
    expect(typeof m.ms365.dirty).toBe('boolean');
    expect(m.ms365.distSha256).toMatch(/^[0-9a-f]{64}$/);
    expect(m.files['error-4-scope_ceiling.json']).toMatchObject({ exitCode: 4, capture: 'error' });
  });

  it('covers every contract shape the hosts parse', () => {
    const errors: ErrorLine[] = (CASES as { file: string; capture: string }[]).filter(c => c.capture === 'error').map(c => JSON.parse(read(c.file)));
    expect(new Set(errors.map(e => e.error.exitCode))).toEqual(new Set([1, 2, 3, 4, 5, 6, 7]));
    expect(errors.map(e => e.error.code)).toEqual(expect.arrayContaining(['scope_ceiling', 'token_rejected', 'cursor_mismatch', 'checkpoint_mismatch']));
    const lines: WatchLine[] = fs.readdirSync(FIXTURE_DIR).filter(f => f.startsWith('watch-')).flatMap(f => read(f).trim().split('\n').map(l => JSON.parse(l)));
    expect(new Set(lines.map(l => l.type))).toEqual(new Set(['checkpoint', 'gap', 'backfill', 'created', 'updated', 'deleted']));
    expect(new Set(lines.filter(l => l.type === 'gap').map(l => l.reason))).toEqual(new Set(['sync_state_expired', 'window_rolled']));
    expect(new Set(lines.filter(l => l.type === 'deleted').map(l => l.removedReason))).toEqual(new Set(['deleted', 'changed']));
    const pages: ListPage[] = ['list-next-cursor.json', 'list-page-limit.json', 'list-search-cap.json'].map(f => JSON.parse(read(f)));
    expect(pages.map(p => p.incompleteReason)).toEqual(['next_cursor', 'page_limit', 'search_cap']);
    expect(read('list-complete.ndjson').trim().split('\n').pop()).toBe('{"_meta":{"complete":true}}');
    expect(read('list-next-cursor.ndjson')).toContain('"_ref":{');
    // Chief's Teams brief (TMS-3): no-profile samples, every item untrusted, bodies dropped by --body none.
    const unfiltered: { commands: { path: string; options: { long: string }[] }[] } = JSON.parse(read('describe-unfiltered.json'));
    expect(unfiltered.commands.find(c => c.path === 'chat list')?.options.map(o => o.long)).toEqual(expect.arrayContaining(['--recent', '--body', '--top']));
    const teams = ['chat-list-recent.json', 'chat-messages.json', 'channel-messages.json'].map(f => JSON.parse(read(f)));
    expect(teams.map(p => p.complete)).toEqual([false, false, true]);
    for (const p of teams) {
      expect(p.value.length).toBeGreaterThan(0);
      for (const item of p.value) {
        expect(item.untrusted).toBe(true);
        expect(item.body ?? item.lastMessagePreview?.body).toBeUndefined();
      }
    }
    expect(teams[1].value.map((m: { messageType: string }) => m.messageType)).toEqual(['message', 'unknownFutureValue']);
    // No host-specific bytes: no temp or home paths.
    for (const f of fs.readdirSync(FIXTURE_DIR)) {
      expect(read(f), f).not.toContain(os.tmpdir());
      expect(read(f), f).not.toMatch(/\/home\/|\/Users\/|\\Users\\/);
    }
  });
});
