import { describe, it, expect, afterEach } from 'vitest';
import { Readable } from 'stream';
import fs from 'fs';
import path from 'path';
import { Command } from 'commander';
import { catalogPaths, commandMeta, explicitUses, requiredScopes, usedScopes } from '../src/catalog.js';
import { activeCeiling, assertWithinCeiling, ceilingSurface, profileCeiling, unmetClause } from '../src/ceiling.js';
import { CliError, ExitCode } from '../src/errors.js';
import { decodeCursor, encodeCursor, setCursor } from '../src/formatter.js';
import { readProfile, saveProfile, validateProfile } from '../src/profiles.js';
import { buildTools, serve } from '../src/mcp-server.js';
import { registerMailCommands } from '../src/commands/mail.js';
import { registerCalendarCommands } from '../src/commands/calendar.js';
import { registerDriveCommands } from '../src/commands/drive.js';
import { registerWatchCommands } from '../src/commands/watch.js';
import { registerWorkflowCommands } from '../src/commands/workflow.js';
import { registerDescribeCommand } from '../src/commands/describe.js';
import { runCommand } from './helpers/cli.js';
import { fakeJwt, runEntry, entryHome } from './helpers/entry.js';
import type { GraphClient } from '../src/graph-client.js';

const GATE2 = ['User.Read', 'Mail.Read', 'Calendars.Read', 'offline_access'];
/** A synthetic tenant GUID (never a real one). */
const TENANT = '11111111-2222-3333-4444-555555555555';
const ceiling = profileCeiling({ type: 'host', tenantId: TENANT, scopes: GATE2 })!;

function refusal(fn: () => void): CliError {
  try { fn(); } catch (e) { return e as CliError; }
  throw new Error('expected a refusal');
}

afterEach(() => { setCursor(null); delete process.env.MS365_CLI_PROFILE; });

describe('catalog uses (AH-1(2), AC 4)', () => {
  it('every entry whose single clause lists several scopes has an explicit uses (15 today)', () => {
    const multi = catalogPaths().filter(p => { const m = commandMeta(p)!; return !m.scopesBy && m.scopes.length > 1; });
    expect(multi).toHaveLength(15);
    expect(multi.filter(p => !explicitUses(p))).toEqual([]);
    // Every explicit uses names only scopes the catalog lists for the command, and covers all of them.
    for (const p of multi) expect(new Set(explicitUses(p)!.flat())).toEqual(new Set(commandMeta(p)!.scopes));
  });

  it("workflow digest uses Tasks.Read as its own clause; user list keeps its alternatives together", () => {
    expect(explicitUses('workflow digest')).toContainEqual(['Tasks.Read']);
    expect(explicitUses('user list')).toEqual([['User.ReadBasic.All', 'User.Read.All']]);
  });

  it('without an explicit uses the selected requiredScopes clauses are used', () => {
    expect(usedScopes('watch', {}, ['files'])).toEqual(requiredScopes('watch', {}, ['files']));
    expect(usedScopes('mail list')).toEqual([{ scopes: ['Mail.Read'] }]);
  });
});

describe('profile ceiling (AH-1, AC 1-2)', () => {
  it('ignores OIDC scopes and the Graph resource prefix; admin profiles and profiles without scopes have none', () => {
    expect([...profileCeiling({ scopes: ['https://graph.microsoft.com/Mail.Read', 'openid', 'profile', 'email', 'offline_access'] })!]).toEqual(['Mail.Read']);
    expect(profileCeiling({ type: 'admin', tenantId: 'contoso.com' })).toBeNull();
    expect(profileCeiling({})).toBeNull();
  });

  it.each([
    ['mail send', [], {}], ['mail mark-read', ['m1'], {}], ['calendar create', [], {}], ['workflow digest', [], {}],
  ] as const)('%s exits 4 scope_ceiling under the Gate-2 ceiling', (cmd, args, opts) => {
    const e = refusal(() => assertWithinCeiling(cmd, usedScopes(cmd, opts, [...args]), ceiling, 'chief'));
    expect(e.code).toBe('scope_ceiling');
    expect(e.exitCode).toBe(ExitCode.PERMISSION);
    expect(e.fix).toContain('chief');
  });

  it('mail list, calendar view and watch events run; watch files is refused', () => {
    for (const [cmd, args] of [['mail list', []], ['calendar view', []], ['watch', ['events']], ['watch', ['mail']]] as const) {
      expect(() => assertWithinCeiling(cmd, usedScopes(cmd, {}, [...args]), ceiling, 'chief')).not.toThrow();
    }
    expect(refusal(() => assertWithinCeiling('watch', usedScopes('watch', {}, ['files']), ceiling, 'chief')).message).toContain('Files.Read');
  });

  it('property: over every catalog command, a run fits iff each uses clause has a scope the ceiling satisfies; widening never refuses more', () => {
    const all = [...new Set(catalogPaths().flatMap(p => commandMeta(p)!.scopes))].sort();
    let seed = 7;
    const rand = () => { seed = (seed * 1103515245 + 12345) % 2 ** 31; return seed / 2 ** 31; };
    for (let i = 0; i < 60; i++) {
      const small = new Set(all.filter(() => rand() < 0.3));
      const wide = new Set([...small, ...all.filter(() => rand() < 0.3)]);
      for (const p of catalogPaths()) {
        const clauses = usedScopes(p);
        const fits = !unmetClause(clauses, small);
        if (fits) expect(unmetClause(clauses, wide)).toBeUndefined();
        if (!clauses.some(c => c.scopes.length)) expect(fits).toBe(true);
      }
    }
    // Fail closed: an empty ceiling refuses every command that calls Graph; the full set allows all.
    for (const p of catalogPaths()) {
      const clauses = usedScopes(p);
      expect(!unmetClause(clauses, new Set())).toBe(!clauses.some(c => c.scopes.length));
      expect(unmetClause(clauses, new Set(all))).toBeUndefined();
    }
  });

  it('activeCeiling reads --profile, else MS365_CLI_PROFILE', () => {
    saveProfile('cap', { type: 'host', tenantId: TENANT, scopes: GATE2 });
    expect(activeCeiling('cap')?.profile).toBe('cap');
    process.env.MS365_CLI_PROFILE = 'cap';
    expect([...activeCeiling(undefined)!.ceiling]).toEqual(['User.Read', 'Mail.Read', 'Calendars.Read']);
    expect(activeCeiling(undefined)).not.toBeNull();
  });
});

describe('host profile (AH-2 type, AHT-1)', () => {
  it('needs the tenant GUID and scopes; account is allowed', () => {
    expect(refusal(() => validateProfile('h', { type: 'host', scopes: ['Mail.Read'] })).code).toBe('invalid_profile');
    expect(refusal(() => validateProfile('h', { type: 'host', tenantId: 'common', scopes: ['Mail.Read'] })).code).toBe('invalid_profile');
    expect(refusal(() => validateProfile('h', { type: 'host', tenantId: 'contoso.com', scopes: ['Mail.Read'] })).message).toContain('GUID');
    expect(refusal(() => validateProfile('h', { type: 'host', tenantId: TENANT })).message).toContain('ceiling');
    expect(() => validateProfile('h', { type: 'host', tenantId: TENANT, scopes: ['Mail.Read'], account: 'ceo@contoso.com' })).not.toThrow();
  });

  it('round-trips through the profile file', () => {
    saveProfile('h2', { type: 'host', tenantId: TENANT, scopes: ['Mail.Read'] });
    expect(readProfile('h2')).toEqual({ type: 'host', tenantId: TENANT, scopes: ['Mail.Read'] });
  });
});

const noClient = (() => { throw new Error('no Graph'); }) as () => GraphClient;
function program(): Command {
  const p = new Command().name('ms365').option('--profile <name>');
  registerMailCommands(p, noClient);
  registerCalendarCommands(p, noClient);
  registerDriveCommands(p, noClient);
  registerWatchCommands(p, noClient);
  registerWorkflowCommands(p, noClient);
  registerDescribeCommand(p);
  return p;
}

describe('effective surface: MCP and describe (AH-1(4)-(5), AC 3)', () => {
  it('mcp lists no tool outside the ceiling; the watch enum and watch resources are mail and events', async () => {
    const tools = buildTools(program(), { ceiling });
    const names = tools.map(t => t.name);
    expect(names).toEqual(expect.arrayContaining(['mail_list', 'calendar_view', 'watch']));
    for (const t of tools) expect(ceilingSurface(t.path, ceiling).allowed).toBe(true);
    for (const n of ['mail_send', 'mail_mark-read', 'calendar_create', 'workflow_digest', 'drive_ls']) expect(names).not.toContain(n);
    expect(tools.find(t => t.name === 'watch')!.inputSchema.properties.resource.enum).toEqual(['mail', 'events']);

    const out: Array<{ result: { resources: Array<{ uri: string }> } }> = [];
    const input = Readable.from([JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'resources/list' }) + '\n']);
    await serve(input, (l) => out.push(JSON.parse(l)), { tools, version: '0', runner: async () => ({ exitCode: 0, stdout: '', stderr: '' }) });
    expect(out[0].result.resources.map((r: { uri: string }) => r.uri)).toEqual(['ms365://watch/mail', 'ms365://watch/events']);
  });

  it('describe --json lists only allowed commands; --all marks the others allowed: false', async () => {
    saveProfile('g2', { type: 'host', tenantId: TENANT, scopes: GATE2 });
    const register = (p: Command) => { p.option('--profile <name>'); registerMailCommands(p, noClient); registerWatchCommands(p, noClient); registerDescribeCommand(p); };
    const narrow = JSON.parse((await runCommand(register, ['--profile', 'g2', 'describe', '--json'])).stdout).commands;
    expect(narrow.map((c: { path: string }) => c.path)).toContain('mail list');
    expect(narrow.map((c: { path: string }) => c.path)).not.toContain('mail send');
    expect(narrow.every((c: { allowed: boolean }) => c.allowed)).toBe(true);
    const all = JSON.parse((await runCommand(register, ['--profile', 'g2', 'describe', '--all', '--json'])).stdout).commands;
    expect(all.find((c: { path: string }) => c.path === 'mail send').allowed).toBe(false);
    const plain = JSON.parse((await runCommand(register, ['describe', '--json'])).stdout).commands;
    expect(plain.some((c: object) => 'allowed' in c)).toBe(false);
  });
});

describe('cursor binding (AH-1(7), AC 28)', () => {
  const v1 = 'https://graph.microsoft.com/v1.0';
  it('a cursor for another path or API version is cursor_mismatch; the same path with another query is fine', () => {
    const e = refusal(() => decodeCursor(encodeCursor(`${v1}/me/drive/root/children`), `${v1}/me/messages`));
    expect(e.code).toBe('cursor_mismatch');
    expect(e.exitCode).toBe(ExitCode.USAGE);
    expect(refusal(() => decodeCursor(encodeCursor(`${v1}/me/messages?$skip=10`), 'https://graph.microsoft.com/beta/me/messages')).code).toBe('cursor_mismatch');
    expect(decodeCursor(encodeCursor(`${v1}/me/messages?$skip=10`), `${v1}/me/messages`)).toBe(`${v1}/me/messages?$skip=10`);
    expect(decodeCursor(encodeCursor(`${v1}/me/mailFolders/AAMk%3D/messages?$skip=10`), `${v1}/me/mailFolders/AAMk=/messages`)).toContain('$skip=10');
  });

  it('mail list --cursor <drive URL> exits 2 with zero fetch calls; its own cursor still works', async () => {
    setCursor(encodeCursor(`${v1}/me/drive/root/children`));
    const bad = await runCommand(registerMailCommands, ['mail', 'list', '-o', 'json']);
    expect(bad.exitCode).toBe(2);
    expect(bad.calls).toEqual([]);
    setCursor(encodeCursor(`${v1}/me/messages?$skip=10`));
    const ok = await runCommand(registerMailCommands, ['mail', 'list', '-o', 'json'], { routes: [['GET', /skip=10/, () => ({ value: [] })]] });
    expect(ok.exitCode, ok.stderr).not.toBe(2);
    expect(ok.calls).toHaveLength(1);
  });
});

describe('through the real entry point (AC 1, spawned)', () => {
  it('scope_ceiling under --dry-run, and with MS365_CLI_PREFLIGHT=0 and an opaque token: zero fetches, before any token use', () => {
    const { home, configDir } = entryHome();
    fs.mkdirSync(path.join(configDir, 'profiles'));
    fs.writeFileSync(path.join(configDir, 'profiles', 'chief.json'), JSON.stringify({ type: 'user', scopes: GATE2 }));
    const send = ['mail', 'send', '--to', 'a@contoso.com', '--subject', 's', '--body', 'b'];
    // A token that holds Mail.Send would pass pre-flight: the ceiling, not the token, refuses.
    const broad = fakeJwt({ scp: 'Mail.Send Mail.Read User.Read', aud: 'https://graph.microsoft.com', exp: Math.floor(Date.now() / 1000) + 3600 });
    const dry = runEntry(['--profile', 'chief', '--dry-run', ...send], { home, configDir, env: { MS365_CLI_TOKEN: broad } });
    expect(dry.status, dry.stderr).toBe(4);
    expect(dry.error?.code).toBe('scope_ceiling');
    expect(dry.stderr).not.toContain('[dry-run]');
    const opaque = runEntry(['--profile', 'chief', ...send], {
      home, configDir, env: { MS365_CLI_PREFLIGHT: '0', MS365_CLI_TOKEN: 'opaque-token', MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000' },
    });
    expect(opaque.status, opaque.stderr).toBe(4);
    expect(opaque.error?.code).toBe('scope_ceiling');
    expect([...dry.fetches, ...opaque.fetches]).toEqual([]);
  }, 30_000);
});
