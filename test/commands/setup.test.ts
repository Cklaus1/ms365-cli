import { describe, it, expect, afterEach, beforeEach, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import readline from 'readline';
import type { Command } from 'commander';
import { runCommand } from '../helpers/cli.js';
import { registerAuthCommands } from '../../src/commands/auth.js';
import { registerSetupCommand, openBrowserCommand } from '../../src/commands/setup.js';
import { registerGenerateSkillsCommand } from '../../src/commands/generate-skills.js';
import { AuthManager, buildScopes } from '../../src/auth.js';

let tmp: string;
const savedEnv: Record<string, string | undefined> = {};
const ENV_KEYS = ['MS365_CLI_CONFIG_DIR', 'MS365_CLI_CLIENT_ID'];

beforeEach(() => {
  tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-setup-'));
  for (const k of ENV_KEYS) { savedEnv[k] = process.env[k]; delete process.env[k]; }
  process.env.MS365_CLI_CONFIG_DIR = path.join(tmp, 'config');
});

afterEach(() => {
  vi.restoreAllMocks();
  for (const k of ENV_KEYS) { if (savedEnv[k] === undefined) delete process.env[k]; else process.env[k] = savedEnv[k]; }
  fs.rmSync(tmp, { recursive: true, force: true });
  process.exitCode = undefined;
});

/** Stub readline so each wizard prompt takes the next answer (or `fallback` once the list runs out). */
function answerPrompts(answers: string[], fallback = '') {
  const asked: string[] = [];
  vi.spyOn(readline, 'createInterface').mockImplementation((() => ({
    question: (q: string, cb: (a: string) => void) => { asked.push(q); cb(answers.length ? answers.shift()! : fallback); },
    close: () => {},
  })) as never);
  return asked;
}

/** The production getAuth: constructing AuthManager throws when no client ID is configured. */
const realGetAuth = vi.fn(() => new AuthManager(buildScopes(false)));

function registerAuth(getAuth: () => AuthManager) {
  return (program: Command) => {
    registerAuthCommands(program, getAuth);
    registerSetupCommand(program, getAuth);
  };
}

function stubLogin() {
  vi.spyOn(AuthManager.prototype, 'initialize').mockResolvedValue();
  const login = vi.spyOn(AuthManager.prototype, 'login').mockResolvedValue(undefined as never);
  vi.spyOn(AuthManager.prototype, 'verifyLogin').mockResolvedValue({ success: true, user: { displayName: 'Ada', email: 'ada@x.com' } });
  return login;
}

const configJson = () => JSON.parse(fs.readFileSync(path.join(tmp, 'config', 'config.json'), 'utf8'));

describe('auth setup on a fresh install (M-2, C-16)', () => {
  it('`auth setup --help` works with no client ID', async () => {
    realGetAuth.mockClear();
    const r = await runCommand(registerAuth(realGetAuth), ['auth', 'setup', '--help']);
    expect(r.exitCode).toBe(0);
    expect(r.stdout).toContain('setup');
    expect(realGetAuth).not.toHaveBeenCalled();
  });

  it('reaches the first prompt without a client ID; option [1] says it is not available', async () => {
    realGetAuth.mockClear();
    const asked = answerPrompts(['1', 'n', '']);
    const r = await runCommand(registerAuth(realGetAuth), ['auth', 'setup']);
    expect(realGetAuth).not.toHaveBeenCalled();
    expect(asked[0]).toMatch(/Choose \[1\/2\]/);
    expect(r.stderr).toMatch(/not available yet/i);
    expect(r.stderr).toMatch(/option \[2\]/);
    expect(asked.some(q => /client\) ID/.test(q))).toBe(true);
    expect(r.exitCode).toBe(2); // empty client ID is a usage error, not process.exit(1)
    expect(r.stdout).toBe('');
  });
});

describe('auth setup saves the chosen scopes (M-12)', () => {
  it('writes config.json and logs in with the entered client ID and chosen scopes', async () => {
    const login = stubLogin();
    // mode 2, no browser, client ID, then defaults (personal yes, org no), Enter.
    answerPrompts(['2', 'n', 'client-123']);
    const r = await runCommand(registerAuth(realGetAuth), ['auth', 'setup']);
    expect(r.exitCode).toBe(0);
    expect(login).toHaveBeenCalledOnce();
    const saved = configJson().scopes as string[];
    expect(saved).toContain('Mail.ReadWrite');
    expect(saved).toContain('Presence.ReadWrite');
    expect(saved).not.toContain('Sites.Read.All');
    // Login (and later runs, even in org mode) request exactly the saved choice.
    expect(new Set(buildScopes(true))).toEqual(new Set(saved.filter(s => s !== 'Presence.Read')));
  });

  it('offers every scope the CLI requests in org mode', async () => {
    stubLogin();
    const all = buildScopes(true);
    answerPrompts(['2', 'n', 'client-123'], 'y');
    await runCommand(registerAuth(realGetAuth), ['auth', 'setup']);
    const saved = new Set(configJson().scopes as string[]);
    for (const s of all) expect(saved.has(s), s).toBe(true);
  });

  it('keeps other keys already in config.json', async () => {
    stubLogin();
    fs.mkdirSync(path.join(tmp, 'config'), { recursive: true });
    fs.writeFileSync(path.join(tmp, 'config', 'config.json'), JSON.stringify({ other: 1 }));
    answerPrompts(['2', 'n', 'client-123']);
    await runCommand(registerAuth(realGetAuth), ['auth', 'setup']);
    expect(configJson().other).toBe(1);
  });
});

describe('auth setup existing login (M-28)', () => {
  it('stops the wizard when the user keeps the existing login', async () => {
    process.env.MS365_CLI_CLIENT_ID = 'env-client';
    const fake = { initialize: async () => {}, verifyLogin: async () => ({ success: true, user: { displayName: 'Ada', email: 'a@x.com' } }) };
    const asked = answerPrompts(['y']);
    const r = await runCommand(registerAuth(() => fake as never), ['auth', 'setup']);
    expect(asked).toHaveLength(1);
    expect(r.stderr).toContain('Using existing credentials');
  });

  it('opens the browser on Windows with start "" "URL"', () => {
    expect(openBrowserCommand('https://x.test/a', 'win32')).toBe('start "" "https://x.test/a"');
    expect(openBrowserCommand('https://x.test/a', 'darwin')).toBe('open "https://x.test/a"');
    expect(openBrowserCommand('https://x.test/a', 'linux')).toBe('xdg-open "https://x.test/a"');
  });
});

describe('auth status / logout (M-28)', () => {
  it('auth status exits 3 when not logged in', async () => {
    const fake = { initialize: async () => {}, verifyLogin: async () => ({ success: false }) };
    const r = await runCommand(registerAuth(() => fake as never), ['auth', 'status']);
    expect(r.exitCode).toBe(3);
  });

  it('auth logout reports success on stderr only', async () => {
    let called = false;
    const fake = { initialize: async () => {}, logout: async () => { called = true; } };
    const r = await runCommand(registerAuth(() => fake as never), ['auth', 'logout']);
    expect(called).toBe(true);
    expect(r.exitCode).toBe(0);
    expect(r.stderr).toContain('Logged out');
    expect(r.stdout).toBe('');
  });
});

describe('buildScopes saved scopes (M-12)', () => {
  it('returns the scopes from config.json when present', () => {
    fs.mkdirSync(path.join(tmp, 'config'), { recursive: true });
    fs.writeFileSync(path.join(tmp, 'config', 'config.json'), JSON.stringify({ scopes: ['User.Read', 'Mail.Read'] }));
    expect(buildScopes(true)).toEqual(['User.Read', 'Mail.Read']);
    expect(buildScopes(false)).toEqual(['User.Read', 'Mail.Read']);
  });

  it('falls back to the built-in list without config.json', () => {
    expect(buildScopes(false)).toContain('Mail.ReadWrite');
    expect(buildScopes(true)).toContain('Place.Read.All');
  });
});

describe('generate-skills (M-28)', () => {
  const gen = (argv: string[]) => runCommand(registerGenerateSkillsCommand as never, ['generate-skills', ...argv]);

  it('does not claim json is the default or that every list command has --filter/--select/--all', async () => {
    const out = path.join(tmp, 'skills');
    await gen(['--output-dir', out]);
    const shared = fs.readFileSync(path.join(out, 'ms365-shared', 'SKILL.md'), 'utf8');
    expect(shared).not.toMatch(/\*\*json\*\*[^\n]*\(default\)/);
    expect(shared).toMatch(/table[^\n]*terminal/i);
    const mail = fs.readFileSync(path.join(out, 'ms365-mail', 'SKILL.md'), 'utf8');
    expect(mail).not.toMatch(/All list commands support/);
  });

  it('refuses to overwrite an existing SKILL.md unless --force', async () => {
    const out = path.join(tmp, 'skills');
    await gen(['--output-dir', out]);
    const file = path.join(out, 'ms365-mail', 'SKILL.md');
    fs.writeFileSync(file, 'hand edited');
    const r = await gen(['--output-dir', out]);
    expect(r.exitCode).not.toBe(0);
    expect(fs.readFileSync(file, 'utf8')).toBe('hand edited');
    const forced = await gen(['--output-dir', out, '--force']);
    expect(forced.exitCode).toBe(0);
    expect(fs.readFileSync(file, 'utf8')).not.toBe('hand edited');
  });
});
