import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { execFileSync, spawnSync } from 'child_process';
import { runCommand } from './helpers/cli.js';
import { registerMailCommands } from '../src/commands/mail.js';
import { registerDescribeCommand } from '../src/commands/describe.js';
import { registerGenerateSkillsCommand } from '../src/commands/generate-skills.js';
import {
  type Catalog, renderCommandReference, renderSkills, replaceGeneratedRegion, checkGeneratedRegion,
  BEGIN_MARKER, END_MARKER,
} from '../src/docs-gen.js';

const ROOT = path.resolve(__dirname, '..');

const opt = (flags: string, extra: Partial<Catalog['commands'][number]['options'][number]> = {}) => {
  const long = flags.match(/--[\w-]+/)?.[0] ?? null;
  return { flags, long, short: null, description: `${long} option`, required: false, takesValue: flags.includes('<'), choices: null, default: null, ...extra };
};

const CATALOG: Catalog = {
  version: '9.9.9',
  globalOptions: [opt('--dry-run', { description: 'Preview the API request without executing it' }), opt('--fields <list>')],
  commands: [
    { path: 'mail send', description: 'Send an email', arguments: [], options: [opt('--to <addresses>', { required: true }), opt('-o, --output <format>', { choices: ['json', 'ndjson', 'table'] })], destructive: false, idempotent: false, scopes: ['Mail.Send'], offline: false },
    { path: 'describe', description: 'Describe every command', arguments: [{ name: 'command', required: false, variadic: true }], options: [opt('--json')], destructive: false, idempotent: true, offline: true, scopes: [] },
    { path: 'mail delete', description: 'Delete a message | permanently', arguments: [{ name: 'messageId', required: true, variadic: false }], options: [], destructive: true, idempotent: true, scopes: ['Mail.ReadWrite'], offline: false },
    { path: 'calendar list', description: 'List calendars', arguments: [], options: [opt('--all'), opt('--page-limit <n>'), opt('--top <n>')], destructive: false, idempotent: true, scopes: ['Calendars.Read'], offline: false },
  ],
};

const shuffled = (c: Catalog): Catalog => ({ ...c, commands: [...c.commands].reverse() });

describe('renderCommandReference', () => {
  it('is deterministic: groups and commands are sorted regardless of input order', () => {
    const md = renderCommandReference(CATALOG);
    expect(renderCommandReference(shuffled(CATALOG))).toBe(md);
    const headings = [...md.matchAll(/^### (\S+)$/gm)].map(m => m[1]);
    expect(headings).toEqual(['calendar', 'describe', 'mail']);
    expect(md.indexOf('`mail delete')).toBeLessThan(md.indexOf('`mail send'));
  });

  it('includes every command with its arguments, key options and notes', () => {
    const md = renderCommandReference(CATALOG);
    for (const c of CATALOG.commands) expect(md).toContain(`\`${c.path}\``);
    const del = md.split('\n').find(l => l.startsWith('| `mail delete`'))!;
    expect(del).toContain('`<messageId>`');
    expect(del).toContain('**destructive**');
    expect(del).toContain('Mail.ReadWrite');
    expect(del).toContain('Delete a message \\| permanently');
    const send = md.split('\n').find(l => l.startsWith('| `mail send`'))!;
    expect(send).toContain('`--to <addresses>` (required)');
    expect(send).not.toContain('--output');
    expect(send).toContain('not idempotent');
    expect(md.split('\n').find(l => l.startsWith('| `describe`'))).toMatch(/`\[command\.\.\.\]`.*offline/);
    const list = md.split('\n').find(l => l.startsWith('| `calendar list`'))!;
    expect(list).toContain('`--top <n>`');
    expect(list).not.toContain('--page-limit');
  });

  it('lists the global options', () => {
    expect(renderCommandReference(CATALOG)).toMatch(/\| `--dry-run` \| Preview the API request without executing it \|/);
  });
});

describe('renderSkills', () => {
  const files = () => new Map(renderSkills(CATALOG));

  it('writes a shared skill, one skill per group and an index, in sorted order', () => {
    expect(renderSkills(CATALOG).map(([p]) => p)).toEqual([
      'ms365-shared/SKILL.md', 'ms365-calendar/SKILL.md', 'ms365-describe/SKILL.md', 'ms365-mail/SKILL.md', 'INDEX.md',
    ]);
    expect(renderSkills(shuffled(CATALOG))).toEqual(renderSkills(CATALOG));
  });

  it('documents the real output, paging, error and safety contract in the shared skill', () => {
    const shared = files().get('ms365-shared/SKILL.md')!;
    expect(shared).toMatch(/table[^\n]*terminal[^\n]*json[^\n]*piped/i);
    expect(shared).not.toMatch(/json[^\n]*\(default\)/);
    for (const s of ['-o ndjson', '--fields', '--cursor', 'nextCursor', '--error-format json', '--dry-run', '--wrap-untrusted', 'untrusted: true']) {
      expect(shared).toContain(s);
    }
    const codes = [[0, 'ok'], [1, 'generic'], [2, 'usage'], [3, 'auth'], [4, 'permission'], [5, 'not found'], [6, 'throttled'], [7, 'confirmation refused']];
    for (const [code, name] of codes) expect(shared).toMatch(new RegExp(`\\| ${code} \\| ${name}`, 'i'));
    expect(shared).toMatch(/destructive[^\n]*(TTY|terminal)/i);
    expect(shared).toContain('version: 9.9.9');
  });

  it('marks destructive commands and their scopes in group skills', () => {
    const mail = files().get('ms365-mail/SKILL.md')!;
    expect(mail).toMatch(/^name: ms365-mail$/m);
    expect(mail).toContain('ms365 mail delete <messageId>');
    expect(mail).toMatch(/\| `mail delete` .*\*\*destructive\*\*/);
    expect(mail).toMatch(/## Destructive commands[\s\S]*`mail delete`/);
    expect(mail).toContain('../ms365-shared/SKILL.md');
    expect(files().get('ms365-calendar/SKILL.md')).not.toContain('## Destructive commands');
    const index = files().get('INDEX.md')!;
    for (const g of ['shared', 'calendar', 'describe', 'mail']) expect(index).toContain(`(ms365-${g}/SKILL.md)`);
  });

  it('limits group skills to the filter', () => {
    expect(renderSkills(CATALOG, /^mail$/).map(([p]) => p)).toEqual(['ms365-shared/SKILL.md', 'ms365-mail/SKILL.md', 'INDEX.md']);
  });
});

describe('README generated region', () => {
  const readme = `# Title\n\nintro\n\n${BEGIN_MARKER}\nold\n${END_MARKER}\n\n## After\n`;

  it('round-trips: replace then check passes; a mutation makes check fail with a diff summary', () => {
    const fresh = replaceGeneratedRegion(readme, CATALOG);
    expect(fresh.startsWith('# Title\n\nintro\n\n')).toBe(true);
    expect(fresh.endsWith('\n\n## After\n')).toBe(true);
    expect(fresh).toContain('`mail delete`');
    expect(replaceGeneratedRegion(fresh, CATALOG)).toBe(fresh);
    expect(checkGeneratedRegion(fresh, CATALOG)).toEqual({ ok: true, summary: '' });

    const stale = fresh.replace('| `mail send`', '| `mail sendx`');
    const r = checkGeneratedRegion(stale, CATALOG);
    expect(r.ok).toBe(false);
    expect(r.summary).toContain('- | `mail sendx`');
    expect(r.summary).toContain('+ | `mail send`');
    expect(checkGeneratedRegion(readme, CATALOG).ok).toBe(false);
  });

  it('refuses a README without markers', () => {
    expect(() => replaceGeneratedRegion('# no markers\n', CATALOG)).toThrow(/BEGIN GENERATED COMMANDS/);
  });
});

describe('generate-skills (live command tree)', () => {
  let tmp: string;
  beforeEach(() => { tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-skills-')); });
  afterEach(() => { fs.rmSync(tmp, { recursive: true, force: true }); });

  const register = (p: Parameters<typeof registerMailCommands>[0], g: Parameters<typeof registerMailCommands>[1]) => {
    registerMailCommands(p, g);
    registerDescribeCommand(p);
    registerGenerateSkillsCommand(p);
  };

  it('builds the same catalogue as describe --json', async () => {
    const { buildCatalog } = await import('../src/commands/generate-skills.js');
    const r = await runCommand((p, g) => {
      register(p, g);
      p.command('capture').action(() => { console.log(JSON.stringify(buildCatalog(p))); });
    }, ['capture']);
    const described = await runCommand(register, ['describe', '--json']);
    const built = JSON.parse(r.stdout);
    const fromDescribe = JSON.parse(described.stdout);
    expect(built.commands.filter((c: { path: string }) => c.path !== 'capture')).toEqual(fromDescribe.commands);
  });

  it('writes skills for the registered groups from the live tree', async () => {
    const out = path.join(tmp, 'nested', 'skills');
    const r = await runCommand(register, ['generate-skills', '--output-dir', out]);
    expect(r.exitCode).toBe(0);
    const mail = fs.readFileSync(path.join(out, 'ms365-mail', 'SKILL.md'), 'utf8');
    expect(mail).toMatch(/\| `mail delete` .*\*\*destructive\*\*/);
    expect(mail).toContain('`mail send-draft`');
    expect(fs.existsSync(path.join(out, 'ms365-shared', 'SKILL.md'))).toBe(true);
    expect(fs.existsSync(path.join(out, 'ms365-describe', 'SKILL.md'))).toBe(true);
    expect(fs.readFileSync(path.join(out, 'INDEX.md'), 'utf8')).toContain('ms365-mail/SKILL.md');
  });

  it('honours --filter and refuses to overwrite without --force', async () => {
    const out = path.join(tmp, 'skills');
    await runCommand(register, ['generate-skills', '--output-dir', out, '--filter', '^mail$']);
    expect(fs.readdirSync(out).sort()).toEqual(['INDEX.md', 'ms365-mail', 'ms365-shared']);
    const file = path.join(out, 'ms365-mail', 'SKILL.md');
    fs.writeFileSync(file, 'hand edited');
    const refused = await runCommand(register, ['generate-skills', '--output-dir', out, '--filter', '^mail$']);
    expect(refused.exitCode).not.toBe(0);
    expect(fs.readFileSync(file, 'utf8')).toBe('hand edited');
    const forced = await runCommand(register, ['generate-skills', '--output-dir', out, '--filter', '^mail$', '--force']);
    expect(forced.exitCode).toBe(0);
    expect(fs.readFileSync(file, 'utf8')).not.toBe('hand edited');
  });
});

describe('scripts/gen-docs.ts', () => {
  let tmp: string;
  beforeEach(() => { tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-docs-')); });
  afterEach(() => { fs.rmSync(tmp, { recursive: true, force: true }); });

  const run = (...args: string[]) => spawnSync(process.execPath, ['--import', 'tsx', 'scripts/gen-docs.ts', ...args], { cwd: ROOT, encoding: 'utf8' });

  it('rewrites the README region from the live catalogue; --check passes, then fails once stale', () => {
    const readme = path.join(tmp, 'nested', 'README.md');
    fs.mkdirSync(path.dirname(readme));
    fs.writeFileSync(readme, `# x\n\n${BEGIN_MARKER}\n${END_MARKER}\n`);

    expect(run('--check', '--readme', readme).status).toBe(1);
    expect(run('--readme', readme).status).toBe(0);
    expect(run('--check', '--readme', readme).status).toBe(0);

    const written = fs.readFileSync(readme, 'utf8');
    const live = JSON.parse(execFileSync(process.execPath, ['--import', 'tsx', 'src/index.ts', 'describe', '--json'], { cwd: ROOT, encoding: 'utf8' })) as Catalog;
    for (const c of live.commands) expect(written).toContain(`| \`${c.path}\` |`);

    fs.writeFileSync(readme, written.replace(`| \`${live.commands[0].path}\` |`, '| `bogus` |'));
    const stale = run('--check', '--readme', readme);
    expect(stale.status).toBe(1);
    expect(stale.stderr).toMatch(/stale/i);
    expect(stale.stderr).toContain('npm run docs');
  }, 60_000);
});
