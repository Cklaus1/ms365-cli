import { describe, it, expect } from 'vitest';
import { readFileSync, readdirSync } from 'node:fs';
import path from 'node:path';
import { spawnSync } from 'node:child_process';
import { runCli, REMOTE_ENTRY, tempDir } from './run-cli.js';

describe('META-2 version comes from package.json', () => {
  const pkg = JSON.parse(readFileSync(new URL('../../package.json', import.meta.url), 'utf8')) as { version: string };

  it('ms365 --version', async () => {
    const r = await runCli(['--version']);
    expect(r.code).toBe(0);
    expect(r.stdout.trim()).toBe(pkg.version);
  });

  it('ms365-remote --version', () => {
    const r = spawnSync(process.execPath, [REMOTE_ENTRY, '--version'], { encoding: 'utf8' });
    expect(r.status).toBe(0);
    expect(r.stdout.trim()).toBe(pkg.version);
  });
});

describe('META-2 generated skills carry the package version', () => {
  it('every SKILL.md front-matter version equals package.json', async () => {
    const pkg = JSON.parse(readFileSync(new URL('../../package.json', import.meta.url), 'utf8')) as { version: string };
    const out = tempDir();
    const r = await runCli(['generate-skills', '--output-dir', out, '--filter', '^(user|shared)$']);
    expect(r.code).toBe(0);
    const files = readdirSync(out, { recursive: true, encoding: 'utf8' }).filter((f) => f.endsWith('SKILL.md'));
    expect(files.length).toBeGreaterThan(0);
    for (const f of files) expect(readFileSync(path.join(out, f), 'utf8')).toMatch(new RegExp(`^version: ${pkg.version.replace(/\./g, '\\.')}$`, 'm'));
  });
});
