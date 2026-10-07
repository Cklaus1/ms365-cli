import { describe, it, expect, beforeAll } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { runCli, tempDir, type RunResult } from '../e2e/run-cli.js';
import { buildCommandTree, type CommandTree } from '../../src/command-tree.js';
import { introspectionProgram } from '../../src/commands/registry.js';
import { endpointKeys } from '../../src/commands/schema.js';

function readTree(dir: string): Record<string, string> {
  const files: Record<string, string> = {};
  for (const entry of fs.readdirSync(dir, { recursive: true, withFileTypes: true })) {
    if (entry.isFile()) files[path.relative(dir, path.join(entry.parentPath, entry.name))] = fs.readFileSync(path.join(entry.parentPath, entry.name), 'utf8');
  }
  return files;
}

describe('SKILLS-1 generate-skills derives the catalog from the command tree', () => {
  let tree: CommandTree;
  let run: RunResult;
  let files: Record<string, string>;
  let all: string;

  beforeAll(async () => {
    tree = buildCommandTree(introspectionProgram());
    const out = tempDir('ms365-skills-');
    run = await runCli(['generate-skills', '--output-dir', out]);
    files = readTree(out);
    all = Object.values(files).join('\n');
  });

  it('exits 0 and keeps stdout free of progress text', () => {
    expect(run.code).toBe(0);
    expect(run.stdout).toBe('');
    expect(run.stderr).toContain('Generated: INDEX.md');
  });

  it('emits a skill for every top-level group, including user and groups', () => {
    const topGroups = tree.groups.filter(g => g.names.length === 1).map(g => g.path);
    expect(topGroups).toEqual(expect.arrayContaining(['user', 'groups', 'mail', 'calendar', 'auth', 'chat']));
    for (const group of topGroups) expect(Object.keys(files)).toContain(`ms365-${group}/SKILL.md`);
    expect(Object.keys(files)).toEqual(expect.arrayContaining(['ms365-shared/SKILL.md', 'INDEX.md']));
  });

  it('documents every leaf command of the tree', () => {
    const missing = tree.leaves.filter(l => !all.includes(`ms365 ${l.names.join(' ')}`)).map(l => l.path);
    expect(missing).toEqual([]);
  });

  it('documents options of a leaf in its group skill', () => {
    expect(files['ms365-calendar/SKILL.md']).toContain('`--start <datetime>`');
    expect(files['ms365-calendar/SKILL.md']).toContain('Alias: `ms365 cal`');
  });

  it('CHOICE-1 lists the values of a comma-list option', () => {
    expect(files['ms365-search/SKILL.md']).toMatch(/`--entity <types>`.*comma-separated list of: driveItem, drive, list, listItem, site, externalItem, message, event, chatMessage, person, acronym, bookmark, qna/);
  });

  it('emits schema hints only for endpoints that exist', () => {
    const keys = new Set(endpointKeys());
    const hints = [...all.matchAll(/ms365 schema ([a-z][\w.-]*)/g)].map(m => m[1]);
    expect(hints.length).toBeGreaterThan(0);
    expect(hints.filter(h => !keys.has(h))).toEqual([]);
  });

  it('documents the global flags declared by the running program', () => {
    expect(files['ms365-shared/SKILL.md']).toContain('`-q, --quiet`');
    expect(files['ms365-shared/SKILL.md']).toContain('`--dry-run`');
  });

  it('--filter limits the group skills', async () => {
    const out = tempDir('ms365-skills-');
    const r = await runCli(['generate-skills', '--output-dir', out, '--filter', '^user$']);
    expect(r.code).toBe(0);
    expect(fs.readdirSync(out).sort()).toEqual(['INDEX.md', 'ms365-shared', 'ms365-user']);
  });

  it('the introspection program registers the same top-level commands as the CLI', async () => {
    const r = await runCli(['--help']);
    const section = r.stdout.split(/^Commands:$/m)[1];
    const cliCommands = [...section.matchAll(/^ {2}([a-z][\w-]*)(?:\|[\w-]+)?\b/gm)].map(m => m[1]).filter(n => n !== 'help');
    const treeTop = [...new Set([...tree.groups, ...tree.leaves].map(n => n.names[0]))];
    expect(cliCommands.sort()).toEqual(treeTop.sort());
  });
});
