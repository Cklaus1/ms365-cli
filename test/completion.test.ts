import { describe, it, expect } from 'vitest';
import { spawnSync } from 'node:child_process';
import { Command } from 'commander';
import { runCommand } from './helpers/cli.js';
import { registerMailCommands } from '../src/commands/mail.js';
import { registerCalendarCommands } from '../src/commands/calendar.js';
import { registerCompletionCommand } from '../src/commands/completion.js';
import type { GraphClient } from '../src/graph-client.js';

function register(program: Command, getClient: () => GraphClient): void {
  program
    .option('--fields <list>', 'Only output these fields')
    .option('--api-version <version>', 'Graph API version')
    .option('--dry-run', 'Preview');
  registerMailCommands(program, getClient);
  registerCalendarCommands(program, getClient);
  const auth = program.command('auth');
  const profile = auth.command('profile');
  profile.command('list');
  profile.command('show').argument('<name>');
  registerCompletionCommand(program);
}

async function script(shell: string): Promise<string> {
  const r = await runCommand(register, ['completion', shell]);
  expect(r.exitCode).not.toBe(2);
  return r.stdout;
}

function has(cmd: string): boolean {
  return spawnSync(cmd, ['-c', 'exit 0'], { stdio: 'ignore' }).status === 0;
}

/** Run the bash completion function for a command line (cursor at the end). */
function bashComplete(src: string, line: string): string[] {
  const words = line.split(' ');
  const driver = [
    src,
    `COMP_WORDS=(${words.map(w => `'${w}'`).join(' ')})`,
    `COMP_CWORD=${words.length - 1}`,
    '_ms365',
    'printf "%s\\n" "${COMPREPLY[@]}"',
  ].join('\n');
  const r = spawnSync('bash', ['--norc', '--noprofile', '-c', driver], { encoding: 'utf8' });
  expect(r.stderr).toBe('');
  return r.stdout.split('\n').filter(Boolean);
}

describe('ms365 completion', () => {
  it('rejects an unknown shell with a usage error', async () => {
    const r = await runCommand(register, ['completion', 'tcsh']);
    expect(r.exitCode).toBe(2);
  });

  it('--help shows install instructions for every shell', async () => {
    const r = await runCommand(register, ['completion', '--help']);
    expect(r.stdout).toContain('ms365 completion bash > ~/.local/share/bash-completion/completions/ms365');
    expect(r.stdout).toContain('_ms365');
    expect(r.stdout).toContain('~/.config/fish/completions/ms365.fish');
    expect(r.stdout).toContain('$PROFILE');
  });

  describe('bash', () => {
    it('is static, valid bash, and lists nested subcommands', async () => {
      const src = await script('bash');
      expect(src).toContain('complete -o default -F _ms365 ms365');
      expect(src).toContain("'mail|list'");
      expect(src).toContain("'auth profile'");
      // static: never runs ms365 at completion time
      expect(src).not.toMatch(/\$\(\s*ms365/);
      if (!has('bash')) return;
      const r = spawnSync('bash', ['-n'], { input: src, encoding: 'utf8' });
      expect(r.status, r.stderr).toBe(0);
    });

    it.runIf(has('bash'))('completes subcommands, options, choices and aliases', async () => {
      const src = await script('bash');
      expect(bashComplete(src, 'ms365 ma')).toEqual(['mail']);
      expect(bashComplete(src, 'ms365 mail li')).toEqual(['list']);
      expect(bashComplete(src, 'ms365 auth profile ')).toEqual(['list', 'show']);
      expect(bashComplete(src, 'ms365 mail list -o ')).toEqual(['json', 'ndjson', 'table', 'csv', 'text', 'yaml']);
      expect(bashComplete(src, 'ms365 mail list --output t')).toEqual(['table', 'text']);
      expect(bashComplete(src, 'ms365 mail list --to')).toContain('--top');
      // global options everywhere, and global option values are skipped when walking
      expect(bashComplete(src, 'ms365 mail list --fie')).toEqual(['--fields']);
      expect(bashComplete(src, 'ms365 mail list --he')).toEqual(['--help']);
      expect(bashComplete(src, 'ms365 --he')).toEqual(['--help']);
      expect(bashComplete(src, 'ms365 --fields id mail li')).toEqual(['list']);
      expect(bashComplete(src, 'ms365 --api-version beta mail list -o j')).toEqual(['json']);
      // alias walks into the canonical command
      expect(bashComplete(src, 'ms365 cal li')).toEqual(['list']);
      // positional choices
      expect(bashComplete(src, 'ms365 completion f')).toEqual(['fish']);
      // a free-form value falls back to file completion (empty reply)
      expect(bashComplete(src, 'ms365 mail list --fields ')).toEqual([]);
    });
  });

  it('zsh has #compdef and the same tables', async () => {
    const src = await script('zsh');
    expect(src.startsWith('#compdef ms365')).toBe(true);
    expect(src).toContain("'mail|list'");
    expect(src).toContain('compdef _ms365 ms365');
    if (!has('zsh')) return;
    const r = spawnSync('zsh', ['-n'], { input: src, encoding: 'utf8' });
    expect(r.status, r.stderr).toBe(0);
  });

  it('fish uses complete -c ms365', async () => {
    const src = await script('fish');
    expect(src).toContain('complete -c ms365');
    expect(src).toContain("case 'mail|list'");
    if (!has('fish')) return;
    const r = spawnSync('fish', ['--no-execute'], { input: src, encoding: 'utf8' });
    expect(r.status, r.stderr).toBe(0);
  });

  it('powershell registers Register-ArgumentCompleter', async () => {
    const src = await script('powershell');
    expect(src).toContain('Register-ArgumentCompleter -Native -CommandName ms365');
    expect(src).toContain("$child['mail|list'] = 'mail list'");
    expect(src).toContain("'json', 'ndjson', 'table'");
  });

  it('includes commands registered after it (generated at action time)', async () => {
    const r = await runCommand((program, getClient) => {
      register(program, getClient);
      program.command('late').command('added');
    }, ['completion', 'bash']);
    expect(r.stdout).toContain("'late|added'");
  });
});
