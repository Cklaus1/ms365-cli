import { describe, it, expect, beforeAll, afterAll, vi } from 'vitest';
import type { Command } from 'commander';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { runCommand } from './helpers/cli.js';
import { MODULES, leafCommands, sampleArgv, sampleOption, writeSampleFiles, type Register } from './helpers/commands.js';

/**
 * Dry-run endpoint snapshot: every Graph-calling command is invoked with synthetic arguments under --dry-run
 * (no TTY: a dry run never prompts), and the requests it would send (method + URL, read from the stderr trace)
 * are recorded. A second pass also sets every optional value flag (DR-21). A diff in the snapshot means a
 * command's wire behavior changed — review it deliberately.
 */
let tmp: string;
const cwd = process.cwd();

beforeAll(() => {
  tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-snap-'));
  writeSampleFiles(tmp);
  process.chdir(tmp);
  vi.stubEnv('MS365_CLI_TIMEZONE', 'America/Los_Angeles');
  vi.setSystemTime(new Date('2026-10-14T17:00:00Z'));
});
afterAll(() => { process.chdir(cwd); vi.unstubAllEnvs(); vi.useRealTimers(); fs.rmSync(tmp, { recursive: true, force: true }); });

const REQUEST = /^\[dry-run\] (GET|POST|PUT|PATCH|DELETE) /;
const TRAILER = /^\[dry-run\] (\d+) request\(s\) previewed; nothing was sent(; bodies use placeholders for values read from Graph)?\.$/;

/** Every optional option that takes a value, sampled like the mandatory ones (--body-file replaces --body and create-rule's --from-address excludes --from, so not both). */
const EXCLUSIVE = new Set(['--body-file', '--from-address']);
const optionalValues = (cmd: Command) => cmd.options.filter(o => !o.mandatory && (o.required || o.optional) && !EXCLUSIVE.has(o.long ?? '')).flatMap(sampleOption);

async function snapshotRun(register: Register, argv: string[]) {
  const r = await runCommand(register, argv, { dryRun: true });
  const stderr = r.stderr.split('\n').filter(Boolean);
  const trace = stderr.filter(l => l.startsWith('[dry-run] ') && !TRAILER.test(l) && !/^\[dry-run\] (would ask|Body|Headers|File)/.test(l));

  // DR AC1: a dry run exits 0, or 2 for a usage error in the synthetic input.
  expect([0, 2], `${argv.join(' ')}\n${r.stderr}`).toContain(r.exitCode);
  // DR AC2: the trace is stderr only; a successful dry run ends with exactly one trailer counting the requests.
  expect(r.stdout).not.toContain('[dry-run]');
  if (r.exitCode === 0) {
    expect(stderr.filter(l => TRAILER.test(l)), argv.join(' ')).toHaveLength(1);
    expect(stderr.at(-1)).toMatch(TRAILER);
    expect(Number(TRAILER.exec(stderr.at(-1)!)![1]), argv.join(' ')).toBe(trace.filter(l => REQUEST.test(l)).length);
  }

  const error = stderr.find(l => /^error:/i.test(l)) ?? '';
  const clean = (s: string) => s.replaceAll(tmp, '<tmp>');
  return { argv: clean(argv.join(' ')), exitCode: r.exitCode, requests: trace.map(clean), error: clean(error) };
}

describe('dry-run endpoint snapshot (every command)', () => {
  for (const register of MODULES) {
    for (const { chain, cmd } of leafCommands(register)) {
      // Sampled inside each test: the sample files exist only after beforeAll.
      it(chain.join(' '), async () => {
        expect(await snapshotRun(register, sampleArgv(chain, cmd))).toMatchSnapshot();
      });
      if (optionalValues(cmd).length > 0) {
        it(`${chain.join(' ')} [optional flags]`, async () => {
          expect(await snapshotRun(register, [...sampleArgv(chain, cmd), ...optionalValues(cmd)])).toMatchSnapshot();
        });
      }
    }
  }
});
