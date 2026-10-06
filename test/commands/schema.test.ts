import { describe, it, expect } from 'vitest';
import { runCli } from '../e2e/run-cli.js';

describe('SCHEMA-1 schema --list and unknown endpoints', () => {
  it('schema --list exits 0 and prints a JSON list of endpoints', async () => {
    const r = await runCli(['schema', '--list']);
    expect(r.code).toBe(0);
    const list = JSON.parse(r.stdout) as Array<{ endpoint: string; method: string; path: string }>;
    expect(list.find(e => e.endpoint === 'mail.list')).toMatchObject({ method: 'GET', path: '/me/messages' });
  });

  it('schema --list honours -o', async () => {
    const r = await runCli(['schema', '--list', '-o', 'csv']);
    expect(r.code).toBe(0);
    expect(r.stdout.split('\n')[0]).toContain('endpoint');
  });

  it('schema <endpoint> prints the endpoint schema as JSON', async () => {
    const r = await runCli(['schema', 'mail.send']);
    expect(r.code).toBe(0);
    expect(JSON.parse(r.stdout)).toMatchObject({ httpMethod: 'POST', path: '/me/sendMail' });
  });

  it('an unknown endpoint exits 1 with suggestions from the same service', async () => {
    const r = await runCli(['schema', 'mail.lst']);
    expect(r.code).toBe(1);
    expect(r.stdout).toBe('');
    expect(r.stderr).toContain('Unknown endpoint: mail.lst');
    expect(r.stderr).toMatch(/Did you mean: .*mail\.list/);
    expect(r.stderr).toContain('ms365 schema --list');
  });

  it('a misspelled service still gets close suggestions', async () => {
    const r = await runCli(['schema', 'contcts.list']);
    expect(r.code).toBe(1);
    expect(r.stderr).toMatch(/Did you mean: contacts\.list/);
  });

  it('no endpoint and no --list exits 1', async () => {
    const r = await runCli(['schema']);
    expect(r.code).toBe(1);
    expect(r.stderr).toContain('schema --list');
  });
});
