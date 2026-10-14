import { describe, it, expect } from 'vitest';
import { runCommand } from './helpers/cli.js';
import { registerTodoCommands } from '../src/commands/todo.js';

describe('command test harness (T07)', () => {
  it('records requests, captures stdout, and reports exit 0', async () => {
    const r = await runCommand(registerTodoCommands, ['todo', 'lists', '-o', 'json'], {
      routes: [['GET', '/me/todo/lists', () => ({ value: [{ id: 'L1', displayName: 'Groceries' }] })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0]).toMatchObject({ method: 'GET', path: '/me/todo/lists' });
    expect(r.calls[0].headers.Authorization).toBe('Bearer test-token');
    expect(JSON.parse(r.stdout).value[0].displayName).toBe('Groceries');
  });

  it('maps a Graph 404 to exit 5', async () => {
    const r = await runCommand(registerTodoCommands, ['todo', 'lists'], {
      routes: [['GET', '/me/todo/lists', () => new Response('{"error":{"message":"nope"}}', { status: 404 })]],
    });
    expect(r.exitCode).toBe(5);
    expect(r.stderr).toContain('nope');
  });

  it('usage errors exit 2', async () => {
    const r = await runCommand(registerTodoCommands, ['todo', 'lists', '--bogus']);
    expect(r.exitCode).toBe(2);
  });

  it('non-TTY confirm exits 7 and sends nothing', async () => {
    const r = await runCommand(registerTodoCommands, ['todo', 'delete', 'L', 'T']);
    expect(r.exitCode).toBe(7);
    expect(r.calls).toHaveLength(0);
  });

  it('TTY confirm "y" proceeds with DELETE', async () => {
    const r = await runCommand(registerTodoCommands, ['todo', 'delete', 'L', 'T'], { tty: true, answer: 'y' });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0]).toMatchObject({ method: 'DELETE', path: '/me/todo/lists/L/tasks/T' });
  });
});
