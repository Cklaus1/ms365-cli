import { describe, it, expect } from 'vitest';
import { buildCommandTree } from '../../src/command-tree.js';
import { introspectionProgram } from '../../src/commands/registry.js';
import { COMMAND_POLICY } from '../../src/serve/capabilities.js';

describe('REMOTE-1c command → capability table', () => {
  const tree = buildCommandTree(introspectionProgram());
  const leaves = tree.leaves.map(l => l.path);

  it('maps every registered command or marks it remote: false', () => {
    expect(leaves.filter(p => !Object.hasOwn(COMMAND_POLICY, p))).toEqual([]);
  });

  it('has no entries for commands that do not exist', () => {
    expect(Object.keys(COMMAND_POLICY).filter(p => !leaves.includes(p))).toEqual([]);
  });

  it('never runs auth, setup, serve or generate-skills remotely', () => {
    const local = leaves.filter(p => p.startsWith('auth.') || ['serve', 'generate-skills'].includes(p));
    expect(local).toContain('auth.setup');
    for (const p of local) expect(COMMAND_POLICY[p], p).toEqual({ remote: false });
  });
});
