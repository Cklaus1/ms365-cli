import { describe, it, expect, afterEach, vi } from 'vitest';
import { GraphClient, graphError } from '../src/graph-client.js';

// A 403 must not suggest --org-mode to someone who is already using it.
describe('403 hint and --org-mode', () => {
  it('graphError suggests --org-mode only when it is off', () => {
    expect(graphError(403, 'Invalid permission.').message).toMatch(/Try running with --org-mode\./);
    expect(graphError(403, 'Invalid permission.', {}, { orgMode: true }).message).not.toMatch(/--org-mode/);
  });

  describe('GraphClient passes its org mode through', () => {
    const original = globalThis.fetch;
    afterEach(() => { globalThis.fetch = original; });
    const auth = { getToken: vi.fn().mockResolvedValue('t'), initialize: vi.fn() };
    const deny = () => vi.fn(async () => new Response(JSON.stringify({ error: { code: 'S2SUnauthorized', message: 'Invalid permission.' } }),
      { status: 403, headers: { 'content-type': 'application/json' } })) as never;

    it('org mode on: no --org-mode suggestion', async () => {
      globalThis.fetch = deny();
      const c = new GraphClient(auth as never, false, 'v1.0', false, { orgMode: true });
      await expect(c.request('/reports/x')).rejects.toThrow(/^(?!.*--org-mode).*Permission denied/s);
    });
    it('org mode off: suggestion kept', async () => {
      globalThis.fetch = deny();
      const c = new GraphClient(auth as never);
      await expect(c.request('/reports/x')).rejects.toThrow(/Try running with --org-mode\./);
    });
  });
});
