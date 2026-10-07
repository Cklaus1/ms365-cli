import { describe, it, expect } from 'vitest';
import { runCli } from './run-cli.js';

const ENV = { MS365_CLI_CLIENT_ID: '00000000-0000-0000-0000-000000000000', MS365_CLI_TOKEN: 'fake' };
const ROUTES = [{ match: '/', status: 200, body: '{"value":[]}' }];

const OPTIONS: { name: string; args: (v: string) => string[]; min: string; outOfRange: string; expected: string }[] = [
  { name: '--top', args: (v) => ['mail', 'list', '--top', v], min: '1', outOfRange: '0', expected: 'Expected an integer >= 1.' },
  { name: 'search query --top', args: (v) => ['search', 'query', 'x', '--top', v], min: '1', outOfRange: '1001', expected: 'Expected an integer from 1 to 1000.' },
  { name: 'serve --port', args: (v) => ['serve', '--app-cap', 'c', '--port', v], min: '0', outOfRange: '65536', expected: 'Expected an integer from 0 to 65535.' },
  { name: 'create-rule --sequence', args: (v) => ['mail', 'create-rule', '--name', 'r', '--sequence', v], min: '0', outOfRange: '-1', expected: 'Expected an integer >= 0.' },
  { name: 'free-busy --interval', args: (v) => ['calendar', 'free-busy', '--emails', 'a@b.cd', '--start', '2026-10-07T09:00', '--end', '2026-10-07T10:00', '--interval', v], min: '5', outOfRange: '1441', expected: 'Expected an integer from 5 to 1440.' },
];

describe('CLEAN-1 integer options share one parser', () => {
  for (const opt of OPTIONS) {
    for (const bad of ['abc', '15x', opt.outOfRange]) {
      it(`${opt.name} ${bad} exits 1 with the shared message and sends nothing`, async () => {
        const r = await runCli(opt.args(bad), { env: ENV, routes: ROUTES });
        expect(r.code).toBe(1);
        expect(r.stderr).toContain(`argument '${bad}' is invalid. ${opt.expected}`);
        expect(r.requests).toHaveLength(0);
      });
    }
  }
});
