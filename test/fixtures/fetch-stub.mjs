// Offline Graph double for spawned CLI runs (E5): `node --import <this file> …`.
// Every fetch is logged to stderr as `[fetch] METHOD URL`, so a test can prove "zero fetch calls".
// Answers come from MS365_TEST_FETCH_ROUTES: a JSON array of {method?, match, status?, body?, headers?, hang?},
// `match` a regex tested against the full URL; the first match answers. Unmatched requests get 404 {}.
// `hang: true` never answers (honours the request's abort signal), for timeout tests.
import { writeSync } from 'node:fs';

const routes = JSON.parse(process.env.MS365_TEST_FETCH_ROUTES || '[]');

globalThis.fetch = async (input, init = {}) => {
  const url = String(input instanceof Request ? input.url : input);
  const method = (init.method || 'GET').toUpperCase();
  writeSync(2, `[fetch] ${method} ${url}\n`);
  const route = routes.find(r => (!r.method || r.method === method) && new RegExp(r.match).test(url));
  if (route?.hang) {
    return new Promise((_, reject) => {
      init.signal?.addEventListener('abort', () => reject(init.signal.reason ?? new Error('aborted')));
    });
  }
  const status = route?.status ?? (route ? 200 : 404);
  const body = route?.body ?? {};
  return new Response(status === 204 || status === 202 ? null : JSON.stringify(body), {
    status, headers: { 'content-type': 'application/json', ...(route?.headers ?? {}) },
  });
};
