// Preloaded with `node --import` by the e2e harness. Replaces global fetch with
// canned responses so the built binary can run offline.
//   MS365_FAKE_GRAPH      path to JSON: [{ match, method?, status?, headers?, body?, bodyBase64?, delayMs? }]
//                         `match` is a regex tested against the full request URL; first match wins.
//   MS365_FAKE_GRAPH_LOG  optional path; one JSON line per request { method, url, headers, body }.
// The same settings can come from the import URL instead (`fake-fetch.mjs?routes=…&log=…&procs=…`), which
// survives into children that inherit `process.execArgv` but not the environment (ms365 serve).
//   procs                 optional path; one JSON line per process start and exit { event, pid, t }.
import fs from 'node:fs';

const params = new URL(import.meta.url).searchParams;
const routesPath = params.get('routes') ?? process.env.MS365_FAKE_GRAPH;
const logPath = params.get('log') ?? process.env.MS365_FAKE_GRAPH_LOG;
const procsPath = params.get('procs');
const routes = routesPath ? JSON.parse(fs.readFileSync(routesPath, 'utf8')) : [];

if (procsPath) {
  const record = (event) =>
    fs.appendFileSync(procsPath, JSON.stringify({ event, pid: process.pid, t: performance.timeOrigin + performance.now() }) + '\n');
  record('start');
  process.on('exit', () => record('exit'));
}

function bodyToString(body) {
  if (body == null) return null;
  if (typeof body === 'string') return body;
  if (Buffer.isBuffer(body) || body instanceof Uint8Array) return `base64:${Buffer.from(body).toString('base64')}`;
  return String(body);
}

globalThis.fetch = async (input, init = {}) => {
  const url = typeof input === 'string' ? input : input.url;
  const method = (init.method || 'GET').toUpperCase();
  if (logPath) {
    const headers = Object.fromEntries(new Headers(init.headers || {}).entries());
    fs.appendFileSync(logPath, JSON.stringify({ method, url, headers, body: bodyToString(init.body) }) + '\n');
  }
  const route = routes.find((r) => new RegExp(r.match).test(url) && (!r.method || r.method === method));
  if (!route) {
    return new Response(JSON.stringify({ error: { code: 'NoFakeRoute', message: `${method} ${url}` } }), {
      status: 404,
      headers: { 'content-type': 'application/json' },
    });
  }
  if (route.delayMs) await new Promise((r) => setTimeout(r, route.delayMs));
  const status = route.status ?? 200;
  const body = route.bodyBase64 !== undefined
    ? Buffer.from(route.bodyBase64, 'base64')
    : route.body === undefined || status === 204 ? null
    : typeof route.body === 'string' ? route.body : JSON.stringify(route.body);
  const headers = { 'content-type': 'application/json', ...(route.headers || {}) };
  return new Response(body, { status, headers });
};
