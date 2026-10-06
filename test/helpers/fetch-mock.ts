import { vi, type Mock } from 'vitest';
import type { AuthManager } from '../../src/auth.js';

export interface RecordedCall {
  url: string;
  method: string;
  headers: Record<string, string>;
  /** Request body as sent: string, or Buffer for binary bodies. */
  body: string | Buffer | undefined;
}

export type Handler = (call: RecordedCall, index: number) => Response | Promise<Response>;

/** Replace global fetch; every call is recorded. Call `restore()` in afterEach. */
export function mockFetch(handler: Handler): { calls: RecordedCall[]; restore: () => void } {
  const original = globalThis.fetch;
  const calls: RecordedCall[] = [];
  const fake = async (input: string | URL | Request, init: RequestInit = {}): Promise<Response> => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    const raw = init.body;
    const body = raw === undefined || raw === null ? undefined
      : typeof raw === 'string' ? raw
      : raw instanceof Uint8Array ? Buffer.from(raw)
      : String(raw);
    const call: RecordedCall = {
      url,
      method: (init.method ?? 'GET').toUpperCase(),
      headers: Object.fromEntries(new Headers(init.headers).entries()),
      body,
    };
    calls.push(call);
    return handler(call, calls.length - 1);
  };
  globalThis.fetch = vi.fn(fake) as unknown as typeof fetch;
  return { calls, restore: () => { globalThis.fetch = original; } };
}

/** Respond in order; the last response repeats. */
export function sequence(...responses: Array<() => Response | Promise<Response>>): Handler {
  return (_call, i) => responses[Math.min(i, responses.length - 1)]();
}

export function jsonResponse(body: unknown, status = 200, headers: Record<string, string> = {}): Response {
  return new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json', ...headers } });
}

export function graphError(status: number, message = 'boom', headers: Record<string, string> = {}): Response {
  return jsonResponse({ error: { code: 'Err', message } }, status, headers);
}

/** An error shaped like the one fetch rejects with when its AbortSignal fires. */
export function abortError(): Error {
  const err = new Error('This operation was aborted');
  err.name = 'AbortError';
  return err;
}

/** The `value` array of a collection payload; throws if the payload has none. */
export function valuesOf(data: unknown): unknown[] {
  if (typeof data === 'object' && data !== null && 'value' in data && Array.isArray(data.value)) return data.value;
  throw new Error(`not a collection: ${JSON.stringify(data)}`);
}

export interface FakeAuth {
  getToken: Mock<(forceRefresh?: boolean) => Promise<string>>;
}

/** Minimal AuthManager stand-in: hands out `token-1`, then `token-2` on forceRefresh. */
export function fakeAuth(): FakeAuth & AuthManager {
  let n = 1;
  const auth = {
    getToken: vi.fn(async (forceRefresh?: boolean) => {
      if (forceRefresh) n++;
      return `token-${n}`;
    }),
  };
  // The client only calls getToken; the cast documents that narrow contract.
  return auth as unknown as FakeAuth & AuthManager;
}
