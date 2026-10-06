import { describe, it, expect } from 'vitest';
import { CapsHeaderError, parseAppCapabilities } from '../../src/serve/app-caps.js';

const CAP = 'example.com/cap/ms365';

describe('REMOTE-1b Tailscale-App-Capabilities parsing', () => {
  it('unions the allow lists of every matching grant', () => {
    const header = JSON.stringify({
      [CAP]: [{ agent: 'chief', allow: ['mail.read'] }, { agent: 'chief', allow: ['calendar.read', 'mail.read'] }],
      'other.example/cap/x': [{ allow: ['mail.send'] }],
    });
    const caller = parseAppCapabilities(header, CAP);
    expect(caller.agent).toBe('chief');
    expect([...caller.allow].sort()).toEqual(['calendar.read', 'mail.read']);
  });

  it('accepts the one-copy array of req.headersDistinct', () => {
    const caller = parseAppCapabilities([JSON.stringify({ [CAP]: [{ agent: 'chief', allow: ['mail.read'] }] })], CAP);
    expect([...caller.allow]).toEqual(['mail.read']);
  });

  it.each([
    ['missing', undefined],
    ['empty', ''],
    ['repeated', ['{}', '{}']],
    ['no copies', []],
    ['{}', '{}'],
    ['RFC 2047', '=?utf-8?b?e30=?='],
    ['malformed JSON', '{"example.com/cap/ms365": ['],
    ['array', '[]'],
    ['other key only', JSON.stringify({ 'example.com/cap/other': [{ allow: ['mail.read'] }] })],
    ['no values', JSON.stringify({ [CAP]: [] })],
    ['empty allow', JSON.stringify({ [CAP]: [{ agent: 'a', allow: [] }] })],
    ['allow not a list', JSON.stringify({ [CAP]: [{ allow: 'mail.read' }] })],
    ['non-string capability', JSON.stringify({ [CAP]: [{ allow: [1] }] })],
    ['non-string agent', JSON.stringify({ [CAP]: [{ agent: 1, allow: ['mail.read'] }] })],
  ])('rejects %s', (_label, header) => {
    expect(() => parseAppCapabilities(header, CAP)).toThrow(CapsHeaderError);
  });
});
