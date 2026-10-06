import { describe, it, expect } from 'vitest';
import { buildScopes } from '../src/auth.js';

describe('buildScopes', () => {
  it('includes personal scopes by default', () => {
    const scopes = buildScopes(false);
    expect(scopes).toContain('User.Read');
    expect(scopes).toContain('Mail.ReadWrite');
    expect(scopes).toContain('Mail.Send');
    expect(scopes).toContain('Calendars.ReadWrite');
    expect(scopes).toContain('Files.ReadWrite');
  });

  it('excludes org scopes when orgMode is false', () => {
    const scopes = buildScopes(false);
    expect(scopes).not.toContain('Chat.Read');
    expect(scopes).not.toContain('Team.ReadBasic.All');
    expect(scopes).not.toContain('Sites.Read.All');
  });

  it('includes org scopes when orgMode is true', () => {
    const scopes = buildScopes(true);
    expect(scopes).toContain('Chat.Read');
    expect(scopes).toContain('Team.ReadBasic.All');
    expect(scopes).toContain('Sites.Read.All');
    expect(scopes).toContain('User.Read.All');
  });

  it('removes redundant lower scopes via hierarchy', () => {
    const scopes = buildScopes(false);
    // Mail.ReadWrite subsumes Mail.Read
    expect(scopes).toContain('Mail.ReadWrite');
    expect(scopes).not.toContain('Mail.Read');
  });

  it('deduplicates Calendars.Read when Calendars.ReadWrite present', () => {
    const scopes = buildScopes(false);
    expect(scopes).toContain('Calendars.ReadWrite');
    expect(scopes).not.toContain('Calendars.Read');
  });

  it('deduplicates Presence.Read when Presence.ReadWrite present', () => {
    const scopes = buildScopes(false);
    expect(scopes).toContain('Presence.ReadWrite');
    expect(scopes).not.toContain('Presence.Read');
  });

  it('deduplicates all ReadWrite/Read pairs', () => {
    const scopes = buildScopes(false);
    const pairs = [
      ['Mail.ReadWrite', 'Mail.Read'],
      ['Calendars.ReadWrite', 'Calendars.Read'],
      ['Files.ReadWrite', 'Files.Read'],
      ['Tasks.ReadWrite', 'Tasks.Read'],
      ['Contacts.ReadWrite', 'Contacts.Read'],
      ['Presence.ReadWrite', 'Presence.Read'],
    ];
    for (const [higher, lower] of pairs) {
      if (scopes.includes(higher)) {
        expect(scopes).not.toContain(lower);
      }
    }
  });

  it('returns no duplicate scopes', () => {
    const scopes = buildScopes(true);
    const unique = new Set(scopes);
    expect(scopes.length).toBe(unique.size);
  });

  it('always includes User.Read', () => {
    expect(buildScopes(false)).toContain('User.Read');
    expect(buildScopes(true)).toContain('User.Read');
  });
});
