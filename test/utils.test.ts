import { describe, it, expect } from 'vitest';
import { parseDate, truncate, validateEmail, parseRecipients, formatSize, buildODataParams } from '../src/utils.js';

describe('parseDate', () => {
  it('accepts valid YYYY-MM-DD dates', () => {
    expect(parseDate('2024-12-25')).toBe('2024-12-25');
    expect(parseDate('2025-01-01')).toBe('2025-01-01');
  });

  it('rejects invalid date strings', () => {
    expect(() => parseDate('not-a-date')).toThrow('Invalid date');
    expect(() => parseDate('2024-13-01')).toThrow('Invalid date');
    expect(() => parseDate('')).toThrow('Invalid date');
  });

  it('accepts partial dates that JS Date parses', () => {
    // JS Date accepts these, so parseDate allows them
    expect(() => parseDate('2024')).not.toThrow();
    expect(() => parseDate('2024-12')).not.toThrow();
  });
});

describe('truncate', () => {
  it('returns short strings unchanged', () => {
    expect(truncate('hello', 10)).toBe('hello');
  });

  it('truncates at exact boundary', () => {
    expect(truncate('hello', 5)).toBe('hello');
  });

  it('adds ellipsis for long strings', () => {
    const result = truncate('hello world', 8);
    expect(result).toBe('hello w…');
    expect(result.length).toBe(8);
  });

  it('handles single-char max', () => {
    expect(truncate('hello', 1)).toBe('…');
  });
});

describe('validateEmail', () => {
  it('accepts valid emails', () => {
    expect(validateEmail('user@example.com')).toBe('user@example.com');
    expect(validateEmail('a@b.co')).toBe('a@b.co');
  });

  it('trims whitespace', () => {
    expect(validateEmail('  user@example.com  ')).toBe('user@example.com');
  });

  it('rejects invalid emails', () => {
    expect(() => validateEmail('not-email')).toThrow('Invalid email');
    expect(() => validateEmail('@no-local.com')).toThrow('Invalid email');
    expect(() => validateEmail('no-domain@')).toThrow('Invalid email');
    expect(() => validateEmail('')).toThrow('Invalid email');
  });
});

describe('parseRecipients', () => {
  it('parses single email', () => {
    const result = parseRecipients('user@example.com');
    expect(result).toHaveLength(1);
    expect(result[0].emailAddress.address).toBe('user@example.com');
  });

  it('parses comma-separated emails', () => {
    const result = parseRecipients('a@b.com, c@d.com');
    expect(result).toHaveLength(2);
    expect(result[0].emailAddress.address).toBe('a@b.com');
    expect(result[1].emailAddress.address).toBe('c@d.com');
  });

  it('throws on invalid email in list', () => {
    expect(() => parseRecipients('a@b.com, invalid')).toThrow('Invalid email');
  });
});

describe('formatSize', () => {
  it('formats bytes', () => {
    expect(formatSize(500)).toBe('500 B');
  });

  it('formats kilobytes', () => {
    expect(formatSize(1024)).toBe('1.0 KB');
    expect(formatSize(1536)).toBe('1.5 KB');
  });

  it('formats megabytes', () => {
    expect(formatSize(1024 * 1024)).toBe('1.0 MB');
  });

  it('formats gigabytes', () => {
    expect(formatSize(1024 * 1024 * 1024)).toBe('1.0 GB');
    expect(formatSize(2.5 * 1024 * 1024 * 1024)).toBe('2.5 GB');
  });

  it('formats zero', () => {
    expect(formatSize(0)).toBe('0 B');
  });
});

describe('buildODataParams', () => {
  it('maps option keys to OData params', () => {
    const opts = { top: '10', filter: "isRead eq false", select: undefined };
    const params = buildODataParams(opts, {
      top: '$top',
      filter: '$filter',
      select: '$select',
    });
    expect(params).toEqual({ '$top': '10', '$filter': 'isRead eq false' });
    expect(params['$select']).toBeUndefined();
  });

  it('returns empty object when no options match', () => {
    expect(buildODataParams({}, { top: '$top' })).toEqual({});
  });
});
