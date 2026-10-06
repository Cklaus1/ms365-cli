import { describe, it, expect, vi } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { AuditLog, type AuditEntry } from '../../src/serve/audit.js';
import { tempDir } from '../e2e/run-cli.js';

const entry: AuditEntry = {
  ts: '2026-01-01T00:00:00.000Z', request_id: 'req_1', agent: null, src: null,
  method: 'GET', path: '/v1/health', command: null, args: [], status: 200, duration_ms: 1,
};

describe('AuditLog', () => {
  it('reports a failed write on stderr instead of throwing', () => {
    const log = new AuditLog(path.join(tempDir(), 'audit.jsonl'));
    log.write(entry);
    expect(fs.readFileSync(log.file, 'utf8')).toBe(JSON.stringify(entry) + '\n');
    log.close(); // the descriptor is gone: the next write fails like a full disk would
    const stderr = vi.spyOn(console, 'error').mockImplementation(() => {});
    try {
      expect(() => log.write(entry)).not.toThrow();
      expect(stderr).toHaveBeenCalledWith(expect.stringContaining('audit log write failed'));
    } finally {
      stderr.mockRestore();
    }
  });
});
