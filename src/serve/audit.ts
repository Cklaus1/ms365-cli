import fs from 'node:fs';
import path from 'node:path';

/** One audit line per request. Holds argument KEYS only: bodies and recipients are sensitive. */
export interface AuditEntry {
  ts: string;
  request_id: string;
  agent: string | null;
  /** `X-Forwarded-For` as set by `tailscale serve` (the caller's tailnet address). */
  src: string | null;
  method: string;
  path: string;
  command: string | null;
  args: string[];
  status: number;
  error?: string;
  duration_ms: number;
}

/** Append-only JSON-lines audit file, mode 0600 (its directory is created 0700 when missing). */
export class AuditLog {
  private readonly fd: number;

  constructor(readonly file: string) {
    fs.mkdirSync(path.dirname(file), { recursive: true, mode: 0o700 });
    this.fd = fs.openSync(file, 'a', 0o600);
    fs.fchmodSync(this.fd, 0o600);
  }

  write(entry: AuditEntry): void {
    fs.writeSync(this.fd, JSON.stringify(entry) + '\n');
  }

  close(): void {
    fs.closeSync(this.fd);
  }
}
