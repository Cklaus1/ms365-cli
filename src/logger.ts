import fs from 'fs';
import path from 'path';
import os from 'os';

export type LogLevel = 'debug' | 'info' | 'warn' | 'error' | 'none';

const LEVEL_ORDER: Record<LogLevel, number> = {
  debug: 0,
  info: 1,
  warn: 2,
  error: 3,
  none: 4,
};

class Logger {
  private level: LogLevel;
  private logFile: string | null;
  private stream: fs.WriteStream | null = null;

  constructor() {
    const envLevel = process.env.MS365_CLI_LOG_LEVEL;
    this.level = (envLevel && envLevel in LEVEL_ORDER) ? envLevel as LogLevel : 'none';
    this.logFile = process.env.MS365_CLI_LOG_FILE || null;

    if (this.logFile) {
      const dir = path.dirname(this.logFile);
      if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
      }
      // Daily rotation: append date to filename
      const date = new Date().toISOString().slice(0, 10);
      const ext = path.extname(this.logFile);
      const base = ext ? this.logFile.slice(0, -ext.length) : this.logFile;
      const rotatedPath = `${base}-${date}${ext || '.log'}`;
      this.stream = fs.createWriteStream(rotatedPath, { flags: 'a' });
    }
  }

  private shouldLog(level: LogLevel): boolean {
    return LEVEL_ORDER[level] >= LEVEL_ORDER[this.level];
  }

  private write(level: LogLevel, message: string, data?: unknown): void {
    if (!this.shouldLog(level)) return;

    const entry = {
      timestamp: new Date().toISOString(),
      level,
      message,
      ...(data !== undefined ? { data } : {}),
    };

    // Write to log file as JSON lines
    if (this.stream) {
      this.stream.write(JSON.stringify(entry) + '\n');
    }

    // Also write to stderr if no log file (or for warn/error)
    if (!this.logFile || level === 'warn' || level === 'error') {
      const prefix = level === 'error' ? 'Error' : level === 'warn' ? 'Warning' : level;
      if (level === 'debug' || level === 'info') {
        // Only to file, not stderr, unless no file configured
        if (!this.logFile) {
          console.error(`[${prefix}] ${message}`);
        }
      } else {
        console.error(`[${prefix}] ${message}`);
      }
    }
  }

  debug(message: string, data?: unknown): void {
    this.write('debug', message, data);
  }

  info(message: string, data?: unknown): void {
    this.write('info', message, data);
  }

  warn(message: string, data?: unknown): void {
    this.write('warn', message, data);
  }

  error(message: string, data?: unknown): void {
    this.write('error', message, data);
  }

  close(): void {
    if (this.stream) {
      this.stream.end();
      this.stream = null;
    }
  }
}

// Singleton
const logger = new Logger();
export default logger;
