import { describe, it, expect } from 'vitest';
import { sourceFiles } from './helpers/commands.js';

/**
 * V spec AC 1 + VT9: user-input validation throws CliError (exit 2 with a code), never a bare Error, and status
 * lines ("Email sent.", "Saved to x") go to stderr through info()/success(), never console.log. Data output —
 * reports, tables, "Logged in as …" (V Q7) — stays on stdout.
 */
const STATUS = /\b(sent|created|deleted|updated|saved|moved|uploaded|downloaded|removed|cancell?ed|forwarded|replied|copied|renamed|cleared|marked|accepted|declined)\b|^(Error|Warning|Note):|^No [\w ]+ found/i;

/** The literal a console.log line opens with ('…', "…" or `…`; ${…} blanked), or null when it logs an expression. */
function loggedLiteral(line: string): string | null {
  const m = /console\.log\(\s*(['"`])((?:\\.|(?!\1).)*)/.exec(line);
  return m ? m[2].replace(/\$\{[^}]*\}/g, 'X').replace(/^\\n/, '') : null;
}

function isStatusLog(line: string): boolean {
  const lit = loggedLiteral(line);
  return lit !== null && STATUS.test(lit.trim());
}

const commandFiles = () => sourceFiles().filter(f => f.rel.startsWith('commands/') || f.rel === 'utils.ts');
const scan = (test: (line: string) => boolean) => commandFiles().flatMap(f => f.lines
  .flatMap((line, i) => test(line) ? [`src/${f.rel}:${i + 1}: ${line.trim()}`] : []));

describe('command output conventions (VT9)', () => {
  it('no `throw new Error(` in src/commands or src/utils.ts (use CliError / usage errors)', () => {
    expect(scan(line => line.includes('throw new Error('))).toEqual([]);
  });

  it('no status line printed with console.log (use info() or success())', () => {
    expect(scan(isStatusLog)).toEqual([]);
  });

  it('the status detector flags status lines and passes data lines', () => {
    expect(isStatusLog("console.log('Email sent.');")).toBe(true);
    expect(isStatusLog('console.log(`Photo saved to ${outPath}`);')).toBe(true);
    expect(isStatusLog("console.log('No messages found.');")).toBe(true);
    expect(isStatusLog('console.log(`Logged in as ${name} (${email})`);')).toBe(false);
    expect(isStatusLog('console.log(`  ${link.webUrl}`);')).toBe(false);
    expect(isStatusLog('console.log(formatOutput(data, format));')).toBe(false);
  });
});
