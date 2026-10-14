import { describe, it, expect } from 'vitest';
import { sourceFiles } from './helpers/commands.js';

/**
 * DR-20 / DR AC10: code tells a preview from a Graph answer with GraphClient.isPreview(res)
 * or the utils isDryRun() flag, never by sniffing the placeholder's `dryRun` property.
 * graph-client.ts builds the placeholder; index.ts and commands/mcp.ts read the CLI option.
 */
const ALLOWED = new Set(['graph-client.ts', 'index.ts', 'commands/mcp.ts']);

describe('dry-run detection (DR-20)', () => {
  it('no .dryRun property reads or `dryRun?: boolean` casts outside the allow-list', () => {
    const hits = sourceFiles().filter(f => !ALLOWED.has(f.rel)).flatMap(f => f.lines
      .flatMap((line, i) => /\.dryRun\b|dryRun\?\s*:\s*boolean/.test(line) ? [`src/${f.rel}:${i + 1}: ${line.trim()}`] : []));
    expect(hits, 'use client.isPreview(res) or isDryRun() from utils').toEqual([]);
  });
});
