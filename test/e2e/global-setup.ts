import { build } from 'tsup';
import path from 'node:path';

/** Build the CLI once per test run into node_modules/.cache so imports resolve against the repo's node_modules. */
export const E2E_OUT_DIR = path.resolve(__dirname, '../../node_modules/.cache/ms365-e2e');

export default async function setup(): Promise<void> {
  await build({
    config: false,
    entry: ['src/index.ts'],
    format: ['esm'],
    target: 'es2020',
    outDir: E2E_OUT_DIR,
    clean: true,
    silent: true,
    banner: { js: '#!/usr/bin/env node' },
  });
}
