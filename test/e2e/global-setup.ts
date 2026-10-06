import { build } from 'tsup';
import path from 'node:path';

/**
 * Build the CLI once per test run into this checkout's .e2e-build/ (not node_modules, which worktrees share
 * via symlink). Package imports resolve upward through the checkout's node_modules.
 */
export const E2E_OUT_DIR = path.resolve(__dirname, '../../.e2e-build');

export default async function setup(): Promise<void> {
  await build({
    config: false,
    entry: ['src/index.ts', 'src/remote.ts'],
    format: ['esm'],
    target: 'node24',
    outDir: E2E_OUT_DIR,
    clean: true,
    silent: true,
    banner: { js: '#!/usr/bin/env node' },
  });
}
