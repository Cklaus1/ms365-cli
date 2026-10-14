import { defineConfig } from 'tsup';

export default defineConfig({
  entry: ['src/index.ts'],
  format: ['esm'],
  target: 'es2020',
  outDir: 'dist',
  clean: true,
  sourcemap: true,
  dts: false,
  // keytar is an optional native dependency: load it from node_modules at runtime, never bundle it
  external: ['keytar'],
  banner: {
    js: '#!/usr/bin/env node',
  },
});
