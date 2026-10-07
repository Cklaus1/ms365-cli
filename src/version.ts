import pkg from '../package.json' with { type: 'json' };

/** The package version (package.json is the single source; tsup inlines it into the bundle). */
export const VERSION: string = pkg.version;
