// `--import` preload: writes `[probe] msal-node loaded` to stderr whenever anything resolves @azure/msal-node.
import { register } from 'node:module';

register(new URL('./msal-probe-hooks.mjs', import.meta.url).href);
