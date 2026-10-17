import { writeSync } from 'node:fs';

export async function resolve(specifier, context, next) {
  if (specifier === '@azure/msal-node' || specifier.startsWith('@azure/msal-node/')) writeSync(2, '[probe] msal-node loaded\n');
  return next(specifier, context);
}
