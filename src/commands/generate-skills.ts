import { Command, Option } from 'commander';
import { handleErrors, CliError, ExitCode } from '../errors.js';
import fs from 'fs';
import path from 'path';
import { info } from '../utils.js';
import { commandMeta } from '../catalog.js';
import { describeSelector } from './describe.js';
import { type Catalog, type CatalogCommand, type CatalogOption, renderSkills } from '../docs-gen.js';

function describeOption(o: Option): CatalogOption {
  return {
    flags: o.flags,
    long: o.long ?? null,
    short: o.short ?? null,
    description: o.description,
    required: o.mandatory,
    takesValue: o.required || o.optional,
    choices: o.argChoices ?? null,
    default: o.defaultValue ?? null,
  };
}

/** The `describe --json` catalogue, built from the live command tree under `program`. */
export function buildCatalog(program: Command): Catalog {
  const commands: CatalogCommand[] = [];
  const walk = (c: Command, chain: string[]) => {
    if (c.commands.length === 0) {
      const p = chain.join(' ');
      const meta = commandMeta(p);
      commands.push({
        path: p,
        description: c.description(),
        arguments: c.registeredArguments.map(a => ({ name: a.name(), required: a.required, variadic: a.variadic })),
        options: c.options.map(describeOption),
        scopes: meta?.scopes ?? null,
        scopesBy: describeSelector(c, meta?.scopesBy),
        destructive: meta?.destructive ?? null,
        idempotent: meta?.idempotent ?? null,
        offline: meta?.offline ?? null,
      });
    }
    for (const sub of c.commands) walk(sub, [...chain, sub.name()]);
  };
  for (const top of program.commands) walk(top, [top.name()]);
  return { version: program.version() ?? null, globalOptions: program.options.map(describeOption), commands };
}

export function registerGenerateSkillsCommand(program: Command): void {
  program
    .command('generate-skills')
    .description('Generate SKILL.md files for AI agent integration')
    .option('--output-dir <dir>', 'Output directory', './skills')
    .option('--filter <pattern>', 'Only generate skills for command groups matching this regular expression')
    .option('--force', 'Overwrite existing skill files')
    .action(handleErrors(async (opts: { outputDir: string; filter?: string; force?: boolean }) => {
      const outDir = path.resolve(opts.outputDir);
      let filter: RegExp | undefined;
      if (opts.filter) {
        try {
          filter = new RegExp(opts.filter, 'i');
        } catch {
          throw new CliError(`Invalid --filter pattern: ${opts.filter}`, { code: 'usage', exitCode: ExitCode.USAGE });
        }
      }
      const files = renderSkills(buildCatalog(program), filter);

      // Check everything first so a refusal leaves no partial regeneration behind.
      if (!opts.force) {
        const existing = files.map(([rel]) => rel).filter(rel => fs.existsSync(path.join(outDir, rel)));
        if (existing.length > 0) {
          throw new CliError(`Refusing to overwrite existing skill files in ${outDir}: ${existing.join(', ')}`,
            { code: 'file_exists', fix: 'pass --force to overwrite' });
        }
      }

      for (const [rel, content] of files) {
        const target = path.join(outDir, rel);
        fs.mkdirSync(path.dirname(target), { recursive: true });
        fs.writeFileSync(target, content);
        info(`  Generated: ${rel}`);
      }

      info(`\n${files.length} skill files written to ${outDir}/`);
    }));
}
