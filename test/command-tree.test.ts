import { describe, it, expect } from 'vitest';
import { Command, Option } from 'commander';
import {
  buildArgv, buildCommandTree, findCommand, InvalidInvocationError, listChoicesOption, validateInvocation, type CommandNode,
} from '../src/command-tree.js';
import { introspectionProgram } from '../src/commands/registry.js';

/** Program shaped like the CLI root: global flag, positional options, one group with an alias. */
function sampleProgram(): Command {
  const program = new Command('ms365').option('--dry-run', 'Preview').enablePositionalOptions().exitOverride();
  const mail = program.command('mail').alias('m').description('Mail');
  mail
    .command('read <messageId> [extra...]')
    .alias('r')
    .description('Read a message')
    .option('--text', 'Plain text')
    .option('--no-notify', 'Do not notify')
    .option('--tag <tags...>', 'Tags')
    .option('--top <n>', 'Count', '10')
    .requiredOption('--folder <id>', 'Folder')
    .addOption(new Option('-o, --output <format>', 'Format').choices(['json', 'table']).default('table'))
    .action(() => {});
  program.command('schema [endpoint]').description('Schema').action(() => {});
  return program;
}

/** Find the Commander object for a leaf, replace its action, and parse `argv` through the root. */
async function parseLeaf(program: Command, leaf: CommandNode, argv: string[]) {
  let cmd = program;
  for (const name of leaf.names) cmd = cmd.commands.find(c => c.name() === name)!;
  let captured: { args: unknown[]; opts: Record<string, unknown> } | undefined;
  cmd.action((...a: unknown[]) => {
    const command = a.at(-1) as Command;
    captured = { args: command.processedArgs, opts: command.opts() };
  });
  const visit = (c: Command): void => { c.exitOverride(); c.commands.forEach(visit); };
  visit(program);
  await program.parseAsync(argv, { from: 'user' });
  return captured!;
}

describe('SKILLS-1 buildCommandTree', () => {
  const tree = buildCommandTree(sampleProgram());
  const read = tree.leaves.find(l => l.path === 'mail.read')!;

  it('describes leaves with path, groups, arguments and options', () => {
    expect(tree.leaves.map(l => l.path)).toEqual(['mail.read', 'schema']);
    expect(tree.groups).toEqual([{ path: 'mail', names: ['mail'], aliases: ['m'], description: 'Mail' }]);
    expect(tree.options.map(o => o.attribute)).toEqual(['dryRun']);
    expect(read).toMatchObject({ names: ['mail', 'read'], aliases: ['r'], groups: ['mail'], description: 'Read a message' });
    expect(read.arguments).toEqual([
      { name: 'messageId', description: '', required: true, variadic: false },
      { name: 'extra', description: '', required: false, variadic: true },
    ]);
    const byAttr = Object.fromEntries(read.options.map(o => [o.attribute, o]));
    expect(byAttr.text).toMatchObject({ long: '--text', takesValue: false, required: false, negate: false });
    expect(byAttr.notify).toMatchObject({ long: '--no-notify', negate: true, takesValue: false });
    expect(byAttr.tag).toMatchObject({ variadic: true, takesValue: true });
    expect(byAttr.top).toMatchObject({ takesValue: true, default: '10', required: false });
    expect(byAttr.folder).toMatchObject({ required: true });
    expect(byAttr.output).toMatchObject({ short: '-o', long: '--output', choices: ['json', 'table'], default: 'table' });
  });

  it('is JSON-serializable', () => {
    expect(JSON.parse(JSON.stringify(tree))).toEqual(tree);
  });

  it('excludes help and covers every registered group of the real CLI', () => {
    const real = buildCommandTree(introspectionProgram());
    expect(real.leaves.some(l => l.names.includes('help'))).toBe(false);
    for (const p of ['user.me', 'groups.list', 'auth.setup', 'calendar.view', 'schema', 'generate-skills']) {
      expect(findCommand(real, p)?.path).toBe(p);
    }
  });
});

describe('SKILLS-1 findCommand', () => {
  const tree = buildCommandTree(sampleProgram());
  it('looks up canonical and alias paths', () => {
    expect(findCommand(tree, 'mail.read')?.path).toBe('mail.read');
    expect(findCommand(tree, 'm.r')?.path).toBe('mail.read');
    expect(findCommand(tree, 'schema')?.path).toBe('schema');
    expect(findCommand(tree, 'mail')).toBeUndefined();
    expect(findCommand(tree, 'mail.nope')).toBeUndefined();
    expect(findCommand(buildCommandTree(introspectionProgram()), 'cal.view')?.path).toBe('calendar.view');
  });
});

describe('SKILLS-1 validateInvocation', () => {
  const read = findCommand(buildCommandTree(sampleProgram()), 'mail.read')!;
  const ok = { positionals: ['id1'], args: { folder: 'f' } };

  it('accepts a valid invocation', () => {
    expect(validateInvocation(read, ok)).toEqual([]);
    expect(validateInvocation(read, { positionals: ['id', 'a', 'b'], args: { folder: 'f', tag: ['x', 'y'], top: 5, text: true, notify: false, output: 'json' } })).toEqual([]);
  });

  it('rejects unknown keys, bad choices, missing required options and wrong types', () => {
    expect(validateInvocation(read, { ...ok, args: { folder: 'f', dryRun: true } })).toEqual(['unknown option: dryRun']);
    expect(validateInvocation(read, { ...ok, args: { folder: 'f', output: 'yaml' } })).toEqual(['option output must be one of: json, table']);
    expect(validateInvocation(read, { positionals: ['id'], args: {} })).toEqual(['missing required option: folder']);
    expect(validateInvocation(read, { ...ok, args: { folder: 'f', text: 'yes' } })).toEqual(['option text must be a boolean']);
    expect(validateInvocation(read, { ...ok, args: { folder: true } })).toEqual(['option folder requires a value']);
    expect(validateInvocation(read, { ...ok, args: { folder: ['a'] } })).toHaveLength(1);
    expect(validateInvocation(read, { ...ok, args: { folder: { a: 1 } } })).toHaveLength(1);
  });

  it('checks positional count and type', () => {
    expect(validateInvocation(read, { args: { folder: 'f' } })).toEqual(['mail.read expects at least 1 positional argument(s), got 0']);
    expect(validateInvocation(read, { positionals: [{}], args: { folder: 'f' } })).toEqual(['positional <messageId> must be a string or number']);
    const schema = findCommand(buildCommandTree(sampleProgram()), 'schema')!;
    expect(validateInvocation(schema, { positionals: ['a', 'b'] })).toEqual(['schema expects 0-1 positional argument(s), got 2']);
  });

  it('buildArgv throws InvalidInvocationError listing the problems', () => {
    expect(() => buildArgv(read, { positionals: [] })).toThrow(InvalidInvocationError);
  });
});

describe('CHOICE-1 comma-list choices', () => {
  function listProgram(): Command {
    const program = new Command('ms365').exitOverride();
    program
      .command('find')
      .addOption(listChoicesOption('--kind <kinds>', 'Kinds', ['a', 'b', 'c']).default('a'))
      .action(() => {});
    program.commands.forEach(c => c.exitOverride());
    return program;
  }
  const find = findCommand(buildCommandTree(listProgram()), 'find')!;

  it('carries the allowed values on the option node (serializable)', () => {
    const kind = find.options.find(o => o.attribute === 'kind')!;
    expect(kind.listChoices).toEqual(['a', 'b', 'c']);
    expect(kind.choices).toBeUndefined();
    expect(JSON.parse(JSON.stringify(kind))).toEqual(kind);
  });

  it('validateInvocation accepts lists of allowed values and rejects others', () => {
    expect(validateInvocation(find, { args: { kind: 'a' } })).toEqual([]);
    expect(validateInvocation(find, { args: { kind: 'a, c' } })).toEqual([]);
    expect(validateInvocation(find, { args: { kind: 'a,x' } })).toEqual(['option kind must be a comma-separated list of: a, b, c']);
    expect(validateInvocation(find, { args: { kind: ' , ' } })).toEqual(['option kind must be a comma-separated list of: a, b, c']);
  });

  it('Commander rejects a bad value with the allowed choices and passes a valid one through', async () => {
    await expect(listProgram().parseAsync(['find', '--kind', 'a,x'], { from: 'user' })).rejects.toThrow(/Allowed choices.*a, b, c/);
    const program = listProgram();
    await program.parseAsync(['find', '--kind', 'b,c'], { from: 'user' });
    expect(program.commands[0].opts().kind).toBe('b,c');
  });
});

describe('SKILLS-1 buildArgv', () => {
  it('rebuilds an argv that Commander parses back to the same values', async () => {
    const program = sampleProgram();
    const read = findCommand(buildCommandTree(program), 'mail.read')!;
    const invocation = {
      positionals: ['--dry-run', '; rm -rf ~', 'a b'],
      args: { folder: '--org-mode', tag: ['x', 'y z'], top: 5, text: true, notify: false, output: 'json' },
    };
    const argv = buildArgv(read, invocation);
    expect(argv).toEqual([
      'mail', 'read', '--text', '--no-notify', '--tag=x', '--tag=y z', '--top=5', '--folder=--org-mode', '--output=json',
      '--', '--dry-run', '; rm -rf ~', 'a b',
    ]);
    const parsed = await parseLeaf(program, read, argv);
    expect(parsed.args).toEqual(['--dry-run', ['; rm -rf ~', 'a b']]);
    expect(parsed.opts).toEqual({ folder: '--org-mode', tag: ['x', 'y z'], top: '5', text: true, notify: false, output: 'json' });
    expect(program.opts().dryRun).toBeUndefined();
  });

  it('omits false booleans without a --no form and the -- separator without positionals', () => {
    const schema = findCommand(buildCommandTree(sampleProgram()), 'schema')!;
    expect(buildArgv(schema, {})).toEqual(['schema']);
    const read = findCommand(buildCommandTree(sampleProgram()), 'mail.read')!;
    expect(buildArgv(read, { positionals: ['1'], args: { folder: 'f', text: false, notify: true } })).toEqual(['mail', 'read', '--folder=f', '--', '1']);
  });

  it('round-trips a real CLI command with a positional, a choice option and a boolean flag', async () => {
    const program = introspectionProgram();
    const tree = buildCommandTree(program);
    const leaf = tree.leaves.find(l =>
      l.arguments.some(a => a.required) &&
      l.options.some(o => o.choices && !o.negate) &&
      l.options.some(o => !o.takesValue && !o.negate))!;
    expect(leaf).toBeDefined();
    const choice = leaf.options.find(o => o.choices && !o.negate)!;
    const flag = leaf.options.find(o => !o.takesValue && !o.negate)!;
    const positionals = leaf.arguments.filter(a => a.required).map((a, i) => `${a.name}-${i} with space`);
    const args: Record<string, unknown> = { [choice.attribute]: choice.choices!.at(-1), [flag.attribute]: true };
    for (const o of leaf.options) if (o.required) args[o.attribute] = `value for ${o.attribute}`;

    const parsed = await parseLeaf(program, leaf, buildArgv(leaf, { positionals, args }));
    expect(parsed.args.slice(0, positionals.length)).toEqual(positionals);
    expect(parsed.opts).toMatchObject(args);
  });
});
