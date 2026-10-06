import type { Argument, Command, Option } from 'commander';

/**
 * Serializable view of a Commander program: every leaf command with its arguments and options.
 * Shared by `generate-skills` (catalog) and the remote executor (lookup, validation, argv rebuild).
 */

export interface ArgumentNode {
  name: string;
  description: string;
  required: boolean;
  variadic: boolean;
  choices?: string[];
  default?: unknown;
}

export interface OptionNode {
  flags: string;
  description: string;
  long?: string;
  short?: string;
  /** Key in the parsed options object (`--page-limit` → `pageLimit`; `--no-notify` → `notify`). */
  attribute: string;
  /** Option accepts a value (`<v>` required or `[v]` optional). */
  takesValue: boolean;
  /** Value is optional (`[v]`): the flag alone is also accepted. */
  valueOptional: boolean;
  variadic: boolean;
  /** Caller must supply it: mandatory and without a default. */
  required: boolean;
  /** `--no-x` form; sets `attribute` to false. */
  negate: boolean;
  choices?: string[];
  default?: unknown;
}

export interface GroupNode {
  /** Dotted path of canonical names, e.g. `calendar`. */
  path: string;
  names: string[];
  aliases: string[];
  description: string;
}

export interface CommandNode {
  /** Dotted path of canonical names, e.g. `calendar.view`. */
  path: string;
  /** Command names from the root (excluded) to this leaf, e.g. `['calendar', 'view']`. */
  names: string[];
  /** Aliases of the leaf itself (group aliases live on the group). */
  aliases: string[];
  description: string;
  arguments: ArgumentNode[];
  options: OptionNode[];
  /** Ancestor group paths, outermost first (`[]` for a top-level leaf). */
  groups: string[];
}

export interface CommandTree {
  name: string;
  description: string;
  /** Options declared on the root program (global flags). */
  options: OptionNode[];
  groups: GroupNode[];
  leaves: CommandNode[];
}

export interface Invocation {
  positionals?: unknown[];
  /** Option values keyed by attribute name. */
  args?: Record<string, unknown>;
}

export class InvalidInvocationError extends Error {
  constructor(readonly problems: string[]) {
    super(problems.join('; '));
    this.name = 'InvalidInvocationError';
  }
}

function serializable(value: unknown): unknown {
  if (value === undefined || typeof value === 'function') return undefined;
  if (value === null || ['string', 'number', 'boolean'].includes(typeof value)) return value;
  if (Array.isArray(value) && value.every(v => ['string', 'number', 'boolean'].includes(typeof v))) return [...value];
  return undefined;
}

function argumentNode(arg: Argument): ArgumentNode {
  const node: ArgumentNode = {
    name: arg.name(),
    description: arg.description,
    required: arg.required,
    variadic: arg.variadic,
  };
  if (arg.argChoices) node.choices = [...arg.argChoices];
  const def = serializable(arg.defaultValue);
  if (def !== undefined) node.default = def;
  return node;
}

function optionNode(opt: Option): OptionNode {
  const def = serializable(opt.defaultValue);
  const node: OptionNode = {
    flags: opt.flags,
    description: opt.description,
    attribute: opt.attributeName(),
    takesValue: opt.required || opt.optional,
    valueOptional: opt.optional,
    variadic: opt.variadic,
    required: opt.mandatory && opt.defaultValue === undefined,
    negate: opt.negate,
  };
  if (opt.long) node.long = opt.long;
  if (opt.short) node.short = opt.short;
  if (opt.argChoices) node.choices = [...opt.argChoices];
  if (def !== undefined) node.default = def;
  return node;
}

/** Walk a Commander program and describe every leaf command (the implicit `help` command excluded). */
export function buildCommandTree(program: Command): CommandTree {
  const tree: CommandTree = {
    name: program.name(),
    description: program.description(),
    options: program.options.map(optionNode),
    groups: [],
    leaves: [],
  };

  const visit = (cmd: Command, parents: string[], groups: string[]) => {
    const names = [...parents, cmd.name()];
    const path = names.join('.');
    const children = cmd.commands.filter(c => c.name() !== 'help');
    if (children.length > 0) {
      tree.groups.push({ path, names, aliases: [...cmd.aliases()], description: cmd.description() });
      for (const child of children) visit(child, names, [...groups, path]);
      return;
    }
    tree.leaves.push({
      path,
      names,
      aliases: [...cmd.aliases()],
      description: cmd.description(),
      arguments: cmd.registeredArguments.map(argumentNode),
      options: cmd.options.map(optionNode),
      groups,
    });
  };

  for (const cmd of program.commands) {
    if (cmd.name() !== 'help') visit(cmd, [], []);
  }
  return tree;
}

/** Names (canonical first, then aliases) accepted for segment `i` of `leaf`'s path. */
function segmentNames(tree: CommandTree, leaf: CommandNode, i: number): string[] {
  if (i === leaf.names.length - 1) return [leaf.names[i], ...leaf.aliases];
  const groupPath = leaf.names.slice(0, i + 1).join('.');
  const group = tree.groups.find(g => g.path === groupPath);
  return [leaf.names[i], ...(group?.aliases ?? [])];
}

/** Find a leaf by dotted path; aliases are accepted per segment (`cal.view` → `calendar.view`). */
export function findCommand(tree: CommandTree, path: string): CommandNode | undefined {
  const segments = path.split('.');
  return tree.leaves.find(leaf =>
    leaf.names.length === segments.length &&
    segments.every((seg, i) => segmentNames(tree, leaf, i).includes(seg)));
}

const isScalar = (v: unknown): v is string | number =>
  typeof v === 'string' || (typeof v === 'number' && Number.isFinite(v));

/** Check an invocation against a leaf. Returns the problems found (empty when valid). */
export function validateInvocation(node: CommandNode, invocation: Invocation): string[] {
  const problems: string[] = [];
  const positionals = invocation.positionals ?? [];
  const args = invocation.args ?? {};

  const min = node.arguments.filter(a => a.required).length;
  const variadic = node.arguments.at(-1)?.variadic ?? false;
  const max = variadic ? Infinity : node.arguments.length;
  if (positionals.length < min || positionals.length > max) {
    const expected = min === max ? `${min}` : max === Infinity ? `at least ${min}` : `${min}-${max}`;
    problems.push(`${node.path} expects ${expected} positional argument(s), got ${positionals.length}`);
  }
  positionals.forEach((value, i) => {
    const arg = node.arguments[Math.min(i, node.arguments.length - 1)];
    if (!arg) return;
    if (!isScalar(value)) {
      problems.push(`positional <${arg.name}> must be a string or number`);
    } else if (arg.choices && !arg.choices.includes(String(value))) {
      problems.push(`positional <${arg.name}> must be one of: ${arg.choices.join(', ')}`);
    }
  });

  for (const [key, value] of Object.entries(args)) {
    if (value === undefined) continue;
    const opts = node.options.filter(o => o.attribute === key);
    if (opts.length === 0) {
      problems.push(`unknown option: ${key}`);
      continue;
    }
    const valued = opts.find(o => o.takesValue && !o.negate);
    if (!valued) {
      if (typeof value !== 'boolean') problems.push(`option ${key} must be a boolean`);
      continue;
    }
    if (typeof value === 'boolean') {
      if (value && !valued.valueOptional) problems.push(`option ${key} requires a value`);
      continue;
    }
    const values = Array.isArray(value) && valued.variadic ? value : [value];
    if (values.length === 0 || !values.every(isScalar)) {
      problems.push(`option ${key} must be a string or number${valued.variadic ? ' (or an array of them)' : ''}`);
      continue;
    }
    if (valued.choices) {
      const bad = values.filter(v => !valued.choices!.includes(String(v)));
      if (bad.length > 0) problems.push(`option ${key} must be one of: ${valued.choices.join(', ')}`);
    }
  }

  for (const opt of node.options) {
    if (opt.required && args[opt.attribute] === undefined) problems.push(`missing required option: ${opt.attribute}`);
  }
  return problems;
}

function flagToken(opt: OptionNode, value?: string): string {
  const flag = opt.long ?? opt.short!;
  if (value === undefined) return flag;
  return opt.long ? `${flag}=${value}` : `${flag}${value}`;
}

/**
 * Rebuild the argv (after the program name) that runs `node` with `invocation`:
 * `[...names, ...options as --opt=value, '--', ...positionals]`. Each value is one argv element,
 * so nothing is shell-interpreted or re-split. Throws InvalidInvocationError when validation fails.
 */
export function buildArgv(node: CommandNode, invocation: Invocation): string[] {
  const problems = validateInvocation(node, invocation);
  if (problems.length > 0) throw new InvalidInvocationError(problems);

  const args = invocation.args ?? {};
  const argv = [...node.names];
  const seen = new Set<string>();
  for (const opt of node.options) {
    const key = opt.attribute;
    if (seen.has(key) || args[key] === undefined) continue;
    seen.add(key);
    const value = args[key];
    const opts = node.options.filter(o => o.attribute === key);
    const positive = opts.find(o => !o.negate);
    const negative = opts.find(o => o.negate);
    if (typeof value === 'boolean') {
      if (value && positive) argv.push(flagToken(positive));
      else if (!value && negative) argv.push(flagToken(negative));
      continue;
    }
    const values = Array.isArray(value) ? value : [value];
    for (const v of values) argv.push(flagToken(positive!, String(v)));
  }

  const positionals = invocation.positionals ?? [];
  if (positionals.length > 0) argv.push('--', ...positionals.map(String));
  return argv;
}
