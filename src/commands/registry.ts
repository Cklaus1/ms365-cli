import { Command } from 'commander';
import type { AuthManager } from '../auth.js';
import type { GraphClient } from '../graph-client.js';
import { registerAuthCommands } from './auth.js';
import { registerMailCommands } from './mail.js';
import { registerCalendarCommands } from './calendar.js';
import { registerDriveCommands } from './drive.js';
import { registerTodoCommands } from './todo.js';
import { registerContactsCommands } from './contacts.js';
import { registerTeamsCommands } from './teams.js';
import { registerSharePointCommands } from './sharepoint.js';
import { registerPlannerCommands } from './planner.js';
import { registerOneNoteCommands } from './onenote.js';
import { registerExcelCommands } from './excel.js';
import { registerSearchCommands } from './search.js';
import { registerUserCommands } from './user.js';
import { registerWorkflowCommands } from './workflow.js';
import { registerGroupsCommands } from './groups.js';
import { registerSchemaCommands } from './schema.js';
import { registerGenerateSkillsCommand } from './generate-skills.js';
import { registerSetupCommand } from './setup.js';

/** The one registration list: the CLI, generate-skills, and remote introspection all use it. */
export function registerAllCommands(program: Command, getAuth: () => AuthManager, getClient: () => GraphClient): void {
  registerAuthCommands(program, getAuth);
  registerMailCommands(program, getClient);
  registerCalendarCommands(program, getClient);
  registerDriveCommands(program, getClient);
  registerTodoCommands(program, getClient);
  registerContactsCommands(program, getClient);
  registerTeamsCommands(program, getClient);
  registerSharePointCommands(program, getClient);
  registerPlannerCommands(program, getClient);
  registerOneNoteCommands(program, getClient);
  registerExcelCommands(program, getClient);
  registerSearchCommands(program, getClient);
  registerUserCommands(program, getClient);
  registerWorkflowCommands(program, getClient);
  registerGroupsCommands(program, getClient);
  registerSchemaCommands(program);
  registerGenerateSkillsCommand(program);
  registerSetupCommand(program, getAuth);
}

/** A program with every command registered, for introspection only: its getters throw if an action ever runs. */
export function introspectionProgram(): Command {
  const program = new Command('ms365').enablePositionalOptions();
  const unavailable = (): never => {
    throw new Error('Commands cannot run from the introspection program');
  };
  registerAllCommands(program, unavailable, unavailable);
  return program;
}
