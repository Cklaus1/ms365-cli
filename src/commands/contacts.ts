import { Command } from 'commander';
import { GraphClient } from '../graph-client.js';
import { formatOutput, type OutputFormat } from '../formatter.js';
import { handleErrors } from '../errors.js';
import { outputOption, confirm, done } from '../utils.js';
import { paginationOptions, listRequest, type ListOptions } from '../list.js';
import { searchQuote } from '../graph-helpers.js';

interface ContactListOptions extends ListOptions {
  filter?: string;
  orderby?: string;
  select?: string;
  search?: string;
}

interface UpdateOptions {
  givenName?: string;
  surname?: string;
  email?: string;
  replaceEmail?: boolean;
  phone?: string;
  company?: string;
  output: string;
}

export function registerContactsCommands(program: Command, getClient: () => GraphClient): void {
  const contacts = program.command('contacts').description('Outlook contacts');

  paginationOptions(
    contacts
      .command('list')
      .description('List contacts')
      .option('--filter <filter>', 'OData filter')
      .option('--orderby <field>', 'Order results by field')
      .option('--select <fields>', 'Fields to return')
      .option('--search <query>', 'Search query'),
  )
    .addOption(outputOption())
    .action(handleErrors(async (opts: ContactListOptions) => {
      const queryParams: Record<string, string> = {};
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.orderby) queryParams['$orderby'] = opts.orderby;
      if (opts.select) queryParams['$select'] = opts.select;
      if (opts.search) queryParams['$search'] = searchQuote(opts.search);
      await listRequest(getClient(), '/me/contacts', { queryParams }, opts);
    }));

  contacts
    .command('get <contactId>')
    .description('Get a specific contact')
    .addOption(outputOption())
    .action(handleErrors(async (contactId: string, opts) => {
      const client = getClient();
      const response = await client.request(`/me/contacts/${encodeURIComponent(contactId)}`);
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  contacts
    .command('create')
    .description('Create a contact')
    .requiredOption('--given-name <name>', 'First name')
    .option('--surname <name>', 'Last name')
    .option('--email <email>', 'Email address')
    .option('--phone <phone>', 'Phone number')
    .option('--company <company>', 'Company name')
    .option('--job-title <title>', 'Job title')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const contact: Record<string, unknown> = {
        givenName: opts.givenName,
      };
      if (opts.surname) contact.surname = opts.surname;
      if (opts.email) contact.emailAddresses = [{ address: opts.email }];
      if (opts.phone) contact.mobilePhone = opts.phone;
      if (opts.company) contact.companyName = opts.company;
      if (opts.jobTitle) contact.jobTitle = opts.jobTitle;

      const response = await client.request('/me/contacts', {
        method: 'POST',
        body: JSON.stringify(contact),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  contacts
    .command('update <contactId>')
    .description('Update a contact')
    .option('--given-name <name>', 'First name')
    .option('--surname <name>', 'Last name')
    .option('--email <email>', 'Email address (adds to existing; use --replace-email to overwrite)')
    .option('--replace-email', 'Replace all email addresses instead of adding')
    .option('--phone <phone>', 'Phone number')
    .option('--company <company>', 'Company name')
    .addOption(outputOption())
    .action(handleErrors(async (contactId: string, opts: UpdateOptions) => {
      if (opts.replaceEmail && !opts.email) throw new Error('--replace-email needs --email.');
      if (!opts.givenName && !opts.surname && !opts.email && !opts.phone && !opts.company) {
        throw new Error('Nothing to update: give at least one of --given-name, --surname, --email, --phone, --company.');
      }
      const client = getClient();
      const patch: Record<string, unknown> = {};
      if (opts.givenName) patch.givenName = opts.givenName;
      if (opts.surname) patch.surname = opts.surname;
      if (opts.email) {
        if (opts.replaceEmail) {
          patch.emailAddresses = [{ address: opts.email }];
        } else {
          // Fetch existing emails and append
          const existing = await client.request(`/me/contacts/${encodeURIComponent(contactId)}`, {
            queryParams: { '$select': 'emailAddresses' },
          });
          const data = existing.data as Record<string, unknown>;
          const emails = (data.emailAddresses as Array<Record<string, unknown>>) || [];
          emails.push({ address: opts.email });
          patch.emailAddresses = emails;
        }
      }
      if (opts.phone) patch.mobilePhone = opts.phone;
      if (opts.company) patch.companyName = opts.company;

      const response = await client.request(`/me/contacts/${encodeURIComponent(contactId)}`, {
        method: 'PATCH',
        body: JSON.stringify(patch),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  contacts
    .command('delete <contactId>')
    .description('Delete a contact')
    .action(handleErrors(async (contactId: string) => {
      await confirm(`Delete contact ${contactId.substring(0, 16)}...?`);
      done(await getClient().request(`/me/contacts/${encodeURIComponent(contactId)}`, { method: 'DELETE' }), 'Contact deleted.');
    }));

  // ── Contact folders ─────────────────────────────────────────
  paginationOptions(
    contacts
      .command('folders')
      .description('List contact folders'),
  )
    .addOption(outputOption())
    .action(handleErrors(async (opts: ListOptions) => {
      await listRequest(getClient(), '/me/contactFolders', {}, opts);
    }));

  contacts
    .command('create-folder')
    .description('Create a contact folder')
    .requiredOption('--name <name>', 'Folder display name')
    .option('--parent <folderId>', 'Parent folder ID')
    .addOption(outputOption())
    .action(handleErrors(async (opts) => {
      const client = getClient();
      const endpoint = opts.parent
        ? `/me/contactFolders/${encodeURIComponent(opts.parent)}/childFolders`
        : '/me/contactFolders';
      const response = await client.request(endpoint, {
        method: 'POST',
        body: JSON.stringify({ displayName: opts.name }),
      });
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));


  contacts
    .command('update-folder <folderId>')
    .description('Rename a contact folder (use "[ARCHIVED] Name" to archive)')
    .requiredOption('--name <name>', 'New display name')
    .addOption(outputOption())
    .action(handleErrors(async (folderId: string, opts) => {
      const client = getClient();
      const response = await client.request(
        `/me/contactFolders/${encodeURIComponent(folderId)}`,
        { method: 'PATCH', body: JSON.stringify({ displayName: opts.name }) }
      );
      console.log(formatOutput(response.data, opts.output as OutputFormat));
    }));

  paginationOptions(
    contacts
      .command('folder-contacts <folderId>')
      .description('List contacts in a specific folder')
      .option('--filter <filter>', 'OData filter')
      .option('--orderby <field>', 'Order results by field'),
  )
    .addOption(outputOption())
    .action(handleErrors(async (folderId: string, opts: ContactListOptions) => {
      const queryParams: Record<string, string> = {};
      if (opts.filter) queryParams['$filter'] = opts.filter;
      if (opts.orderby) queryParams['$orderby'] = opts.orderby;
      await listRequest(getClient(), `/me/contactFolders/${encodeURIComponent(folderId)}/contacts`, { queryParams }, opts);
    }));
}
