import { describe, it, expect, afterEach } from 'vitest';
import { detectShape, defaultColumns } from '../src/columns.js';
import { formatOutput, setFields } from '../src/formatter.js';

const message = {
  '@odata.etag': 'W/"CQAAABYAAAA"',
  id: 'AAMkAGI2TG93AAA=',
  createdDateTime: '2026-10-01T09:00:00Z',
  lastModifiedDateTime: '2026-10-01T09:00:02Z',
  changeKey: 'CQAAABYAAAA',
  categories: [],
  receivedDateTime: '2026-10-01T09:00:00Z',
  sentDateTime: '2026-10-01T08:59:58Z',
  hasAttachments: true,
  internetMessageId: '<abc@contoso.com>',
  subject: 'Q3 numbers',
  bodyPreview: 'Hi team, attached are the Q3 numbers for review before Friday.',
  importance: 'normal',
  parentFolderId: 'AQMkADAwATM0MDAAMS1i',
  conversationId: 'AAQkAGI2TG93',
  isRead: false,
  isDraft: false,
  webLink: 'https://outlook.office365.com/owa/?ItemID=AAMkAGI2TG93AAA%3D&exvsurl=1&viewmodel=ReadMessageItem',
  inferenceClassification: 'focused',
  body: { contentType: 'html', content: '<html><body>Hi team</body></html>' },
  sender: { emailAddress: { name: 'Ada Lovelace', address: 'ada@contoso.com' } },
  from: { emailAddress: { name: 'Ada Lovelace', address: 'ada@contoso.com' } },
  toRecipients: [{ emailAddress: { name: 'Team', address: 'team@contoso.com' } }],
  ccRecipients: [],
  flag: { flagStatus: 'notFlagged' },
};

const event = {
  '@odata.etag': 'W/"ZlnW4RIAV06KYYwlrfNZvQAAKGWwbw=="',
  id: 'AAMkAGIAAAoZDOFAAA=',
  createdDateTime: '2026-09-30T12:00:00Z',
  subject: 'Planning sync',
  bodyPreview: 'Agenda to follow',
  isAllDay: false,
  isCancelled: false,
  showAs: 'busy',
  type: 'singleInstance',
  webLink: 'https://outlook.office365.com/owa/?itemid=AAMkAGIAAAoZDOFAAA%3D&exvsurl=1&path=/calendar/item',
  onlineMeeting: { joinUrl: 'https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc' },
  start: { dateTime: '2026-10-15T09:00:00.0000000', timeZone: 'UTC' },
  end: { dateTime: '2026-10-15T09:30:00.0000000', timeZone: 'UTC' },
  location: { displayName: 'Room 4', locationType: 'default' },
  attendees: [{ type: 'required', emailAddress: { name: 'Bob', address: 'bob@contoso.com' } }],
  organizer: { emailAddress: { name: 'Ada Lovelace', address: 'ada@contoso.com' } },
};

const driveItem = {
  '@microsoft.graph.downloadUrl': 'https://contoso-my.sharepoint.com/download.aspx?UniqueId=123',
  createdDateTime: '2026-09-01T10:00:00Z',
  eTag: '"{1234},1"',
  id: '01BYE5RZ6QN3ZWBTUFOFD3GSPGOHDJD36K',
  lastModifiedDateTime: '2026-09-02T10:00:00Z',
  name: 'report.docx',
  webUrl: 'https://contoso-my.sharepoint.com/personal/ada/Documents/report.docx',
  cTag: '"c:{1234},1"',
  size: 20480,
  createdBy: { user: { displayName: 'Ada Lovelace', email: 'ada@contoso.com' } },
  parentReference: { driveType: 'business', driveId: 'b!abc', id: '01BYE5RZ56Y2GOVW7725BZO354PWSELRRZ', path: '/drive/root:' },
  file: { mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', hashes: { quickXorHash: 'abc=' } },
};

const todoTask = {
  '@odata.etag': 'W/"xzyPKP0BiUGgld+lMKXwbQAAgdhkVw=="',
  importance: 'high',
  isReminderOn: false,
  status: 'notStarted',
  title: 'File expenses',
  createdDateTime: '2026-10-01T08:00:00Z',
  lastModifiedDateTime: '2026-10-01T08:00:00Z',
  hasAttachments: false,
  categories: [],
  id: 'AAkALgAAAAAAHYQDEapmEc2byACqAC-EWg0',
  body: { content: '', contentType: 'text' },
  dueDateTime: { dateTime: '2026-10-20T00:00:00.0000000', timeZone: 'UTC' },
};

const plannerTask = {
  '@odata.etag': 'W/"JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBAWCc="',
  planId: 'xqQg5FS2LkCp935s-FIFm2QAFkHM',
  bucketId: 'gcrYAaAkgU2EQUvpkNNXLGQAGTtu',
  title: 'Draft launch post',
  orderHint: '8586',
  assigneePriority: '',
  percentComplete: 50,
  priority: 5,
  startDateTime: null,
  createdDateTime: '2026-09-20T10:00:00Z',
  dueDateTime: '2026-10-25T00:00:00Z',
  hasDescription: false,
  completedDateTime: null,
  referenceCount: 0,
  checklistItemCount: 0,
  activeChecklistItemCount: 0,
  conversationThreadId: null,
  id: '01gzSlKkIUSUl0DsqhxF6ZYAJzHm',
  createdBy: { user: { displayName: null, id: '6463a5ce-2119-4198-9f2a-628761df4a62' } },
  appliedCategories: {},
  assignments: {},
};

const user = {
  businessPhones: ['+1 425 555 0109'],
  displayName: 'Adele Vance',
  givenName: 'Adele',
  jobTitle: 'Retail Manager',
  mail: 'AdeleV@contoso.com',
  mobilePhone: null,
  officeLocation: '18/2111',
  preferredLanguage: 'en-US',
  surname: 'Vance',
  userPrincipalName: 'AdeleV@contoso.com',
  id: '87d349ed-44d7-43e1-9a83-5f2406dee5bd',
};

const group = {
  id: '02bd9fd6-8f93-4758-87c3-1fb73740a315',
  createdDateTime: '2017-07-31T18:56:16Z',
  description: 'Welcome to the HR Taskforce team.',
  displayName: 'HR Taskforce',
  groupTypes: ['Unified'],
  mail: 'HRTaskforce@contoso.com',
  mailEnabled: true,
  mailNickname: 'HRTaskforce',
  securityEnabled: false,
  visibility: 'Private',
};

const chatMessage = {
  id: '1616990032035',
  replyToId: null,
  etag: '1616990032035',
  messageType: 'message',
  createdDateTime: '2026-10-01T03:53:52.035Z',
  lastModifiedDateTime: '2026-10-01T03:53:52.035Z',
  importance: 'normal',
  chatId: '19:2da4c29f6d7041eca70b638b43d45437@thread.v2',
  from: { application: null, device: null, user: { id: '8ea0e38b', displayName: 'Robin Kline', userIdentityType: 'aadUser' } },
  body: { contentType: 'text', content: 'Hello world' },
  attachments: [],
  mentions: [],
  reactions: [],
};

const chat = {
  id: '19:meeting_abc@thread.v2',
  topic: 'Design review',
  createdDateTime: '2026-09-01T10:00:00Z',
  lastUpdatedDateTime: '2026-10-01T10:00:00Z',
  chatType: 'group',
  webUrl: 'https://teams.microsoft.com/l/chat/19%3Ameeting_abc%40thread.v2/0',
  tenantId: 'b33cbe9f-8ebe-4f2a-912b-7e2a427f477f',
};

const team = {
  id: '172b0cce-e65d-44ce-9a49-91d9f2e8493a',
  displayName: 'Contoso Team',
  description: 'This is a Contoso team',
  isArchived: false,
  tenantId: 'b33cbe9f-8ebe-4f2a-912b-7e2a427f477f',
};

const channel = {
  id: '19:561fbdbbfca848a484f0a6f00ce9dbbd@thread.tacv2',
  createdDateTime: '2020-05-27T19:22:25.692Z',
  displayName: 'General',
  description: 'AutoTestTeam_20200527192221',
  membershipType: 'standard',
  email: 'team@contoso.com',
  webUrl: 'https://teams.microsoft.com/l/channel/19%3a561fbdbbfca848a484f0a6f00ce9dbbd%40thread.tacv2/General',
};

const contact = {
  id: 'AAMkAGI2THk0AAA=',
  createdDateTime: '2026-01-01T00:00:00Z',
  displayName: 'Pavel Bansky',
  givenName: 'Pavel',
  surname: 'Bansky',
  companyName: 'Fabrikam',
  mobilePhone: '+1 206 555 0100',
  businessPhones: [],
  emailAddresses: [{ address: 'pavelb@fabrikam.onmicrosoft.com', name: 'Pavel Bansky' }],
};

const plan = {
  '@odata.etag': 'W/"JzEtUGxhbiAgQEBAQEBAQEBAQEBAQEBAWCc="',
  createdDateTime: '2026-08-01T10:00:00Z',
  owner: 'ebf3b108-5234-4e22-b93d-656d7dae5874',
  title: 'Launch',
  id: 'xqQg5FS2LkCp935s-FIFm2QAFkHM',
  createdBy: { user: { id: '95e27074-6c4a-447a-aa24-9d718a0b86fa' } },
  container: { containerId: 'ebf3b108-5234-4e22-b93d-656d7dae5874', type: 'group', url: 'https://graph.microsoft.com/v1.0/groups/ebf3b108' },
};

const notebook = {
  id: '1-d6ab2e8c-7b2f-4b3e-a1f1-a6d4c3e58f1a',
  self: 'https://graph.microsoft.com/v1.0/users/ada/onenote/notebooks/1-d6ab2e8c',
  createdDateTime: '2026-01-01T00:00:00Z',
  displayName: 'Work Notes',
  lastModifiedDateTime: '2026-10-01T00:00:00Z',
  isDefault: true,
  userRole: 'Owner',
  isShared: false,
  sectionsUrl: 'https://graph.microsoft.com/v1.0/users/ada/onenote/notebooks/1-d6ab2e8c/sections',
  sectionGroupsUrl: 'https://graph.microsoft.com/v1.0/users/ada/onenote/notebooks/1-d6ab2e8c/sectionGroups',
  links: { oneNoteClientUrl: { href: 'onenote:https://contoso/Work%20Notes' }, oneNoteWebUrl: { href: 'https://contoso/Work%20Notes' } },
};

afterEach(() => {
  setFields(null);
  delete process.env.MS365_CLI_TABLE_ALL_COLUMNS;
});

describe('detectShape', () => {
  it.each([
    ['message', message],
    ['event', event],
    ['driveItem', driveItem],
    ['todoTask', todoTask],
    ['plannerTask', plannerTask],
    ['user', user],
    ['group', group],
    ['chatMessage', chatMessage],
    ['chat', chat],
    ['team', team],
    ['channel', channel],
    ['contact', contact],
    ['plannerPlan', plan],
    ['notebook', notebook],
  ])('detects %s', (name, item) => {
    expect(detectShape(item)?.name).toBe(name);
  });

  it('prefers @odata.type over key heuristics', () => {
    expect(detectShape({ '@odata.type': '#microsoft.graph.user', displayName: 'X', groupTypes: [] })?.name).toBe('user');
    expect(detectShape({ '@odata.type': '#microsoft.graph.eventMessageRequest', subject: 'Invite', receivedDateTime: 'x' })?.name).toBe('message');
  });

  it('returns undefined for unknown shapes and non-objects', () => {
    expect(detectShape({ foo: 1, bar: 2 })).toBeUndefined();
    expect(detectShape('text')).toBeUndefined();
    expect(detectShape(null)).toBeUndefined();
    expect(detectShape([1, 2])).toBeUndefined();
  });
});

describe('defaultColumns', () => {
  it('uses the first object item', () => {
    expect(defaultColumns([message, message])).toEqual(['receivedDateTime', 'from.emailAddress.address', 'subject', 'isRead', 'hasAttachments']);
    expect(defaultColumns([event])).toEqual(['subject', 'start.dateTime', 'end.dateTime', 'location.displayName', 'organizer.emailAddress.name']);
    expect(defaultColumns([driveItem])).toEqual(['name', 'size', 'lastModifiedDateTime', 'webUrl']);
    expect(defaultColumns([todoTask])).toEqual(['title', 'status', 'dueDateTime.dateTime', 'importance']);
    expect(defaultColumns([user])).toEqual(['displayName', 'mail', 'jobTitle']);
    expect(defaultColumns([group])).toEqual(['displayName', 'mail', 'groupTypes']);
  });

  it('is null for unknown or empty input', () => {
    expect(defaultColumns([])).toBeNull();
    expect(defaultColumns([{ a: 1 }])).toBeNull();
  });
});

const header = (table: string) => table.split('\n')[0].trim().split(/\s{2,}/);

describe('table output uses default columns', () => {
  it('projects messages to the default columns, in order', () => {
    const out = formatOutput({ value: [message, { ...message, subject: 'Re: Q3', isRead: true }] }, 'table');
    expect(header(out)).toEqual(['receivedDateTime', 'from.emailAddress.address', 'subject', 'isRead', 'hasAttachments']);
    expect(out).toContain('ada@contoso.com');
    expect(out).not.toContain('bodyPreview');
    expect(out.split('\n')).toHaveLength(4);
  });

  it('projects events, drive items, groups (array column) and contacts', () => {
    expect(header(formatOutput({ value: [event] }, 'table'))).toEqual(['subject', 'start.dateTime', 'end.dateTime', 'location.displayName', 'organizer.emailAddress.name']);
    expect(header(formatOutput([driveItem], 'table'))).toEqual(['name', 'size', 'lastModifiedDateTime', 'webUrl']);
    const groups = formatOutput({ value: [group] }, 'table');
    expect(header(groups)).toEqual(['displayName', 'mail', 'groupTypes']);
    expect(groups).toContain('Unified');
    expect(formatOutput({ value: [contact] }, 'table')).toContain('pavelb@fabrikam.onmicrosoft.com');
  });

  it('keeps a column present in a later item, and drops one no item has', () => {
    const noLocation = { ...event, location: undefined };
    const out = formatOutput({ value: [noLocation, event] }, 'table');
    expect(header(out)).toContain('location.displayName');
    const none = formatOutput({ value: [noLocation] }, 'table');
    expect(header(none)).not.toContain('location.displayName');
  });

  it('applies to a single object too', () => {
    expect(header(formatOutput(user, 'table'))).toEqual(['displayName', 'mail', 'jobTitle']);
  });

  it('--fields wins over the defaults', () => {
    setFields(['id', 'subject']);
    expect(header(formatOutput({ value: [message] }, 'table'))).toEqual(['id', 'subject']);
  });

  it('MS365_CLI_TABLE_ALL_COLUMNS=1 restores every column', () => {
    process.env.MS365_CLI_TABLE_ALL_COLUMNS = '1';
    const cols = header(formatOutput({ value: [message] }, 'table'));
    expect(cols).toContain('bodyPreview');
    expect(cols).toContain('webLink');
  });

  it('unknown shapes keep every column', () => {
    expect(header(formatOutput({ value: [{ a: '1', b: '2', c: '3' }] }, 'table'))).toEqual(['a', 'b', 'c']);
  });

  it('json, ndjson, csv, text and yaml are unaffected', () => {
    const data = { value: [message] };
    expect(JSON.parse(formatOutput(data, 'json'))).toEqual(data);
    expect(JSON.parse(formatOutput(data, 'ndjson'))).toEqual(message);
    expect(formatOutput(data, 'csv').split('\n')[0]).toContain('bodyPreview');
    expect(formatOutput(data, 'text')).toContain('bodyPreview');
    expect(formatOutput(data, 'yaml')).toContain('bodyPreview');
  });
});

describe('chat and planner plan columns (X-13, X-14)', () => {
  const members = [
    { '@odata.type': '#microsoft.graph.aadUserConversationMember', id: 'm1', displayName: 'Ada Lovelace' },
    { '@odata.type': '#microsoft.graph.aadUserConversationMember', id: 'm2', displayName: 'Bob Builder' },
  ];

  it('a chat table shows the members (from $expand=members) so 1:1 chats can be told apart', () => {
    const oneOnOne = { ...chat, topic: null, chatType: 'oneOnOne', members };
    const out = formatOutput({ value: [oneOnOne] }, 'table');
    expect(header(out)).toEqual(['topic', 'members', 'chatType', 'lastUpdatedDateTime']);
    expect(out).toContain('Ada Lovelace, Bob Builder');
  });

  it('a chat without members keeps the other columns', () => {
    expect(header(formatOutput({ value: [chat] }, 'table'))).toEqual(['topic', 'chatType', 'lastUpdatedDateTime']);
  });

  it('planner plans show container.type, not the deprecated owner GUID', () => {
    expect(defaultColumns([plan])).toEqual(['title', 'container.type', 'createdDateTime']);
    const out = formatOutput({ value: [plan] }, 'table');
    expect(header(out)).toEqual(['title', 'container.type', 'createdDateTime']);
    expect(out).not.toContain('ebf3b108');
    const { container: _c, ...ownerOnly } = plan;
    expect(detectShape(ownerOnly)?.name).toBe('plannerPlan');
  });

  it('no default shape has an id column (F9)', async () => {
    const { SHAPES } = await import('../src/columns.js');
    expect(SHAPES.filter(s => s.columns.some(c => c === 'id' || c.endsWith('.id')))).toEqual([]);
  });
});
