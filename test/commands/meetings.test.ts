import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import { runCommand, type RecordedCall } from '../helpers/cli.js';
import { registerMeetingsCommands, vttToText } from '../../src/commands/meetings.js';

const TZ = 'America/Los_Angeles';
const env = { MS365_CLI_TIMEZONE: TZ };
const query = (call: RecordedCall) => new URLSearchParams(call.path.split('?')[1] ?? '');
const JOIN = "https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc%40thread.v2/0?context=%7b%22Tid%22%3a%22t'1%22%7d";

const VTT = [
  'WEBVTT',
  '',
  '1d3c/12-0',
  '00:00:00.000 --> 00:00:03.500',
  '<v Alice Smith>Hello everyone.</v>',
  '',
  '1d3c/13-0',
  '00:00:03.500 --> 00:00:06.000',
  '<v Alice Smith>Let\'s start.</v>',
  '',
  '00:00:06.000 --> 00:00:08.000',
  '<v Bob>Sounds <b>good</b>.</v>',
  '',
  'NOTE a comment',
  '',
  '00:00:08.000 --> 00:00:09.000',
  'No speaker line',
  '',
].join('\n');

const vttResponse = () => new Response(VTT, { status: 200, headers: { 'content-type': 'text/vtt' } });

let dir: string;
beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ms365-meet-'));
  vi.useFakeTimers({ toFake: ['Date'] });
  vi.setSystemTime(new Date('2026-10-14T17:00:00Z')); // Wed 10:00 PDT
});
afterEach(() => {
  vi.useRealTimers();
  fs.rmSync(dir, { recursive: true, force: true });
});

describe('vttToText (PM-4)', () => {
  it('strips timestamps/tags and merges consecutive same-speaker cues', () => {
    expect(vttToText(VTT)).toBe("Alice Smith: Hello everyone. Let's start.\nBob: Sounds good.\nNo speaker line\n");
  });
  it('handles CRLF input', () => {
    expect(vttToText(VTT.replace(/\n/g, '\r\n'))).toBe("Alice Smith: Hello everyone. Let's start.\nBob: Sounds good.\nNo speaker line\n");
  });
});

describe('meetings list (PM-4)', () => {
  it('queries calendarView for online meetings in the user zone', async () => {
    const r = await runCommand(registerMeetingsCommands,
      ['meet', 'list', '--start', '2026-10-15T00:00', '--end', '2026-10-16T00:00'], {
        env, routes: [['GET', '/me/calendarView', () => ({ value: [{ subject: 'Standup', onlineMeeting: { joinUrl: JOIN } }] })]],
      });
    expect(r.exitCode).toBe(0);
    const call = r.calls[0];
    expect(call.path.startsWith('/me/calendarView?')).toBe(true);
    const q = query(call);
    expect(q.get('$filter')).toBe('isOnlineMeeting eq true');
    expect(q.get('startDateTime')).toBe('2026-10-15T07:00:00.000Z');
    expect(q.get('endDateTime')).toBe('2026-10-16T07:00:00.000Z');
    expect(q.get('$select')).toContain('onlineMeeting');
    expect(call.headers.Prefer).toBe(`outlook.timezone="${TZ}"`);
    expect(JSON.parse(r.stdout).value[0].onlineMeeting.joinUrl).toBe(JOIN);
  });

  it('defaults to the next 7 days', async () => {
    const r = await runCommand(registerMeetingsCommands, ['meetings', 'list'], { env });
    const q = query(r.calls[0]);
    expect(q.get('startDateTime')).toBe('2026-10-14T17:00:00.000Z');
    expect(q.get('endDateTime')).toBe('2026-10-21T17:00:00.000Z');
  });
});

describe('meetings get (PM-4)', () => {
  it('resolves a join URL via $filter JoinWebUrl with odata quoting', async () => {
    const r = await runCommand(registerMeetingsCommands, ['meet', 'get', JOIN], {
      routes: [['GET', '/me/onlineMeetings', () => ({ value: [{ id: 'M1', subject: 'Standup' }] })]],
    });
    expect(r.exitCode).toBe(0);
    expect(new URL(r.calls[0].url).searchParams.get('$filter')).toBe(`JoinWebUrl eq '${JOIN.replace(/'/g, "''")}'`);
    expect(JSON.parse(r.stdout)).toEqual({ id: 'M1', subject: 'Standup' });
  });

  it('fails NOT_FOUND when no meeting matches the URL', async () => {
    const r = await runCommand(registerMeetingsCommands, ['meet', 'get', JOIN], {
      routes: [['GET', '/me/onlineMeetings', () => ({ value: [] })]],
    });
    expect(r.exitCode).toBe(5);
  });

  it('gets by id', async () => {
    const r = await runCommand(registerMeetingsCommands, ['meet', 'get', 'M/1'], {
      routes: [['GET', '/me/onlineMeetings/M', () => ({ id: 'M/1' })]],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[0].url).toContain('/me/onlineMeetings/M%2F1');
  });
});

describe('meetings transcripts / transcript (PM-4, G-7)', () => {
  it('lists transcripts; accepts a join URL as the meeting', async () => {
    const r = await runCommand(registerMeetingsCommands, ['meet', 'transcripts', JOIN], {
      routes: [
        ['GET', /^\/me\/onlineMeetings\?/, () => ({ value: [{ id: 'M1' }] })],
        ['GET', '/me/onlineMeetings/M1/transcripts', () => ({ value: [{ id: 'T1' }] })],
      ],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[1].path).toBe('/me/onlineMeetings/M1/transcripts');
    expect(JSON.parse(r.stdout).value[0].id).toBe('T1');
  });

  it('defaults to the latest transcript and prints VTT on stdout', async () => {
    const r = await runCommand(registerMeetingsCommands, ['meet', 'transcript', 'M1'], {
      routes: [
        ['GET', /\/transcripts$/, () => ({ value: [
          { id: 'T-old', createdDateTime: '2026-10-01T10:00:00Z' },
          { id: 'T-new', createdDateTime: '2026-10-02T10:00:00Z' },
          { id: 'T-mid', createdDateTime: '2026-10-01T12:00:00Z' },
        ] })],
        ['GET', '/me/onlineMeetings/M1/transcripts/T-new/content', vttResponse],
      ],
    });
    expect(r.exitCode).toBe(0);
    expect(r.calls[1].path).toBe('/me/onlineMeetings/M1/transcripts/T-new/content?$format=text/vtt');
    expect(r.stdout).toBe(VTT);
  });

  it('--format txt converts and --out writes a file (refusing overwrite without --force)', async () => {
    const out = path.join(dir, 'm.txt');
    const routes: [string, RegExp | string, () => Response][] = [['GET', '/me/onlineMeetings/M1/transcripts/T1/content', vttResponse]];
    const r = await runCommand(registerMeetingsCommands, ['meet', 'transcript', 'M1', 'T1', '--format', 'txt', '--out', out], { routes });
    expect(r.exitCode).toBe(0);
    expect(r.stdout).toBe('');
    expect(fs.readFileSync(out, 'utf8')).toBe("Alice Smith: Hello everyone. Let's start.\nBob: Sounds good.\nNo speaker line\n");
    expect(r.stderr).toContain(out);

    const again = await runCommand(registerMeetingsCommands, ['meet', 'transcript', 'M1', 'T1', '--out', out], { routes });
    expect(again.exitCode).toBe(1);
    expect(again.calls).toHaveLength(0);

    const forced = await runCommand(registerMeetingsCommands, ['meet', 'transcript', 'M1', 'T1', '--out', out, '--force'], { routes });
    expect(forced.exitCode).toBe(0);
    expect(fs.readFileSync(out, 'utf8')).toBe(VTT);
  });

  it('fails NOT_FOUND when the meeting has no transcripts', async () => {
    const r = await runCommand(registerMeetingsCommands, ['meet', 'transcript', 'M1'], {
      routes: [['GET', /\/transcripts$/, () => ({ value: [] })]],
    });
    expect(r.exitCode).toBe(5);
  });

  it('rejects an unknown --format', async () => {
    const r = await runCommand(registerMeetingsCommands, ['meet', 'transcript', 'M1', 'T1', '--format', 'docx']);
    expect(r.exitCode).toBe(2);
    expect(r.calls).toHaveLength(0);
  });

  it('works under --dry-run', async () => {
    const r = await runCommand(registerMeetingsCommands, ['meet', 'transcript', 'M1', '--format', 'txt'], { dryRun: true });
    expect(r.exitCode).toBe(0);
    expect(r.stdout).toContain('"dryRun": true');
  });
});

describe('meetings recordings / recording (PM-4, G-7)', () => {
  const MP4 = Buffer.from([0x00, 0x00, 0x00, 0x18, 0x66, 0x74, 0x79, 0x70, 0xff, 0xfe, 0x80]);
  const mp4 = () => new Response(MP4, { status: 200, headers: { 'content-type': 'video/mp4' } });

  it('lists recordings', async () => {
    const r = await runCommand(registerMeetingsCommands, ['meet', 'recordings', 'M1']);
    expect(r.calls[0].path).toBe('/me/onlineMeetings/M1/recordings');
  });

  it('streams the latest recording byte-exact to --out', async () => {
    const out = path.join(dir, 'rec.mp4');
    const r = await runCommand(registerMeetingsCommands, ['meet', 'recording', 'M1', '--out', out], {
      routes: [
        ['GET', /\/recordings$/, () => ({ value: [
          { id: 'R1', createdDateTime: '2026-10-02T10:00:00Z' },
          { id: 'R0', createdDateTime: '2026-10-01T10:00:00Z' },
        ] })],
        ['GET', '/me/onlineMeetings/M1/recordings/R1/content', mp4],
      ],
    });
    expect(r.exitCode).toBe(0);
    expect(fs.readFileSync(out).equals(MP4)).toBe(true);
    expect(r.stdout).toBe('');
  });

  it('requires --out', async () => {
    const r = await runCommand(registerMeetingsCommands, ['meet', 'recording', 'M1', 'R1']);
    expect(r.exitCode).toBe(2);
  });
});

describe('meetings all-transcripts (G-7, beta)', () => {
  it('fails with a usage error unless the client is on beta', async () => {
    const r = await runCommand(registerMeetingsCommands, ['meet', 'all-transcripts']);
    expect(r.exitCode).toBe(2);
    expect(r.stderr).toContain('--api-version beta');
    expect(r.calls).toHaveLength(0);
  });

  it('calls getAllTranscripts for the signed-in organizer on beta', async () => {
    const r = await runCommand(registerMeetingsCommands,
      ['meet', 'all-transcripts', '--start', '2026-10-01T00:00', '--end', '2026-10-08T00:00'], {
        env, apiVersion: 'beta',
        routes: [
          ['GET', /^\/me\?/, () => ({ id: 'U1' })],
          ['GET', '/users/U1/onlineMeetings/getAllTranscripts', () => ({ value: [{ id: 'T1' }] })],
        ],
      });
    expect(r.exitCode).toBe(0);
    expect(r.calls[1].url.startsWith('https://graph.microsoft.com/beta/')).toBe(true);
    expect(r.calls[1].path).toBe(
      "/users/U1/onlineMeetings/getAllTranscripts(meetingOrganizerUserId='U1',startDateTime=2026-10-01T07:00:00.000Z,endDateTime=2026-10-08T07:00:00.000Z)");
    expect(JSON.parse(r.stdout).value[0].id).toBe('T1');
  });
});
