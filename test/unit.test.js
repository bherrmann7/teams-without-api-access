'use strict';

const test = require('node:test');
const assert = require('node:assert');
const { parseArgv, intOpt, isTeamsHost, safeUrl, pickChat, deployInfo, COMMANDS, ExitError } = require('../send.js');

test('isTeamsHost accepts the work client hosts only', () => {
  assert.ok(isTeamsHost('teams.microsoft.com'));
  assert.ok(isTeamsHost('teams.cloud.microsoft'));
  assert.ok(!isTeamsHost('teams.live.com'));
  assert.ok(!isTeamsHost('login.microsoftonline.com'));
  assert.ok(!isTeamsHost('example.okta.com'));
  assert.ok(!isTeamsHost('teams.microsoft.com.evil.example'));
  assert.ok(!isTeamsHost('notteams.microsoft.com'));
  assert.ok(!isTeamsHost(''));
});

test('safeUrl drops the query and fragment', () => {
  assert.strictEqual(
    safeUrl('https://login.microsoftonline.com/common/oauth2/authorize?code=SECRET#id_token=SECRET'),
    'https://login.microsoftonline.com/common/oauth2/authorize'
  );
  assert.strictEqual(safeUrl('not a url'), '(unparseable url)');
});

test('parseArgv consumes option values regardless of order', () => {
  const a = parseArgv(['--limit', '5', 'probe', '--quiet', '--filter', 'compose']);
  assert.deepStrictEqual(a.positional, ['probe']);
  assert.deepStrictEqual(a.opts, { '--limit': '5', '--filter': 'compose' });
  assert.ok(a.bools.has('--quiet'));
});

test('parseArgv rejects a value flag with no value', () => {
  assert.throws(() => parseArgv(['probe', '--limit']), ExitError);
});

test('intOpt validates', () => {
  assert.strictEqual(intOpt({}, '--limit', 7), 7);
  assert.strictEqual(intOpt({ '--limit': '3' }, '--limit', 7), 3);
  assert.throws(() => intOpt({ '--limit': 'x' }, '--limit', 7), ExitError);
});

test('every command resolves', () => {
  for (const name of ['login', 'status', 'shot', 'probe', 'chats', 'list']) {
    assert.strictEqual(typeof COMMANDS[name], 'function', name);
  }
});

test('pickChat prefers exact, then an unambiguous substring', () => {
  const names = ['Sam Owner (You)', 'Joe Example', 'Pat, Lee, Joe, +2', 'Pat Sample', 'Joe Example'];
  assert.strictEqual(pickChat(names, 'Joe Example').index, 1);
  assert.strictEqual(pickChat(names, 'joe example').index, 1);
  assert.strictEqual(pickChat(names, 'example').index, 1);
  assert.strictEqual(pickChat(names, 'me').index, 0);
  assert.strictEqual(pickChat(names, 'Sam Owner').index, 0);
  assert.strictEqual(pickChat(names, 'joe').error, 'ambiguous');
  assert.strictEqual(pickChat(names, 'pat').error, 'ambiguous');
  assert.strictEqual(pickChat(names, 'nobody').error, 'none');
  assert.strictEqual(pickChat(names, '  ').error, 'empty');
});

test('deployInfo is exported', () => {
  assert.strictEqual(typeof deployInfo, 'function');
});
