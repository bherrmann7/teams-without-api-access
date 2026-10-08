#!/usr/bin/env node
// send.js — read and post Microsoft Teams chats as yourself by driving the Teams
// web client with a persistent Playwright profile. No Graph app registration.
//
// The session lives in a Chromium profile directory (~/.teams-send/profile-*),
// so `login` is an interactive step and every later command is headless. Sign-in
// is federated (Okta -> Entra ID -> Teams), so how long "later" lasts is set by
// the identity provider's session policy, not by this tool.
//
// Teams' DOM is React with hashed class names. The durable hooks are the
// data-tid attributes it ships for its own tests plus a few structured ids, so
// every selector below is one that `probe` showed on the live client. When one
// misses we screenshot and fail loudly rather than typing into whatever has focus.

'use strict';

const path = require('path');
const os = require('os');
const fs = require('fs');

const HOME = os.homedir();
const STATE_DIR = process.env.TEAMS_SEND_HOME || path.join(HOME, '.teams-send');
const SHOT_DIR = path.join(STATE_DIR, 'shots');
const CONFIG_FILE = path.join(STATE_DIR, 'config.json');
// The released edition unattended callers run — see ./install. Deliberately NOT
// this file.
const DEPLOY_DIR = path.join(STATE_DIR, 'deploy');
// Session cookies carried from one launch to the next — see the browser section.
const SESSION_COOKIES_FILE = path.join(STATE_DIR, 'session-cookies.json');

function loadConfig() {
  try {
    return JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf8'));
  } catch (_) {
    return {};
  }
}

const CONFIG = loadConfig();
// Where headless commands land. Teams redirects through the identity provider
// and back on its own when the profile still holds a live session.
const CLIENT_URL = process.env.TEAMS_SEND_URL || CONFIG.url || 'https://teams.cloud.microsoft/';
// Where `login` starts. Point this at the SSO dashboard (config.json
// "loginUrl") to sign in the way you do by hand — dashboard, then the Teams
// tile — instead of answering Microsoft's "which account?" prompt first.
const LOGIN_URL = process.env.TEAMS_SEND_LOGIN_URL || CONFIG.loginUrl || CLIENT_URL;
// Which browser build drives the profile. 'chromium' is Playwright's bundled
// full Chromium; 'chrome' or 'msedge' use the installed browser, which matters
// if the tenant only admits a browser that can present a device identity.
const CHANNEL = process.env.TEAMS_SEND_CHANNEL || CONFIG.channel || 'chromium';
// One profile per channel. A cookie store written by one browser build cannot be
// decrypted by another (different OSCrypt/Keychain identity), and Chromium's
// answer to a store it cannot read is to discard it — which would silently
// destroy the login session the moment the channel changed.
const PROFILE_DIR = path.join(STATE_DIR, `profile-${CHANNEL.replace(/[^a-z0-9-]/gi, '_')}`);

const NAV_TIMEOUT = 60000;
const UI_TIMEOUT = 20000;
const LOGIN_TIMEOUT = 600000;

// ---------------------------------------------------------------- selectors

const SEL = {
  // Signed-in shell: the left app bar only exists once the client has booted.
  shell: '[data-tid="app-bar-wrapper"]',
  // The left rail is one ARIA tree: level-1 items are sections (Favorites,
  // Teams and channels, Chats), level-2 items are the conversations.
  rail: '[data-tid="simple-collab-dnd-rail"]',
  // A conversation row's name. Rows carry no data-tid of their own; this id
  // prefix is the only thing separating a chat from "Mentions" or a team.
  chatTitle: '[id^="title-chat-list-item_"]',
  // Conversation header — used to PROVE which chat is open before typing.
  header: '[data-tid="chat-title"]',
  // One rendered message. data-mid is the message id and keys the sibling
  // elements: author-<mid>, timestamp-<mid>, content-<mid>.
  msg: '[data-tid="chat-pane-message"]',
  // The composer is CKEditor. role=textbox keeps this off the search box,
  // which is an <input role="combobox">.
  composer: '[data-tid="ckeditor"][role="textbox"]',
  // Clicking Send beats a key chord: it cannot be hijacked by whatever holds
  // focus and does not depend on the Enter-vs-Cmd+Enter preference.
  sendButton: '[data-tid="sendMessageCommands-send"]',
};

// ---------------------------------------------------------------- utilities

// Thrown rather than exiting on the spot. process.exit() inside a command body
// skips every `finally`, so the browser context would be abandoned, not closed.
class ExitError extends Error {
  constructor(msg, code) {
    super(msg);
    this.name = 'ExitError';
    this.exitCode = code;
  }
}

function die(msg, code = 1) {
  throw new ExitError(msg, code);
}

let QUIET = false;

function log(msg) {
  if (!QUIET) console.error(`teams-send: ${msg}`);
}

function ensureDirs() {
  for (const d of [STATE_DIR, PROFILE_DIR, SHOT_DIR]) {
    fs.mkdirSync(d, { recursive: true, mode: 0o700 });
  }
  // The profile holds a live corporate session. Keep it owner-only.
  try { fs.chmodSync(STATE_DIR, 0o700); } catch (_) {}
}

async function shot(page, label) {
  try {
    const p = path.join(SHOT_DIR, `${label}-${Date.now()}.png`);
    await page.screenshot({ path: p, fullPage: false });
    return p;
  } catch (_) {
    return null;
  }
}

async function failWithShot(page, msg) {
  const p = await shot(page, 'error');
  die(`${msg}${p ? `\n  screenshot: ${p}` : ''}`);
}

function norm(s) {
  return (s || '').replace(/\s+/g, ' ').trim().toLowerCase();
}

// Choose a chat row by name. Exact match first — where "me" and your own bare
// name both mean the "<name> (You)" self chat — then a substring match only if
// it is unambiguous. Returns { index } or { error, candidates }.
function pickChat(names, query) {
  const q = norm(query);
  if (!q) return { error: 'empty', candidates: [] };
  const n = names.map(norm);
  const hits = (pred) => n.map((x, i) => (pred(x) ? i : -1)).filter((i) => i >= 0);
  const self = (x) => /\(you\)$/.test(x);
  let m = hits((x) => x === q || (self(x) && (q === 'me' || x === q + ' (you)')));
  if (m.length === 0) m = hits((x) => x.includes(q));
  if (m.length === 0) return { error: 'none', candidates: [] };
  // The same chat can be listed twice (Favorites and Chats); that is one chat.
  const distinct = [...new Set(m.map((i) => n[i]))];
  if (distinct.length > 1) return { error: 'ambiguous', candidates: [...new Set(m.map((i) => names[i]))] };
  return { index: m[0] };
}

// What is deployed, and whether the working tree has moved on. The whole point
// of ./install is that unattended callers do not run the file you are editing,
// so `status` has to say which edition is live or you are back to guessing from
// mtimes.
function deployInfo() {
  let meta = null;
  try {
    meta = JSON.parse(fs.readFileSync(path.join(DEPLOY_DIR, 'VERSION.json'), 'utf8'));
  } catch (_) {
    return null;
  }
  let sourceHash = null;
  try {
    const crypto = require('crypto');
    const buf = fs.readFileSync(meta.source);
    sourceHash = crypto.createHash('sha256').update(buf).digest('hex').slice(0, 12);
  } catch (_) {}
  // An unreadable source is NOT a match — it means the project moved or was
  // deleted, and claiming "matches deploy" there would assert something this
  // function just failed to check.
  return {
    ...meta,
    sourceHash,
    stale: !!(sourceHash && sourceHash !== meta.hash),
    sourceMissing: !sourceHash,
  };
}

// The hosts the work/school Teams web client is served from. teams.live.com is
// the consumer product and is deliberately not here.
function isTeamsHost(hostname) {
  return /(^|\.)teams\.(microsoft\.com|cloud\.microsoft)$/i.test(hostname || '');
}

// Redirect URLs in a federated sign-in carry codes and tokens in the query and
// fragment. Anything printed or logged goes through this first.
function safeUrl(raw) {
  try {
    const u = new URL(raw);
    return u.origin + u.pathname;
  } catch (_) {
    return '(unparseable url)';
  }
}

// ---------------------------------------------------------------- browser

// Chromium reports "HeadlessChrome" in its UA, which web clients commonly answer
// with an "unsupported browser" page. Present the normal Chrome UA instead.
// Bump the major version when the bundled Chromium moves (see `status` output).
const UA_MAJOR = '140';
const USER_AGENT =
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 ' +
  `(KHTML, like Gecko) Chrome/${UA_MAJOR}.0.0.0 Safari/537.36`;

async function launch({ headless }) {
  ensureDirs();
  const { chromium } = require('playwright');
  // An explicit channel selects the full browser build running Chrome's new
  // headless mode. Omitting it gets chromium_headless_shell, a different binary
  // that cannot read this profile's cookie store (see PROFILE_DIR above).
  const ctx = await chromium.launchPersistentContext(PROFILE_DIR, {
    headless,
    channel: CHANNEL,
    viewport: { width: 1512, height: 945 },
    locale: 'en-US',
    timezoneId: Intl.DateTimeFormat().resolvedOptions().timeZone,
    // Only the bundled build needs the spoof; an installed Chrome or Edge
    // should present its own UA, which is what device checks expect to see.
    userAgent: headless && CHANNEL === 'chromium' ? USER_AGENT : undefined,
    args: ['--disable-blink-features=AutomationControlled'],
  });
  ctx.setDefaultTimeout(UI_TIMEOUT);
  ctx.setDefaultNavigationTimeout(NAV_TIMEOUT);
  await ctx.addInitScript(() => {
    try {
      Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
    } catch (_) {}
  });
  await restoreSessionCookies(ctx);
  const page = ctx.pages()[0] || (await ctx.newPage());
  return { ctx, page };
}

// A browser forgets session cookies (the ones with no expiry) when it exits, and
// the SSO provider's sign-in session is exactly such a cookie. Every command
// here is a fresh browser launch, so without this the provider would see a
// brand-new visitor each time and ask for the password again. Carrying them
// across makes consecutive launches look like one browser that stayed open; the
// provider's own server-side idle and lifetime limits still apply untouched.
//
// Chromium's "continue where you left off" preference was tried first and does
// not restore them under Playwright, hence doing it by hand.
//
// This file IS the session, exactly as sensitive as the profile. Written 0600
// and never logged.
async function saveSessionCookies(ctx) {
  const all = await ctx.cookies();
  const session = all.filter((c) => c.expires === -1);
  const tmp = SESSION_COOKIES_FILE + '.tmp';
  fs.writeFileSync(tmp, JSON.stringify(session), { mode: 0o600 });
  fs.renameSync(tmp, SESSION_COOKIES_FILE);
}

async function restoreSessionCookies(ctx) {
  let saved;
  try {
    saved = JSON.parse(fs.readFileSync(SESSION_COOKIES_FILE, 'utf8'));
  } catch (_) {
    return;
  }
  if (!Array.isArray(saved) || saved.length === 0) return;
  // One malformed entry must not cost the rest, so add them one at a time.
  for (const c of saved) await ctx.addCookies([c]).catch(() => {});
}

// Every command closes its context through here so the session cookies are
// captured first. A failed save must not leave the browser running.
async function closeContext(ctx) {
  try {
    await saveSessionCookies(ctx);
  } catch (e) {
    log('could not save session cookies (' + String(e.message).split('\n')[0] + ')');
  } finally {
    await ctx.close();
  }
}

async function isSignedIn(page) {
  let host;
  try {
    host = new URL(page.url()).hostname;
  } catch (_) {
    return false;
  }
  if (!isTeamsHost(host)) return false;
  const n = await page.locator(SEL.shell).count().catch(() => 0);
  return n > 0;
}

// Teams keeps its token cache encrypted under a session cookie, so every fresh
// browser launch has to go back to Entra for tokens. Entra still holds the
// persistent session, but asks "Pick an account" first — one click on the
// remembered tile, no credentials. This makes that click, and nothing else: it
// never touches "Use another account" and never types into a sign-in form.
const ACCOUNT = process.env.TEAMS_SEND_ACCOUNT || CONFIG.account || '';
const MAX_ACCOUNT_PICKS = 3;
const ACCOUNT_TILE = 'div.table[role="button"][data-test-id]';

async function pickRememberedAccount(page) {
  let host;
  try {
    host = new URL(page.url()).hostname;
  } catch (_) {
    return false;
  }
  if (host !== 'login.microsoftonline.com') return false;
  const tiles = page.locator(ACCOUNT_TILE);
  const ids = await tiles.evaluateAll((els) => els.map((e) => e.getAttribute('data-test-id'))).catch(() => []);
  if (ids.length === 0) return false;
  // With an account pinned, take that one; otherwise only act when there is
  // exactly one remembered account, so there is nothing to choose between.
  const want = ACCOUNT ? ids.findIndex((id) => (id || '').toLowerCase() === ACCOUNT.toLowerCase()) : (ids.length === 1 ? 0 : -1);
  if (want < 0) return false;
  const clicked = await tiles.nth(want).click({ timeout: 5000 }).then(() => true, () => false);
  if (clicked) {
    log('picked the remembered account on the Microsoft account picker');
    await page.waitForTimeout(1500);
  }
  return clicked;
}

// Sign-in hops across several origins and an SSO dashboard tile opens Teams in
// a NEW tab, so watch every page in the context rather than the one we opened.
async function waitForSignedIn(ctx, timeout, { autoPick = false } = {}) {
  const deadline = Date.now() + timeout;
  let picks = 0;
  while (Date.now() < deadline) {
    for (const p of ctx.pages()) {
      if (p.isClosed()) continue;
      if (await isSignedIn(p)) return p;
      // Capped: a tile that keeps reappearing means Entra wants something a
      // click cannot give, and hammering it would only bury that in redirects.
      if (autoPick && picks < MAX_ACCOUNT_PICKS && (await pickRememberedAccount(p))) picks++;
    }
    await new Promise((r) => setTimeout(r, 500));
  }
  return null;
}

async function gotoClient(page) {
  await page.goto(CLIENT_URL, { waitUntil: 'domcontentloaded', timeout: NAV_TIMEOUT }).catch((e) => {
    // A redirect chain can abort the original navigation; where it ends up is
    // what matters, and waitForSignedIn decides that.
    log('navigation interrupted (' + String(e.message).split('\n')[0] + ')');
  });
}

// ---------------------------------------------------------------- chats

// Land on the signed-in client or die saying how to fix it.
async function openClient(ctx, page) {
  await gotoClient(page);
  const landed = await waitForSignedIn(ctx, NAV_TIMEOUT, { autoPick: true });
  if (!landed) {
    const p = await shot(page, 'signed-out');
    die('not signed in (stopped at ' + safeUrl(page.url()) + ') — run `teams-send login`.' +
      (p ? `\n  screenshot: ${p}` : ''), 3);
  }
  await landed.locator(SEL.rail).first().waitFor({ state: 'visible', timeout: NAV_TIMEOUT });
  return landed;
}

// The Chats section starts collapsed, and a collapsed section's rows are not in
// the DOM at all. Only the first page of chats is listed; older ones sit behind
// "See more", which this does not page through.
async function chatRows(page) {
  const rail = page.locator(SEL.rail).first();
  const collapsed = rail
    .locator('[role="treeitem"][aria-level="1"][aria-expanded="false"]')
    .filter({ hasText: /^Chats/ });
  if ((await collapsed.count()) > 0) {
    await collapsed.first().click();
    await rail.locator(SEL.chatTitle).nth(1).waitFor({ state: 'visible', timeout: UI_TIMEOUT }).catch(() => {});
  }
  const rows = rail.locator('[role="treeitem"]').filter({ has: page.locator(SEL.chatTitle) });
  const names = await rows.evaluateAll((els) =>
    els.map((el) => (el.querySelector('[id^="title-chat-list-item_"]').textContent || '').trim()));
  return { rows, names };
}

// Open a chat from the left rail and PROVE which one landed. Returns its title.
async function openChat(page, to, force) {
  const { rows, names } = await chatRows(page);
  const pick = pickChat(names, to);
  if (pick.error === 'ambiguous') {
    die(`"${to}" matches several chats: ${pick.candidates.join(' / ')}. Be more specific.`, 2);
  }
  if (pick.error) {
    die(`no chat matching "${to}" in the left rail (run \`teams-send chats\` to see what is listed).`, 2);
  }
  const want = names[pick.index];
  if (norm(want) !== norm(to)) log(`"${to}" -> "${want}"`);
  await rows.nth(pick.index).click();
  // The header is the proof. Opening is asynchronous and the previous chat's
  // header stays up until the new one renders, so poll rather than read once.
  const header = page.locator(SEL.header).first();
  const deadline = Date.now() + UI_TIMEOUT;
  let got = '';
  while (Date.now() < deadline) {
    got = (await header.innerText().catch(() => '')).trim();
    if (norm(got) === norm(want)) break;
    await page.waitForTimeout(250);
  }
  if (norm(got) !== norm(want) && !force) {
    await failWithShot(page, `clicked "${want}" but the open chat is titled "${got}". Refusing to continue (--force overrides).`);
  }
  await page.locator(SEL.composer).first().waitFor({ state: 'visible', timeout: UI_TIMEOUT });
  // Let the message list settle; an empty list right after open means "not
  // rendered yet" far more often than "no messages".
  await page.locator(SEL.msg).first().waitFor({ state: 'visible', timeout: 6000 }).catch(() => {});
  await page.waitForTimeout(800);
  return got || want;
}

// The messages currently rendered, oldest first. The list is virtualised, so
// this is the recent tail of the chat, not its whole history.
async function listMessages(page) {
  return page.locator(SEL.msg).evaluateAll((els) => els.map((el) => {
    const mid = el.getAttribute('data-mid') || '';
    const byId = (p) => document.getElementById(p + mid);
    const t = byId('timestamp-');
    const c = byId('content-');
    return {
      mid,
      author: ((byId('author-') || {}).textContent || '').trim(),
      time: t ? t.getAttribute('datetime') || '' : '',
      text: c ? (c.innerText || '').trim() : '',
    };
  }).filter((m) => m.mid));
}

function formatMessage(m) {
  const d = new Date(m.time);
  const pad = (n) => String(n).padStart(2, '0');
  const when = isNaN(d) ? '????-??-?? ??:??'
    : `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
  return `${when}  ${m.author || '?'}: ${m.text.replace(/\n/g, '\n    ')}`;
}

async function composerText(composer) {
  return ((await composer.innerText().catch(() => '')) || '').replace(/\u200b|\ufeff/g, '').trim();
}

async function clearComposer(composer) {
  await composer.press('Meta+a').catch(() => {});
  await composer.press('Backspace').catch(() => {});
}

// ---------------------------------------------------------------- commands

async function cmdChats() {
  const { ctx, page } = await launch({ headless: true });
  try {
    const p = await openClient(ctx, page);
    const { names } = await chatRows(p);
    for (const n of [...new Set(names)]) console.log(n);
  } finally {
    await closeContext(ctx);
  }
}

async function cmdList(to, { limit, force }) {
  if (!to) die('list: which chat? (see `teams-send chats`)', 2);
  const { ctx, page } = await launch({ headless: true });
  try {
    const p = await openClient(ctx, page);
    const title = await openChat(p, to, force);
    const msgs = await listMessages(p);
    log(`${title}: showing ${Math.min(limit, msgs.length)} of ${msgs.length} rendered messages`);
    for (const m of msgs.slice(-limit)) console.log(formatMessage(m));
  } finally {
    await closeContext(ctx);
  }
}

async function cmdSend({ to, text, force, dryRun }) {
  const { ctx, page } = await launch({ headless: true });
  try {
    const p = await openClient(ctx, page);
    const title = await openChat(p, to, force);
    const composer = p.locator(SEL.composer).first();
    await composer.click();
    // Teams keeps unsent drafts per chat. Typing on top of one would send text
    // you wrote by hand and never meant to send, so stop instead.
    const draft = await composerText(composer);
    if (draft) die(`"${title}" already has an unsent draft in the composer; send or discard it in Teams first.`);
    // By id, not by "newer than": a just-sent row carries a client-side id that
    // is not a timestamp and does not order against the ids of older rows.
    const before = new Set((await listMessages(p)).map((m) => m.mid));

    // Keys go through the composer locator, never page.keyboard, so they cannot
    // land in the search box. A bare Enter would send; Shift+Enter is a newline.
    const lines = text.split('\n');
    for (let i = 0; i < lines.length; i++) {
      if (i > 0) await composer.press('Shift+Enter');
      if (lines[i]) await composer.pressSequentially(lines[i], { delay: 5 });
    }
    // The editor autoformats as you type (lists, emoji, mentions). Read back
    // what it holds and refuse to send anything that is not what was asked for.
    const typed = await composerText(composer);
    if (norm(typed) !== norm(text)) {
      await clearComposer(composer);
      die(`the composer rewrote the message (got ${JSON.stringify(typed.slice(0, 80))}); nothing was sent.`);
    }
    if (dryRun) {
      await clearComposer(composer);
      await p.waitForTimeout(1500);
      log(`dry run: typed into "${title}" and cleared it; nothing was sent.`);
      return;
    }
    await p.locator(SEL.sendButton).first().click();

    // Proof of delivery is the message appearing in the list as a new row.
    const deadline = Date.now() + UI_TIMEOUT;
    let sent = null;
    while (Date.now() < deadline && !sent) {
      await p.waitForTimeout(400);
      sent = (await listMessages(p)).find((m) => !before.has(m.mid) && norm(m.text) === norm(text)) || null;
    }
    if (!sent) await failWithShot(p, `pressed Send in "${title}" but the message never appeared in the chat. Check Teams before retrying.`);
    // Give the client a moment to flush before the browser goes away.
    await p.waitForTimeout(1500);
    console.log(`sent to ${title} (id ${sent.mid})`);
  } finally {
    await closeContext(ctx);
  }
}

async function cmdLogin() {
  const { ctx, page } = await launch({ headless: false });
  try {
    await page.goto(LOGIN_URL, { waitUntil: 'domcontentloaded', timeout: NAV_TIMEOUT }).catch(() => {});
    console.error(
      'teams-send: a browser window is open. Sign in and get to Teams there\n' +
      'teams-send: (via the SSO dashboard tile if that is where you landed).\n' +
      'teams-send: if Microsoft asks "Stay signed in?", answer Yes.\n' +
      'teams-send: this window IS the saved session — do not close it manually;\n' +
      'teams-send: it closes itself once Teams has loaded (waiting up to 10 min).'
    );
    const landed = await waitForSignedIn(ctx, LOGIN_TIMEOUT);
    if (!landed) {
      const where = ctx.pages().filter((p) => !p.isClosed()).map((p) => safeUrl(p.url())).join(', ');
      die('Teams never loaded. Last seen at: ' + (where || '(no open page)'));
    }
    // Let the client finish booting and flush its session to the profile.
    await landed.waitForTimeout(8000);
    log('signed in at ' + safeUrl(landed.url()));
    log('session saved to ' + PROFILE_DIR);
  } finally {
    await closeContext(ctx);
  }
}

async function cmdStatus() {
  const { ctx, page } = await launch({ headless: true });
  try {
    await gotoClient(page);
    const landed = await waitForSignedIn(ctx, NAV_TIMEOUT, { autoPick: true });
    console.log('signed in : ' + (landed ? 'yes' : 'no'));
    console.log('url       : ' + safeUrl((landed || page).url()));
    console.log('channel   : ' + CHANNEL);
    console.log('profile   : ' + PROFILE_DIR);
    console.log('ua        : ' + (await (landed || page).evaluate(() => navigator.userAgent).catch(() => '(unavailable)')));
    console.log('config    : ' + CONFIG_FILE);
    const dep = deployInfo();
    if (dep) {
      console.log(`deployed  : ${dep.hash}  ${dep.deployedAt}${dep.git ? '  git ' + dep.git : ''}   <- what teams-send runs`);
      const verdict = dep.sourceMissing
        ? `source not readable at ${dep.source} — moved? re-run ./install there`
        : dep.stale
          ? 'AHEAD of deploy — run ./install'
          : 'matches deploy';
      console.log(`source    : ${dep.sourceHash || '(missing)'}  ${verdict}`);
    } else {
      console.log('deployed  : nothing — run ./install in the project directory');
    }
    if (!landed) {
      // Where the redirect chain stopped is the diagnosis: an Okta host means
      // the SSO session lapsed, a Microsoft one means Entra wants something.
      const p = await shot(page, 'status');
      if (p) console.log('screenshot: ' + p);
      console.log('\nRun `teams-send login` to authenticate.');
      process.exitCode = 3;
    }
  } finally {
    await closeContext(ctx);
  }
}

async function cmdShot() {
  const { ctx, page } = await launch({ headless: true });
  try {
    await gotoClient(page);
    const landed = await waitForSignedIn(ctx, NAV_TIMEOUT, { autoPick: true });
    const target = landed || page;
    if (landed) await target.waitForTimeout(5000);
    const p = await shot(target, landed ? 'shot' : 'signed-out');
    if (!p) die('could not take a screenshot.');
    console.log(p);
    if (!landed) process.exitCode = 3;
  } finally {
    await closeContext(ctx);
  }
}

// Selector discovery. Dumps the hooks the live client actually exposes so
// selectors are written from evidence. Read-only: it navigates and reads.
async function cmdProbe({ filter, limit, wait }) {
  const { ctx, page } = await launch({ headless: true });
  try {
    await gotoClient(page);
    const landed = await waitForSignedIn(ctx, NAV_TIMEOUT, { autoPick: true });
    if (!landed) die('not signed in — run `teams-send login`.', 3);
    await landed.waitForTimeout(wait);
    const rows = await landed.evaluate(() => {
      const out = [];
      const seen = new Set();
      const nodes = document.querySelectorAll('[data-tid], [role="textbox"], [role="combobox"], [contenteditable="true"]');
      for (const el of nodes) {
        const r = el.getBoundingClientRect();
        const row = {
          tag: el.tagName.toLowerCase(),
          tid: el.getAttribute('data-tid') || '',
          role: el.getAttribute('role') || '',
          aria: (el.getAttribute('aria-label') || '').slice(0, 80),
          editable: el.getAttribute('contenteditable') === 'true',
          visible: r.width > 0 && r.height > 0,
          text: el.children.length === 0 ? (el.textContent || '').trim().slice(0, 60) : '',
        };
        // Lists repeat one tid per row; one sample of each shape is enough.
        const key = [row.tag, row.tid.replace(/[0-9a-f-]{8,}|\d+/gi, '#'), row.role, row.editable].join('|');
        if (seen.has(key)) continue;
        seen.add(key);
        out.push(row);
      }
      return out;
    });
    const f = (filter || '').toLowerCase();
    const kept = rows
      .filter((r) => !f || JSON.stringify(r).toLowerCase().includes(f))
      .slice(0, limit);
    for (const r of kept) console.log(JSON.stringify(r));
    log(`${kept.length} of ${rows.length} distinct hooks at ${safeUrl(landed.url())}`);
  } finally {
    await closeContext(ctx);
  }
}

// ---------------------------------------------------------------- argv

const VALUE_FLAGS = new Set(['--filter', '--limit', '--wait']);

function parseArgv(argv) {
  const positional = [];
  const opts = {};
  const bools = new Set();
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (VALUE_FLAGS.has(a)) {
      if (i + 1 >= argv.length) die(`${a} needs a value`, 2);
      opts[a] = argv[++i];
    } else if (a.startsWith('--') || a === '-h') {
      bools.add(a);
    } else {
      positional.push(a);
    }
  }
  return { positional, opts, bools };
}

function intOpt(opts, name, def) {
  if (opts[name] === undefined) return def;
  const n = Number(opts[name]);
  if (!Number.isInteger(n) || n < 0) die(`${name}: expected a non-negative integer, got "${opts[name]}"`, 2);
  return n;
}

// ---------------------------------------------------------------- cli

function usage() {
  console.error(`usage:
  teams-send <chat> "<message>"        send a message (headless)
  teams-send <chat>                    read the message from stdin
  teams-send list <chat> [--limit N]   print the most recent messages (default 20)
  teams-send chats                     list the chats in the left rail
  teams-send login                     interactive sign-in (opens a window)
  teams-send status                    report session state (headless)
  teams-send shot                      screenshot the client (read-only)
  teams-send probe [--filter s] [--limit N] [--wait ms]
                                       dump the DOM hooks the client exposes

options:
  --dry-run    (send) type the message, then clear it without sending
  --force      continue even if the opened chat's title does not match
  --quiet      suppress progress chatter on stderr

notes:
  <chat> is a name from the left rail: exact, or an unambiguous part of one.
  "me" is your own chat. Opening a chat marks it read, as it would by hand.

state    : ${STATE_DIR}
channel  : ${CHANNEL}
client   : ${CLIENT_URL}
login at : ${LOGIN_URL}`);
}

const COMMANDS = {
  login: () => cmdLogin(),
  status: () => cmdStatus(),
  shot: () => cmdShot(),
  chats: () => cmdChats(),
  list: (a) => cmdList(a.positional[1], {
    limit: intOpt(a.opts, '--limit', 20),
    force: a.bools.has('--force'),
  }),
  probe: (a) => cmdProbe({
    filter: a.opts['--filter'] || '',
    limit: intOpt(a.opts, '--limit', 400),
    wait: intOpt(a.opts, '--wait', 8000),
  }),
};

async function main() {
  const a = parseArgv(process.argv.slice(2));
  QUIET = a.bools.has('--quiet');
  const help = a.bools.has('--help') || a.bools.has('-h');
  if (help || a.positional.length === 0) {
    usage();
    process.exit(help ? 0 : 2);
  }
  const sub = a.positional[0];
  if (Object.prototype.hasOwnProperty.call(COMMANDS, sub)) return COMMANDS[sub](a);

  // Default form: teams-send <chat> "<message>"
  const to = a.positional[0];
  let text = a.positional.slice(1).join(' ');
  if (!text) {
    if (process.stdin.isTTY) die('no message given (pass one as an argument or pipe it in).', 2);
    text = fs.readFileSync(0, 'utf8').replace(/\n+$/, '');
  }
  if (!text.trim()) die('refusing to send an empty message.', 2);
  // A leading slash opens the composer's command menu instead of typing text.
  if (text.trimStart().startsWith('/') && !a.bools.has('--force')) {
    die('message starts with "/" — Teams would treat it as a command. Re-run with --force if that is intended.', 2);
  }
  return cmdSend({ to, text, force: a.bools.has('--force'), dryRun: a.bools.has('--dry-run') });
}

// Only run when executed directly, so tests can require() the pure helpers.
if (require.main === module) {
  main().catch((e) => {
    if (e instanceof ExitError) {
      console.error(`teams-send: ${e.message}`);
      process.exit(e.exitCode);
    }
    console.error(`teams-send: ${e && e.stack ? e.stack : String(e)}`);
    process.exit(1);
  });
}

module.exports = {
  parseArgv, intOpt, isTeamsHost, safeUrl, norm, pickChat, formatMessage, COMMANDS, ExitError,
  openClient, chatRows, openChat, listMessages, deployInfo,
  launch, gotoClient, isSignedIn, waitForSignedIn, pickRememberedAccount, closeContext, log, die, STATE_DIR,
};
