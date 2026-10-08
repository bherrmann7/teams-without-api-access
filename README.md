# teams-without-api-access

Read and post Microsoft Teams chats **as yourself**, from the command line, on an account
where you cannot register a Graph application.

Most Teams automation assumes an app registration in Entra ID and an admin willing to consent
to it. Plenty of tenants don't offer that. This tool exists for that case. It signs in the way
you do, in a real browser, through whatever single sign-on your organisation puts in front of
Teams, and then drives the Teams web client with your own session. No app, no bot, no token to
request from anyone.

Everything it does, it does **as you**. Messages come from your account, not from a bot with
your avatar.

```console
$ teams-send "Alice Example" "year-end run finished, 0 differences"
$ teams-send list "Alice Example" --limit 30
$ teams-send me "note to self"
```

It is the sibling of [slack-without-api-access](https://github.com/bherrmann7/slack-without-api-access)
and follows the same layout: one file, `send.js`, plus `install`.

## How it works

One path: the browser. Every command launches Chromium headlessly with
[Playwright](https://playwright.dev) against a persistent profile, opens the Teams web client,
and reads or types there. Expect 15 to 30 seconds per command.

There is no HTTP fast path like the Slack tool has. Teams' chat service authenticates with a
short-lived derived token, so a harvested credential would need the browser to refresh it
anyway.

Three things make a headless relaunch land back in Teams without a password prompt:

- **A persistent profile** holds the long-lived cookies.
- **Session cookies are carried across launches by hand.** The SSO provider's sign-in session
  is a cookie with no expiry, which a browser forgets on exit, and every command here is a new
  browser. The tool saves those cookies on close and restores them at the next launch, so
  consecutive commands look like one browser that stayed open. The provider's own idle and
  lifetime limits still apply.
- **Microsoft's "Pick an account" tile is clicked automatically** when exactly one account is
  remembered (or the one pinned in config). That is the only sign-in step the tool performs; it
  never types a credential.

## Please read this before using it

This drives Teams with **your own user session**, which is not the same thing as an approved
integration:

- **It is very likely outside your organisation's acceptable-use policy, and possibly
  Microsoft's terms.** Automating a user session is not a supported use of the product.
- **This is a corporate identity, not just a chat login.** The same session reaches whatever
  else your SSO covers. Sign-ins from an automated browser appear in your tenant's sign-in logs,
  and whether that is fine or a serious problem depends on your organisation. It is on you to
  know which.
- **Actions are indistinguishable from you doing them by hand,** because that is what they are.
  There is no bot label and no audit trail saying a script did it.
- **Reading has a side effect.** Opening a chat marks it read, exactly as clicking it would.

It was written for one person automating their own workflow on their own account. That is the
use it is fit for.

## Install

```console
git clone https://github.com/bherrmann7/teams-without-api-access
cd teams-without-api-access
npm install
./install             # publishes a snapshot and writes ~/bin/teams-send
teams-send login      # opens a browser; sign in there
```

`login` opens a real browser window and waits (up to 10 minutes) for Teams to load. Sign in
however you normally do, including any SSO dashboard tile that opens Teams in a new tab; the
tool watches every tab. If Microsoft asks "Stay signed in?", answer Yes. That window **is** the
saved session: don't close it manually; it closes itself once Teams has loaded.

To start `login` at your SSO dashboard instead of at Teams, set `loginUrl` in the config file
(see [State](#state)).

Re-run `teams-send login` whenever the session expires. `teams-send status` tells you where you
stand, and exits 3 when you are signed out.

### Why `./install` and not `npm link`

`./install` publishes a snapshot, so a script that announces a finished job never runs a
half-finished edit of your working copy. It runs the tests, checks the syntax, executes
`send.js --help` (parsing is not running), confirms Playwright resolves from the deploy
directory, and only then copies `send.js` to `~/.teams-send/deploy/` and writes a `teams-send`
wrapper into `~/bin` pointing there. It copies rather than symlinks, deliberately.

If `~/bin/teams-send` already exists and was not written by `install`, the first run keeps it
at `~/bin/teams-send.bak`.

`teams-send status` shows which edition is live:

```
deployed  : d74884d0f216  2026-10-07 22:56:22  git 67bc5d8   <- what teams-send runs
source    : 0680ab670455  AHEAD of deploy — run ./install
```

To exercise an uncommitted edit, run `node send.js <args>` from the project directory.

## Commands

| command | what it does |
|---|---|
| `teams-send <chat> "<message>"` | post a message (also reads stdin) |
| `teams-send list <chat> [--limit N]` | print the most recent messages with timestamps |
| `teams-send chats` | list the chats that can be addressed |
| `teams-send login` / `status` | sign in; report session and deploy state |
| `teams-send shot` | screenshot the client (read-only, for debugging) |
| `teams-send probe [--filter s]` | dump the DOM hooks the live client exposes |

`<chat>` is a name from the Teams left rail. `me` is your own chat.

### Safety

The chat is checked, not guessed. A name must match a left-rail row exactly, or be a part of
exactly one row's name; an ambiguous name is refused outright instead of picked. After clicking
the row, the tool reads the open chat's header and refuses to type unless it matches.

A send is verified twice: the composer's contents are read back and must equal the message
before Send is clicked (the editor autoformats as you type), and the command only reports
success once the message appears in the chat as a new row.

It refuses to type into a chat that already holds an unsent draft, and refuses a message that
starts with `/`.

## Flags

```
--dry-run    (send) type the message, then clear it without sending
--limit N    how many messages to print (list), or hooks (probe)
--force      continue even if the opened chat's title does not match; also allows a
             message starting with "/"
--quiet      suppress progress output on stderr
```

## State

Everything lives in `~/.teams-send/` (override with `TEAMS_SEND_HOME`):

| file | contents |
|---|---|
| `profile-<channel>/` | the browser profile — **this is your login session** |
| `session-cookies.json` | the no-expiry cookies carried between launches, mode `600` |
| `config.json` | optional settings (below) |
| `deploy/` | the released edition `teams-send` runs |
| `shots/` | screenshots written on failure |

`config.json` keys, each also settable by environment variable:

| key | env | meaning |
|---|---|---|
| `loginUrl` | `TEAMS_SEND_LOGIN_URL` | where `login` starts, e.g. your SSO dashboard |
| `url` | `TEAMS_SEND_URL` | the Teams client URL (default `https://teams.cloud.microsoft/`) |
| `account` | `TEAMS_SEND_ACCOUNT` | which remembered account to pick, if you have several |
| `channel` | `TEAMS_SEND_CHANNEL` | `chromium` (bundled, default), `chrome` or `msedge` |

Use `chrome` or `msedge` if your tenant only admits a browser that can present a managed-device
identity. Each channel gets its own profile directory, so changing it means signing in again.

`profile-*/` and `session-cookies.json` are as sensitive as your password. They are written
owner-only; keep them that way, and don't copy them anywhere.

## Limitations

- **It depends on Teams' DOM.** It leads with the `data-tid` attributes Teams ships for its own
  tests, which are the most durable hooks available, but Microsoft can still change them. Every
  non-obvious selector carries a comment saying why. `teams-send probe` shows what the live
  client currently exposes.
- **`list` shows only the recent tail.** It prints the messages Teams renders on opening the
  chat and does not scroll back through history.
- **Only the first page of chats is addressable.** Chats behind "See more" in the left rail are
  not reached. Channels in a team are not addressable at all.
- **Images and attachments are not read.** They come through as empty messages.
- **No reactions, edits, deletes or threads.**
- **The session lasts as long as your SSO provider allows.** Expect to re-run `login`
  periodically.
- Tested on macOS, against one tenant behind Okta. Other identity providers may add sign-in
  steps this has never seen.

## Development

```console
npm test              # unit tests, no browser
node --check send.js  # syntax
```

The tests cover the pure logic (argument parsing, chat-name matching, URL and host checks).
Anything that touches Playwright is verified by hand: if you change a selector, drive it
against a real account, and prefer your own chat (`me`) as the target.

Helpers are exported at the bottom of `send.js` so tests and throwaway probe scripts can reach
them. None carries a stability promise.

## License

MIT
