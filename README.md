# Microsoft 365 for Gladys Assistant

An external integration for [Gladys Assistant](https://gladysassistant.com) that brings the
**Outlook calendar** and **Microsoft To Do** of a Microsoft account (personal or work) into Gladys,
through the official [Microsoft Graph](https://learn.microsoft.com/graph/overview) API.

> **Developed without the hardware: feedback welcome.** Here the "hardware" is a Microsoft account:
> the sign-in was exercised against the real Microsoft identity platform up to the refusal of an
> invented application, and Graph from Microsoft's documented answers, but no real account has used
> it yet. Please report on the Gladys forum how it goes.
>
> Not affiliated with Microsoft. Microsoft, Microsoft 365, Outlook and To Do are trademarks of the
> Microsoft group of companies.

## What it does

- **One device per Outlook calendar**: _Busy_ (busy, tentative or out of office), _Out of office_,
  _Current event_ and _Next event_.
- **One device per To Do list**: _Open tasks_, _Overdue tasks_, _Tasks due today_.
- **Scene triggers**: an Outlook event starts, ends, or starts in 5 to 120 minutes (filters:
  calendar, keyword from a configured list); a To Do task becomes overdue.
- **Scene actions**: add a To Do task (title, list, due date), complete a To Do task (by title), add
  an Outlook event.
- **Dashboard widget** "Agenda and tasks": what is left of today, the tasks of a list, and an
  "Add a task" button with a form.

The user documentation, in English, French, German and Spanish, is in [`docs/`](docs/) — Gladys shows
it from the Configuration screen.

## Requirements

- Gladys **5.2.0** or later (the widget's form button).
- A Microsoft account with a mailbox (Outlook.com, or Exchange Online for a work account).
- A free **app registration** of your own in Microsoft Entra: public client ("Allow public client
  flows"), accounts in any directory and personal accounts, delegated permissions
  `Calendars.ReadWrite`, `Tasks.ReadWrite`, `offline_access` (plus the default `User.Read`). The
  sign-in is the OAuth 2.0 **device code** flow: Gladys shows a code, the user types it at
  `https://login.microsoft.com/device`. No client secret, no redirect URI. Step by step:
  [`docs/en.md`](docs/en.md#create-your-app-registration).

## Installation

In Gladys: **Integrations → Install from GitHub**, with this repository's address
(`https://github.com/guim31/gladys-microsoft365`), or from the store once it is listed. Then fill in
the application (client) ID, save, run **Sign in**, and add the calendars and lists from the
**Discover** tab.

## How it works

- Sign-in: `src/auth.js` (device code, refresh token rotation). The refresh token is kept in
  `/data/tokens.json` (mode 600), the access token in memory only.
- Graph: `src/graph.js`, one request at a time, `Retry-After` honoured (429/503/504), Graph paging
  links only. Outlook allows 10,000 requests per 10 minutes and 4 at a time per app and mailbox; the
  integration makes one request per added device and period (1 to 15 minutes).
- Refresh: the container runs its own 30-second timer (`src/monitor.js`) — the Gladys core does not
  reliably poll external devices (GladysAssistant/Gladys#3198); the core's `onPoll` stays as a
  fallback. States are recomputed from what was read every 30 seconds, so triggers fire on the
  minute whatever the Graph period.
- Logs are quiet: nothing at the info level for a normal cycle; `LOG_LEVEL=debug` for the rest.

## Development

```bash
npm ci
npm run format:check && npm run lint && npm test
npm run probe   # real Microsoft endpoints, up to the refusal of an invented client ID
```

The tests run against the SDK's fake Gladys (`@gladysassistant/integration-sdk/testing`) and a fake
Microsoft served through an injected `fetch`: no network. Fixtures: the examples of the Microsoft
Graph reference and answers recorded from the identity platform (`test/fixtures/`). Rules for
contributors and coding assistants: [`CLAUDE.md`](CLAUDE.md).

Releases go through **Actions → Release** (see `CHANGELOG.md`); never edit the versions by hand.

## Limits

- One Microsoft account per Gladys installation (the administrator's).
- The calendar is read from yesterday to two weeks ahead; events do not appear in the Gladys calendar
  view (use the "ICS Calendars" integration for that).
- Keyword filters are decided by the integration from a configured list: the Gladys scene engine only
  compares for equality.
- Group calendars, reminders and attendees are not handled.

## Credits

- Microsoft Graph and Microsoft identity platform documentation, and the examples of
  [microsoftgraph/microsoft-graph-docs-contrib](https://github.com/microsoftgraph/microsoft-graph-docs-contrib)
  (© Microsoft, CC BY 4.0) used as test fixtures.
- Home Assistant has no Microsoft To Do or Outlook integration in its core (its `microsoft`
  integration is text-to-speech): this one follows the Microsoft documentation, no code was copied.
- Built on [`@gladysassistant/integration-sdk`](https://github.com/GladysAssistant/integration-sdk-js)
  and the official integration template.

## License

[Apache-2.0](LICENSE).
