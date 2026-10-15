# Microsoft 365

Your **Outlook calendar** and your **Microsoft To Do** lists in Gladys, through the official Microsoft
Graph API: busy or free, the current and next event, open and overdue tasks, scenes that start with
an event or add a task, and a dashboard widget with today's agenda.

Works with a personal Microsoft account (Outlook.com, Hotmail, Live) and with a work or school
account (Microsoft 365). It complements the "ICS Calendars" integration: no public calendar link
here, and Gladys can write too (tasks, events).

> **Developed without a test account: feedback welcome.** The sign-in and every Microsoft answer
> were built from Microsoft's documentation and recorded answers, not tested on a real account yet.
> Tell us on the Gladys forum how it goes.
>
> This integration is not affiliated with Microsoft. Microsoft, Microsoft 365, Outlook and To Do
> are trademarks of the Microsoft group of companies.

## What you need

- Gladys 5.2.0 or later.
- A Microsoft account with a mailbox: Outlook.com for a personal account, Exchange Online for a work
  account (To Do and the calendar live in the mailbox).
- Your own **app registration** in Microsoft Entra (free, 5 minutes, see below). Microsoft asks every
  application reading a calendar to be registered; with your own, nobody else holds an access to
  your data, and Microsoft's quotas are yours alone.

## Create your app registration

1. Open [App registrations](https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
   in Microsoft Entra and sign in. With a personal account, Microsoft may first ask you to create a
   free directory (a free Azure account): the registration itself costs nothing.
2. **New registration**: any name (for example "Gladys"). Supported account types: **Accounts in any
   organizational directory and personal Microsoft accounts**. Leave the redirect URI empty. Register.
3. On the **Overview** page, copy the **Application (client) ID**.
4. **Authentication** → Advanced settings → **Allow public client flows**: **Yes**, then Save. Without
   it, Microsoft refuses the sign-in by code (AADSTS7000218).
5. **API permissions** → Add a permission → Microsoft Graph → **Delegated permissions**: tick
   `Calendars.ReadWrite`, `Tasks.ReadWrite` and `offline_access` (`User.Read` is already there).
   No client secret, no certificate: the integration is a public client.

A work account may need an administrator to grant consent to these permissions for the organization
(**Grant admin consent**). If your organization only allows its own accounts, choose "Accounts in this
organizational directory only" in step 2 and enter the directory (tenant) ID in the **Tenant**
setting.

## Configuration

In Gladys, **Integrations → Microsoft 365 → Configuration**:

- **Application (client) ID**: the ID copied in step 3.
- **Tenant**: `common` (personal and work accounts). `consumers` for personal accounts only,
  `organizations` for work accounts only, or your directory ID.
- **Read Microsoft 365 every**: 1 to 15 minutes (5 by default). The triggers keep their minute
  precision whatever the choice.
- **Keywords for the scene triggers**: words to find in event titles, separated by commas (see the
  scene triggers).
- **Language of the device texts**: names of the features (set when a device is added), texts of the
  current and next event, messages. French by default.

Save, then use the actions under the form:

- **Sign in**: shows a code and a Microsoft page (`https://login.microsoft.com/device`). Open the page
  on any device, enter the code within 15 minutes and sign in with your Microsoft account. The status
  of the Configuration screen shows the code meanwhile, then turns to connected. The devices appear
  in the **Discover** tab.
- **Check the connection**: tells which account is signed in and how many calendars and To Do lists
  it holds, and refreshes the Discover tab.
- **Sign out**: deletes the Microsoft tokens from Gladys. Do it before signing in with another
  account. To revoke the access on Microsoft's side too, remove the application from
  [your account's apps](https://account.live.com/consent/Manage) or from My Apps (work account).

The session lasts as long as it is used: Microsoft ends it after 90 days without use, or when the
password changes or the access is revoked. The status then says to sign in again.

## Devices

The **Discover** tab lists one device per Outlook calendar and one per To Do list of the account.
Add the ones you want: only those are read.

### Outlook calendar

| Feature       | Value                                                                                         |
| ------------- | --------------------------------------------------------------------------------------------- |
| Busy          | 1 during an event shown as busy, tentative or out of office; 0 when free or working elsewhere |
| Out of office | 1 during an event shown as out of office (holidays, appointments)                             |
| Current event | title of the event in progress (a timed one first), or "None"                                 |
| Next event    | title and start of the next event within two weeks ("Dentist · tomorrow 14:00"), or "None"    |

Cancelled events and invitations you declined are ignored. All-day events cover their whole day.

### To Do list

| Feature         | Value                                     |
| --------------- | ----------------------------------------- |
| Open tasks      | tasks not completed                       |
| Overdue tasks   | open tasks whose due date is before today |
| Tasks due today | open tasks due today                      |

A task due today is not overdue before tomorrow: the counts change at midnight.

## Dashboard widget

**Agenda and tasks**: what is left of today in Outlook (8 events at most, "Now" on the one in
progress), the open and overdue counts of a To Do list as live tiles, its first five open tasks with
their due date (red when overdue, orange when due today), and an **Add a task** button that opens a
small form (title, due date). Settings: the calendar (empty: every calendar added) and the To Do list
(empty: the list, if only one was added).

## Scene triggers

The core compares the filters for equality only: "the title contains a word" is decided by the
integration, from the **Keywords for the scene triggers** setting. An event whose title contains one
of these words (case and accents ignored) carries the first one found; a scene filters on it by
typing the keyword as written in the settings.

Each event trigger gives the scene: the event title, location, start, end, the calendar and the
keyword found.

### Outlook event starts

When an event of an added calendar begins. Filters: calendar, keyword.

### Outlook event ends

When an event is over. Filters: calendar, keyword.

### Outlook event in N minutes

5, 10, 15, 30, 60 or 120 minutes before an event begins (all-day events excluded). Filters: minutes
(required), calendar, keyword.

### To Do task overdue

When an open task passes its due date (at the first check after midnight). Filter: the To Do list.
Gives the task, its due date and the list. Tasks already overdue when the list is first read are not
announced, and a task is announced once.

Triggers are computed every 30 seconds from the events already read: an event fires at its time even
if Microsoft 365 is read every 15 minutes. A start missed by more than 5 minutes (Gladys or the
integration stopped) is not announced late, and nothing is announced from data more than 2 hours
old.

## Scene actions

### Add a To Do task

To a To Do list added to Gladys, with a title (scene variables allowed) and an optional due date in
days (0: today, 1: tomorrow). Returns the title.

### Complete a To Do task

Completes the open task whose title is the text (case and accents ignored), or the only one whose
title contains it. Several partial matches: nothing is completed. Returns "Task completed (1 yes,
0 no)" and the task title.

### Add an Outlook event

Adds a busy event to a calendar, now or in up to 7 days (in minutes), lasting 5 minutes to 24 hours.
Returns the start time.

## Scene examples

- **Remote work**: keyword `Remote work`. Trigger "Outlook event starts", keyword `Remote work` → turn
  the office heating to comfort; "Outlook event ends" → back to eco.
- **Leave on time**: trigger "Outlook event in N minutes", 30 minutes, keyword `Dentist` → message
  "{{triggerEvent.data.title}} at {{triggerEvent.data.start}}, {{triggerEvent.data.location}}".
- **Do not disturb**: device feature "Busy" of the work calendar goes to 1 → doorbell muted, a light
  outside the office turns red; back to 0 → normal.
- **Shopping list**: a button or a voice command → "Add a To Do task" "Milk" to the "Groceries" list.
- **Chores**: trigger "To Do task overdue" on the "Home" list → notification "{{triggerEvent.data.title}}
  is overdue".

## Refresh and Microsoft limits

The integration reads Microsoft Graph itself, on its own timer (it does not rely on the Gladys poll):
one request per added calendar or list per period, plus one per 200 events or 100 tasks. Microsoft
allows 10,000 requests per 10 minutes and 4 at a time for one application and one mailbox: with a
1-minute period and 10 devices, the integration uses 100. Requests go one at a time. When Microsoft
asks to slow down (HTTP 429 with `Retry-After`), the integration waits the time asked, without any
request meanwhile.

Values are published only when they change, plus every 6 hours so that Gladys never shows them as
stale. The use of Microsoft Graph follows the
[Microsoft APIs Terms of Use](https://learn.microsoft.com/legal/microsoft-apis/terms-of-use).

## Troubleshooting

- **"Microsoft refused the sign-in (invalid_request (AADSTS50059))"**, or **AADSTS700016**: the client ID is unknown, or
  the app is registered for one organization while the tenant is `common`: check the ID, or enter the
  directory ID as tenant.
- **AADSTS7000218**: "Allow public client flows" is not enabled (step 4).
- **AADSTS65001 or "consent"**: a work account whose administrator must grant the permissions.
- **The code expired**: it lasts 15 minutes; run "Sign in" again.
- **"The Microsoft session expired or was revoked"**: run "Sign in" again.
- **A To Do list or a work calendar gives an error**: the account needs an Exchange Online mailbox
  (To Do of a work account without mailbox is not reachable through Graph).
- **No device in the Discover tab**: run "Check the connection", then the scan of the Discover tab.
- **Logs**: the integration logs only its life cycle and the problems (once when they start, once
  when they end). `LOG_LEVEL=debug` shows the rest. Tokens and codes are never logged.

## Limits

- One Microsoft account per installation (the administrator's), not one per Gladys user.
- The calendar is read from yesterday to two weeks ahead; no event of the Gladys calendar view
  (the "ICS Calendars" integration does that).
- Feature names are fixed when a device is added; changing the language later renames nothing.
- Shared calendars appear if Microsoft lists them in the account's calendars; group calendars do not.
- No reminders, no attendees, no event edition beyond "Add an Outlook event".
- Developed without a test account: the real answers of a Microsoft account may still differ on some
  detail.

## Privacy and security

- The Microsoft refresh token stays in Gladys, in `/data/tokens.json` (readable by the container
  only); the one-hour access token is kept in memory. Nothing else is stored: no event, no task.
- The integration talks only to `login.microsoftonline.com` and `graph.microsoft.com`.
- Requested permissions: `Calendars.ReadWrite`, `Tasks.ReadWrite`, `offline_access` and `User.Read`
  (your name and account id, to label the session).
