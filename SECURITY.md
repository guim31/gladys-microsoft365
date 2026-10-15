# Security policy

## Reporting a vulnerability

Please do **not** report a security problem in a public issue, pull request or
forum post.

Report it privately through GitHub instead: open the **Security** tab of this
repository, then **Report a vulnerability**.

No such button? Private reporting is not enabled yet: open an issue that only
asks the maintainers for a private contact, without any detail of the problem,
and send the report there.

Include, when you can:

- the version of the integration and of Gladys;
- what an attacker can do, and the steps to reproduce it;
- the log lines that help (remove tokens, passwords and addresses first).

The fix ships as a new version of the integration, and the advisory is
published once users can update.

## Supported versions

Only the latest release receives security fixes: update the integration from
Gladys before reporting a problem.

## Scope

In scope: the code of this repository and the Docker image it publishes, in particular:

- what is stored in `/data`: `tokens.json` (mode 600) holds the Microsoft **refresh token**, the client
  ID and tenant it was issued for, and the account id and display name; `state.json` holds the Graph
  ids of the calendars and lists and the ids of the tasks already announced as overdue. No event and
  no task content is stored. The access token lives in memory only;
- secrets in logs: no token, device code or client ID may ever be logged, at any level;
- network: the container only talks to `login.microsoftonline.com` and `graph.microsoft.com` over
  HTTPS, and follows Graph paging links only when they point to `graph.microsoft.com`;
- the permissions requested from Microsoft (`Calendars.ReadWrite`, `Tasks.ReadWrite`, `offline_access`,
  `User.Read`).

Out of scope: Gladys Assistant itself (see its
[security policy](https://github.com/GladysAssistant/Gladys/security/policy)),
Microsoft's services, and the app registration the user creates in Microsoft Entra.
