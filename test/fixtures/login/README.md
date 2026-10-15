# Microsoft identity platform answers

Recorded on 15 October 2026 from `login.microsoftonline.com` (`/common/oauth2/v2.0/devicecode` and
`/token`), with an invented client ID for the refusals. The trace and correlation ids are zeroed; the
user and device codes of `devicecode.json` are invented (the real ones expired 15 minutes later).

- `devicecode.json`: a device code (HTTP 200). Note `verification_uri`: `https://login.microsoft.com/device`.
- `authorization-pending.json`: a token request before the user typed the code (HTTP 400).
- `unknown-client.json`: a device code request with an unknown client ID on `common` (HTTP 400).
- `bad-device-code.json`: a token request with an unknown device code (HTTP 400).
- `invalid-refresh-token.json`: a refresh with a refresh token Microsoft does not know (HTTP 400).
