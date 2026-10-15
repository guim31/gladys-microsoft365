# Changelog

All notable changes to this integration are documented in this file. The format
follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/) and the project
uses [semantic versioning](https://semver.org/).

Describe each change under `## [Unreleased]` as you make it. The Release
workflow moves that section under the version it ships, and the section becomes
the notes of the version's GitHub Release.

## [Unreleased]

### Added

- Sign-in to a Microsoft account (personal or work) with the device code flow and your own Azure app
  registration: "Sign in", "Check the connection" and "Sign out" actions, the code shown in the
  connection status, the refresh token kept in `/data`.
- One device per Outlook calendar: busy, out of office, current event, next event.
- One device per Microsoft To Do list: open tasks, overdue tasks, tasks due today.
- Scene triggers: an Outlook event starts, ends, or starts in N minutes (filtered by calendar and by
  keywords listed in the settings); a To Do task becomes overdue.
- Scene actions: add a To Do task, complete a To Do task, add an Outlook event.
- Dashboard widget "Agenda and tasks", with an "Add a task" form (Gladys 5.2.0).
- Refresh on the integration's own timer, Microsoft Graph throttling (`Retry-After`) honoured.
- Documentation in English, French, German and Spanish.
