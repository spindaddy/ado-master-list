# ADO Master List

Desktop app (Electron) for a personal master list of Azure DevOps work items assigned to you, plus Outlook and Teams in the same window.

Latest installer: [GitHub Releases](https://github.com/spindaddy/ado-master-list/releases) (current **v1.0.6**).

## Features

- **Multi-org Azure DevOps** — one PAT per organization; skip orgs with Active off
- **Sync assigned to `@Me`** — keeps tickets you already have, updates them, adds newly assigned active items, and drops ones that close or leave your list
- Failed org syncs (for example Azure 503) **leave that org’s tickets in place**
- **Auto-sync** — off, or every 2 / 5 / 15 / 30 / 60 minutes (also refreshes Outlook and Teams unread badges)
- **Now** — pin at most three work items or local notes; this is the launch screen
- **Notes** — local tasks that are not in Azure DevOps (separate from work items)
- **Work detail** — description, then discussion underneath; open the item in ADO in the browser
- **Outlook** — Outlook on the web (work week) inside the app
- **Teams** — Teams web, presence keep-alive, screen share, Mute for calls
- **Custom https tabs**
- **Loud alerts** — new tickets, Outlook mail, Teams messages (calendar reminders are not treated as mail); Mute on the top bar
- **Light / Dark** in Settings
- PATs encrypted at rest on this Mac

Work items no longer have a local notes field on the ticket. Use the **Notes** tab instead.

## Setup

```bash
npm install
npm start
```

`npm start` builds and launches `ADO Master Electron` from `/Applications`.

## Azure DevOps PAT

Create a Personal Access Token with at least:

- **Work Items** — Read

In Settings, add a row per organization:

- **Organization** — e.g. `contoso` (from `https://dev.azure.com/contoso`)
- **PAT** — your token

The app queries work items where `Assigned To = @Me`, excluding closed / done / backlog-style states.

## Scripts

| Command | Description |
|---------|-------------|
| `npm start` | Build and launch the desktop app |
| `npm run build` | Production build only |
| `npm run dist` | Package for the current OS (`release/`) |

Installers: macOS `.zip` (`arm64` or `x64`), Windows portable `.exe`, Linux `.AppImage`. If macOS blocks the app: right-click → Open, or `xattr -cr` on the `.app`.
