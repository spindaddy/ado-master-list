# ADO Master List

Electron app for browsing Azure DevOps projects and maintaining a personal master working list.

## Features

- Connect to an Azure DevOps organization with a PAT
- Sync **work items assigned to you** (`@Me`)
- Add / remove items from your master working list
- Local notes, status, and priority
- Push title/state/description updates back to Azure DevOps
- Open any work item in the browser

## Setup

```bash
npm install
npm run dev
```

## Azure DevOps PAT

Create a Personal Access Token with at least:

- **Work Items** — Read (required)
- **Work Items** — Read & write (to push updates)

In Settings, enter:

- **Organization** — e.g. `contoso` (from `https://dev.azure.com/contoso`)
- **PAT** — your token

The app queries: work items where `Assigned To = @Me`, excluding Closed / Removed / Done.

## Scripts

| Command | Description |
|---------|-------------|
| `npm run dev` | Start Electron in development |
| `npm run build` | Production build |
| `npm start` | Preview production build |
| `npm run dist` | Package for the current OS (`release/`) |

Installers are published on the [GitHub Releases](https://github.com/spindaddy/ado-master-list/releases) page (macOS `.app` zip, Windows `.exe`, Linux `.AppImage`).
