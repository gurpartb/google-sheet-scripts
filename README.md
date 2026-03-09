# google-sheet-scripts

Monorepo of Google Apps Script projects for Google Sheets.

## Projects

| Project                      | Description                                                               |
| ---------------------------- | ------------------------------------------------------------------------- |
| [trip-tools](./trip-tools/)  | Dispatch trip entry dialog for parsing and writing trip data to a sheet   |

## Setup

**Prerequisites:** Node.js, npm

```bash
# Install shared dev dependencies (GAS type hints)
npm install

# Install clasp globally
npm install -g @google/clasp

# Authenticate with Google
clasp login
```

## Per-project workflow

```bash
cd <project>/

# First time — link to existing GAS project
clasp clone <script-id>

# Deploy local changes to GAS
clasp push

# Auto-push on file save
clasp push --watch

# Open GAS editor in browser
clasp open
```

> **Script ID:** found in GAS editor under Project Settings.

## Testing

Testing is done manually via the GAS editor. Use `clasp open` to open the editor, then run functions directly from there.

## Adding a new project

1. Create a new directory: `mkdir <project-name>`
2. Copy boilerplate: `.claspignore`, `jsconfig.json` from an existing project
3. `cd <project-name> && clasp clone <script-id>`
