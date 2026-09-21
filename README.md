# Sheets Sync

An Obsidian plugin that keeps a markdown table in a note bidirectionally in sync with a range in Google Sheets. Edit in either place; push or pull to reconcile.

**Desktop only** (uses Node APIs for the OAuth loopback flow).

## Features

- **Per-note mapping** — each note declares its own spreadsheet/range in frontmatter, so any number of notes can sync to different sheets (or different ranges of the same sheet)
- **Pull** — fetch the sheet range and render it as a GitHub-flavored markdown table between `<!-- sheets-sync:start/end -->` markers, preserving the rest of the note
- **Push** — write the note's table back to the sheet; pads the range with empty cells so stale remote rows/columns are cleared
- **Conflict detection** — a per-note baseline hash detects when both sides changed since the last sync and asks before overwriting
- **Auto-pull** — optional poll interval sweeps every configured note in the vault
- **Loopback OAuth** — browser-based Google sign-in with a local redirect; token auto-refreshes, no code pasting
- Cell-safe round-tripping: pipes and newlines inside cells are escaped (`\|`, `<br>`)

## Installation

### From source

```bash
git clone https://github.com/personofnorank/sheets-sync-plugin.git
cd sheets-sync-plugin
npm install
npm run build
mkdir -p /path/to/vault/.obsidian/plugins/sheets-sync
cp main.js manifest.json /path/to/vault/.obsidian/plugins/sheets-sync/
```

Then enable **Sheets Sync** in Settings → Community plugins.

## Google OAuth setup

The plugin talks to the Google Sheets API as an installed (desktop) app, so you need an OAuth 2.0 **Desktop** client:

1. In [Google Cloud Console](https://console.cloud.google.com/), create or pick a project.
2. Enable the **Google Sheets API**.
3. Configure the OAuth consent screen (External is fine; add yourself as a test user while the app is in "testing" status).
4. Create credentials → **OAuth client ID** → application type **Desktop app**.
5. In the plugin settings, either paste the client ID and secret, or click **Import** and select the downloaded `client_secret.json`.
6. Run the command **Sheets Sync: Authenticate with Google** and approve access in the browser. The token is stored in the plugin's data and refreshes automatically.

No redirect URIs need to be registered — the flow uses a loopback server on `127.0.0.1` with a random port.

## Usage

### Configure a note

Add a `sheets-sync` block to the note's frontmatter:

```yaml
---
sheets-sync:
  spreadsheetId: 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms
  sheet: Sheet1
  range: A1:F50
---
```

- `spreadsheetId` — from the sheet's URL (the part between `/d/` and `/edit`)
- `sheet` — the tab title
- `range` — A1 notation

The command **Sheets Sync: Insert sheets-sync config into this note** does this for you (merging into existing frontmatter) and appends the sync markers.

### Sync

- **Sheets Sync: Pull table from Sheets (active note)**
- **Sheets Sync: Push table to Sheets (active note)**

The table lives between markers in the note:

```markdown
<!-- sheets-sync:start -->

| Task | Owner | Status |
| ---- | ----- | ------ |
| ...  |       |        |

<!-- sheets-sync:end -->
```

Everything outside the markers is untouched. If both the note and the sheet changed since the last sync, a dialog asks which side should win.

### Settings

| Setting                 | Purpose                                                           |
| ----------------------- | ----------------------------------------------------------------- |
| Client ID / secret      | Google OAuth desktop credentials (or import `client_secret.json`) |
| Default spreadsheet ID  | Prefills the template when inserting config into a note           |
| Poll interval (seconds) | If > 0, auto-pulls every configured note in the vault             |

## Development

```bash
npm run build    # typecheck + bundle to main.js
npm run format   # prettier
```

Source is `src/main.ts` (single file). `main.js` is the committed build artifact for manual installs.

## License

MIT
