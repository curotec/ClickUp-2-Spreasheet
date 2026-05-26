# ClickUp Time Entries → Google Sheets

A Google Apps Script that imports ClickUp time entries into a Google Sheet
for a chosen List and date range, with optional two-way sync for
description, tags, and billable status — gated by an explicit confirm step
so nothing leaves the sheet by accident.

## Features

- **One-click refresh** of time entries for a chosen List and date range.
- **Date presets**: current/previous month, current/previous quarter, custom.
- **Billable filter**: all / billable only / non-billable only.
- **Custom Task IDs** (e.g. `CTK-10334`) when the ClickApp is enabled,
  with fallback to internal task IDs.
- **Multi-select tag dropdown** sourced from your workspace's tag list.
- **Confirm-before-sync** for description, tags, and billable changes —
  edits are flagged but never sent to ClickUp until you tick a row and
  click "Sync pending changes".
- **Full change log** of every sync attempt (success and failure).
- **List discovery helper** to find the correct List ID by name.

## Requirements

- A Google account with access to Google Sheets.
- A ClickUp workspace and a **personal API token** (`pk_...`):
  ClickUp avatar (bottom left) → **Settings → Apps → API Token → Generate**.
- Your ClickUp **Team ID** (the number in the URL when you're in the workspace).
- The **List ID** of the List you want to track. The script's "List all Lists
  with time entries" helper makes this easy to find.

## Installation

1. Create a new Google Sheet (or open the one you want to use).
2. **Extensions → Apps Script** opens the script editor.
3. Delete the default `Code.gs` contents and paste in the full contents of
   [`Code.gs`](./Code.gs). Save (Ctrl/Cmd + S) and **reload the sheet**.
4. A new **ClickUp** menu appears in the toolbar.

## First-time setup

In order:

1. **ClickUp → Setup config sheet** — creates the `Config` sheet with all
   the settings rows. Re-running later is non-destructive: it preserves
   your existing values.
2. Fill in `Config`:
   - **API Token** — your `pk_...` token.
   - **Team ID** — the workspace ID.
   - **List ID** — leave blank for now if you don't know it.
   - **Preset** — pick from the dropdown.
   - **Custom start/end dates** — only used when Preset = Custom.
   - **Billable filter** — All / Billable only / Non-billable only.
3. **ClickUp → List all Lists with time entries** (only if you don't know
   your List ID) — populates the `Lists Found` sheet so you can locate
   "your List" and copy its ID into the Config.
4. **ClickUp → Refresh tag list** — populates the `Tags` sheet (protected
   against manual edits) and enables the multi-select dropdown on the
   Labels column for future refreshes.
5. **ClickUp → Setup two-way sync** — installs the `onEdit` trigger.
   Google will ask for additional permissions (the script needs to act on
   your behalf so it can run on edits); accept.
6. **ClickUp → Refresh time entries** — pulls the data and applies
   checkboxes, dropdowns, and column widths.

## Daily use

### Read-only view

Just refresh whenever you want fresh data:

- **ClickUp → Refresh time entries**

### Editing values back to ClickUp

The Time Entries sheet has three editable columns: **Description**,
**Labels (Tags)**, and **Billable**. Editing them does **not** write to
ClickUp on the spot — instead:

1. Edit any of those cells. The **Pending** column shows what's changed
   (`Desc`, `Tags`, `Billable`, or a combination). If you edit a value
   back to the original, Pending auto-clears.
2. Tick the **Confirm** checkbox on each row you want to push.
3. **ClickUp → Sync pending changes**. A dialog appears listing up to
   ten changes as `old → new`. Confirm to push.
4. Successful rows: snapshot refreshes, Pending clears, Confirm un-ticks,
   row flashes green briefly. The `Change Log` sheet records each
   operation.
5. Failed rows: snapshot and Pending stay so you can retry; the error
   appears in `Change Log` with `Status = Failure`.

To revert pending edits without syncing:
- **ClickUp → Discard pending changes** — confirms, then reverts all
  pending rows to their snapshot values.

If you refresh while changes are pending, a three-way dialog asks whether
to sync first, discard and refresh, or cancel.

## Sheets created by the script

| Sheet         | Purpose                                                      |
|---------------|--------------------------------------------------------------|
| `Config`      | Settings (token, IDs, preset, filters). Editable.            |
| `Time Entries`| Main data. Refreshed on demand.                              |
| `Tags`        | Workspace tag list. Protected — managed by the script.       |
| `Lists Found` | Helper output of "List all Lists with time entries".         |
| `Change Log`  | Every sync attempt (success and failure). Capped at 5000 rows. |

## Menu reference

| Menu item                            | What it does                                         |
|--------------------------------------|------------------------------------------------------|
| Refresh time entries                 | Re-fetches data for the current Config range/filters |
| Refresh tag list                     | Re-fetches workspace tags into the `Tags` sheet      |
| List all Lists with time entries     | Populates `Lists Found` for List ID discovery        |
| Sync pending changes                 | Confirms and pushes ticked Pending rows to ClickUp   |
| Discard pending changes              | Reverts ticked / Pending rows to snapshot values     |
| Setup config sheet                   | Creates / refreshes `Config` non-destructively       |
| Setup two-way sync                   | Installs the `onEdit` trigger (one-time)             |

## Known limitations

- **API token in plaintext.** The token lives in cell B2 of the `Config`
  sheet. Fine for personal use; if you share the sheet with view access,
  anyone who can see it can read the token. Moving the token to
  `PropertiesService` is straightforward if needed.
- **Multi-cell edits skipped.** Paste-over-many-cells and fill-down on the
  three editable columns are intentionally ignored to prevent accidental
  bulk writes. Edit one cell at a time when changes need to sync.
- **No real-time pull from ClickUp.** The sheet is a snapshot. If someone
  else edits a time entry in ClickUp after your last refresh, you won't
  see their change until you refresh.
- **Race conditions.** Two users editing the same row both ticking Confirm
  and running Sync results in last-write-wins on ClickUp's side.
- **Tags are diffed against the local snapshot.** If a tag is removed from
  ClickUp between refresh and sync, the DELETE will succeed silently;
  if it's renamed, the tag is treated as removed and added.
- **Custom IDs.** `task.custom_id` can be `null` for tasks created before
  the Custom Task IDs ClickApp was enabled. Those rows show the internal
  task ID instead.
- **5000-row Change Log cap.** Older rows roll off automatically. If you
  need long-term history, copy the sheet contents elsewhere periodically.

## Configuration reference

The constants at the top of `Code.gs` are designed to be tweakable
without touching the rest of the code:

```javascript
const COLUMNS = [...];          // Column headers in display order
const COLUMN_WIDTHS = [...];    // Pixel widths matching COLUMNS
const WRAP_COLUMNS = [3, 4];    // Columns that wrap long text
const CHANGE_LOG_MAX_ROWS = 5000;
```

The "magic numbers" defining column positions (e.g. `BILLABLE_COL = 8`)
are kept in one place near the top. If you change `COLUMNS`, update the
matching position constants and `COLUMN_WIDTHS` together.

## License

MIT.

## Changelog

See [CHANGELOG.md](./CHANGELOG.md).
