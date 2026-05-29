# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.8.0] — 2026-05-29

### Added
- **List ID dropdown on Config sheet.** The `List ID` cell (B4) is now a
  dropdown populated from the `Lists Found` sheet. Display labels show the
  full path: `"Support LCI (Client Engagements > LCI Paper)"`.
- New `Display Label` column (column 8) in the `Lists Found` sheet, used as
  the dropdown source.
- `resolveListId_()` helper resolves dropdown labels back to numeric List
  IDs via `Lists Found` lookup. Raw numeric IDs still work for backward
  compatibility.
- `buildListLabel_()` and `applyListIdDropdown_()` helper functions.
- **"Sync & Reload" menu item.** Syncs all pending rows (ignoring Confirm
  checkbox), skips the confirmation dialog, then automatically refreshes
  time entries. One-click push-and-pull.
- `collectChanges_(requireConfirm)` and `executeSyncChanges_(changes, sheet)`
  extracted from `syncPendingChanges` so both the dialog and silent paths
  share the same logic.

### Changed
- `listAllListsWithEntries` now also refreshes the Config B4 dropdown after
  writing to the `Lists Found` sheet.
- `readConfig` reads 9 rows (was 8) to cover the full Config spec.
- `refreshTimeEntries` accepts an optional `skipPendingCheck` parameter so
  `syncAndReload` can bypass the pending-changes dialog.
- Config notes for List ID updated to say "Run 'List all Lists' first, then
  pick from dropdown."
- Dropdown on B4 rejects manual input (`setAllowInvalid(false)`).

## [1.7.0] — 2026-05-26

### Changed
- Replaced the `Sync Errors` sheet with a unified `Change Log` sheet that
  records every sync attempt (both successes and failures).
- `Change Log` columns: Timestamp, Status, Entry ID, Task ID, Task Name,
  Field, Old value, New value, Error.
- Each PUT/POST/DELETE is logged independently, so a single confirmed row can
  produce multiple log entries (one per field or per tag added/removed).
- Retention capped at 5000 most recent data rows; oldest rows roll off
  automatically.
- "User" column width increased from 140 → 200 px.

### Removed
- `Sync Errors` sheet (replaced by `Change Log`). Existing `Sync Errors` tabs
  are no longer written to but are not deleted automatically.

## [1.6.0] — 2026-05-26

### Added
- Confirm-before-sync workflow for all editable fields (Billable,
  Description, Labels).
- New `Pending` column showing which fields have unsynced edits
  (e.g. `Desc`, `Tags`, `Billable`, or combinations).
- New `Confirm` checkbox column to mark rows for the next sync.
- Hidden `Snapshot` column storing the row's original ClickUp values as
  JSON, enabling diff display and discard-to-original.
- Menu item **Sync pending changes** — opens a confirmation dialog showing
  up to 10 changes with `old → new` previews (plus a count for the rest).
- Menu item **Discard pending changes** — reverts all pending rows to
  their snapshot values after confirmation.
- Three-way dialog when refreshing with pending changes:
  *Sync first / Refresh anyway / Cancel*.
- Auto-clear of Pending status when the user manually edits a value back
  to its original.

### Changed
- Edit handler no longer calls ClickUp directly. It now only marks rows as
  Pending, leaving the API calls for the explicit Sync step.
- Best-effort tag sync: individual tag add/remove failures do not block
  other operations on the same row.

## [1.5.0] — 2026-05-26

### Added
- New `Tags` sheet listing all workspace time-entry tags, fetched via
  `GET /team/{team_id}/time_entries/tags`.
- Menu item **Refresh tag list** populates the `Tags` sheet and protects
  it against manual edits.
- Multi-select dropdown applied to the Labels column on each refresh,
  sourced from the `Tags` sheet.
- Two-way sync for the Labels column: diffs old vs new tag sets and
  issues individual POST/DELETE calls per tag.

### Changed
- Tag display no longer applies Title Case — original ClickUp casing is
  preserved to ensure reliable round-tripping with the API.
- `setupConfigSheet` is now non-destructive: re-running preserves existing
  values and only adds missing rows / validations.

### Removed
- `toTitleCase` helper (no longer used).

## [1.4.0] — 2026-05-26

### Added
- Two-way sync for the Billable column. Toggling the checkbox issues a
  `PUT /team/{team_id}/time_entries/{id}` with the new `billable` value.
- Hidden `Entry ID` column to identify which ClickUp record to update.
- Menu item **Setup two-way sync** installs an installable `onEdit`
  trigger (required for API calls from edit events).
- `Sync Errors` sheet auto-created on first failure to record sync issues.
- Automatic revert of the checkbox if the API call fails.
- Toast notification on success and failure.

### Changed
- Multi-cell edits (paste, fill-down) are intentionally ignored to prevent
  bulk accidental writes.

## [1.3.0] — 2026-05-26

### Added
- `Billable filter` setting in Config (`All` / `Billable only` /
  `Non-billable only`). Filtering is applied client-side after fetching
  since ClickUp's API does not expose a billable filter parameter.

### Changed
- Sync-complete toast notes how many entries were kept vs filtered out
  when a filter is active.

## [1.2.0] — 2026-05-26

### Added
- Fixed pixel widths per column (configurable via `COLUMN_WIDTHS`).
- Wrapping enabled for Task Name and Description columns.

### Changed
- Replaced `autoResizeColumns` with explicit width application.
- Title Case applied to tag display (later reverted in 1.5.0).

## [1.1.0] — 2026-05-26

### Added
- Custom ID support: the Task ID column now shows the human-readable
  `task.custom_id` (e.g. `CTK-10334`) when available, falling back to
  the internal task ID when not.

### Changed
- Output column set trimmed to seven columns:
  Date, Task ID, Task Name, Description, Time (hours), User, Labels (Tags).

## [1.0.0] — 2026-05-26

### Added
- Initial release: Google Apps Script bound to a Google Sheet that
  imports ClickUp time entries.
- `Config` sheet with token, Team ID, List ID, date-range preset
  (Current month / Previous month / Current quarter / Previous quarter /
  Custom), and Include-subtasks toggle.
- Date presets resolved against calendar boundaries; custom ranges
  inclusive of the end date.
- Multi-member fetch: pulls the team roster via
  `GET /team/{team_id}` and passes all assignees to the time-entry call.
- List-scoped query: server-side filtering via `list_id` parameter.
- `Time Entries` output sheet with 16 columns (Entry ID, dates, user,
  task metadata, list/folder/space, tags, billable, description, URL).
- Menu item **List all Lists with time entries** — populates a
  `Lists Found` sheet showing every List with time entries in the
  selected date range, including name, ID, folder, space, entry count,
  and total hours. Used to discover the correct List ID.
- Custom **ClickUp** menu with refresh and setup actions.

### Fixed
- Removed an unsupported `include_subtask_time` parameter that did
  nothing on the time-entries endpoint.
- Read location info from `task_location` (the actual field returned by
  the API), so List/Folder/Space columns populate correctly.
