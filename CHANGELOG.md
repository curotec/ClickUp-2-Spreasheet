# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.6.0] — 2026-10-01

### Added
- **Extra Users tab** for including time logged by people who are no longer
  active workspace members (deactivated users, guests). ClickUp's time-entries
  endpoint only returns entries for the assignee IDs it is given, and
  `/team/{id}` lists **active** members only — so before this release, time
  logged by deactivated users was silently missing from every report.
  Put their numeric ClickUp user IDs in column A of the `Extra Users` tab and
  they are merged into the assignee list.
- Applies to both **Refresh time entries** and **List all Lists**, so List
  discovery counts match what the Report will actually fetch.
- The `Extra Users` tab is **auto-created** (headers + a comment row) on the
  first Refresh / List all Lists if it doesn't exist. It is not protected —
  it is user-managed.
- Completion toast reports how many extra users were included (only when > 0).
- **Invalid ID warning:** non-numeric values in column A are skipped and listed
  (row + value, first 5) in the completion toast, which stays up longer (15s)
  when warnings are present. Refresh still completes.
- `VERSION` constant in `Code.gs`, kept in sync with the header comment.

### Notes
- Column A rules: blank → skipped; starting with `#` → comment, skipped;
  digits only → used (deduped, and deduped against active members);
  anything else → warned and skipped. Column B is a free-text note, never read.
- Verified against the live API before release: querying `time_entries` by a
  deactivated user's ID returns their entries.

### Upgrade note
- Replace `Code.gs` and reload the sheet. Run **Refresh time entries** once to
  create the `Extra Users` tab, then add the user IDs of any deactivated people
  whose time should appear, and refresh again.

## [2.5.0] — 2026-09-03

### Added
- **Conditional "List" column on the Report.** When more than one ClickUp List is
  selected in the Config `List ID` cell (via Google Sheets smart chips or a
  comma-separated value), the Report now shows a leading **`List`** column so it's
  clear which List each time entry belongs to. The column displays the List name
  (e.g. `Support LCI`).
- Multi-list selection resolver (`resolveListSelection_`): splits the `List ID` cell
  into parts, resolves each to a numeric List ID (deduped), and builds an
  ID → List-name map from the `Lists Found` sheet to populate the new column without
  extra API calls (names come from each entry's `task_location`).
- `getLayout_()` — a single source of truth for the Report's column layout. Every
  sheet-touching function (refresh, summary block, diff/pending, sync, discard,
  edit trigger, category dropdown) now derives its column indices from here.

### Changed
- The Report layout is now computed dynamically rather than from fixed column
  constants. In multi-list mode all columns shift right by one and the summary
  block's formulas retarget the Billed Hours column automatically (E → F). The
  hidden `Entry ID` / `Snapshot` columns, checkboxes, protection-free editable
  cells, and the two-way sync/diff all track the shift.

### Notes
- **Single-list Reports are unchanged** — byte-for-byte the same layout as 2.4.0
  (no `List` column). The new column appears only when 2+ Lists are selected.
- Detection of "which mode" is based on the Report header row (first cell = `List`),
  so an existing sheet's sync/diff keeps working without a re-refresh.
- Multi-list fetching itself already worked in 2.4.0 (ClickUp's time-entries endpoint
  accepts comma-separated `list_id`); this release adds the disambiguating column and
  makes the whole two-way workflow layout-aware.

## [2.4.0] — 2026-09-01

### Changed
- **"List all Lists with time entries" is now "List all Lists."** The feature previously
  inferred Lists from time entries in the selected period, so any List with no logged
  time for that range never appeared. It now walks the ClickUp hierarchy
  (Spaces → Folders → Lists, plus folderless Lists) and returns **every non-archived
  List** in the workspace, regardless of activity.
- The menu item and its function were renamed accordingly
  (`listAllListsWithEntries` → `listAllLists`); all in-code references and error
  messages were updated to match.

### Added
- Hierarchy API helpers: `getSpaces_`, `getFolders_`, `getFolderLists_`,
  `getFolderlessLists_`, and `getAllListsHierarchy_`. All pass `archived=false`.
- A second entry scan still runs after the hierarchy fetch to populate the
  **`# entries`** and **`Total hours`** columns for the selected period. Lists with no
  entries in the range show blank in those two columns (rather than `0`).
- The completion toast now reports both the total List count and how many had entries
  in the period.

### Notes
- **Archived Lists and Folders are excluded entirely** (not shown, not in the dropdown).
- Sort order is unchanged in spirit — Lists with entries appear first (by count, desc);
  the remaining zero-activity Lists follow, ordered alphabetically by display label.
- Listing all Lists makes more API calls than before (one per Space for folders and
  folderless Lists, plus the entry scan), so the operation is slightly slower on large
  workspaces.

## [2.3.4] — 2026-07-01

### Fixed
- **Rounding drift in the summary block.** The Total Support Hours formula
  is now `=ROUND(SUM(...),2)` instead of a bare `SUM(...)`. With many
  fractional time entries, floating-point accumulation could make the
  displayed total (and the Rate/Overage/Total Due figures derived from it)
  disagree with the visible sum of the Billed Hours column by a cent or a
  hundredth of an hour. Rounding the total at the source fixes every
  downstream figure, since they all reference this cell.

## [2.3.3] — 2026-07-01

### Changed
- **Rate** value cell (standard summary mode) now has a Dark Gray 1
  (`#b7b7b7`) background; number format unchanged (`$0.00`).

## [2.3.2] — 2026-07-01

### Changed
- **Column widths** adjusted: `[100, 100, 280, 400, 110, 200, 200, 80,
  130, 80, 120, 120]`.
- **Text wrap** now applied to the entire Report sheet (all cells), rather
  than only the Issue summary and Work Description columns.

### Removed
- `WRAP_COLUMNS` constant (obsolete now that wrap is applied sheet-wide).

## [2.3.1] — 2026-07-01

### Added
- `Version:` line in the `Code.gs` header comment.

### Changed
- **Overage (hrs)** value cell now uses bold white text on the Dark Gray 2
  (`#999999`) background (was default black, non-bold).

## [2.3.0] — 2026-07-01

### Added
- **Target Contract Hours** config field (Config row 12, optional). When
  set, the Report summary block switches to an **overage billing** layout;
  when blank, the standard Total Due block is used (fully backward
  compatible).
- Overage summary block (shown only when Target Contract Hours is set):
  1. **Total Support Hours for the Month** — `=SUM(...)`, `0.00`.
  2. **Target Contract Hours - {N} hrs** — target value; `{N}` is the
     whole-number target embedded in the label. Value cell `0.00`.
  3. **Overage (hrs)** — live formula `=MAX(0, Total − Target)` so hours
     under target show `0`; value cell shaded **Dark Gray 2 (`#999999`)`.
  4. **Overage (${rate}/hr)** — live formula `Overage hrs × Rate`,
     `$0.00`, styled bold black/white like the old Total Due row.

### Changed
- `readConfig` now reads 11 setting rows (was 10) to include Target
  Contract Hours, and exposes `cfg.targetHours` (`null` when blank).
- Summary-block builder branches on `cfg.targetHours`; outer border now
  spans the correct number of rows for whichever layout renders.

## [2.2.0] — 2026-06-03

### Added
- Entire Report sheet (header, data, summary block) now uses **Anek Tamil
  11pt** as the font.

### Changed
- **Summary block restyled:**
  - Outer border only (removed internal grid lines).
  - Row 1 (Total Support Hours) and Row 2 (Rate) no longer bold —
    only Total Due row is bold.
  - Rate and Total Due now display as currency (`$0.00`); Total Hours
    stays as plain `0.00` (it's hours, not money).
  - Total Due row keeps its black background + white text treatment.

## [2.1.0] — 2026-06-03

### Added
- Summary block now has a thin black border around the 3x2 cell area,
  matching the client report style.
- Total Due row in the summary block is styled black background + white
  bold text (mirrors the header row treatment).
- Billed Hours column on each data row is forced to 2-decimal display
  format (`0.00`).

### Changed
- Separation between last data row and the summary block increased from
  1 blank row to 3 blank rows for visual breathing room.

### Fixed
- Stale data validations (checkboxes, dropdowns) and borders from previous
  refreshes are now wiped from the entire sheet before re-population, so
  empty rows no longer show leftover checkboxes or category dropdowns.

## [2.0.0] — 2026-06-03

### Added
- **Tag mapping system.** The Tags sheet now has a "Display Name" column.
  Only tags with a display name appear in the dropdown and in the report.
  Unmapped tags are hidden from the cell entirely. Mappings are preserved
  across "Refresh tag list" runs.
- **Report summary block.** After the last data row: Total Support Hours
  (SUM formula), Rate (from Config), Total Due (Hours × Rate). Rebuilt on
  every refresh.
- **Rate setting** added to the Config sheet (row 11).
- **Black header with white text** on all script-managed sheets (Report,
  Lists Found, Tags, Change Log).

### Changed
- **BREAKING:** Output sheet renamed from `Time Entries` to `Report`.
  Existing `Time Entries` sheets are not migrated — delete or rename them.
- **BREAKING:** Column renames to match client report format:
  Task ID → Issue Key, Task Name → Issue summary,
  Description → Work Description, Time (hours) → Billed Hours,
  User → Full name, Labels (Tags) → Task Category.
- Tags sheet protection changed: column A (tag names) is protected;
  column B (display names) is editable.
- `LABELS_COL` renamed to `CATEGORY_COL` throughout.
- Tag sync now reverse-maps display names → ClickUp tag names before
  making API calls.
- Pending status shows "Category" instead of "Tags" for tag changes.
- `readConfig` reads 10 rows (was 9) to include Rate.
- `entryToRow` accepts a tag forward-map and uses it for display.
- `applyCategoryDropdown_` replaces `applyLabelsDropdown` — sources display
  names from the mapped Tags sheet.

### Removed
- Title Case tag display (already removed in 1.5.0; confirmed gone).
- Full sheet protection on Tags sheet (replaced with column-A-only).

## [1.8.1] — 2026-05-29

### Added
- `Lists Found` and `Change Log` sheets are now protected against manual
  edits, same as the `Tags` sheet.
- Shared `protectSheet_()` helper replaces inline protection boilerplate
  across all three protected sheets.

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
