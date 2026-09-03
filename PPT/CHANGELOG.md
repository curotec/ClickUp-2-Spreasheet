# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

---

## [1.8.0] — 2026-09-02

### Added

- **Optional "List" column in the Report (multi-list mode).** When more than one List is selected, a **List** column is prepended as the first column of the Report, showing which List each time entry belongs to — so merged sprints are distinguishable at a glance. When a single List is selected the column is absent and the Report is laid out exactly as before.
  - The List cell shows the **List name only** (e.g. `Sprint 2`).
  - The column is **display-only**: it is not editable, not part of the Snapshot, and never participates in two-way sync (a time entry's List has no editable ClickUp field).

### Changed (internal)

- All Report column references now shift by a computed **List-column offset** (0 or 1). The write path derives the offset from Config (whether >1 List is selected); the **sync, edit-tracking, revert, and Dashboard read paths self-detect it from the Report's actual header row** (`A1 === "List"`). This means edits are read/written against the layout actually on the sheet, so a Report generated in one mode is still synced correctly even if the selection changed afterward.
- Cost and subtotal formulas, checkbox insertion, hidden columns, wrap columns, number formats, the Task Category dropdown, and row highlighting all shift with the offset. Single-list output is byte-identical to 1.7.1.

### Notes

- No Config layout change. Existing single-List reports and their two-way sync are unaffected.
- To see the List column, select two or more Lists in the `List ID` cell and run **Refresh time entries**.

---

## [1.7.1] — 2026-09-02

### Fixed

- **The "Lists Found" tab is no longer protected.** A leftover sheet/range protection (carried over from an earlier version) left the tab showing a lock icon. `sheet.clear()` does not remove protections, so the lock persisted across runs. **"List all Lists"** now strips any existing protection on the Lists Found tab before rewriting it, clearing the lock and self-healing if it ever reappears. Run **ClickUp → List all Lists** once to remove the lock.

---

## [1.7.0] — 2026-09-02

### Changed

- **"List all Lists" now enumerates every List in the workspace**, not only Lists with logged time in the selected period. The scan walks the ClickUp hierarchy (Spaces → Folders → Lists, plus folderless Lists) to build the full List set, then overlays entry counts / hours for the configured period. Lists with no entries in the period appear with a **blank count** so they can still be selected — useful for picking a sprint before any time is logged to it. (Ported from the upstream ClickUp-2-Spreadsheet list-discovery approach.)
  - The menu item **"List all Lists with time entries"** is renamed to **"List all Lists"** (and the underlying function `listAllListsWithEntries` → `listAllLists`).
  - **Archived** Spaces, Folders, and Lists are excluded from the dropdown.
- **No more "unknown" List row.** The previous scan derived Lists from time entries, so an entry with a missing task location produced an `unknown` List in the dropdown. The hierarchy walk only lists real Lists; entries whose List isn't in the hierarchy are ignored when counting, so no `unknown` row is created.

### Added

- A prominent in-code reminder above `applyListIdDropdown_` documenting that native multi-select ("chip") mode is a one-time manual UI step on the List ID cell (the Apps Script API has no `setMultiSelect()`), so it isn't accidentally re-added.
- List hierarchy helpers: `getSpaces_`, `getFolders_`, `getFolderLists_`, `getFolderlessLists_`, `getAllListsHierarchy_`.

### Upgrade note

- After installing, run **ClickUp → List all Lists** to repopulate the Lists Found tab and the List ID dropdown with the full workspace List set. (No Config layout change in this release.)

---

## [1.6.0] — 2026-09-02

### Changed

- **Multi-list selection.** The `List ID` cell can now hold several Lists at once (e.g. multiple sprints), and a single sync pulls time entries from all of them. This supersedes the `Additional List IDs` row introduced in 1.5.0.
  - Entries from all selected Lists are merged and **de-duplicated by Entry ID**. The configured **date range still applies to every List**.
  - The **first selected List** is used anywhere a single List is needed (Dashboard title, Roles refresh).
  - Parsing is **comma-safe**: the cell value is matched against the known List labels from the Lists Found sheet (longest-first), so a List label that itself contains a comma is not split incorrectly. Raw numeric IDs fall back to a plain comma split.

### Enabling multi-select (one-time UI step)

The Apps Script data-validation API cannot enable native multi-select ("chip") dropdowns from code, so the script builds a **single-select** dropdown on the `List ID` cell. To pick more than one List, turn on multi-select once in the Sheets UI:

1. Select the `List ID` cell → **Data → Data validation** (or right-click → Data validation).
2. Set **Display style** to **Chip** and enable **Allow multiple selections**.
3. Click **Done**.

From then on you can select multiple Lists as chips. The script reads the multi-selected value correctly whether or not chip mode is on — a single selection still works as a plain dropdown.

### Removed

- The **Additional List IDs** Config row from 1.5.0 is removed. The Config layout returns to the 1.4.3 row positions, so `LAST_SYNCED_ROW` and the value-dropdown cells revert to their prior rows. This also fixes the data-validation misalignment that appeared when 1.5.0's inserted row shifted the dropdowns.

### Upgrade note

- After installing, run **ClickUp → Setup config sheet** once (re-lays the Config rows, preserves values by name), then **List all Lists with time entries** to (re)apply the List ID dropdown. Optionally enable chip multi-select via the UI step above, then select your Lists.
- If you set anything in the 1.5.0 **Additional List IDs** cell, it is no longer read — select those Lists in the List ID cell instead.

---

## [1.5.0] — 2026-09-02

### Added

- **Multi-list fetch.** A new optional Config row, **Additional List IDs**, lets a single sync pull time entries from several Lists at once — e.g. selecting multiple sprints for one project. The primary **List ID** works exactly as before; any Lists named in **Additional List IDs** (comma-separated, using the same dropdown labels or raw numeric IDs) are fetched in addition to it.
  - The ClickUp `time_entries` endpoint accepts only one `list_id` per request, so each List is fetched in turn and the results are merged.
  - Merged entries are **de-duplicated by Entry ID**, so a time entry that surfaces through more than one List (overlapping tasks) is counted once. One ticket with many time entries keeps all of its rows.
  - The configured **date range still applies to every List** — results are "these Lists, within this window."
  - Two-way sync is unaffected: each entry carries its own task/entry IDs, so edits push back regardless of which List an entry came from.

### Changed

- `getTimeEntries` now accepts either a single List ID (unchanged behavior) or an array of List IDs. A single-element array is byte-identical to the old single-List path, so existing single-List configs behave exactly as before.
- The Config sheet layout gains one row (**Additional List IDs**, inserted directly under **List ID**). Internal fixed-row references were shifted accordingly (`LAST_SYNCED_ROW` 11 → 12; Config value dropdowns; config read ranges).

### Upgrade note

- After installing, run **ClickUp → Setup config sheet** once. It re-lays the Config rows and preserves your existing values by setting name. Then run **List all Lists with time entries** so the new **Additional List IDs** cell picks up the dropdown of available Lists.

---

## [1.4.3] — 2026-07-09

### Fixed

- **Rounding reconciliation for fractional hours.** Each row's Cost formula now rounds to 2 decimals (`=ROUND(IF(J,E*F,0),2)`) instead of storing full precision. Previously, fractional time entries could make the subtotal (`SUM` of full-precision costs) disagree by a cent with the sum of the displayed, penny-rounded line items. Rounding at the row level means every printed line item is the true stored value and the total always reconciles with them.
- The informational **credit value** (grayed-out, All filter) now rounds per row the same way (`SUMPRODUCT` over `ROUND(E*F,2)`), and the **Dashboard** `totalCost`/`creditCost` accumulate per-row-rounded costs, so Dashboard and Report agree to the penny.

---

## [1.4.2] — 2026-07-03

### Changed

- **Billable-only KPI strip:** instead of reflowing to three cards, the Dashboard now keeps all four card slots and simply **blanks the Credit card content** (the F:G slot renders empty). This keeps `% billable` in its original position (col I) and preserves the warm band layout, which reads better than a reflowed three-card strip. Revises the 1.4.1 behavior.

---

## [1.4.1] — 2026-07-03

### Changed

**Billable-only refinements (Report + Dashboard)**
- When `Billable filter = Billable only`, the Report subtotal block collapses to a **single row** labeled `Sub total Hours / Total Amount Due` (hours in col E, amount in col G). The `Total Hours Credit` and separate `Total Amount Due` rows are omitted — every row is billable, so they'd be redundant.
- When `Billable filter = Billable only`, the Dashboard **drops the Credit hours KPI card** and reflows the strip to three cards: Total (B:C), Billable (D:E), % billable (F:G). The col-H spacer bridge is only painted in the full four-card layout.
- When `Billable filter = Billable only` (Per Task), the "Hours by person" section drops the **Credit** and separate **Total** columns → `Person · Total` (billable hours), with a single-series chart.

**Per Role header label**
- In `Rate Mode = Per Role`, the Report's column I header now reads **`Role`** instead of `Task Category` (header text only; `COLUMNS` is not mutated, column position and behavior unchanged).

### Notes

- These conditions stack cleanly: Billable only + Per Role yields the `Role` header, the collapsed single-row subtotal, the three-card KPI strip, and the already-billable role section.
- Per Task with all/other filters and Per Role's core behavior are otherwise unchanged from 1.4.0.

---

## [1.4.0] — 2026-07-03

### Changed

**"Per Developer" rate mode reworked into "Per Role"**

The 1.3.1 per-developer feature is generalized into a role-based billing profile. Rate Mode now offers `Per Task` (default, unchanged) and `Per Role`. The two modes are fully separated — Per Task retains the original fork behavior exactly; Per Role is a distinct profile that does not affect Per Task.

- **Sheet renamed** `Developers` → **`Roles`** with a new middle column:
  `Full Name` (col A, read-only) · **`Roles`** (col B, user-edited) · `Rate ($/hr)` (col C, user-edited).
  Roles **and** Rates are preserved across refreshes. The writer tolerates an older 2-column layout and migrates it (col B treated as Rate, Role left blank).
- **Menu item renamed** `Refresh developers list` → **`Refresh roles list`**.
- **Config option relabeled** `Per Developer` → **`Per Role`**.

**Per Role mode behavior**
- **Rate** looked up from the Roles sheet by Full Name (blank/unknown = `$0`).
- **Task Category (col I)** shows the person's **Role** (blank if no role) instead of mapped tag Display Names.
- **All rows are visible** — tag mapping no longer governs row visibility in this mode.
- **Task Category is excluded from two-way sync** — a Role has no ClickUp tag equivalent, so editing col I never marks the row Pending and never pushes. Work Description and Billable still sync normally. Enforced in both the edit handler and the Pending recompute (`syncEditableCols_`, mode-guarded diff).
- The Task Category dropdown (tag Display Names) is applied only in Per Task mode; in Per Role mode any stale validation on col I is cleared.

**Dashboard (Per Role only; Per Task dashboard unchanged)**
- "Hours by person — billable vs credit" → **"Hours by role — billable"**: rows are roles, a single `Total` column = billable hours only (credit excluded), and the chart becomes single-series.
- **"Hours by task category" section hidden.**
- **Top 10 issues** drops the **Type** column (Issue · Hours only).

### Notes

- Per-task rate logic (first raw tag → Tags rate) remains distinct from upstream and unchanged.
- The KPI cards, Cost formula, and credit-value tracking are identical in both modes.

---

## [1.3.1] — 2026-07-03

### Added

**Per-developer rate mode**
- New **Developers** sheet mirroring the Tags sheet pattern: `Full Name` (col A, read-only, script-managed) · `Rate ($/hr)` (col B, user-edited). Rates are preserved across refreshes.
- New Config option **Rate Mode**: `Per Task` (default) or `Per Developer`.
  - `Per Task` — unchanged fork behavior: rate looked up from the Tags sheet by the entry's first raw ClickUp tag.
  - `Per Developer` — rate looked up from the Developers sheet by Full Name. A developer with no rate resolves to `$0` and their rows still surface.
- New menu item **Refresh developers list** (under *Refresh tag list*): standalone scan of who logged time on the selected List.
- The Developers sheet is also auto-populated on every normal sync (top-up union — never drops existing people or rates).
- **`SCRIPT_VERSION` constant** added below the header (`const SCRIPT_VERSION = '1.3.1'`) so the running version is greppable and identifiable when testing multiple copies. Standing rule: keep it in sync with the header comment + CHANGELOG on every release.

### Notes

- Only the *rate source* changes between modes. Cost formula (`=IF(J{n}, E{n}*F{n}, 0)`), credit-value tracking, and Task Category display are identical in both modes.
- **The Tags sheet still governs Report row visibility in both modes** — unmapped-tag rows stay hidden regardless of Rate Mode.
- Per-task rate logic (first raw tag → Tags rate) remains distinct from upstream and unchanged.

### Fixed

- `LAST_SYNCED_ROW` bumped 10 → 11 to track the Config row shift caused by inserting the Rate Mode row, so the sync timestamp no longer risks overwriting the Rate Mode cell. `readConfig` range widened 12 → 13 rows so the shifted Skip IDs row is still read.

---

## [1.3.0] — 2026-06-03

### Changed

**Tags sheet: Display Name mapping (ported from upstream)**
- Tags sheet now has three columns: `Tag name` (col A, read-only), `Display Name` (col B, user-edited), `Rate ($/hr)` (col C, user-edited).
- Tags without a populated Display Name are **hidden from the Report's Task Category column**. This is the opt-in mechanism — fill in a Display Name to surface a tag.
- The Task Category dropdown is built only from populated Display Names, sorted alphabetically.
- Two-way sync reverses Display Name → ClickUp tag name before calling the ClickUp tags API. Users edit friendly display names; the API still receives raw tag names.
- `refreshTagList` preserves both Display Name and Rate values across re-syncs.
- Per-task rate logic is unchanged — Rate is still looked up by the entry's first raw ClickUp tag.

**Report sheet styling**
- Font: **Anek Tamil 11pt** across header + data + subtotal rows. *(Note: Anek Tamil is a Google Font but not built-in to Sheets — add it once via Format → Font → More fonts. Falls back to default if unavailable.)*
- Header row background changed from `#4a4a8a` (purple) → `#000000` (black), text remains bold white.
- Full name column (col H) widened from 160pt → 200pt.

**Dashboard styling**
- Section heading bars (Hours by person / Hours by task category / Top 10 issues by hours) no longer have the dark navy fill — they now render as plain bold 12pt dark text on the default sheet background, blending in with their surroundings.
- Col I widened from 105pt → 125pt to give "% billable" and "Of total hours worked" more breathing room.
- Top 10 issues `Type` column (col 4) now wraps text so long comma-joined category lists fit cleanly.
- KPI strip: col H (rows 2-4) now shares the warm `#F4F2EC` background of the surrounding cards, so the spacer between "Credit hours" and "% billable" reads as one continuous band instead of a visible white gap.
- Section spacing: each section now advances the cursor past its associated chart's bottom edge (`max(cursor, chartAnchor + 14) + 2`) so charts no longer overlap with the next section's heading regardless of how few rows the table has.

---

## [1.2.0] — 2026-06-01

### Changed

**Billing: non-billable rows now show $0 cost**
- Cost formula (col G) changed from `=E*F` to `=IF(J=TRUE, E*F, 0)`. Non-billable rows always display `$0.00` in the sheet.
- Rate (col F) is still populated for non-billable rows so the Dashboard can compute the credit value independently.

**Dashboard: billable/credit classification now driven by Billable flag**
- Previously classified by `rate > 0`. Now reads the Billable checkbox (col J) directly, matching the actual ClickUp billable flag.
- Added `creditCost` — the aggregate value of non-billable hours at their tag rate. Shown in the Dashboard as **Credit value** alongside the billable amount.
- Billing summary row now has two dollar figures: **Billable amount** (invoiced) and **Credit value** (informational).

**Report subtotals block improved**
- `Total Hours Credit` row now has a live `SUMIF(J, FALSE, E)` formula for hours and a `SUMPRODUCT` formula for the credit cost value.
- `Total Amount Due` row now has a live `SUMIF(J, TRUE, E)` formula for billable hours.
- Credit cost cell is grayed out to signal it is informational and not invoiced.

---

## [1.1.0] — 2026-06-01

### Added

**Billing fork — initial release**

**Tags sheet**
- Added `Rate ($/hr)` column (col B) to the Tags sheet.
- Rate values are preserved across `refreshTagList` runs — re-syncing tags never overwrites manually entered rates.
- Tag names in col A are protected (warning-only) to prevent accidental edits; rate column remains freely editable.

**Report sheet** *(renamed from "Time Entries")*
- New column structure: `Date, Issue Key, Issue summary, Work Description, Hours, Rate, Cost, Full name, Task Category, Billable, Pending, Confirm, Entry ID, Snapshot`.
- Rate (col F) auto-filled from Tags sheet using the entry's first tag as the lookup key.
- Cost (col G) written as a live `=Hours × Rate` formula (updated to billable-conditional in v1.2.0).
- Subtotals block at the bottom: Sub total Support Hours, Total Hours Credit, Total Amount Due.

**Dashboard sheet**
- Auto-built on every `refreshTimeEntries` run; also available standalone via **Rebuild Dashboard**.
- Pinned as the first tab automatically.
- Sections: title, KPI cards (total/billable/credit hours, % billable), billing summary row, hours by person (billable vs credit), hours by task category, top 10 issues by hours.
- Matches the ParadigmPT reference layout.

**Config sheet additions**
- Two new settings: `Client Name` (used in Dashboard title) and `Month Label` (used in Dashboard header and billing summary row).

**Menu additions**
- `Rebuild Dashboard` — rebuilds Dashboard from current Report data without a ClickUp API call.

---

## [1.0.0] — upstream baseline

### Added *(upstream — all features carried forward into this fork)*

**Data import**
- Pulls time entries from ClickUp `/team/{id}/time_entries` for a configurable date range and list.
- Supports presets: Current month, Previous month, Current quarter, Previous quarter, Custom.
- Filters by billable status: All, Billable only, Non-billable only.
- Paginates assignees in chunks of 100 to stay within API limits.

**Config sheet**
- Single-sheet configuration with dropdowns for Preset, Billable filter, Include subtasks.
- `setupConfigSheet` preserves existing values when re-run.
- `List ID` cell accepts a display label selected from the Lists Found dropdown.

**List discovery**
- `listAllListsWithEntries` scans all workspace entries and writes a Lists Found sheet with list name, ID, folder, space, entry count, total hours, date range, and display label.
- Populates a dropdown on Config B4 automatically.

**Tags sheet**
- `refreshTagList` syncs all workspace time-entry tags, sorted alphabetically.
- Sheet is protected (warning-only) to prevent accidental edits.
- `applyLabelsDropdown` applies a multi-select-style validation to the Labels column on every refresh.

**Two-way sync**
- Editable columns: Work Description (D), Task Category / Labels (I), Billable (J).
- Snapshot stored as hidden JSON (col N) on import.
- `onClickUpEdit` installable trigger detects changes and marks rows as `Pending` (col K) with a diff label (`Desc`, `Tags`, `Billable`).
- `syncPendingChanges` — syncs only rows with Confirm = TRUE, with a preview dialog before sending.
- `syncAndReload` — syncs all pending rows regardless of Confirm, then refreshes.
- `discardPendingChanges` — reverts sheet edits to snapshot values without touching ClickUp.
- Tags synced via dedicated POST/DELETE endpoints (not the main entry PUT).
- `setupTwoWaySync` installs the installable trigger; removes any duplicate triggers first.

**Change Log sheet**
- Every sync operation logged: timestamp, status, entry ID, task ID, task name, field, old value, new value, error message.
- Capped at 5,000 data rows; oldest rows trimmed automatically.

**UX**
- Pending check on refresh — prompts to sync first, refresh anyway, or cancel.
- Flash-green row animation on successful sync.
- `Last synced` timestamp written to Config after every sync.
- Toast notifications throughout for progress and results.
