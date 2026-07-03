# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
