# ClickUp Time Entries → Google Sheets — Billing Fork

A Google Apps Script that pulls time entries from ClickUp into a Google Sheet, extended with per-tag billing rates, a live cost column, and a summary dashboard. Built as a fork of the upstream confirm-before-sync importer.

---

## Sheets overview

| Sheet | Purpose |
|---|---|
| **Dashboard** | KPI cards + hours by person + by category + top 10 issues. Auto-rebuilt on every refresh. |
| **Report** | One row per time entry. Includes Rate, Cost, and all billing columns. |
| **Tags** | Workspace tag list with editable Display Name (col B) and Rate ($/hr) (col C) columns. Tags without a Display Name are hidden from the Report. |
| **Roles** | Person list (col A, script-managed) with editable Roles (col B) and Rate ($/hr) (col C). Used when Rate Mode = `Per Role`. Roles and rates preserved across refreshes. |
| **Config** | API credentials and run settings. |
| **Lists Found** | Output of the list discovery scan. Used to populate the List ID dropdown. |
| **Change Log** | Audit trail of every sync operation (success and failure). Capped at 5,000 rows. |

---

## First-time setup

### 1. Create the Apps Script project

1. Open [script.google.com](https://script.google.com) and create a new project.
2. Rename the default file to `Code.gs` and paste the contents of `Code.gs` from this repository.
3. Save the project.

### 2. Initialize the Config sheet

Open the spreadsheet, then go to **ClickUp → Setup config sheet**. Fill in the values:

| Setting | Description |
|---|---|
| API Token | Your ClickUp personal API token (`pk_...`). Found under Profile → Apps. |
| Team ID | Your ClickUp workspace ID. Visible in the URL: `app.clickup.com/{team_id}/...` |
| List ID | Leave blank for now — populated via the list discovery step below. |
| Preset | Time range to pull. Defaults to `Previous month`. |
| Custom start date | Used only when Preset = `Custom`. Format: `YYYY-MM-DD`. |
| Custom end date | Used only when Preset = `Custom`. Inclusive. Format: `YYYY-MM-DD`. |
| Include subtasks | `Yes` or `No`. |
| Billable filter | `All`, `Billable only`, or `Non-billable only`. |
| Rate Mode | `Per Task` (rate from Tags sheet by first tag) or `Per Role` (rate + role from Roles sheet by Full Name). Defaults to `Per Task`. |
| Client Name | Shown in the Dashboard title (e.g. `Acme Corp`). |
| Month Label | Shown in the Dashboard header (e.g. `April 2026`). |

### 3. Discover your lists

Run **ClickUp → List all Lists with time entries**. This scans all entries in the configured date range and writes a `Lists Found` tab. It also populates the `List ID` dropdown in Config automatically.

Go to Config and select your list from the `List ID` dropdown.

### 4. Set up tag mappings and billing rates

Run **ClickUp → Refresh tag list**. The Tags sheet will be populated with all workspace tags in column A.

The Tags sheet has three columns:

| Col | Field | Purpose |
|---|---|---|
| A | Tag name | The raw ClickUp tag name (read-only, managed by script) |
| B | Display Name | Friendly name shown in the Report's Task Category column. **Tags with an empty Display Name are hidden entirely.** |
| C | Rate ($/hr) | Hourly rate used to compute Cost via first-tag lookup |

Fill in Display Name for every tag you want to appear in the Report, and Rate for billable tags. Both columns are preserved across `Refresh tag list` runs — re-syncing tags only updates column A, never overwrites your manual entries.

> **Tag visibility:** Leaving Display Name blank is the way to exclude noisy or internal tags from the Report. The Task Category dropdown is built from populated Display Names only.

> **Rate lookup:** Cost is calculated from the **first tag** on each time entry. The raw ClickUp tag (not the Display Name) is used for the rate lookup, so renaming via Display Name doesn't break billing.

### Rate Mode: Per Task vs Per Role

Config → **Rate Mode** selects one of two fully separate profiles. Per Task is the original behavior, untouched. Per Role changes the rate source, the Task Category, row visibility, the two-way sync, and the Dashboard.

| | `Per Task` *(default)* | `Per Role` |
|---|---|---|
| **Rate source** | Tags sheet, by the entry's first raw ClickUp tag | Roles sheet, by Full Name |
| **Task Category (col I)** | Mapped tag Display Names (header: `Task Category`) | The person's **Role** from the Roles sheet, blank if no role (header: `Role`) |
| **Row visibility** | Rows with unmapped tags are hidden | **All rows visible** (tag mapping ignored) |
| **Two-way sync of Task Category** | Yes — Display Name reverse-maps to a ClickUp tag | **No** — a Role has no ClickUp equivalent, so the column doesn't sync |
| **Dashboard: first section** | Hours by person — billable vs credit | **Hours by role — billable** (single `Total` column = billable hours) |
| **Dashboard: Hours by task category** | Shown | **Hidden** |
| **Dashboard: Top 10 issues** | Issue · Hours · Type | Issue · Hours (**Type column removed**) |

In **Per Role** mode, run **ClickUp → Refresh roles list** (or just sync — the Roles sheet is auto-populated from each sync). Fill in **Roles** (col B) and **Rate ($/hr)** (col C) per person. A person with no rate is treated as `$0`; a person with no role shows a blank Task Category. Both columns are preserved across refreshes.

What's identical across both modes: the Cost formula (`=IF(J{n}, E{n}*F{n}, 0)`), credit-value tracking, Work Description and Billable two-way sync, and the KPI cards. Only the items in the table above differ.

Because Rate (and, in Per Role mode, the Role shown in Task Category) is written into the Report at sync time, changing a rate or role on the Roles sheet requires a re-sync to take effect (same behavior as tag rates).

### 5. Install the two-way sync trigger

Run **ClickUp → Setup two-way sync**. This installs an installable `onEdit` trigger that watches for changes to the editable columns (Work Description, Task Category, Billable) and marks affected rows as `Pending`.

Google will ask for authorization on first run.

### 6. Refresh entries

Run **ClickUp → Refresh time entries**. This will:

- Pull all time entries for the configured list and date range
- Look up the rate for each entry's first tag
- Write the Report sheet
- Write `Cost = Hours × Rate` formulas (zero for non-billable rows)
- Rebuild the Dashboard

---

## How billing works

### Billable vs credit

Classification is driven entirely by the **Billable checkbox** (column J) on each row, not by whether a rate exists.

| Row | Billable checkbox | Cost shown | Dashboard classification |
|---|---|---|---|
| Has rate, Billable = ✅ | TRUE | `Hours × Rate` | Billable hours + billable cost |
| Has rate, Billable = ☐ | FALSE | `$0.00` | Credit hours + credit value (informational) |
| No rate, Billable = ☐ | FALSE | `$0.00` | Credit hours |

**Credit value** is what non-billable hours *would* have cost at the tag rate. It appears in the Dashboard for reference but is not included in the invoiced amount.

### Cost formula

Column G uses: `=ROUND(IF(J{row}=TRUE, E{row}*F{row}, 0), 2)`

Non-billable rows always display `$0` in the sheet. The rate in column F is still stored so the Dashboard can compute the credit value.

The per-row `ROUND(…, 2)` ensures the stored Cost is the true penny amount, so the subtotal always reconciles with the sum of the displayed line items — including when hours are fractional. (Without it, summing full-precision costs could differ by a cent from adding up the rounded line items by hand.)

### Report subtotals

With `Billable filter = All` (or `Non-billable only`), three rows appear below the data:

| Row | Hours | Cost |
|---|---|---|
| Sub total Support Hours | `SUM` of all hours | `SUM` of all Cost (G) |
| Total Hours Credit | `SUMIF` where Billable = FALSE | `SUMPRODUCT` of non-billable hours × rate *(grayed out)* |
| Total Amount Due | `SUMIF` where Billable = TRUE | `SUM` of Cost (G) |

With `Billable filter = Billable only`, every row is billable, so the block collapses to a **single row** — `Sub total Hours / Total Amount Due` (hours in E, amount in G). The credit row is omitted.

### Billable only: what changes

Selecting `Billable filter = Billable only` also simplifies the outputs, since credit is always zero:

- **Report subtotals** collapse to the single row described above.
- **Dashboard KPI strip** blanks the Credit hours card (the F:G slot renders empty), keeping the four-card band and `% billable` in its original position.
- **Dashboard "Hours by person"** (Per Task) drops the Credit and separate Total columns → `Person · Total` with a single-series chart. (In Per Role mode the role section is already billable-only.)

These stack with Rate Mode: Billable only + Per Role gives the `Role` header, the collapsed subtotal, the KPI strip with a blank Credit card, and the role-based first section.

---

## Two-way sync workflow

1. **Refresh** — loads data and takes a snapshot of Description, Task Category, and Billable for each row.
2. **Edit** — change any of those three fields. The `Pending` column updates automatically to show what changed (e.g. `Desc`, `Tags`, `Billable`).
3. **Confirm** — tick the `Confirm` checkbox on rows you want to push.
4. **Sync** — run **ClickUp → Sync pending changes**. A dialog summarizes what will be sent. Confirm to proceed.

Alternative flows:

- **Sync & Reload** — syncs all pending rows (ignores Confirm), then refreshes from ClickUp.
- **Discard pending changes** — reverts edits back to snapshot values without touching ClickUp.

Every sync operation is logged in the **Change Log** sheet.

---

## Menu reference

| Menu item | What it does |
|---|---|
| Refresh time entries | Pull from ClickUp, write Report, rebuild Dashboard |
| Refresh tag list | Sync tags from ClickUp; preserves existing rates |
| Refresh roles list | Scan who logged time on the selected List; preserves existing per-person Roles and rates |
| Rebuild Dashboard | Rebuild Dashboard from current Report data (no API call) |
| List all Lists with time entries | Discover List IDs and populate the Config dropdown |
| Sync pending changes | Push confirmed edits to ClickUp |
| Sync & Reload | Sync all pending rows then refresh |
| Discard pending changes | Revert edits to snapshot values |
| Setup config sheet | Create or refresh the Config sheet |
| Setup two-way sync | Install the onEdit trigger |

---

## Report column reference

| # | Column | Source | Editable |
|---|---|---|---|
| A | Date | Entry start date | — |
| B | Issue Key | Task custom ID or task ID | — |
| C | Issue summary | Task name | — |
| D | Work Description | Time entry description | ✅ syncs to ClickUp |
| E | Hours | Entry duration in hours | — |
| F | Rate | Looked up from Tags sheet by first tag | — |
| G | Cost | `=IF(Billable=TRUE, Hours×Rate, 0)` | — |
| H | Full name | ClickUp user username or email | — |
| I | Task Category | Display Names of mapped tags (unmapped tags hidden) | ✅ syncs to ClickUp |
| J | Billable | ClickUp billable flag | ✅ syncs to ClickUp |
| K | Pending | Auto-computed diff status | — |
| L | Confirm | Checkbox to approve sync | ✅ manual |
| M | Entry ID | Hidden — used for API sync | — |
| N | Snapshot | Hidden — JSON of original values | — |

---

## Dashboard layout

| Section | Content |
|---|---|
| Title | `{Client Name} — {Month Label} Support Hours Dashboard` |
| KPI row | Total hours · Billable hours · Credit hours · % billable |
| Billing row | Month label · Billable amount ($) · Credit value ($) |
| Hours by person | Billable vs Credit vs Total per team member |
| Hours by task category | All categories sorted by hours descending |
| Top 10 issues | Issue name · Hours · Type/category |

---

## Notes

- Re-running **Refresh tag list** never overwrites Display Names or Rates you've already entered — both are keyed by tag name and survive re-syncs.
- **Tags without a Display Name are hidden from the Report.** This is the opt-in / opt-out mechanism for noisy or internal tags.
- **Rebuild Dashboard** re-reads the existing Report sheet without calling the ClickUp API — useful when you've manually adjusted rates or billable flags.
- The Change Log retains the most recent 5,000 rows. Older rows are trimmed automatically.
- If you change a rate in the Tags sheet, run **Refresh time entries** to repopulate column F and recalculate costs, then **Rebuild Dashboard** if you only need the summary updated.
- The Report uses **Anek Tamil 11pt**. It's a Google Font but not built into Sheets — if text appears in a fallback font, add it once via **Format → Font → More fonts** in any sheet and it will apply going forward.
