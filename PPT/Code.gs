/**
 * ClickUp Time Entries → Google Sheet  |  Billing Fork
 * Version: 1.10.0  (2026-09-15)
 *
 * Fork of the upstream confirm-before-sync importer. Highlights vs upstream:
 *   - Tags sheet has three columns: Tag name (read-only) · Display Name · Rate ($/hr)
 *     Tags without a Display Name are hidden from the Report's Task Category.
 *     Both Display Name and Rate are preserved across refreshes.
 *   - Main sheet "Report" with billing columns:
 *       Date, Issue Key, Issue summary, Work Description, Hours,
 *       Rate, Cost, Full name, Task Category, Billable
 *   - Rate looked up from Tags sheet by the entry's first raw ClickUp tag
 *   - Rate Mode (Config): "Per Task" uses Tags-sheet rate by first raw tag and
 *     shows mapped tag Display Names in Task Category (Tags govern row visibility);
 *     "Per Role" uses the Roles sheet (Full Name · Roles · Rate) keyed by Full Name —
 *     rate from the Rate column, Task Category shows the person's Role, all rows are
 *     visible regardless of tag mapping, and Task Category is excluded from two-way sync.
 *   - Cost = IF(Billable, Hours × Rate, 0) — live formula
 *   - Dashboard sheet: KPI cards (warm-card aesthetic), hours-by-person,
 *     hours-by-category, top-10 issues, with horizontal bar charts beside each table
 *   - Report sheet typography: Anek Tamil 11pt, black header row
 *   - Config supports Client Name, Month Label, and Skip IDs (custom task IDs to exclude)
 *   - Multi-list fetch: the Config "List ID" cell can hold several Lists at once
 *     (e.g. multiple sprints). The script builds a single-select dropdown; enabling
 *     native multi-select ("chip") mode is a one-time manual step in the Sheets UI.
 *     Whichever way Lists are chosen, readConfig parses them, pulls entries from all,
 *     de-duplicates by Entry ID, and applies the date range to every List. The first
 *     selected List drives the Dashboard title / Roles refresh. When more than one
 *     List is selected, a display-only "List" column is prepended to the Report so
 *     each entry's List is visible; it never participates in two-way sync.
 *   - Extra Users: ClickUp only returns time entries for the assignee IDs passed, and
 *     /team/{id} lists ACTIVE members only — so time logged by deactivated/offboarded
 *     users or guests is invisible unless their IDs are supplied. Put those numeric
 *     user IDs in column A of the "Extra Users" tab (auto-created on first fetch);
 *     they are merged into the assignee list for every fetch, deduped against the
 *     active roster.
 *
 * Two-way sync (kept from upstream):
 *   - Confirm-before-sync editing of Work Description, Task Category, Billable
 *   - Snapshot diffing + Pending / Confirm columns
 *   - Change Log sheet with audit trail (capped at 5,000 rows)
 *   - Display Name ↔ ClickUp tag name reverse mapping on push
 *
 * See CHANGELOG.md for version history.
 */

// ---------- Constants ----------

const SCRIPT_VERSION  = '1.10.0';           // keep in sync with header comment + CHANGELOG on every release

const CONFIG_SHEET    = 'Config';
const DATA_SHEET      = 'Report';          // renamed from "Time Entries"
const LISTS_SHEET     = 'Lists Found';
const TAGS_SHEET      = 'Tags';
const DEVS_SHEET      = 'Roles';   // Full Name · Roles · Rate ($/hr); used when Rate Mode = Per Role
const CHANGE_LOG_SHEET = 'Change Log';
const DASHBOARD_SHEET = 'Dashboard';
const EXTRA_USERS_SHEET = 'Extra Users';   // User ID · Name/Note — deactivated/guest assignees to also fetch
const CHANGE_LOG_MAX_ROWS = 5000;
const CLICKUP_BASE    = 'https://api.clickup.com/api/v2';

const PRESETS = ['Current month', 'Previous month', 'Current quarter', 'Previous quarter', 'Custom'];
const BILLABLE_FILTERS = ['All', 'Billable only', 'Non-billable only'];
const RATE_MODES = ['Per Task', 'Per Role'];

// ── Report columns ────────────────────────────────────────────
const COLUMNS = [
  'Date',              // 1
  'Issue Key',         // 2
  'Issue summary',     // 3
  'Work Description',  // 4  editable, syncable
  'Hours',             // 5
  'Rate',              // 6  auto-filled from Tags sheet
  'Cost',              // 7  formula: =Hours*Rate
  'Full name',         // 8
  'Task Category',     // 9
  'Billable',          // 10 editable, syncable
  'Pending',           // 11 read-only status
  'Confirm',           // 12 checkbox
  'Entry ID',          // 13 hidden
  'Snapshot',          // 14 hidden, JSON
];

const COLUMN_WIDTHS   = [100, 110, 260, 380, 80, 70, 90, 200, 180, 80, 110, 80, 120, 120];
const WRAP_COLUMNS    = [3, 4]; // Issue summary, Work Description

const DESCRIPTION_COL = 4;
const RATE_COL        = 6;
const COST_COL        = 7;
const FULLNAME_COL    = 8;
const BILLABLE_COL    = 10;
const PENDING_COL     = 11;
const CONFIRM_COL     = 12;
const ENTRY_ID_COL    = 13;
const SNAPSHOT_COL    = 14;
const LABELS_COL      = 9;   // Task Category doubles as Labels for sync purposes

const EDITABLE_COLS   = [DESCRIPTION_COL, LABELS_COL, BILLABLE_COL];

// ---------- Optional leading "List" column (multi-list mode) ----------
//
// When more than one List is selected, a "List" column is prepended as column 1 of
// the Report so each row shows which List its time entry belongs to. When a single
// List is selected the column is absent and the Report layout is byte-identical to
// prior versions.
//
// Because the column shifts EVERY other column by +1, all column references below are
// treated as BASE (single-list) indices and shifted through col_() / colLetter_() by
// an offset of 0 or 1. The write path knows the offset from Config; the sync / onEdit
// / recompute paths SELF-DETECT it from the Report's actual header row (A1 === 'List'),
// so edits stay correct even if the layout changed between a sync and the next run.
const LIST_COL_HEADER = 'List';

/** Shift a BASE (single-list) column index by the given offset (0 or 1). */
function col_(baseIndex, offset) { return baseIndex + (offset || 0); }

/** Convert a 1-based column index to its A1 letter(s), e.g. 1->A, 27->AA. */
function colLetter_(index) {
  var s = '';
  while (index > 0) { var m = (index - 1) % 26; s = String.fromCharCode(65 + m) + s; index = (index - m - 1) / 26; }
  return s;
}

/**
 * Detect whether the Report currently has the leading "List" column, by reading its
 * header row A1. Returns 1 if present, 0 otherwise. Used by all read/edit paths so
 * they operate on the layout actually on the sheet (not on current Config).
 */
function reportListOffset_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return 0;
  var a1 = String(sheet.getRange(1, 1).getValue() || '').trim();
  return (a1 === LIST_COL_HEADER) ? 1 : 0;
}

/**
 * Editable columns that actually participate in two-way sync for the given mode.
 * In Per Role mode the Task Category (LABELS_COL) holds a Role, which has no
 * ClickUp equivalent — so it is excluded from sync entirely (no Pending, no push).
 */
function syncEditableCols_(rateMode) {
  if (rateMode === 'Per Role') return [DESCRIPTION_COL, BILLABLE_COL];
  return EDITABLE_COLS;
}

/**
 * Read Rate Mode without the full readConfig validation (safe inside onEdit,
 * where token/team may be blank). Falls back to 'Per Task'.
 */
function readRateModeSafe_() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
    if (!sheet) return 'Per Task';
    var values = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 2).getValues();
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0] || '').trim() === 'Rate Mode') {
        var m = String(values[i][1] || 'Per Task').trim();
        return RATE_MODES.indexOf(m) === -1 ? 'Per Task' : m;
      }
    }
  } catch (err) {}
  return 'Per Task';
}

const LAST_SYNCED_ROW = 11;  // spreadsheet row of "Last synced" value cell (shifted +1 by Rate Mode row in v1.3.1)

// Report sheet typography (Anek Tamil 11pt — may require adding the font via
// Format → Font → More fonts the first time. Falls back to default if unavailable)
const REPORT_FONT      = 'Anek Tamil';
const REPORT_FONT_SIZE = 11;

// ---------- Menu ----------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ClickUp')
    .addItem('Refresh time entries',            'refreshTimeEntries')
    .addItem('Refresh tag list',                'refreshTagList')
    .addItem('Refresh roles list',              'refreshRolesList')
    .addItem('Rebuild Dashboard',               'rebuildDashboard')
    .addItem('Toggle Dashboard build (on/off)', 'toggleBuildDashboard')
    .addItem('List all Lists',                  'listAllLists')
    .addSeparator()
    .addItem('Sync pending changes',  'syncPendingChanges')
    .addItem('Sync & Reload',         'syncAndReload')
    .addItem('Discard pending changes','discardPendingChanges')
    .addSeparator()
    .addItem('Setup config sheet',    'setupConfigSheet')
    .addItem('Setup two-way sync',    'setupTwoWaySync')
    .addToUi();
}

// ---------- Config ----------

/**
 * Toggle the Config "Build Dashboard" value between Yes and No. The Config cell is
 * the single source of truth read by refreshTimeEntries; this menu item just flips
 * it (and creates the row if an older Config predates it).
 */
function toggleBuildDashboard() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG_SHEET);
  if (!sheet) { SpreadsheetApp.getUi().alert('Run "Setup config sheet" first.'); return; }
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(1, 1, lastRow, 2).getValues();
  var rowNum = -1;
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0] || '').trim() === 'Build Dashboard') { rowNum = i + 1; break; }
  }
  if (rowNum === -1) {
    // Older Config without the row — append it.
    rowNum = lastRow + 1;
    sheet.getRange(rowNum, 1, 1, 3).setValues([['Build Dashboard', 'Yes',
      'Yes / No — build/update the Dashboard tab on refresh (manual "Rebuild Dashboard" always works)']]);
    sheet.getRange('B' + rowNum).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build());
  }
  var current = String(sheet.getRange(rowNum, 2).getValue() || 'Yes').trim().toLowerCase();
  var next = (current === 'no') ? 'Yes' : 'No';
  sheet.getRange(rowNum, 2).setValue(next);
  SpreadsheetApp.getActive().toast('Build Dashboard set to ' + next + '.', 'ClickUp', 4);
}
function setupConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_SHEET);
  const isNew = !sheet;
  if (isNew) sheet = ss.insertSheet(CONFIG_SHEET);

  const SPEC = [
    ['Setting', 'Value', 'Notes'],
    ['API Token',         '', 'Your ClickUp personal API token (pk_...)'],
    ['Team ID',           '', 'Workspace ID from app.clickup.com/{team_id}/...'],
    ['List ID',           '', 'Run "List all Lists" first, then pick one or more from the dropdown'],
    ['Preset',            'Previous month', 'Pick from dropdown'],
    ['Custom start date', '', 'Only used if Preset = Custom (YYYY-MM-DD)'],
    ['Custom end date',   '', 'Only used if Preset = Custom (YYYY-MM-DD), inclusive'],
    ['Include subtasks',  'Yes', 'Yes / No'],
    ['Billable filter',   'All', 'All / Billable only / Non-billable only'],
    ['Rate Mode',         'Per Task', 'Per Task (rate from Tags) / Per Role (rate + role from Roles sheet)'],
    ['Last synced',       '', 'Auto-updated after a successful sync'],
    ['Client Name',       '', 'Appears in Dashboard title'],
    ['Month Label',       '', 'e.g. April 2026 — appears in Dashboard'],
    ['Skip IDs',          '', 'Comma-separated custom task IDs to exclude from Report and Dashboard'],
    ['Build Dashboard',   'Yes', 'Yes / No — build/update the Dashboard tab on refresh (manual "Rebuild Dashboard" always works)'],
  ];

  const existing = {};
  if (!isNew && sheet.getLastRow() > 0) {
    const data = sheet.getRange(1, 1, sheet.getLastRow(), 3).getValues();
    data.forEach(r => { const k = String(r[0] || '').trim(); if (k) existing[k] = { value: r[1] }; });
  }

  const merged = SPEC.map((row, i) => {
    if (i === 0) return row;
    const key = row[0];
    return existing[key] !== undefined ? [key, existing[key].value, row[2]] : row;
  });

  sheet.clear();
  sheet.getRange(1, 1, merged.length, 3).setValues(merged);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  sheet.setColumnWidth(1, 160); sheet.setColumnWidth(2, 240); sheet.setColumnWidth(3, 380);

  // Apply dropdowns by SETTING NAME, not by hardcoded row. merged may not be in a
  // fixed row order (older sheets, future appends), so we look up each setting's row
  // from what we just wrote and place its validation on the matching B-cell. This
  // keeps every dropdown next to its label regardless of layout.
  var rowByName = {};
  merged.forEach(function(r, i){ var k = String(r[0] || '').trim(); if (k) rowByName[k] = i + 1; });
  function applyDropdown_(settingName, options) {
    var rn = rowByName[settingName];
    if (!rn) return;
    sheet.getRange(rn, 2).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(options, true).build()
    );
  }
  applyDropdown_('Preset',          PRESETS);
  applyDropdown_('Include subtasks', ['Yes', 'No']);
  applyDropdown_('Billable filter', BILLABLE_FILTERS);
  applyDropdown_('Rate Mode',       RATE_MODES);
  applyDropdown_('Build Dashboard', ['Yes', 'No']);

  // Ensure the Extra Users tab exists (for offboarded/deactivated/guest assignee IDs).
  ensureExtraUsersSheet_();

  const preserved = Object.keys(existing).filter(k => k !== 'Setting').length;
  SpreadsheetApp.getActive().toast(
    isNew ? 'Config sheet created. Fill in the values.' : 'Config refreshed. Preserved ' + preserved + ' existing value(s).',
    'ClickUp'
  );
}

function readConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  if (!sheet) throw new Error('No Config sheet. Run "Setup config sheet" first.');
  const values = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 2).getValues();
  const map = {};
  values.forEach(([k, v]) => { map[k] = v; });

  var rawListId = String(map['List ID'] || '').trim();

  // The "List ID" cell is a native multi-select (chip) dropdown. When several
  // Lists are chosen, Sheets stores them as a comma-joined string of labels.
  // Split that value into individual List selections (comma-safe against labels
  // that may themselves contain commas), resolve each to a numeric List ID, and
  // de-duplicate while preserving selection order.
  var selectedLabels = splitMultiSelect_(rawListId);
  var listIds = [];
  selectedLabels.forEach(function(label) {
    var resolved = String(resolveListId_(label));
    if (resolved && listIds.indexOf(resolved) === -1) listIds.push(resolved);
  });
  // First selected List drives anywhere a single List is needed (Dashboard title,
  // Roles refresh). listLabel keeps the first label for display.
  var primaryListId    = listIds.length > 0 ? listIds[0] : '';
  var primaryListLabel = selectedLabels.length > 0 ? selectedLabels[0] : rawListId;

  // Parse Skip IDs: comma-separated, trim whitespace, lowercase for case-insensitive match
  var skipIdsRaw = String(map['Skip IDs'] || '').trim();
  var skipIds = skipIdsRaw
    ? skipIdsRaw.split(',').map(function(s){ return s.trim().toLowerCase(); }).filter(Boolean)
    : [];

  const cfg = {
    token:           String(map['API Token'] || '').trim(),
    teamId:          String(map['Team ID'] || '').trim(),
    listId:          primaryListId,
    listIds:         listIds,
    listLabel:       primaryListLabel,
    preset:          String(map['Preset'] || '').trim(),
    customStart:     map['Custom start date'],
    customEnd:       map['Custom end date'],
    includeSubtasks: String(map['Include subtasks'] || 'Yes').trim().toLowerCase() === 'yes',
    billableFilter:  String(map['Billable filter'] || 'All').trim(),
    rateMode:        (function(){ var m = String(map['Rate Mode'] || 'Per Task').trim(); return RATE_MODES.indexOf(m) === -1 ? 'Per Task' : m; })(),
    clientName:      String(map['Client Name'] || 'Client').trim(),
    monthLabel:      String(map['Month Label'] || '').trim(),
    skipIds:         skipIds,
    buildDashboard:  String(map['Build Dashboard'] || 'Yes').trim().toLowerCase() !== 'no',
  };

  if (!cfg.token)  throw new Error('Missing API Token in Config.');
  if (!cfg.teamId) throw new Error('Missing Team ID in Config.');
  if (!PRESETS.includes(cfg.preset)) throw new Error('Preset must be one of: ' + PRESETS.join(', '));
  if (!BILLABLE_FILTERS.includes(cfg.billableFilter)) throw new Error('Billable filter must be one of: ' + BILLABLE_FILTERS.join(', '));
  return cfg;
}

function resolveListId_(raw) {
  if (!raw) return '';
  if (/^\d+$/.test(raw)) return raw;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var listsSheet = ss.getSheetByName(LISTS_SHEET);
  if (!listsSheet || listsSheet.getLastRow() < 2)
    throw new Error('Lists Found sheet is empty. Run "List all Lists" first.');
  var lastRow = listsSheet.getLastRow();
  var data = listsSheet.getRange(2, 1, lastRow - 1, 8).getValues();
  for (var i = 0; i < data.length; i++) {
    var label = String(data[i][7] || '');
    if (label === raw) return String(data[i][1]);
  }
  throw new Error('Could not find a matching List for "' + raw + '" in the Lists Found sheet.');
}

/**
 * Split a native multi-select cell value into individual List selections.
 *
 * Google Sheets stores a multi-select chip cell as its chosen values joined by
 * ", ". List labels can themselves contain commas, so a naive split would break
 * them apart. To stay comma-safe we match the cell value against the set of known
 * List labels from the Lists Found sheet (longest first), peeling off each label
 * it starts with. Anything left that isn't a known label (e.g. a raw numeric ID,
 * or a single un-listed label) is comma-split as a fallback.
 *
 * Returns an array of trimmed selection tokens (labels or IDs), in order, no blanks.
 */
function splitMultiSelect_(raw) {
  raw = String(raw || '').trim();
  if (!raw) return [];

  var labels = getKnownListLabels_();               // known labels, longest first
  var out = [];
  var rest = raw;

  while (rest.length > 0) {
    // Trim any leading separator / whitespace between chips.
    rest = rest.replace(/^[\s,]+/, '');
    if (!rest.length) break;

    var matched = null;
    for (var i = 0; i < labels.length; i++) {
      if (labels[i] && rest.indexOf(labels[i]) === 0) { matched = labels[i]; break; }
    }
    if (matched) {
      out.push(matched);
      rest = rest.slice(matched.length);
    } else {
      // No known label matches here — fall back to comma-splitting the remainder.
      rest.split(',').forEach(function(tok) {
        var t = tok.trim();
        if (t) out.push(t);
      });
      break;
    }
  }
  return out;
}

/** Labels from the Lists Found sheet (Display Label col), sorted longest-first. */
function getKnownListLabels_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var listsSheet = ss.getSheetByName(LISTS_SHEET);
  if (!listsSheet || listsSheet.getLastRow() < 2) return [];
  var lastRow = listsSheet.getLastRow();
  var data = listsSheet.getRange(2, 1, lastRow - 1, 8).getValues();
  var labels = [];
  for (var i = 0; i < data.length; i++) {
    var label = String(data[i][7] || '').trim();
    if (label) labels.push(label);
  }
  labels.sort(function(a, b){ return b.length - a.length; });  // longest first
  return labels;
}

// ---------- Date range ----------

function resolveDateRange(cfg) {
  const tz  = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const now = new Date();
  let start, end;
  switch (cfg.preset) {
    case 'Current month':
      start = new Date(now.getFullYear(), now.getMonth(), 1);
      end   = new Date(now.getFullYear(), now.getMonth() + 1, 1);
      break;
    case 'Previous month':
      start = new Date(now.getFullYear(), now.getMonth() - 1, 1);
      end   = new Date(now.getFullYear(), now.getMonth(), 1);
      break;
    case 'Current quarter': {
      const q = Math.floor(now.getMonth() / 3);
      start = new Date(now.getFullYear(), q * 3, 1);
      end   = new Date(now.getFullYear(), q * 3 + 3, 1);
      break;
    }
    case 'Previous quarter': {
      const q    = Math.floor(now.getMonth() / 3) - 1;
      const year = q < 0 ? now.getFullYear() - 1 : now.getFullYear();
      const qIdx = (q + 4) % 4;
      start = new Date(year, qIdx * 3, 1);
      end   = new Date(year, qIdx * 3 + 3, 1);
      break;
    }
    case 'Custom':
      if (!cfg.customStart || !cfg.customEnd) throw new Error('Custom preset requires start and end dates.');
      start = cfg.customStart instanceof Date ? cfg.customStart : new Date(cfg.customStart);
      const incEnd = cfg.customEnd instanceof Date ? cfg.customEnd : new Date(cfg.customEnd);
      end = new Date(incEnd.getFullYear(), incEnd.getMonth(), incEnd.getDate() + 1);
      break;
  }
  return {
    startMs: start.getTime(),
    endMs:   end.getTime() - 1,
    label:   Utilities.formatDate(start, tz, 'yyyy-MM-dd') + ' → ' + Utilities.formatDate(new Date(end.getTime() - 1), tz, 'yyyy-MM-dd'),
  };
}

// ---------- ClickUp API ----------

function cuFetch(path, token, params) {
  var qs = '';
  if (params) {
    var parts = [];
    Object.keys(params).forEach(function(k) {
      if (params[k] !== undefined && params[k] !== null && params[k] !== '')
        parts.push(encodeURIComponent(k) + '=' + encodeURIComponent(params[k]));
    });
    if (parts.length) qs = '?' + parts.join('&');
  }
  return cuRequest('get', path + qs, token);
}

function cuPut(path, token, payload)    { return cuRequest('put',    path, token, payload); }
function cuPost(path, token, payload)   { return cuRequest('post',   path, token, payload); }
function cuDelete(path, token, payload) { return cuRequest('delete', path, token, payload); }

function cuRequest(method, path, token, payload) {
  var opts = { method: method, headers: { Authorization: token }, muteHttpExceptions: true };
  if (payload !== undefined) { opts.contentType = 'application/json'; opts.payload = JSON.stringify(payload); }
  var res  = UrlFetchApp.fetch(CLICKUP_BASE + path, opts);
  var code = res.getResponseCode();
  var body = res.getContentText();
  if (code < 200 || code >= 300)
    throw new Error('ClickUp API ' + code + ' on ' + method.toUpperCase() + ' ' + path + ': ' + body);
  return body ? JSON.parse(body) : {};
}

function getTeamMemberIds(token, teamId) {
  var data    = cuFetch('/team/' + teamId, token);
  var members = (data.team && data.team.members) || [];
  return members.map(function(m){ return m.user && m.user.id; }).filter(Boolean);
}

// ---------- Extra Users (deactivated / guest / offboarded assignees) ----------
//
// ClickUp's /team/{id} returns ACTIVE members only, and the time_entries endpoint
// only returns entries for the assignee IDs you pass. So time logged by someone who
// has since been offboarded/deactivated (or a guest not on the roster) is invisible
// unless their numeric user ID is supplied explicitly. The "Extra Users" tab holds
// those IDs; they are merged into the assignee list for every fetch.

/**
 * Ensure the "Extra Users" tab exists. Created with headers + one comment row
 * (col A starting with "#" is treated as a comment). Not protected: user-managed.
 * Returns { sheet, created }.
 */
function ensureExtraUsersSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(EXTRA_USERS_SHEET);
  if (sheet) return { sheet: sheet, created: false };
  sheet = ss.insertSheet(EXTRA_USERS_SHEET);
  sheet.getRange(1, 1, 1, 2).setValues([['User ID', 'Name / Note']])
    .setFontWeight('bold').setBackground('#000000').setFontColor('#ffffff');
  sheet.getRange(2, 1, 1, 2).setValues([[
    '# Add numeric ClickUp user IDs below (one per row)',
    'Column B is a free-text note; only column A is read',
  ]]).setFontColor('#888888');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 320);
  sheet.setColumnWidth(2, 360);
  return { sheet: sheet, created: true };
}

/**
 * Parse raw column-A values into valid / invalid IDs. Pure (no Sheets calls)
 * so it can be unit-tested.
 *   - blank -> skipped silently
 *   - starts with "#" -> comment, skipped silently
 *   - digits only -> valid (deduped, order preserved)
 *   - anything else -> invalid, reported with its sheet row number
 * values: array of cell values from row `firstRow` downward.
 */
function parseExtraUserIds_(values, firstRow) {
  var ids = [], seen = {}, invalid = [];
  (values || []).forEach(function(v, i) {
    var raw = (v == null) ? '' : String(v).trim();
    if (raw === '' || raw.charAt(0) === '#') return;
    if (/^\d+$/.test(raw)) {
      if (!seen[raw]) { seen[raw] = true; ids.push(raw); }
    } else {
      invalid.push({ row: firstRow + i, value: raw });
    }
  });
  return { ids: ids, invalid: invalid };
}

/**
 * Read the Extra Users tab (creating it if absent).
 * Returns { ids: [string], invalid: [{row, value}], created: bool }.
 */
function getExtraUserIds_() {
  var res = ensureExtraUsersSheet_();
  var sheet = res.sheet;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ids: [], invalid: [], created: res.created };
  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(function(r){ return r[0]; });
  var parsed = parseExtraUserIds_(values, 2);
  return { ids: parsed.ids, invalid: parsed.invalid, created: res.created };
}

/**
 * Merge active member IDs with extra IDs, deduped (string compare).
 * Returns { ids, addedCount } where addedCount = extras not already active.
 */
function mergeAssigneeIds_(activeIds, extraIds) {
  var seen = {}, out = [];
  (activeIds || []).forEach(function(id){ var k = String(id); if (!seen[k]) { seen[k] = true; out.push(k); } });
  var added = 0;
  (extraIds || []).forEach(function(id){ var k = String(id); if (!seen[k]) { seen[k] = true; out.push(k); added++; } });
  return { ids: out, addedCount: added };
}

/**
 * One-line toast suffix describing Extra Users effects ('' when nothing to say).
 */
function extraUsersNote_(extra, addedCount) {
  var parts = [];
  if (extra.created) parts.push('"' + EXTRA_USERS_SHEET + '" tab created');
  if (addedCount > 0) parts.push(addedCount + ' extra user(s) included');
  if (extra.invalid.length > 0) {
    var shown = extra.invalid.slice(0, 5).map(function(x){ return 'row ' + x.row + ' "' + truncate_(x.value, 20) + '"'; });
    parts.push('\u26A0 ' + extra.invalid.length + ' invalid Extra Users ID(s) skipped: ' + shown.join(', ') +
               (extra.invalid.length > 5 ? ', \u2026' : ''));
  }
  return parts.length ? ' ' + parts.join('. ') + '.' : '';
}

function getAllWorkspaceTags(token, teamId) {
  var data = cuFetch('/team/' + teamId + '/time_entries/tags', token);
  return Array.isArray(data.data) ? data.data : [];
}

// ---------- List hierarchy (Spaces -> Folders -> Lists) ----------

function getSpaces_(token, teamId) {
  var data = cuFetch('/team/' + teamId + '/space', token, { archived: 'false' });
  return Array.isArray(data.spaces) ? data.spaces : [];
}

function getFolders_(token, spaceId) {
  var data = cuFetch('/space/' + spaceId + '/folder', token, { archived: 'false' });
  return Array.isArray(data.folders) ? data.folders : [];
}

function getFolderLists_(token, folderId) {
  var data = cuFetch('/folder/' + folderId + '/list', token, { archived: 'false' });
  return Array.isArray(data.lists) ? data.lists : [];
}

function getFolderlessLists_(token, spaceId) {
  var data = cuFetch('/space/' + spaceId + '/list', token, { archived: 'false' });
  return Array.isArray(data.lists) ? data.lists : [];
}

/**
 * Walk Spaces -> Folders -> Lists and Spaces -> folderless Lists.
 * Returns { id, name, folder, space } for every non-archived List in the workspace,
 * regardless of whether it has any logged time. Archived Spaces/Folders/Lists are
 * excluded (the API is asked for archived:false, and any archived flag is filtered).
 */
function getAllListsHierarchy_(token, teamId) {
  var out = [];
  var spaces = getSpaces_(token, teamId);
  spaces.forEach(function(sp) {
    var spaceName = sp.name || '';
    var folders = getFolders_(token, sp.id);
    folders.forEach(function(fo) {
      if (fo.archived) return;
      var folderName = fo.name || '';
      // Folder payloads usually embed their lists; fall back to a fetch if absent.
      var lists = Array.isArray(fo.lists) ? fo.lists.filter(function(l){ return !l.archived; })
                                          : getFolderLists_(token, fo.id);
      lists.forEach(function(l) {
        if (l.archived) return;
        out.push({ id: String(l.id), name: l.name || '(unnamed)', folder: folderName, space: spaceName });
      });
    });
    var folderless = getFolderlessLists_(token, sp.id);
    folderless.forEach(function(l) {
      if (l.archived) return;
      out.push({ id: String(l.id), name: l.name || '(unnamed)', folder: '', space: spaceName });
    });
  });
  return out;
}

/**
 * Fetch time entries for one or more Lists.
 * `listSpec` may be:
 *   - a single List ID (string/number)  -> one List (or all Lists if falsy)
 *   - an array of List IDs               -> fetched sequentially and merged
 * The ClickUp time_entries endpoint accepts only one list_id per request, so
 * multiple Lists are looped. Entries are de-duplicated by Entry ID (a single
 * time entry can surface through more than one List query when tasks overlap).
 */
function getTimeEntries(token, teamId, listSpec, startMs, endMs, assigneeIds) {
  // Normalise listSpec into an array of List targets. A falsy value means
  // "all Lists" (no list_id filter) — represented as a single null target.
  var listTargets;
  if (Array.isArray(listSpec)) {
    listTargets = listSpec.length > 0 ? listSpec : [null];
  } else {
    listTargets = [listSpec || null];
  }

  var chunkSize = 100;
  var seen = {};      // Entry ID -> true, for de-duplication across Lists
  var all  = [];

  listTargets.forEach(function(listId) {
    for (var i = 0; i < assigneeIds.length; i += chunkSize) {
      var chunk  = assigneeIds.slice(i, i + chunkSize);
      var params = {
        start_date: startMs, end_date: endMs,
        assignee: chunk.join(','),
        include_task_tags: 'true', include_location_names: 'true',
      };
      if (listId) params.list_id = listId;
      var data = cuFetch('/team/' + teamId + '/time_entries', token, params);
      if (data.data) {
        data.data.forEach(function(entry) {
          var id = entry && entry.id != null ? String(entry.id) : '';
          if (id && seen[id]) return;   // skip true duplicate across Lists
          if (id) seen[id] = true;
          all.push(entry);
        });
      }
    }
  });
  return all;
}

// ---------- Tags sheet with Display Name + Rate columns ----------

/**
 * Sync workspace tags. Tags sheet has three columns:
 *   A: Tag name (from ClickUp, read-only)
 *   B: Display Name (user-edited — empty means tag is hidden from Report)
 *   C: Rate ($/hr) (user-edited — drives per-entry Cost via first-tag lookup)
 * Both Display Name and Rate are preserved across refreshes.
 */
function refreshTagList() {
  var cfg  = readConfig();
  SpreadsheetApp.getActive().toast('Fetching workspace tags...', 'ClickUp');
  var tags = getAllWorkspaceTags(cfg.token, cfg.teamId);

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TAGS_SHEET);
  if (!sheet) sheet = ss.insertSheet(TAGS_SHEET);

  // Preserve existing Display Name AND Rate values
  var existingDisplay = {};
  var existingRates   = {};
  if (sheet.getLastRow() > 1) {
    var existing = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    existing.forEach(function(r) {
      var tagName = String(r[0] || '').trim();
      if (!tagName) return;
      var displayName = String(r[1] || '').trim();
      var rate        = r[2];
      if (displayName) existingDisplay[tagName] = displayName;
      if (rate !== '' && rate !== null && rate !== undefined) existingRates[tagName] = rate;
    });
  }

  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p){ p.remove(); });
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p){ p.remove(); });
  sheet.clear();
  sheet.getRange(1, 1, 1, 3)
    .setValues([['Tag name', 'Display Name', 'Rate ($/hr)']])
    .setFontWeight('bold')
    .setBackground('#000000').setFontColor('#ffffff');

  var names = tags.map(function(t){ return t.name; })
    .filter(function(n){ return n && n.length > 0; })
    .sort(function(a, b){ return a.toLowerCase().localeCompare(b.toLowerCase()); });

  if (names.length > 0) {
    var rows = names.map(function(n) {
      return [n, existingDisplay[n] || '', existingRates[n] !== undefined ? existingRates[n] : ''];
    });
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    sheet.getRange(2, 3, rows.length, 1).setNumberFormat('$#,##0.00');
  }

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 240);
  sheet.setColumnWidth(2, 240);
  sheet.setColumnWidth(3, 120);

  // Protect column A (Tag name) only — leave B and C editable
  var colARange = sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 1), 1);
  var protection = colARange.protect().setDescription('Tag names — managed by script');
  protection.setWarningOnly(false);
  var me = Session.getEffectiveUser();
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors().filter(function(u){ return u.getEmail() !== me.getEmail(); }));
  if (protection.canDomainEdit()) protection.setDomainEdit(false);

  var mappedCount = Object.keys(existingDisplay).length;
  var ratedCount  = Object.keys(existingRates).length;
  SpreadsheetApp.getActive().toast(
    'Loaded ' + names.length + ' tag(s). Preserved ' + mappedCount + ' mapping(s) and ' + ratedCount + ' rate(s). ' +
    'Tags without a Display Name will be hidden from the Report.',
    'ClickUp', 10
  );
}

/**
 * Returns { tagNameLower: rate } for first-tag rate lookup.
 * Keyed by raw ClickUp tag name (lowercased) — rate logic is unchanged.
 */
function getTagRateMap_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TAGS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return {};
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  var map  = {};
  data.forEach(function(r){ if (r[0]) map[String(r[0]).toLowerCase().trim()] = Number(r[2]) || 0; });
  return map;
}

/**
 * Returns { forward, reverse } tag mapping:
 *   forward: ClickUp tag name → Display Name
 *   reverse: Display Name      → ClickUp tag name
 * Only tags with a populated Display Name are included in either map.
 */
function getTagMaps_() {
  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = ss.getSheetByName(TAGS_SHEET);
  var forward = {};
  var reverse = {};
  if (!sheet || sheet.getLastRow() < 2) return { forward: forward, reverse: reverse };
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  data.forEach(function(r) {
    var clickupName = String(r[0] || '').trim();
    var displayName = String(r[1] || '').trim();
    if (clickupName && displayName) {
      forward[clickupName] = displayName;
      reverse[displayName] = clickupName;
    }
  });
  return { forward: forward, reverse: reverse };
}

/**
 * Convert ClickUp tag array → display string. Unmapped tags are hidden.
 */
/**
 * Build the Task Category string for an entry.
 * Shows EVERY tag on the entry so multi-tag entries are visible: a tag with a
 * Display Name in the Tags sheet is shown as that Display Name; a tag without one
 * is shown as its raw ClickUp name. Comma-joined, order preserved.
 *
 * (Previously only mapped tags were shown, which hid extra unmapped tags — that
 * concealment is what made a multi-tag entry look single-tagged while its rate,
 * keyed off the first raw tag, silently resolved to 0.)
 */
function mapTagsForDisplay_(clickupTags, forwardMap) {
  if (!Array.isArray(clickupTags)) return '';
  var out = [];
  clickupTags.forEach(function(t) {
    var name = (t && t.name) ? t.name : t;
    if (!name) return;
    out.push(forwardMap[name] || name);   // Display Name if mapped, else raw name
  });
  return out.join(', ');
}

/**
 * Apply dropdown to Task Category column using only Display Names from the Tags sheet.
 */
function applyLabelsDropdown(dataSheet, firstRow, numRows, listOffset) {
  var ss       = SpreadsheetApp.getActiveSpreadsheet();
  var tagSheet = ss.getSheetByName(TAGS_SHEET);
  if (!tagSheet || tagSheet.getLastRow() < 2) return;
  var displayNames = tagSheet.getRange(2, 2, tagSheet.getLastRow() - 1, 1).getValues()
    .map(function(r){ return String(r[0] || '').trim(); })
    .filter(Boolean);
  if (displayNames.length === 0) return;
  displayNames.sort(function(a, b){ return a.toLowerCase().localeCompare(b.toLowerCase()); });
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(displayNames, true)
    .setAllowInvalid(true)
    .build();
  dataSheet.getRange(firstRow, col_(LABELS_COL, listOffset), numRows, 1).setDataValidation(rule);
}

// ---------- Roles sheet with per-person Role + Rate ----------

/**
 * Roles sheet mirrors the Tags sheet pattern, with three columns:
 *   A: Full Name  (read-only, script-managed)
 *   B: Roles      (user-edited — the person's role label, e.g. "Project Delivery Lead")
 *   C: Rate ($/hr) (user-edited — drives per-entry Cost when Rate Mode = Per Role)
 * Roles AND Rates are preserved across refreshes. A person without a rate is treated
 * as 0; a person without a role shows a blank Task Category in the Report.
 * Used only when Config → Rate Mode = "Per Role". In that mode ALL rows are visible
 * (tag mapping no longer governs visibility) and Task Category shows the Role.
 */

/**
 * Menu action: full standalone scan of who logged time on the selected List,
 * then write/merge the Roles sheet.
 */
function refreshRolesList() {
  var cfg = readConfig();
  if (!cfg.listId) throw new Error('Missing List ID in Config.');
  SpreadsheetApp.getActive().toast('Scanning time entries for people...', 'ClickUp');

  var range     = resolveDateRange(cfg);
  var memberIds = getTeamMemberIds(cfg.token, cfg.teamId);
  if (memberIds.length === 0) throw new Error('No team members found for this Team ID.');
  var extra  = getExtraUserIds_();
  var merged = mergeAssigneeIds_(memberIds, extra.ids);
  var entries   = getTimeEntries(cfg.token, cfg.teamId, cfg.listId, range.startMs, range.endMs, merged.ids);

  var names  = collectPersonNames_(entries);
  var result = writeRolesSheet_(names);
  var extraNote = extraUsersNote_(extra, merged.addedCount);
  SpreadsheetApp.getActive().toast(
    'Loaded ' + names.length + ' person(s). Preserved ' + result.preservedRoles + ' role(s) and ' +
    result.preservedRates + ' rate(s). Fill in Roles and Rate; blank rate = $0 when Rate Mode is Per Role.' + extraNote,
    'ClickUp', extra.invalid.length > 0 ? 15 : 10
  );
}

/**
 * Extract distinct Full Names from a set of time entries, sorted alphabetically.
 * Full name uses the same field the Report does (username || email).
 */
function collectPersonNames_(entries) {
  var seen = {};
  (entries || []).forEach(function(e) {
    var name = (e.user && (e.user.username || e.user.email)) || '';
    name = String(name).trim();
    if (name) seen[name] = true;
  });
  return Object.keys(seen).sort(function(a, b){ return a.toLowerCase().localeCompare(b.toLowerCase()); });
}

/**
 * Write the Roles sheet, preserving existing Roles AND Rates. Merges any incoming
 * names with names already present (so sync-time top-up never drops people).
 * Always ensures the 3-column header (Full Name · Roles · Rate). Manually entered
 * Roles values survive every refresh.
 * Returns { preservedRoles: n, preservedRates: n }.
 */
function writeRolesSheet_(names) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DEVS_SHEET);
  if (!sheet) sheet = ss.insertSheet(DEVS_SHEET);

  // Preserve existing Roles, Rates, and names (union, so we never lose a person).
  // Tolerant of an older 2-column layout (Full Name · Rate): if only 2 cols exist,
  // treat col B as Rate and leave Role blank.
  var existingRoles = {};
  var existingRates = {};
  if (sheet.getLastRow() > 1) {
    var lastCol  = Math.max(sheet.getLastColumn(), 1);
    var width    = Math.min(Math.max(lastCol, 2), 3);
    var existing = sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues();
    existing.forEach(function(r) {
      var nm = String(r[0] || '').trim();
      if (!nm) return;
      if (width >= 3) {
        var role = String(r[1] || '').trim();
        var rate = r[2];
        if (role) existingRoles[nm] = role;
        if (rate !== '' && rate !== null && rate !== undefined) existingRates[nm] = rate;
      } else { // legacy 2-col: [Full Name, Rate]
        var rate2 = r[1];
        if (rate2 !== '' && rate2 !== null && rate2 !== undefined) existingRates[nm] = rate2;
      }
    });
  }

  // Union incoming names with any already on the sheet
  var union = {};
  (names || []).forEach(function(n){ if (n) union[n] = true; });
  Object.keys(existingRoles).forEach(function(n){ union[n] = true; });
  Object.keys(existingRates).forEach(function(n){ union[n] = true; });
  var allNames = Object.keys(union).sort(function(a, b){ return a.toLowerCase().localeCompare(b.toLowerCase()); });

  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p){ p.remove(); });
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p){ p.remove(); });
  sheet.clear();
  sheet.getRange(1, 1, 1, 3)
    .setValues([['Full Name', 'Roles', 'Rate ($/hr)']])
    .setFontWeight('bold')
    .setBackground('#000000').setFontColor('#ffffff');

  if (allNames.length > 0) {
    var rows = allNames.map(function(n) {
      return [n, existingRoles[n] || '', existingRates[n] !== undefined ? existingRates[n] : ''];
    });
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    sheet.getRange(2, 3, rows.length, 1).setNumberFormat('$#,##0.00');
  }

  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 240);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 120);

  // Protect column A (Full Name) only — leave B (Roles) and C (Rate) editable
  var colARange  = sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 1), 1);
  var protection = colARange.protect().setDescription('Person names — managed by script');
  protection.setWarningOnly(false);
  var me = Session.getEffectiveUser();
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors().filter(function(u){ return u.getEmail() !== me.getEmail(); }));
  if (protection.canDomainEdit()) protection.setDomainEdit(false);

  return { preservedRoles: Object.keys(existingRoles).length, preservedRates: Object.keys(existingRates).length };
}

/**
 * ADD-ONLY population of the Roles sheet, used by "Refresh time entries".
 *
 * Appends any incoming people who aren't already on the sheet to the BOTTOM, with a
 * blank Role/Rate. It never rewrites, re-sorts, clears, or re-protects existing rows,
 * so Roles (col B) and Rate (col C) that you've already filled in are physically left
 * alone — a refresh can't wipe them.
 *
 * "Already present" is judged case-insensitively and trimmed, so accented or
 * whitespace-variant names (e.g. "Santiago Muñoz") don't create duplicates.
 *
 * The full rebuild/merge (writeRolesSheet_) is reserved for the explicit
 * "Refresh roles list" menu action.
 *
 * Returns the array of newly-added names (empty if none).
 */
function appendNewPeopleToRoles_(names) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DEVS_SHEET);

  // If the sheet doesn't exist yet, fall back to a full build once (first run).
  if (!sheet) { writeRolesSheet_(names); return []; }

  // Build a set of existing names (lowercased + trimmed) from col A.
  var existing = {};
  if (sheet.getLastRow() > 1) {
    var colA = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    colA.forEach(function(r){
      var nm = String(r[0] || '').toLowerCase().trim();
      if (nm) existing[nm] = true;
    });
  } else if (sheet.getLastRow() === 0) {
    // Empty sheet (no header) — establish the header via a full build, then continue.
    writeRolesSheet_(names);
    return [];
  }

  // Figure out which incoming names are genuinely new (dedup incoming list too).
  var seen  = {};
  var toAdd = [];
  (names || []).forEach(function(n){
    var nm = String(n || '').trim();
    if (!nm) return;
    var key = nm.toLowerCase();
    if (existing[key] || seen[key]) return;
    seen[key] = true;
    toAdd.push(nm);
  });
  if (toAdd.length === 0) return [];

  // Append at the bottom — existing rows are never touched or moved.
  var startRow = sheet.getLastRow() + 1;
  var rows = toAdd.map(function(n){ return [n, '', '']; });
  sheet.getRange(startRow, 1, rows.length, 3).setValues(rows);
  sheet.getRange(startRow, 3, rows.length, 1).setNumberFormat('$#,##0.00');
  return toAdd;
}

/**
 * Returns { fullNameLower: rate } for per-role rate lookup.
 * Keyed by Full Name (lowercased). Missing/blank rate → resolves to 0 at lookup.
 */
function getRoleRateMap_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DEVS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return {};
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  var map  = {};
  data.forEach(function(r){ if (r[0]) map[String(r[0]).toLowerCase().trim()] = Number(r[2]) || 0; });
  return map;
}

/**
 * Returns { fullNameLower: role } for per-role Task Category display.
 * Keyed by Full Name (lowercased). Missing role → absent (blank category).
 */
function getRoleNameMap_() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DEVS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return {};
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  var map  = {};
  data.forEach(function(r){
    var nm = String(r[0] || '').toLowerCase().trim();
    var role = String(r[1] || '').trim();
    if (nm && role) map[nm] = role;
  });
  return map;
}

// ---------- Transform ----------

function entryToRow(e, tz, rateMap, tagForwardMap, rateMode, roleRateMap, roleNameMap, listOffset) {
  var startDate = new Date(Number(e.start));
  var durHours  = Math.round((Number(e.duration || 0) / 3600000) * 100) / 100;
  var task      = e.task || {};
  var user      = (e.user && (e.user.username || e.user.email)) || '';
  var displayId = task.custom_id || task.id || '';
  var billable  = e.billable === true;

  // Rate source + Task Category depend on Config → Rate Mode. The two modes are
  // fully separated: Per Task is the fork's original behavior; Per Role is new.
  var rate;
  var category;
  if (rateMode === 'Per Role') {
    // Per Role: rate keyed by Full Name (blank/unknown = 0); Task Category = the
    // person's Role from the Roles sheet (blank if no role). Tag mapping is NOT
    // consulted for visibility or category — all rows surface.
    var userKey = user.toLowerCase().trim();
    rate     = (roleRateMap && userKey) ? (roleRateMap[userKey] || 0) : 0;
    category = (roleNameMap && userKey) ? (roleNameMap[userKey] || '') : '';
  } else {
    // Per Task: first raw ClickUp tag → Tags sheet rate; Task Category = mapped tag
    // Display Names only (unmapped tags hidden). Original fork logic, unchanged.
    var rawTagNames = Array.isArray(e.tags) ? e.tags.map(function(t){ return t.name; }) : [];
    var firstTag    = rawTagNames[0] || '';
    rate     = rateMap && firstTag ? (rateMap[firstTag.toLowerCase().trim()] || 0) : 0;
    category = mapTagsForDisplay_(e.tags || [], tagForwardMap || {});
  }

  var snapshot = JSON.stringify({ description: e.description || '', tags: category, billable: billable });

  var rowArr = [
    Utilities.formatDate(startDate, tz, 'yyyy-MM-dd'), // 1 Date
    displayId,                                          // 2 Issue Key
    task.name || '',                                    // 3 Issue summary
    e.description || '',                                // 4 Work Description
    durHours,                                           // 5 Hours
    rate,                                               // 6 Rate
    '',                                                 // 7 Cost — formula written separately
    user,                                               // 8 Full name
    category,                                           // 9 Task Category (tag Display Names, or Role in Per Role mode)
    billable,                                           // 10 Billable
    '',                                                 // 11 Pending
    false,                                              // 12 Confirm
    e.id || '',                                         // 13 Entry ID
    snapshot,                                           // 14 Snapshot
  ];

  // Multi-list mode: prepend the List name (display-only, not in Snapshot).
  if (listOffset) {
    var loc      = e.task_location || {};
    var listName = loc.list_name
                || (e.task && e.task.list && e.task.list.name)
                || '';
    rowArr.unshift(listName);
  }
  return rowArr;
}

// ---------- Refresh ----------


function refreshTimeEntries(skipPendingCheck) {
  if (!skipPendingCheck) {
    var pendingCount = countPendingRows_();
    if (pendingCount > 0) {
      var ui   = SpreadsheetApp.getUi();
      var resp = ui.alert(
        'Pending changes',
        'You have ' + pendingCount + ' row(s) with pending changes. Refreshing will discard them.\n\n' +
        'YES = Sync first.\nNO = Refresh anyway (discards edits).\nCANCEL = Do nothing.',
        ui.ButtonSet.YES_NO_CANCEL
      );
      if (resp === ui.Button.YES)  { syncPendingChanges(); return; }
      if (resp !== ui.Button.NO)   return;
    }
  }

  var cfg    = readConfig();
  if (!cfg.listIds || cfg.listIds.length === 0) throw new Error('Missing List ID in Config.');
  var range  = resolveDateRange(cfg);
  var tz     = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  SpreadsheetApp.getActive().toast('Fetching ' + range.label + '...', 'ClickUp');

  var memberIds = getTeamMemberIds(cfg.token, cfg.teamId);
  if (memberIds.length === 0) throw new Error('No team members found for this Team ID.');
  var extra  = getExtraUserIds_();
  var merged = mergeAssigneeIds_(memberIds, extra.ids);

  var entries = getTimeEntries(cfg.token, cfg.teamId, cfg.listIds, range.startMs, range.endMs, merged.ids);
  var totalFetched = entries.length;
  if (cfg.billableFilter === 'Billable only')     entries = entries.filter(function(e){ return e.billable === true; });
  else if (cfg.billableFilter === 'Non-billable only') entries = entries.filter(function(e){ return e.billable !== true; });

  // Skip IDs — exclude entries whose task custom_id or task id matches any entry in cfg.skipIds
  if (cfg.skipIds.length > 0) {
    entries = entries.filter(function(e) {
      var task = e.task || {};
      var id   = (task.custom_id || task.id || '').toLowerCase().trim();
      return cfg.skipIds.indexOf(id) === -1;
    });
  }
  entries.sort(function(a, b){ return Number(a.start) - Number(b.start); });

  var rateMap = getTagRateMap_();
  var tagMaps = getTagMaps_();

  // Add-only top-up of the Roles sheet: append any NEW people from this sync's
  // entries to the bottom, never rewriting existing rows (so Roles/Rate you filled
  // in are never wiped). The full rebuild lives in the "Refresh roles list" menu
  // item. Per Role maps below read the Roles sheet directly, so this stays correct.
  appendNewPeopleToRoles_(collectPersonNames_(entries));
  var roleRateMap = getRoleRateMap_();
  var roleNameMap = getRoleNameMap_();

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DATA_SHEET);
  if (!sheet) sheet = ss.insertSheet(DATA_SHEET);
  sheet.clear();
  // sheet.clear() removes content + formats but NOT data validations. When the layout
  // changes (single↔multi list, or Task Category hidden in Per Role), stale dropdowns
  // would otherwise linger on cells that no longer mean the same thing — and on rows
  // below the new data. Clear ALL validations across the whole sheet up front, then
  // re-apply only the ones the current layout needs.
  if (sheet.getMaxRows() > 0 && sheet.getMaxColumns() > 0) {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearDataValidations();
  }

  // Multi-list: prepend a "List" column (offset = 1) only when >1 List is selected.
  var off      = (cfg.listIds && cfg.listIds.length > 1) ? 1 : 0;
  var totalCols = COLUMNS.length + off;
  // Column letters for formulas, shifted by the offset.
  var HOURS_L = colLetter_(col_(5,  off));   // E (single-list)
  var RATE_L  = colLetter_(col_(6,  off));   // F
  var COST_L  = colLetter_(col_(COST_COL, off));
  var BILL_L  = colLetter_(col_(BILLABLE_COL, off)); // J
  var LABEL_COLNUM = col_(DESCRIPTION_COL, off);     // subtotal labels sit under Work Description (col 4)

  // Header — black background, white bold text, Anek Tamil
  var hdrRange = sheet.getRange(1, 1, 1, totalCols);
  // Header labels are mode-aware: in Per Role mode the Task Category col reads "Role".
  // COLUMNS is not mutated; the optional "List" header is prepended when off === 1.
  var headerLabels = COLUMNS.slice();
  if (cfg.rateMode === 'Per Role') headerLabels[LABELS_COL - 1] = 'Role';
  if (off) headerLabels.unshift(LIST_COL_HEADER);
  hdrRange.setValues([headerLabels])
    .setFontWeight('bold')
    .setBackground('#000000').setFontColor('#ffffff')
    .setFontFamily(REPORT_FONT).setFontSize(REPORT_FONT_SIZE);

  if (entries.length > 0) {
    var rows = entries.map(function(e){ return entryToRow(e, tz, rateMap, tagMaps.forward, cfg.rateMode, roleRateMap, roleNameMap, off); });
    sheet.getRange(2, 1, rows.length, totalCols).setValues(rows);

    sheet.getRange(2, col_(BILLABLE_COL, off), rows.length, 1).insertCheckboxes();
    sheet.getRange(2, col_(CONFIRM_COL,  off), rows.length, 1).insertCheckboxes();

    // Cost formula written AFTER insertCheckboxes so the Billable col is fully
    // initialised as boolean checkboxes before the IF references it.
    // ROUND(...,2) at the row level so the stored Cost is the true penny amount and
    // the subtotal always reconciles with the printed line items even for fractional
    // hours. Column letters are shifted by the List-column offset.
    for (var i = 0; i < rows.length; i++) {
      var rn = i + 2;
      sheet.getRange(rn, col_(COST_COL, off)).setFormula(
        '=ROUND(IF(' + BILL_L + rn + ',' + HOURS_L + rn + '*' + RATE_L + rn + ',0),2)'
      );
    }

    // Task Category dropdown (tag Display Names) applies only in Per Task mode.
    // In Per Role mode the column holds a Role, doesn't sync, and is hidden below —
    // all validations were already cleared for the whole sheet above, so nothing to do.
    if (cfg.rateMode !== 'Per Role') {
      applyLabelsDropdown(sheet, 2, rows.length, off);
    }

    // Subtotals block
    var dataEnd = entries.length + 1; // last data row number
    var subRow  = entries.length + 3;
    var billableOnly = (cfg.billableFilter === 'Billable only');
    var HR = HOURS_L, RT = RATE_L, CT = COST_L, BL = BILL_L;  // shorthands

    if (billableOnly) {
      // Billable only: every row is billable, so collapse to a single line.
      sheet.getRange(subRow, LABEL_COLNUM).setValue('Sub total Hours / Total Amount Due');
      sheet.getRange(subRow, col_(5, off)).setFormula('=SUM(' + HR + '2:' + HR + dataEnd + ')');
      sheet.getRange(subRow, col_(COST_COL, off)).setFormula('=SUM(' + CT + '2:' + CT + dataEnd + ')');
      sheet.getRange(subRow, LABEL_COLNUM, 1, 1).setFontWeight('bold');
      sheet.getRange(subRow, col_(5, off), 1, 1).setFontWeight('bold');
      sheet.getRange(subRow, col_(COST_COL, off), 1, 1).setFontWeight('bold');
      sheet.getRange(subRow, col_(COST_COL, off), 1, 1).setNumberFormat('$#,##0.00');
    } else {
      sheet.getRange(subRow,     LABEL_COLNUM).setValue('Sub total Support Hours');
      sheet.getRange(subRow,     col_(5, off)).setFormula('=SUM(' + HR + '2:' + HR + dataEnd + ')');
      sheet.getRange(subRow,     col_(COST_COL, off)).setFormula('=SUM(' + CT + '2:' + CT + dataEnd + ')');
      sheet.getRange(subRow + 1, LABEL_COLNUM).setValue('Total Hours Credit');
      sheet.getRange(subRow + 1, col_(5, off)).setFormula('=SUMIF(' + BL + '2:' + BL + dataEnd + ',FALSE,' + HR + '2:' + HR + dataEnd + ')');
      sheet.getRange(subRow + 1, col_(COST_COL, off)).setFormula('=SUMPRODUCT((' + BL + '2:' + BL + dataEnd + '=FALSE)*ROUND(' + HR + '2:' + HR + dataEnd + '*' + RT + '2:' + RT + dataEnd + ',2))');
      sheet.getRange(subRow + 2, LABEL_COLNUM).setValue('Total Amount Due');
      sheet.getRange(subRow + 2, col_(5, off)).setFormula('=SUMIF(' + BL + '2:' + BL + dataEnd + ',TRUE,' + HR + '2:' + HR + dataEnd + ')');
      sheet.getRange(subRow + 2, col_(COST_COL, off)).setFormula('=SUM(' + CT + '2:' + CT + dataEnd + ')');
      sheet.getRange(subRow, LABEL_COLNUM, 3, 1).setFontWeight('bold');
      sheet.getRange(subRow, col_(5, off), 3, 1).setFontWeight('bold');
      sheet.getRange(subRow, col_(COST_COL, off), 3, 1).setFontWeight('bold');
      sheet.getRange(subRow, col_(COST_COL, off), 3, 1).setNumberFormat('$#,##0.00');
      sheet.getRange(subRow + 1, col_(COST_COL, off)).setFontColor('#888888'); // credit value grayed out
    }
  }

  sheet.setFrozenRows(1);
  if (off) sheet.setColumnWidth(1, 140);  // List column width
  for (var c = 0; c < COLUMN_WIDTHS.length; c++) sheet.setColumnWidth(c + 1 + off, COLUMN_WIDTHS[c]);
  sheet.hideColumns(col_(ENTRY_ID_COL, off));
  sheet.hideColumns(col_(SNAPSHOT_COL, off));
  // Per Role mode: the Task Category column holds the Role (already shown on the
  // Dashboard and derivable from the Roles sheet) and doesn't sync — hide it to keep
  // the Report clean. sheet.clear() does NOT reset column visibility, so we must
  // explicitly show it again in Per Task mode (in case a prior Per Role run hid it).
  // The column and all offset math stay intact; only its visibility changes.
  if (cfg.rateMode === 'Per Role') {
    sheet.hideColumns(col_(LABELS_COL, off));
  } else {
    sheet.showColumns(col_(LABELS_COL, off));
  }
  WRAP_COLUMNS.forEach(function(colBase){
    sheet.getRange(1, col_(colBase, off), Math.max(sheet.getMaxRows(), 1), 1).setWrap(true);
  });

  // Number formats + body font (Anek Tamil 11pt across all data + subtotal rows)
  if (entries.length > 0) {
    sheet.getRange(2, col_(5, off),        entries.length, 1).setNumberFormat('0.00');       // Hours
    sheet.getRange(2, col_(RATE_COL, off), entries.length, 1).setNumberFormat('$#,##0.00');  // Rate
    sheet.getRange(2, col_(COST_COL, off), entries.length, 1).setNumberFormat('$#,##0.00');  // Cost

    // Apply font across data rows + subtotal block (1 row if Billable only, else 3)
    var subtotalRows = (cfg.billableFilter === 'Billable only') ? 1 : 3;
    var bodyEndRow = entries.length + 1 + subtotalRows;
    sheet.getRange(2, 1, bodyEndRow - 1, totalCols)
      .setFontFamily(REPORT_FONT).setFontSize(REPORT_FONT_SIZE);
  }

  // Rebuild dependents
  // Rebuild dependents — Dashboard build/update is skipped when Config
  // "Build Dashboard" = No (the manual "Rebuild Dashboard" menu item still works).
  // When skipped, any existing Dashboard tab is left as-is (may be stale).
  if (cfg.buildDashboard) rebuildDashboard();

  var dashNote   = cfg.buildDashboard ? '' : ' [Dashboard skipped]';
  var filterNote = cfg.billableFilter !== 'All' ? ' [' + cfg.billableFilter + ': ' + entries.length + '/' + totalFetched + ']' : '';
  var extraNote  = extraUsersNote_(extra, merged.addedCount);
  SpreadsheetApp.getActive().toast(entries.length + ' entries loaded (' + range.label + ')' + filterNote + dashNote + '.' + extraNote,
    'ClickUp', extra.invalid.length > 0 ? 15 : 5);
}

// ---------- Dashboard ----------

// Style constants — exact values from ParadigmPT reference sheet
var DS = {
  DARK:       '#2C3E50',   // title / section heading / col header bg
  KPI_BG:     '#F4F2EC',   // KPI card warm off-white
  TOTAL_BG:   '#EBE8DF',   // Total row slightly darker warm tone
  BORDER:     '#D5D2C8',   // thin border on all table cells
  WHITE:      '#FFFFFF',
  LABEL_GREY: '#555555',   // KPI label text
  SUB_GREY:   '#777777',   // KPI sub-label text
  FONT:       'Arial',
};

function rebuildDashboard() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var report    = ss.getSheetByName(DATA_SHEET);
  var dashboard = ss.getSheetByName(DASHBOARD_SHEET);
  if (!dashboard) dashboard = ss.insertSheet(DASHBOARD_SHEET);

  ss.setActiveSheet(dashboard);
  ss.moveActiveSheet(1);
  dashboard.clearContents();
  dashboard.clearFormats();
  dashboard.setHiddenGridlines(true);  // no gridlines — matches reference

  if (!report || report.getLastRow() < 2) {
    dashboard.getRange(1, 1).setValue('No data. Run "Refresh time entries" first.');
    return;
  }

  var cfg = readConfig();
  var lastReportRow = report.getLastRow();
  // Read including any leading "List" column, then drop it so all downstream
  // 0-based indices (r[1]=Issue Key, r[4]=Hours, r[7]=Full name, etc.) are unchanged
  // whether or not multi-list mode added the column.
  var dashOff = reportListOffset_(report);
  var rawData = report.getRange(2, 1, lastReportRow - 1, COLUMNS.length + dashOff).getValues();
  if (dashOff) rawData = rawData.map(function(r){ return r.slice(dashOff); });
  var data = rawData.filter(function(r){ return r[1] !== ''; });

  // Apply Skip IDs — filter out rows whose Issue Key matches the skip list
  if (cfg.skipIds && cfg.skipIds.length > 0) {
    data = data.filter(function(r) {
      var id = String(r[1] || '').toLowerCase().trim();
      return cfg.skipIds.indexOf(id) === -1;
    });
  }

  // ── Compute KPIs ──────────────────────────────────────────────
  var totalHours = 0, billableHours = 0, creditHours = 0, totalCost = 0, creditCost = 0;
  var personMap = {}, categoryMap = {}, taskMap = {}, roleMap = {};

  // Per Role dashboard needs role labels keyed by Full Name. In Per Role mode the
  // Report's Task Category already holds the Role, but we key off the Roles sheet
  // directly so blank-category rows still attribute to the right role.
  var roleNameMapDash = (cfg.rateMode === 'Per Role') ? getRoleNameMap_() : {};

  data.forEach(function(r) {
    var hours    = Number(r[4]) || 0;
    var rate     = Number(r[5]) || 0;
    var isBill   = r[9] === true;
    var person   = String(r[7] || '');
    var category = String(r[8] || 'Uncategorized');
    var issueKey = String(r[1] || '');
    var issueName= String(r[2] || '');

    totalHours += hours;
    // Round each row's cost to 2 dp before accumulating, matching the Report's
    // per-row ROUND(...,2) so Dashboard totals reconcile with the sheet to the penny.
    var rowCost = Math.round(hours * rate * 100) / 100;
    if (isBill) { billableHours += hours; totalCost  += rowCost; }
    else        { creditHours  += hours; creditCost += rowCost; }

    if (person) {
      if (!personMap[person]) personMap[person] = { billable: 0, credit: 0 };
      if (isBill) personMap[person].billable += hours;
      else        personMap[person].credit   += hours;
    }
    // Role aggregation (Per Role dashboard) — billable hours only, grouped by role.
    if (cfg.rateMode === 'Per Role' && isBill) {
      var role = roleNameMapDash[person.toLowerCase().trim()] || '(no role)';
      if (!roleMap[role]) roleMap[role] = 0;
      roleMap[role] += hours;
    }
    if (category) {
      category.split(',').forEach(function(cat) {
        cat = cat.trim(); if (!cat) return;
        if (!categoryMap[cat]) categoryMap[cat] = 0;
        categoryMap[cat] += hours;
      });
    }
    if (issueKey) {
      if (!taskMap[issueKey]) taskMap[issueKey] = { name: issueName, hours: 0, category: category };
      taskMap[issueKey].hours += hours;
    }
  });

  // ── Column layout (matches reference: col A = narrow margin, col B = wide labels) ──
  // Reference: A=2char, B=38char, C=14, D=13, E=13, F-J=13 each
  dashboard.setColumnWidth(1, 18);   // col A — narrow left margin
  dashboard.setColumnWidth(2, 285);  // col B — names / labels
  dashboard.setColumnWidth(3, 105);  // col C — Billable
  dashboard.setColumnWidth(4, 100);  // col D — Credit
  dashboard.setColumnWidth(5, 100);  // col E — Total
  dashboard.setColumnWidth(6, 100);  // col F — (KPI / spacer)
  dashboard.setColumnWidth(7, 100);  // col G
  dashboard.setColumnWidth(8, 30);   // col H — thin spacer
  dashboard.setColumnWidth(9, 125);  // col I — % billable / Of total

  var r = 1;
  var clientName = cfg.clientName || 'Client';
  var monthLabel = cfg.monthLabel || '';

  // ── Title row — B2:I2 merged, dark bg, white 16pt bold, centred ──
  dashboard.getRange(r, 2, 1, 8).merge()
    .setValue(clientName + ' — ' + monthLabel + ' Support Hours Dashboard')
    .setFontFamily(DS.FONT).setFontSize(16).setFontWeight('bold')
    .setFontColor(DS.WHITE).setBackground(DS.DARK)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  dashboard.setRowHeight(r, 28);
  r++;

  // ── KPI block ─────────────────────────────────────────────────
  // Full: Total (B:C), Billable (D:E), Credit (F:G), % billable (I).
  // Billable only: keep all four card slots (so the warm band and % billable stay
  // in place), but blank the Credit card content — the F:G slot renders empty.
  var billableOnlyDash = (cfg.billableFilter === 'Billable only');
  var pct = totalHours > 0 ? billableHours / totalHours : 0;

  var kpiStartCols = [2, 4, 6, 9];   // B, D, F, I  (col I for % billable which spans more)
  var kpiSpans     = [2, 2, 2, 1];   // how many cols each KPI merges
  var kpiLabels, kpiValues, kpiSubs, kpiNumFmts;
  if (billableOnlyDash) {
    kpiLabels  = ['Total hours', 'Billable hours', '', '% billable'];
    kpiValues  = [Math.round(totalHours*100)/100, Math.round(billableHours*100)/100, '', pct];
    kpiSubs    = [monthLabel, '$' + totalCost.toFixed(2), '', 'Of total hours worked'];
    kpiNumFmts = ['0.0', '0.00', '0.00', '0.0%'];
  } else {
    kpiLabels  = ['Total hours', 'Billable hours', 'Credit hours', '% billable'];
    kpiValues  = [Math.round(totalHours*100)/100, Math.round(billableHours*100)/100, Math.round(creditHours*100)/100, pct];
    kpiSubs    = [monthLabel, '$' + totalCost.toFixed(2), '$' + creditCost.toFixed(2), 'Of total hours worked'];
    kpiNumFmts = ['0.0', '0.00', '0.00', '0.0%'];
  }
  var kpiCount = kpiLabels.length;

  // Paint the col-H spacer with the KPI background so the gap between the Credit
  // card (F:G) and the % billable card (I) blends in as one band (both layouts).
  dashboard.getRange(r, 8, 3, 1).setBackground(DS.KPI_BG);

  // Label row
  for (var i = 0; i < kpiCount; i++) {
    var rng = dashboard.getRange(r, kpiStartCols[i], 1, kpiSpans[i]);
    if (kpiSpans[i] > 1) rng.merge();
    rng.setValue(kpiLabels[i])
      .setFontFamily(DS.FONT).setFontSize(10).setFontWeight('normal')
      .setFontColor(DS.LABEL_GREY).setBackground(DS.KPI_BG)
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
  }
  dashboard.setRowHeight(r, 18); r++;

  // Value row
  for (var i = 0; i < kpiCount; i++) {
    var rng = dashboard.getRange(r, kpiStartCols[i], 1, kpiSpans[i]);
    if (kpiSpans[i] > 1) rng.merge();
    rng.setValue(kpiValues[i])
      .setFontFamily(DS.FONT).setFontSize(18).setFontWeight('bold')
      .setBackground(DS.KPI_BG).setNumberFormat(kpiNumFmts[i])
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
  }
  dashboard.setRowHeight(r, 18); r++;

  // Sub-label row
  for (var i = 0; i < kpiCount; i++) {
    var rng = dashboard.getRange(r, kpiStartCols[i], 1, kpiSpans[i]);
    if (kpiSpans[i] > 1) rng.merge();
    rng.setValue(kpiSubs[i])
      .setFontFamily(DS.FONT).setFontSize(9).setFontWeight('normal')
      .setFontColor(DS.SUB_GREY).setBackground(DS.KPI_BG)
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
  }
  dashboard.setRowHeight(r, 18); r += 2;

  // ── Hours by person (Per Task) / Hours by role (Per Role) ─────
  var isRoleMode = (cfg.rateMode === 'Per Role');
  var personDataFirstRow, personRows;

  if (isRoleMode) {
    // Per Role: "Hours by role — billable". Single value column labeled "Total"
    // (= billable hours only). Credit is excluded entirely from this section.
    dbBand_(dashboard, r, 8, 'Hours by role — billable'); r++;
    dbHeaders_(dashboard, r, ['Role', 'Total'], [2, 3]); r++;

    personRows = Object.keys(roleMap).map(function(n) {
      return [n, Math.round(roleMap[n]*100)/100];
    }).sort(function(a,b){ return b[1]-a[1]; });

    personDataFirstRow = r;
    personRows.forEach(function(row) {
      dbDataRow_(dashboard, r, [
        {col:2, val:row[0], align:'left',  fmt:'',     bold:false},
        {col:3, val:row[1], align:'right', fmt:'0.00', bold:true},
      ]); r++;
    });

    var tRole = personRows.reduce(function(s,v){return s+v[1];},0);
    dbTotalRow_(dashboard, r, 2, ['Total', Math.round(tRole*100)/100],
      [false, true], ['','0.00'], ['left','right']);
    r++;
  } else if (billableOnlyDash) {
    // Per Task + Billable only: no credit exists, so drop the Credit + separate
    // Total columns. Show Person · Total (= billable hours), single-series chart.
    dbBand_(dashboard, r, 8, 'Hours by person — billable'); r++;
    dbHeaders_(dashboard, r, ['Person', 'Total'], [2, 3]); r++;

    personRows = Object.keys(personMap).map(function(n) {
      var p = personMap[n];
      return [n, Math.round(p.billable*100)/100];
    }).sort(function(a,b){ return b[1]-a[1]; });

    personDataFirstRow = r;
    personRows.forEach(function(row) {
      dbDataRow_(dashboard, r, [
        {col:2, val:row[0], align:'left',  fmt:'',     bold:false},
        {col:3, val:row[1], align:'right', fmt:'0.00', bold:true},
      ]); r++;
    });

    var tPers = personRows.reduce(function(s,v){return s+v[1];},0);
    dbTotalRow_(dashboard, r, 2, ['Total', Math.round(tPers*100)/100],
      [false, true], ['','0.00'], ['left','right']);
    r++;
  } else {
    // Per Task: original "Hours by person — billable vs credit" (unchanged).
    dbBand_(dashboard, r, 8, 'Hours by person — billable vs credit'); r++;
    dbHeaders_(dashboard, r, ['Person', 'Billable', 'Credit', 'Total'], [2, 3, 4, 5]); r++;

    personRows = Object.keys(personMap).map(function(n) {
      var p = personMap[n];
      return [n, Math.round(p.billable*100)/100, Math.round(p.credit*100)/100, Math.round((p.billable+p.credit)*100)/100];
    }).sort(function(a,b){ return b[3]-a[3]; });

    personDataFirstRow = r;
    personRows.forEach(function(row) {
      dbDataRow_(dashboard, r, [
        {col:2, val:row[0], align:'left',  fmt:'',     bold:false},
        {col:3, val:row[1], align:'right', fmt:'0.00', bold:false},
        {col:4, val:row[2], align:'right', fmt:'0.00', bold:false},
        {col:5, val:row[3], align:'right', fmt:'0.00', bold:true},
      ]); r++;
    });

    var bT=personRows.reduce(function(s,v){return s+v[1];},0);
    var cT=personRows.reduce(function(s,v){return s+v[2];},0);
    var tT=personRows.reduce(function(s,v){return s+v[3];},0);
    dbTotalRow_(dashboard, r, 4, ['Total', Math.round(bT*100)/100, Math.round(cT*100)/100, Math.round(tT*100)/100],
      [false, false, false, true], ['','0.00','0.00','0.00'], ['left','right','right','right']);
    r++;
  }

  // Chart anchors at (personDataFirstRow - 2) and is ~280px tall (~14 rows).
  // Push r past the chart bottom + 2-row buffer so the next section heading sits clear of it.
  r = Math.max(r, personDataFirstRow - 2 + 14) + 2;

  // ── Hours by task category (Per Task only; hidden in Per Role) ──
  var catRows = [], catDataFirstRow = null;
  if (!isRoleMode) {
    dbBand_(dashboard, r, 8, 'Hours by task category'); r++;
    dbHeaders_(dashboard, r, ['Category', 'Hours'], [2, 3]); r++;

    catRows = Object.keys(categoryMap).map(function(cat) {
      return [cat, Math.round(categoryMap[cat]*100)/100];
    }).sort(function(a,b){ return b[1]-a[1]; });

    catDataFirstRow = r;
    catRows.forEach(function(row) {
      dbDataRow_(dashboard, r, [
        {col:2, val:row[0], align:'left',  fmt:'',     bold:false},
        {col:3, val:row[1], align:'right', fmt:'0.00', bold:false},
      ]); r++;
    });
    // Category chart anchors at (catDataFirstRow - 2) and is ~280px tall (~14 rows).
    r = Math.max(r, catDataFirstRow - 2 + 14) + 2;
  }

  // ── Top 10 issues ──────────────────────────────────────────────
  // Per Task: Issue · Hours · Type.  Per Role: Issue · Hours (Type column removed).
  dbBand_(dashboard, r, 8, 'Top 10 issues by hours'); r++;
  if (isRoleMode) {
    dbHeaders_(dashboard, r, ['Issue', 'Hours'], [2, 3]); r++;
  } else {
    dbHeaders_(dashboard, r, ['Issue', 'Hours', 'Type'], [2, 3, 4]); r++;
  }

  var top10 = Object.keys(taskMap).map(function(k) {
    return [taskMap[k].name, Math.round(taskMap[k].hours*100)/100, taskMap[k].category];
  }).sort(function(a,b){ return b[1]-a[1]; }).slice(0,10);

  var top10DataFirstRow = r;
  top10.forEach(function(row) {
    var cells = [
      {col:2, val:row[0], align:'left',   fmt:'',     bold:false},
      {col:3, val:row[1], align:'right',  fmt:'0.00', bold:false},
    ];
    if (!isRoleMode) cells.push({col:4, val:row[2], align:'center', fmt:'', bold:false});
    dbDataRow_(dashboard, r, cells); r++;
  });

  // Wrap text on the Type column (col 4) for all top10 rows so long category lists fit (Per Task only)
  if (!isRoleMode && top10.length > 0) {
    dashboard.getRange(top10DataFirstRow, 4, top10.length, 1).setWrap(true);
  }

  // ── Charts ────────────────────────────────────────────────────
  dashboard.getCharts().forEach(function(c){ dashboard.removeChart(c); });
  if (personRows.length > 0) {
    if (isRoleMode)            addRoleChart_(dashboard, personDataFirstRow, personRows.length, 'Hours by role — billable');
    else if (billableOnlyDash) addRoleChart_(dashboard, personDataFirstRow, personRows.length, 'Hours by person — billable');
    else                       addPersonChart_(dashboard, personDataFirstRow, personRows.length);
  }
  if (!isRoleMode && catRows.length > 0) addCategoryChart_(dashboard, catDataFirstRow, catRows.length);
  if (top10.length > 0)                  addTop10Chart_(dashboard, top10DataFirstRow, top10.length);

  SpreadsheetApp.getActive().toast('Dashboard rebuilt.', 'ClickUp', 4);
}

// ── Dashboard style helpers ───────────────────────────────────

/** Section heading: bold 12pt text on default sheet background (blends in) */
function dbBand_(sheet, row, numCols, text) {
  sheet.getRange(row, 2, 1, numCols).merge()
    .setValue(text)
    .setFontFamily(DS.FONT).setFontSize(12).setFontWeight('bold')
    .setFontColor(DS.DARK)
    .setHorizontalAlignment('left').setVerticalAlignment('middle');
  sheet.setRowHeight(row, 22);
}

/** Column header row: dark bg, white bold 11pt, WITH thin border */
function dbHeaders_(sheet, row, labels, cols) {
  cols.forEach(function(col, i) {
    dbCell_(sheet, row, col, labels[i], DS.DARK, DS.WHITE, 11, true, 'center', '');
  });
  sheet.setRowHeight(row, 16);
}

/** Single data row */
function dbDataRow_(sheet, row, cells) {
  cells.forEach(function(c) {
    dbCell_(sheet, row, c.col, c.val, null, null, 10, c.bold, c.align, c.fmt);
  });
  sheet.setRowHeight(row, 16);
}

/** Total row: TOTAL_BG background, bold */
function dbTotalRow_(sheet, row, numCols, values, bolded, fmts, aligns) {
  for (var i = 0; i < numCols; i++) {
    dbCell_(sheet, row, i+2, values[i], DS.TOTAL_BG, null, 10, true, aligns[i], fmts[i]);
  }
  sheet.setRowHeight(row, 16);
}

/**
 * Write a single cell with full styling + thin border.
 * bg=null → no fill; fc=null → default text color.
 */
function dbCell_(sheet, row, col, value, bg, fc, fontSize, bold, hAlign, numFmt) {
  var cell = sheet.getRange(row, col);
  cell.setValue(value)
    .setFontFamily(DS.FONT).setFontSize(fontSize).setFontWeight(bold ? 'bold' : 'normal')
    .setHorizontalAlignment(hAlign).setVerticalAlignment('middle');
  if (bg) cell.setBackground(bg);
  if (fc) cell.setFontColor(fc);
  if (numFmt) cell.setNumberFormat(numFmt);
  // Thin border on every table cell
  cell.setBorder(true, true, true, true, false, false, DS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
}

// styleHeader_ kept for legacy calls
function styleHeader_(sheet, row, numCols) {
  sheet.getRange(row, 1, 1, numCols)
    .setFontWeight('bold').setFontFamily('Arial').setFontSize(11)
    .setBackground('#2C3E50').setFontColor('#FFFFFF');
}

// ── Chart helpers ─────────────────────────────────────────────

/** Chart 1: Hours by person — horizontal stacked bar */
function addPersonChart_(sheet, dataFirstRow, numPeople) {
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sheet.getRange(dataFirstRow, 2, numPeople, 1))
    .addRange(sheet.getRange(dataFirstRow, 3, numPeople, 1))
    .addRange(sheet.getRange(dataFirstRow, 4, numPeople, 1))
    .setOption('title', 'Hours by person')
    .setOption('isStacked', true)
    .setOption('legend', { position: 'bottom' })
    .setOption('hAxis', { title: 'Hours' })
    .setOption('vAxis', { textStyle: { fontSize: 11 } })
    .setOption('colors', ['#3c78d8', '#e06666'])
    .setOption('series', { 0: { labelInLegend: 'Billable' }, 1: { labelInLegend: 'Credit' } })
    .setPosition(dataFirstRow - 2, 6, 0, 0)
    .setOption('width', 520).setOption('height', 280)
    .build();
  sheet.insertChart(chart);
}

/** Chart 1b: single-series horizontal bar for the first section (Per Role, or Per Task+Billable only) */
function addRoleChart_(sheet, dataFirstRow, numRows, title) {
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sheet.getRange(dataFirstRow, 2, numRows, 1))
    .addRange(sheet.getRange(dataFirstRow, 3, numRows, 1))
    .setOption('title', title || 'Hours by role — billable')
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { title: 'Hours' })
    .setOption('vAxis', { textStyle: { fontSize: 11 } })
    .setOption('colors', ['#3c78d8'])
    .setPosition(dataFirstRow - 2, 6, 0, 0)
    .setOption('width', 520).setOption('height', 280)
    .build();
  sheet.insertChart(chart);
}

/** Chart 2: Hours by category — horizontal bar */
function addCategoryChart_(sheet, dataFirstRow, numCats) {
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sheet.getRange(dataFirstRow, 2, numCats, 1))
    .addRange(sheet.getRange(dataFirstRow, 3, numCats, 1))
    .setOption('title', 'Hours by category')
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { title: 'Hours' })
    .setOption('vAxis', { textStyle: { fontSize: 11 } })
    .setOption('colors', ['#3c78d8'])
    .setPosition(dataFirstRow - 2, 6, 0, 0)
    .setOption('width', 520).setOption('height', 280)
    .build();
  sheet.insertChart(chart);
}

/** Chart 3: Top 10 issues — horizontal bar */
function addTop10Chart_(sheet, dataFirstRow, numIssues) {
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sheet.getRange(dataFirstRow, 2, numIssues, 1))
    .addRange(sheet.getRange(dataFirstRow, 3, numIssues, 1))
    .setOption('title', 'Top 10 issues by hours')
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { title: 'Hours' })
    .setOption('vAxis', { textStyle: { fontSize: 10 } })
    .setOption('colors', ['#6aa84f'])
    .setPosition(dataFirstRow - 2, 6, 0, 0)
    .setOption('width', 520).setOption('height', 340)
    .build();
  sheet.insertChart(chart);
}

// ---------- List discovery ----------

/**
 * List every non-archived List in the workspace (whether or not it has logged
 * time in the selected period), then overlay entry counts / hours for the period.
 * Lists with no entries appear with a blank count. Entries whose List is not in the
 * hierarchy (e.g. the old "unknown" bucket) are ignored, so no "unknown" row is
 * produced. Populates the "List ID" Config dropdown.
 */
function listAllLists() {
  var cfg   = readConfig();
  var range = resolveDateRange(cfg);
  SpreadsheetApp.getActive().toast('Fetching all Lists from the workspace...', 'ClickUp');

  // 1. Pull every non-archived List from the Spaces -> Folders -> Lists hierarchy.
  var hier  = getAllListsHierarchy_(cfg.token, cfg.teamId);
  var lists = {};
  hier.forEach(function(l) {
    lists[l.id] = { id: l.id, name: l.name, folder: l.folder, space: l.space, count: 0, hours: 0 };
  });

  // 2. Second pass: overlay entry counts / hours for the configured period.
  SpreadsheetApp.getActive().toast('Scanning entries ' + range.label + ' for counts...', 'ClickUp');
  var memberIds = getTeamMemberIds(cfg.token, cfg.teamId);
  if (memberIds.length === 0) throw new Error('No team members found for this Team ID.');
  var extra  = getExtraUserIds_();
  var merged = mergeAssigneeIds_(memberIds, extra.ids);
  var entries = getTimeEntries(cfg.token, cfg.teamId, null, range.startMs, range.endMs, merged.ids);
  entries.forEach(function(e) {
    var loc = e.task_location || {};
    var id  = String(loc.list_id || (e.task && e.task.list && e.task.list.id) || '');
    if (!id || !lists[id]) return;   // ignore entries whose List isn't in the hierarchy (no "unknown" row)
    lists[id].count += 1;
    lists[id].hours += Number(e.duration || 0) / 3600000;
  });

  // Sort: Lists with entries first (desc by count), then alphabetically by label.
  var rows = Object.keys(lists).map(function(k){ return lists[k]; }).sort(function(a, b){
    if (b.count !== a.count) return b.count - a.count;
    return buildListLabel_(a).localeCompare(buildListLabel_(b));
  });

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LISTS_SHEET);
  if (!sheet) sheet = ss.insertSheet(LISTS_SHEET);
  // The Lists Found tab is script-managed output and is intentionally left editable.
  // sheet.clear() does NOT remove protections, so strip any stale sheet/range lock
  // (e.g. carried over from an earlier version) so the tab never shows as protected.
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p){ p.remove(); });
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p){ p.remove(); });
  sheet.clear();
  var header = ['List name', 'List ID', 'Folder', 'Space', '# entries', 'Total hours', 'Range', 'Display Label'];
  sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  if (rows.length > 0) {
    var data = rows.map(function(r){
      var count = r.count > 0 ? r.count : '';                              // blank when no entries
      var hours = r.count > 0 ? Math.round(r.hours * 100) / 100 : '';
      return [r.name, r.id, r.folder, r.space, count, hours, range.label, buildListLabel_(r)];
    });
    sheet.getRange(2, 1, data.length, header.length).setValues(data);
  }
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, header.length);
  applyListIdDropdown_(rows);

  var activeCount = rows.filter(function(r){ return r.count > 0; }).length;
  var extraNote = extraUsersNote_(extra, merged.addedCount);
  SpreadsheetApp.getActive().toast(
    'Found ' + rows.length + ' Lists (' + activeCount + ' with entries ' + range.label + '). Config "List ID" dropdown updated.' + extraNote,
    'ClickUp', extra.invalid.length > 0 ? 15 : 6
  );
}

function buildListLabel_(r) {
  var path = [];
  if (r.space)  path.push(r.space);
  if (r.folder) path.push(r.folder);
  return path.length > 0 ? r.name + ' (' + path.join(' > ') + ')' : r.name;
}

// ┌─────────────────────────────────────────────────────────────────────────┐
// │ REMINDER: NATIVE MULTI-SELECT (CHIPS) IS A ONE-TIME MANUAL UI STEP.        │
// │ The Apps Script DataValidationBuilder has NO setMultiSelect() method, so   │
// │ this builds a single-select dropdown. To select multiple Lists, enable it  │
// │ once in the Sheets UI on the "List ID" (B4) cell:                          │
// │   Data → Data validation → Display style: Chip → Allow multiple selections │
// │ readConfig()/splitMultiSelect_ already parse the comma-joined chip value,  │
// │ so multi-List sync works once that toggle is on. Do NOT re-add             │
// │ .setMultiSelect(true) here — it throws "setMultiSelect is not a function". │
// └─────────────────────────────────────────────────────────────────────────┘
function applyListIdDropdown_(rows) {
  if (!rows || rows.length === 0) return;
  var labels      = rows.map(function(r){ return buildListLabel_(r); });
  var configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  if (!configSheet) return;
  // Locate the "List ID" row by name (drift-proof), rather than assuming a fixed row.
  var listIdRow = -1;
  if (configSheet.getLastRow() > 0) {
    var names = configSheet.getRange(1, 1, configSheet.getLastRow(), 1).getValues();
    for (var i = 0; i < names.length; i++) {
      if (String(names[i][0] || '').trim() === 'List ID') { listIdRow = i + 1; break; }
    }
  }
  if (listIdRow === -1) return;
  // Single-select dropdown built from the available List labels. The Apps Script
  // DataValidationBuilder has no method to enable native multi-select ("chip") mode,
  // so that is turned on manually once in the Sheets UI (List ID cell → data
  // validation → Display style: Chip / Allow multiple selections). When multi-select
  // is on, Sheets stores the chosen Lists comma-joined; readConfig()/splitMultiSelect_
  // parse that back into individual Lists, so multiple-List sync works regardless.
  configSheet.getRange(listIdRow, 2).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(labels, true)
      .setAllowInvalid(false)
      .build()
  );
}

// ---------- Edit handler (marks rows as Pending) ----------

function onClickUpEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== DATA_SHEET) return;
  var col = e.range.getColumn(), row = e.range.getRow();
  if (row < 2) return;
  if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;
  // Self-detect the List-column offset from the header, then map the edited column
  // back to its BASE index before checking whether it's an editable/syncable column.
  var off = reportListOffset_(sheet);
  var baseCol = col - off;
  if (baseCol < 1) return;  // the List column itself is display-only, never syncable
  if (syncEditableCols_(readRateModeSafe_()).indexOf(baseCol) === -1) return;
  recomputePendingForRow_(sheet, row);
}

function recomputePendingForRow_(sheet, row) {
  var off        = reportListOffset_(sheet);
  var rowValues  = sheet.getRange(row, 1, 1, COLUMNS.length + off).getValues()[0];
  var snapshotJson = rowValues[col_(SNAPSHOT_COL, off) - 1];
  if (!snapshotJson) { sheet.getRange(row, col_(PENDING_COL, off)).setValue(''); return; }
  var snap;
  try { snap = JSON.parse(snapshotJson); } catch (err) { snap = null; }
  if (!snap) { sheet.getRange(row, col_(PENDING_COL, off)).setValue('?'); return; }

  var currentDesc     = String(rowValues[col_(DESCRIPTION_COL, off) - 1] || '');
  var currentCategory = String(rowValues[col_(LABELS_COL, off) - 1]      || '');
  var currentBillable = rowValues[col_(BILLABLE_COL, off) - 1] === true;
  var isPerRole = (readRateModeSafe_() === 'Per Role');
  var diffs = [];
  if (currentDesc     !== String(snap.description || ''))           diffs.push('Desc');
  // Task Category only participates in sync in Per Task mode; in Per Role it's a
  // Role with no ClickUp equivalent, so never flag it as a pending Tags change.
  if (!isPerRole &&
      normalizeTagString_(currentCategory) !== normalizeTagString_(snap.tags || '')) diffs.push('Tags');
  if (currentBillable !== (snap.billable === true))                  diffs.push('Billable');

  // Multi-tag conflict guard (Per Task only): if the row has a real change to push
  // AND its Task Category resolves to more than one tag, prefix "Multi-tag". The row
  // is then held from sync (see collectChanges_) until it is edited down to one tag.
  // A multi-tag row with no pending change is left alone (no flag).
  if (!isPerRole && diffs.length > 0 && isMultiTagConflict_(currentCategory)) {
    diffs.unshift('Multi-tag');
  }
  sheet.getRange(row, col_(PENDING_COL, off)).setValue(diffs.join(', '));
}

function normalizeTagString_(s) {
  return String(s || '').split(',').map(function(t){ return t.trim(); }).filter(Boolean).sort().join(',');
}

/**
 * Count the tags in a Task Category string, comma-safely.
 *
 * Because both raw tag names and Display Names can themselves contain commas, a
 * naive comma split can miscount. This matches the string against the known set of
 * labels (Tags sheet Display Names + raw ClickUp tag names), longest-first, peeling
 * off each recognised label. Anything left over is comma-split as a fallback.
 * Used only to detect the multi-tag conflict (>1) — not for sync.
 */
function countCategoryTags_(categoryStr) {
  var raw = String(categoryStr || '').trim();
  if (!raw) return 0;

  // Known labels: every Display Name and every raw tag name we know about.
  var maps = getTagMaps_();
  var known = {};
  Object.keys(maps.forward).forEach(function(rawName){ known[rawName] = true; });     // raw ClickUp names
  Object.keys(maps.reverse).forEach(function(displayName){ known[displayName] = true; }); // Display Names
  var labels = Object.keys(known).sort(function(a, b){ return b.length - a.length; });    // longest first

  var count = 0;
  var rest  = raw;
  while (rest.length > 0) {
    rest = rest.replace(/^[\s,]+/, '');
    if (!rest.length) break;
    var matched = null;
    for (var i = 0; i < labels.length; i++) {
      if (labels[i] && rest.indexOf(labels[i]) === 0) { matched = labels[i]; break; }
    }
    if (matched) { count += 1; rest = rest.slice(matched.length); }
    else {
      // Unrecognised remainder — fall back to comma split for whatever is left.
      rest.split(',').forEach(function(tok){ if (tok.trim()) count += 1; });
      break;
    }
  }
  return count;
}

/** True when a Task Category string resolves to more than one tag (a conflict). */
function isMultiTagConflict_(categoryStr) {
  return countCategoryTags_(categoryStr) > 1;
}

function countPendingRows_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  var off = reportListOffset_(sheet);
  var pending = sheet.getRange(2, col_(PENDING_COL, off), sheet.getLastRow() - 1, 1).getValues();
  return pending.filter(function(r){ return r[0] && String(r[0]).length > 0; }).length;
}

// ---------- Sync ----------

function collectChanges_(requireConfirm) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DATA_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return { sheet: sheet, changes: [] };

  var off     = reportListOffset_(sheet);
  var lastRow = sheet.getLastRow();
  var data    = sheet.getRange(2, 1, lastRow - 1, COLUMNS.length + off).getValues();
  var changes = [];
  var blocked = [];
  data.forEach(function(r, idx) {
    var pending = String(r[col_(PENDING_COL, off) - 1] || '').trim();
    if (!pending) return;
    // Rows with an unresolved multi-tag conflict are held from sync entirely
    // (whole row: Desc/Billable/Category). Resolve by editing the Task Category
    // cell down to a single tag, then sync again.
    if (pending.indexOf('Multi-tag') !== -1) {
      blocked.push({
        rowInSheet: idx + 2,
        entryId:    String(r[col_(ENTRY_ID_COL, off) - 1] || '').trim(),
        taskId:     String(r[col_(2, off) - 1] || ''),
        category:   String(r[col_(LABELS_COL, off) - 1] || ''),
      });
      return;
    }
    if (requireConfirm && r[col_(CONFIRM_COL, off) - 1] !== true) return;
    var snap = null;
    try { snap = JSON.parse(r[col_(SNAPSHOT_COL, off) - 1] || '{}'); } catch (err) { snap = null; }
    changes.push({
      rowInSheet:  idx + 2,
      entryId:     String(r[col_(ENTRY_ID_COL, off) - 1] || '').trim(),
      taskId:      String(r[col_(2, off) - 1] || ''),   // Issue Key
      taskName:    String(r[col_(3, off) - 1] || ''),   // Issue summary
      pending:     pending,
      snap:        snap,
      newDesc:     String(r[col_(DESCRIPTION_COL, off) - 1] || ''),
      newTags:     String(r[col_(LABELS_COL, off) - 1]      || ''),
      newBillable: r[col_(BILLABLE_COL, off) - 1] === true,
    });
  });
  return { sheet: sheet, changes: changes, blocked: blocked };
}

function executeSyncChanges_(changes, sheet) {
  var cfg     = readConfig();
  var tagMaps = getTagMaps_();
  var off     = reportListOffset_(sheet);
  var successCount = 0, failCount = 0;
  changes.forEach(function(c) {
    if (!c.entryId) {
      failCount++;
      logChange_('Failure', c.entryId, c.taskId, c.taskName, 'all', '', '(skipped)', 'Missing Entry ID');
      return;
    }
    var parts = c.pending.split(',').map(function(s){ return s.trim(); });
    var rowOK = true;
    var put   = {};
    if (parts.indexOf('Billable') !== -1) put.billable    = c.newBillable;
    if (parts.indexOf('Desc')     !== -1) put.description = c.newDesc;
    if (Object.keys(put).length > 0) {
      var putErr = null;
      try { cuPut('/team/' + cfg.teamId + '/time_entries/' + c.entryId, cfg.token, put); }
      catch (err) { rowOK = false; putErr = err.message; }
      if (put.billable    !== undefined) logChange_(putErr ? 'Failure' : 'Success', c.entryId, c.taskId, c.taskName, 'Billable',     c.snap ? (c.snap.billable ? 'Yes' : 'No') : '', put.billable ? 'Yes' : 'No', putErr || '');
      if (put.description !== undefined) logChange_(putErr ? 'Failure' : 'Success', c.entryId, c.taskId, c.taskName, 'Description',  c.snap ? (c.snap.description || '') : '',       put.description,               putErr || '');
    }
    if (parts.indexOf('Tags') !== -1) {
      // Sheet stores Display Names; reverse-map → ClickUp tag names before calling API
      var oldDisplay = parseTagList_(c.snap ? c.snap.tags : '');
      var newDisplay = parseTagList_(c.newTags);
      var oldTags = oldDisplay.map(function(d){ return tagMaps.reverse[d] || d; });
      var newTags = newDisplay.map(function(d){ return tagMaps.reverse[d] || d; });
      var oldSet  = {}; oldTags.forEach(function(t){ oldSet[t] = true; });
      var newSet  = {}; newTags.forEach(function(t){ newSet[t] = true; });
      var toRemove = oldTags.filter(function(t){ return !newSet[t]; });
      var toAdd    = newTags.filter(function(t){ return !oldSet[t]; });
      toRemove.forEach(function(tag) {
        var tErr = null;
        try { cuDelete('/team/' + cfg.teamId + '/time_entries/tags', cfg.token, { time_entry_ids: [c.entryId], tags: [{ name: tag }] }); }
        catch (err) { rowOK = false; tErr = err.message; }
        logChange_(tErr ? 'Failure' : 'Success', c.entryId, c.taskId, c.taskName, 'Labels (remove)', tag, '', tErr || '');
      });
      toAdd.forEach(function(tag) {
        var tErr = null;
        try { cuPost('/team/' + cfg.teamId + '/time_entries/tags', cfg.token, { time_entry_ids: [c.entryId], tags: [{ name: tag }] }); }
        catch (err) { rowOK = false; tErr = err.message; }
        logChange_(tErr ? 'Failure' : 'Success', c.entryId, c.taskId, c.taskName, 'Labels (add)', '', tag, tErr || '');
      });
    }
    if (rowOK) {
      successCount++;
      var newSnap = JSON.stringify({ description: c.newDesc, tags: c.newTags, billable: c.newBillable });
      sheet.getRange(c.rowInSheet, col_(SNAPSHOT_COL, off)).setValue(newSnap);
      sheet.getRange(c.rowInSheet, col_(PENDING_COL, off)).setValue('');
      sheet.getRange(c.rowInSheet, col_(CONFIRM_COL, off)).setValue(false);
      flashRow_(sheet, c.rowInSheet);
    } else {
      failCount++;
    }
  });
  updateLastSynced_();
  return { successCount: successCount, failCount: failCount };
}

function syncPendingChanges() {
  var result = collectChanges_(true);
  var blocked = result.blocked || [];
  if (result.changes.length === 0) {
    if (blocked.length > 0) {
      var brows = blocked.map(function(b){ return 'Row ' + b.rowInSheet + ' (' + b.entryId + '): [' + b.category + ']'; });
      SpreadsheetApp.getUi().alert(
        blocked.length + ' row(s) blocked — multi-tag conflict',
        'These rows have more than one tag and will NOT sync until resolved. Edit each Task Category cell down to a single tag, then sync again:\n\n' + brows.join('\n'),
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    SpreadsheetApp.getActive().toast('No confirmed pending changes.', 'ClickUp');
    return;
  }
  var ui = SpreadsheetApp.getUi();
  var lines = [];
  result.changes.slice(0, 10).forEach(function(c) {
    var parts   = c.pending.split(',').map(function(s){ return s.trim(); });
    var summary = parts.map(function(p) {
      if (p === 'Desc')     return 'Desc: "' + truncate_(c.snap ? c.snap.description : '', 40) + '" → "' + truncate_(c.newDesc, 40) + '"';
      if (p === 'Tags')     return 'Tags: [' + (c.snap ? c.snap.tags || '(none)' : '?') + '] → [' + (c.newTags || '(none)') + ']';
      if (p === 'Billable') return 'Billable: ' + (c.snap ? (c.snap.billable ? 'Yes' : 'No') : '?') + ' → ' + (c.newBillable ? 'Yes' : 'No');
      return p;
    }).join('; ');
    lines.push('Row ' + c.rowInSheet + ' (' + c.entryId + '): ' + summary);
  });
  if (result.changes.length > 10) lines.push('... and ' + (result.changes.length - 10) + ' more.');
  if (blocked.length > 0) {
    lines.push('');
    lines.push('⚠ ' + blocked.length + ' row(s) BLOCKED (multi-tag) and will not sync — resolve to one tag each: ' +
      blocked.map(function(b){ return 'Row ' + b.rowInSheet; }).join(', '));
  }
  var resp = ui.alert('Sync ' + result.changes.length + ' change(s) to ClickUp?', lines.join('\n\n'), ui.ButtonSet.OK_CANCEL);
  if (resp !== ui.Button.OK) { SpreadsheetApp.getActive().toast('Sync cancelled.', 'ClickUp'); return; }
  var outcome = executeSyncChanges_(result.changes, result.sheet);
  var msg = 'Sync complete: ' + outcome.successCount + ' succeeded, ' + outcome.failCount + ' failed.';
  if (blocked.length > 0) msg += ' ' + blocked.length + ' blocked (multi-tag).';
  if (outcome.failCount > 0) msg += ' See "' + CHANGE_LOG_SHEET + '" tab.';
  SpreadsheetApp.getActive().toast(msg, 'ClickUp', 8);
}

function syncAndReload() {
  var result = collectChanges_(false);
  var blocked = result.blocked || [];
  if (result.changes.length > 0) {
    SpreadsheetApp.getActive().toast('Syncing ' + result.changes.length + ' change(s)...', 'ClickUp');
    var outcome = executeSyncChanges_(result.changes, result.sheet);
    var t = outcome.successCount + ' synced, ' + outcome.failCount + ' failed';
    if (blocked.length > 0) t += ', ' + blocked.length + ' blocked (multi-tag)';
    SpreadsheetApp.getActive().toast(t + '. Refreshing...', 'ClickUp', 3);
  } else if (blocked.length > 0) {
    SpreadsheetApp.getActive().toast(blocked.length + ' row(s) blocked (multi-tag) — resolve to one tag. Refreshing...', 'ClickUp', 4);
  }
  refreshTimeEntries(true);
}

function discardPendingChanges() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return;
  var pendingCount = countPendingRows_();
  if (pendingCount === 0) { SpreadsheetApp.getActive().toast('Nothing to discard.', 'ClickUp'); return; }
  var ui   = SpreadsheetApp.getUi();
  var resp = ui.alert('Discard ' + pendingCount + ' pending change(s)?',
    'Edits will be reverted to snapshot values. Cannot be undone.', ui.ButtonSet.OK_CANCEL);
  if (resp !== ui.Button.OK) return;
  var off      = reportListOffset_(sheet);
  var lastRow  = sheet.getLastRow();
  var data     = sheet.getRange(2, 1, lastRow - 1, COLUMNS.length + off).getValues();
  var reverted = 0;
  data.forEach(function(r, idx) {
    var pending = String(r[col_(PENDING_COL, off) - 1] || '').trim();
    if (!pending) return;
    var snap;
    try { snap = JSON.parse(r[col_(SNAPSHOT_COL, off) - 1] || '{}'); } catch (err) { snap = null; }
    if (!snap) return;
    var rowNum = idx + 2;
    sheet.getRange(rowNum, col_(DESCRIPTION_COL, off)).setValue(snap.description || '');
    sheet.getRange(rowNum, col_(LABELS_COL, off)).setValue(snap.tags || '');
    sheet.getRange(rowNum, col_(BILLABLE_COL, off)).setValue(snap.billable === true);
    sheet.getRange(rowNum, col_(PENDING_COL, off)).setValue('');
    sheet.getRange(rowNum, col_(CONFIRM_COL, off)).setValue(false);
    reverted++;
  });
  SpreadsheetApp.getActive().toast('Discarded ' + reverted + ' pending change(s).', 'ClickUp', 5);
}

// ---------- Utilities ----------

function parseTagList_(raw) {
  if (raw == null || raw === '') return [];
  return String(raw).split(',').map(function(s){ return s.trim(); }).filter(function(s){ return s.length > 0; });
}

function truncate_(s, n) {
  s = String(s == null ? '' : s);
  return s.length > n ? s.substr(0, n - 1) + '…' : s;
}

function flashRow_(sheet, row) {
  try {
    var off   = reportListOffset_(sheet);
    var range = sheet.getRange(row, 1, 1, col_(CONFIRM_COL, off));
    var prev  = range.getBackgrounds();
    range.setBackground('#d9ead3');
    SpreadsheetApp.flush();
    Utilities.sleep(500);
    range.setBackgrounds(prev);
  } catch (err) { /* cosmetic */ }
}

function updateLastSynced_() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
    if (!sheet) return;
    var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    sheet.getRange(LAST_SYNCED_ROW, 2).setValue(Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss'));
  } catch (err) { /* ignore */ }
}

function logChange_(status, entryId, taskId, taskName, field, oldVal, newVal, message) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CHANGE_LOG_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(CHANGE_LOG_SHEET);
      sheet.getRange(1, 1, 1, 9).setValues([[
        'Timestamp', 'Status', 'Entry ID', 'Task ID', 'Task Name', 'Field', 'Old value', 'New value', 'Error',
      ]]).setFontWeight('bold');
      sheet.setFrozenRows(1);
      [150, 80, 110, 110, 260, 110, 280, 280, 280].forEach(function(w, i){ sheet.setColumnWidth(i + 1, w); });
    }
    var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    sheet.appendRow([
      Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss'),
      status, entryId || '', taskId || '', taskName || '', field || '',
      String(oldVal == null ? '' : oldVal), String(newVal == null ? '' : newVal), message || '',
    ]);
    var dataRows = sheet.getLastRow() - 1;
    if (dataRows > CHANGE_LOG_MAX_ROWS) sheet.deleteRows(2, dataRows - CHANGE_LOG_MAX_ROWS);
  } catch (err) { /* never break a sync */ }
}

// ---------- Trigger setup ----------

function setupTwoWaySync() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed  = 0;
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'onClickUpEdit') { ScriptApp.deleteTrigger(t); removed++; }
  });
  ScriptApp.newTrigger('onClickUpEdit').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  SpreadsheetApp.getActive().toast(
    'Trigger installed (removed ' + removed + ' old). Edits now mark rows as Pending.',
    'ClickUp', 6
  );
}
