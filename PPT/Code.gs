/**
 * ClickUp Time Entries → Google Sheet  |  Billing Fork
 *
 * Fork of the upstream confirm-before-sync importer. Changes:
 *   - Tags sheet gains a "Rate ($/hr)" column (preserved across refreshes)
 *   - Main sheet renamed to "Report" with billing columns:
 *       Date, Issue Key, Issue summary, Work Description, Hours,
 *       Rate, Cost, Full name, Task Category
 *   - Rate is looked up from Tags sheet by the entry's first tag
 *   - Cost = Hours × Rate (live formula in sheet)
 *   - Dashboard sheet: KPIs + hours by person + by category + top 10 issues
 *
 * Upstream features kept intact:
 *   - Confirm-before-sync two-way editing (Description, Labels, Billable)
 *   - Snapshot diffing + Pending / Confirm columns
 *   - Change Log sheet
 *   - List discovery + Config dropdown
 *   - Sync & Reload, Discard pending changes
 */

// ---------- Constants ----------

const CONFIG_SHEET    = 'Config';
const DATA_SHEET      = 'Report';          // renamed from "Time Entries"
const LISTS_SHEET     = 'Lists Found';
const TAGS_SHEET      = 'Tags';
const CHANGE_LOG_SHEET = 'Change Log';
const DASHBOARD_SHEET = 'Dashboard';
const CHANGE_LOG_MAX_ROWS = 5000;
const CLICKUP_BASE    = 'https://api.clickup.com/api/v2';

const PRESETS = ['Current month', 'Previous month', 'Current quarter', 'Previous quarter', 'Custom'];
const BILLABLE_FILTERS = ['All', 'Billable only', 'Non-billable only'];

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
const BILLABLE_COL    = 10;
const PENDING_COL     = 11;
const CONFIRM_COL     = 12;
const ENTRY_ID_COL    = 13;
const SNAPSHOT_COL    = 14;
const LABELS_COL      = 9;   // Task Category doubles as Labels for sync purposes

const EDITABLE_COLS   = [DESCRIPTION_COL, LABELS_COL, BILLABLE_COL];

const LAST_SYNCED_ROW = 10;

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
    .addItem('Rebuild Dashboard',               'rebuildDashboard')
    .addItem('List all Lists with time entries','listAllListsWithEntries')
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

function setupConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_SHEET);
  const isNew = !sheet;
  if (isNew) sheet = ss.insertSheet(CONFIG_SHEET);

  const SPEC = [
    ['Setting', 'Value', 'Notes'],
    ['API Token',         '', 'Your ClickUp personal API token (pk_...)'],
    ['Team ID',           '', 'Workspace ID from app.clickup.com/{team_id}/...'],
    ['List ID',           '', 'Run "List all Lists" first, then pick from dropdown'],
    ['Preset',            'Previous month', 'Pick from dropdown'],
    ['Custom start date', '', 'Only used if Preset = Custom (YYYY-MM-DD)'],
    ['Custom end date',   '', 'Only used if Preset = Custom (YYYY-MM-DD), inclusive'],
    ['Include subtasks',  'Yes', 'Yes / No'],
    ['Billable filter',   'All', 'All / Billable only / Non-billable only'],
    ['Last synced',       '', 'Auto-updated after a successful sync'],
    ['Client Name',       '', 'Appears in Dashboard title'],
    ['Month Label',       '', 'e.g. April 2026 — appears in Dashboard'],
    ['Skip IDs',          '', 'Comma-separated custom task IDs to exclude from Report and Dashboard'],
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

  sheet.getRange('B5').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(PRESETS, true).build());
  sheet.getRange('B8').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build());
  sheet.getRange('B9').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(BILLABLE_FILTERS, true).build());

  const preserved = Object.keys(existing).filter(k => k !== 'Setting').length;
  SpreadsheetApp.getActive().toast(
    isNew ? 'Config sheet created. Fill in the values.' : 'Config refreshed. Preserved ' + preserved + ' existing value(s).',
    'ClickUp'
  );
}

function readConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  if (!sheet) throw new Error('No Config sheet. Run "Setup config sheet" first.');
  const values = sheet.getRange(2, 1, 12, 2).getValues();
  const map = {};
  values.forEach(([k, v]) => { map[k] = v; });

  var rawListId = String(map['List ID'] || '').trim();

  // Parse Skip IDs: comma-separated, trim whitespace, lowercase for case-insensitive match
  var skipIdsRaw = String(map['Skip IDs'] || '').trim();
  var skipIds = skipIdsRaw
    ? skipIdsRaw.split(',').map(function(s){ return s.trim().toLowerCase(); }).filter(Boolean)
    : [];

  const cfg = {
    token:           String(map['API Token'] || '').trim(),
    teamId:          String(map['Team ID'] || '').trim(),
    listId:          resolveListId_(rawListId),
    listLabel:       rawListId,
    preset:          String(map['Preset'] || '').trim(),
    customStart:     map['Custom start date'],
    customEnd:       map['Custom end date'],
    includeSubtasks: String(map['Include subtasks'] || 'Yes').trim().toLowerCase() === 'yes',
    billableFilter:  String(map['Billable filter'] || 'All').trim(),
    clientName:      String(map['Client Name'] || 'Client').trim(),
    monthLabel:      String(map['Month Label'] || '').trim(),
    skipIds:         skipIds,
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
    throw new Error('Lists Found sheet is empty. Run "List all Lists with time entries" first.');
  var lastRow = listsSheet.getLastRow();
  var data = listsSheet.getRange(2, 1, lastRow - 1, 8).getValues();
  for (var i = 0; i < data.length; i++) {
    var label = String(data[i][7] || '');
    if (label === raw) return String(data[i][1]);
  }
  throw new Error('Could not find a matching List for "' + raw + '" in the Lists Found sheet.');
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

function getAllWorkspaceTags(token, teamId) {
  var data = cuFetch('/team/' + teamId + '/time_entries/tags', token);
  return Array.isArray(data.data) ? data.data : [];
}

function getTimeEntries(token, teamId, listId, startMs, endMs, assigneeIds) {
  var chunkSize = 100, all = [];
  for (var i = 0; i < assigneeIds.length; i += chunkSize) {
    var chunk  = assigneeIds.slice(i, i + chunkSize);
    var params = {
      start_date: startMs, end_date: endMs,
      assignee: chunk.join(','),
      include_task_tags: 'true', include_location_names: 'true',
    };
    if (listId) params.list_id = listId;
    var data = cuFetch('/team/' + teamId + '/time_entries', token, params);
    if (data.data) all = all.concat(data.data);
  }
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
function mapTagsForDisplay_(clickupTags, forwardMap) {
  if (!Array.isArray(clickupTags)) return '';
  var mapped = [];
  clickupTags.forEach(function(t) {
    var name    = t.name || t;
    var display = forwardMap[name];
    if (display) mapped.push(display);
  });
  return mapped.join(', ');
}

/**
 * Apply dropdown to Task Category column using only Display Names from the Tags sheet.
 */
function applyLabelsDropdown(dataSheet, firstRow, numRows) {
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
  dataSheet.getRange(firstRow, LABELS_COL, numRows, 1).setDataValidation(rule);
}

// ---------- Transform ----------

function entryToRow(e, tz, rateMap, tagForwardMap) {
  var startDate = new Date(Number(e.start));
  var durHours  = Math.round((Number(e.duration || 0) / 3600000) * 100) / 100;
  var task      = e.task || {};
  var user      = (e.user && (e.user.username || e.user.email)) || '';
  var displayId = task.custom_id || task.id || '';
  var billable  = e.billable === true;

  // Rate: keep using raw first ClickUp tag for lookup (per-task rate logic preserved)
  var rawTagNames = Array.isArray(e.tags) ? e.tags.map(function(t){ return t.name; }) : [];
  var firstTag    = rawTagNames[0] || '';
  var rate        = rateMap && firstTag ? (rateMap[firstTag.toLowerCase().trim()] || 0) : 0;

  // Task Category: only mapped tags, shown by Display Name; unmapped tags hidden
  var displayTags = mapTagsForDisplay_(e.tags || [], tagForwardMap || {});

  var snapshot = JSON.stringify({ description: e.description || '', tags: displayTags, billable: billable });

  return [
    Utilities.formatDate(startDate, tz, 'yyyy-MM-dd'), // 1 Date
    displayId,                                          // 2 Issue Key
    task.name || '',                                    // 3 Issue summary
    e.description || '',                                // 4 Work Description
    durHours,                                           // 5 Hours
    rate,                                               // 6 Rate
    '',                                                 // 7 Cost — formula written separately
    user,                                               // 8 Full name
    displayTags,                                        // 9 Task Category (display names only)
    billable,                                           // 10 Billable
    '',                                                 // 11 Pending
    false,                                              // 12 Confirm
    e.id || '',                                         // 13 Entry ID
    snapshot,                                           // 14 Snapshot
  ];
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
  if (!cfg.listId) throw new Error('Missing List ID in Config.');
  var range  = resolveDateRange(cfg);
  var tz     = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  SpreadsheetApp.getActive().toast('Fetching ' + range.label + '...', 'ClickUp');

  var memberIds = getTeamMemberIds(cfg.token, cfg.teamId);
  if (memberIds.length === 0) throw new Error('No team members found for this Team ID.');

  var entries = getTimeEntries(cfg.token, cfg.teamId, cfg.listId, range.startMs, range.endMs, memberIds);
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

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DATA_SHEET);
  if (!sheet) sheet = ss.insertSheet(DATA_SHEET);
  sheet.clear();

  // Header — black background, white bold text, Anek Tamil
  var hdrRange = sheet.getRange(1, 1, 1, COLUMNS.length);
  hdrRange.setValues([COLUMNS])
    .setFontWeight('bold')
    .setBackground('#000000').setFontColor('#ffffff')
    .setFontFamily(REPORT_FONT).setFontSize(REPORT_FONT_SIZE);

  if (entries.length > 0) {
    var rows = entries.map(function(e){ return entryToRow(e, tz, rateMap, tagMaps.forward); });
    sheet.getRange(2, 1, rows.length, COLUMNS.length).setValues(rows);

    sheet.getRange(2, BILLABLE_COL, rows.length, 1).insertCheckboxes();
    sheet.getRange(2, CONFIRM_COL,  rows.length, 1).insertCheckboxes();

    // Cost formula written AFTER insertCheckboxes so col J (Billable) is fully
    // initialised as boolean checkboxes before the IF references it.
    // Using =IF(J{n}, ...) — checkbox value is directly truthy, avoids filter edge cases.
    for (var i = 0; i < rows.length; i++) {
      var rn = i + 2;
      sheet.getRange(rn, COST_COL).setFormula('=IF(J' + rn + ',E' + rn + '*F' + rn + ',0)');
    }

    applyLabelsDropdown(sheet, 2, rows.length);

    // Subtotals block
    var dataEnd = entries.length + 1; // last data row number
    var subRow  = entries.length + 3;
    sheet.getRange(subRow,     4).setValue('Sub total Support Hours');
    sheet.getRange(subRow,     5).setFormula('=SUM(E2:E' + dataEnd + ')');
    sheet.getRange(subRow,     7).setFormula('=SUM(G2:G' + dataEnd + ')');
    sheet.getRange(subRow + 1, 4).setValue('Total Hours Credit');
    sheet.getRange(subRow + 1, 5).setFormula('=SUMIF(J2:J' + dataEnd + ',FALSE,E2:E' + dataEnd + ')');
    sheet.getRange(subRow + 1, 7).setFormula('=SUMPRODUCT((J2:J' + dataEnd + '=FALSE)*E2:E' + dataEnd + '*F2:F' + dataEnd + ')');
    sheet.getRange(subRow + 2, 4).setValue('Total Amount Due');
    sheet.getRange(subRow + 2, 5).setFormula('=SUMIF(J2:J' + dataEnd + ',TRUE,E2:E' + dataEnd + ')');
    sheet.getRange(subRow + 2, 7).setFormula('=SUM(G2:G' + dataEnd + ')');
    sheet.getRange(subRow, 4, 3, 1).setFontWeight('bold');
    sheet.getRange(subRow, 5, 3, 1).setFontWeight('bold');
    sheet.getRange(subRow, 7, 3, 1).setFontWeight('bold');
    // Format subtotal cost cells
    sheet.getRange(subRow, 7, 3, 1).setNumberFormat('$#,##0.00');
    sheet.getRange(subRow + 1, 7).setFontColor('#888888'); // credit value grayed out (informational)
  }

  sheet.setFrozenRows(1);
  for (var c = 0; c < COLUMN_WIDTHS.length; c++) sheet.setColumnWidth(c + 1, COLUMN_WIDTHS[c]);
  sheet.hideColumns(ENTRY_ID_COL);
  sheet.hideColumns(SNAPSHOT_COL);
  WRAP_COLUMNS.forEach(function(col){
    sheet.getRange(1, col, Math.max(sheet.getMaxRows(), 1), 1).setWrap(true);
  });

  // Number formats + body font (Anek Tamil 11pt across all data + subtotal rows)
  if (entries.length > 0) {
    sheet.getRange(2, 5, entries.length, 1).setNumberFormat('0.00');       // Hours
    sheet.getRange(2, 6, entries.length, 1).setNumberFormat('$#,##0.00');  // Rate
    sheet.getRange(2, 7, entries.length, 1).setNumberFormat('$#,##0.00');  // Cost

    // Apply font across data rows + 3-row subtotal block
    var bodyEndRow = entries.length + 1 + 3;
    sheet.getRange(2, 1, bodyEndRow - 1, COLUMNS.length)
      .setFontFamily(REPORT_FONT).setFontSize(REPORT_FONT_SIZE);
  }

  // Rebuild dependents
  rebuildDashboard();

  var filterNote = cfg.billableFilter !== 'All' ? ' [' + cfg.billableFilter + ': ' + entries.length + '/' + totalFetched + ']' : '';
  SpreadsheetApp.getActive().toast(entries.length + ' entries loaded (' + range.label + ')' + filterNote + '.', 'ClickUp', 5);
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
  var rawData = report.getRange(2, 1, lastReportRow - 1, COLUMNS.length).getValues();
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
  var personMap = {}, categoryMap = {}, taskMap = {};

  data.forEach(function(r) {
    var hours    = Number(r[4]) || 0;
    var rate     = Number(r[5]) || 0;
    var isBill   = r[9] === true;
    var person   = String(r[7] || '');
    var category = String(r[8] || 'Uncategorized');
    var issueKey = String(r[1] || '');
    var issueName= String(r[2] || '');

    totalHours += hours;
    if (isBill) { billableHours += hours; totalCost  += hours * rate; }
    else        { creditHours  += hours; creditCost += hours * rate; }

    if (person) {
      if (!personMap[person]) personMap[person] = { billable: 0, credit: 0 };
      if (isBill) personMap[person].billable += hours;
      else        personMap[person].credit   += hours;
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
  // Layout mirrors reference merged-cell pairs: B:C, D:E, F:G, H:I (skip A)
  // KPI 1=Total hours (B), KPI 2=Billable hours (D), KPI 3=Credit hours (F), KPI 4=% billable (H+)
  var pct = totalHours > 0 ? billableHours / totalHours : 0;
  var kpiStartCols = [2, 4, 6, 9];   // B, D, F, I  (col I for % billable which spans more)
  var kpiSpans     = [2, 2, 2, 1];   // how many cols each KPI merges
  var kpiLabels    = ['Total hours', 'Billable hours', 'Credit hours', '% billable'];
  var kpiValues    = [Math.round(totalHours*100)/100, Math.round(billableHours*100)/100, Math.round(creditHours*100)/100, pct];
  var kpiSubs      = [monthLabel, '$' + totalCost.toFixed(2), '$' + creditCost.toFixed(2), 'Of total hours worked'];
  var kpiNumFmts   = ['0.0', '0.00', '0.00', '0.0%'];

  // Paint col H (rows r, r+1, r+2) with the KPI background so the gap between
  // the Credit hours card (F:G) and the % billable card (I) blends in seamlessly
  dashboard.getRange(r, 8, 3, 1).setBackground(DS.KPI_BG);

  // Label row
  for (var i = 0; i < 4; i++) {
    var rng = dashboard.getRange(r, kpiStartCols[i], 1, kpiSpans[i]);
    if (kpiSpans[i] > 1) rng.merge();
    rng.setValue(kpiLabels[i])
      .setFontFamily(DS.FONT).setFontSize(10).setFontWeight('normal')
      .setFontColor(DS.LABEL_GREY).setBackground(DS.KPI_BG)
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
  }
  dashboard.setRowHeight(r, 18); r++;

  // Value row
  for (var i = 0; i < 4; i++) {
    var rng = dashboard.getRange(r, kpiStartCols[i], 1, kpiSpans[i]);
    if (kpiSpans[i] > 1) rng.merge();
    rng.setValue(kpiValues[i])
      .setFontFamily(DS.FONT).setFontSize(18).setFontWeight('bold')
      .setBackground(DS.KPI_BG).setNumberFormat(kpiNumFmts[i])
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
  }
  dashboard.setRowHeight(r, 18); r++;

  // Sub-label row
  for (var i = 0; i < 4; i++) {
    var rng = dashboard.getRange(r, kpiStartCols[i], 1, kpiSpans[i]);
    if (kpiSpans[i] > 1) rng.merge();
    rng.setValue(kpiSubs[i])
      .setFontFamily(DS.FONT).setFontSize(9).setFontWeight('normal')
      .setFontColor(DS.SUB_GREY).setBackground(DS.KPI_BG)
      .setHorizontalAlignment('left').setVerticalAlignment('middle');
  }
  dashboard.setRowHeight(r, 18); r += 2;

  // ── Hours by person ───────────────────────────────────────────
  dbBand_(dashboard, r, 8, 'Hours by person — billable vs credit'); r++;
  dbHeaders_(dashboard, r, ['Person', 'Billable', 'Credit', 'Total'], [2, 3, 4, 5]); r++;

  var personRows = Object.keys(personMap).map(function(n) {
    var p = personMap[n];
    return [n, Math.round(p.billable*100)/100, Math.round(p.credit*100)/100, Math.round((p.billable+p.credit)*100)/100];
  }).sort(function(a,b){ return b[3]-a[3]; });

  var personDataFirstRow = r;
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

  // Person chart anchors at (personDataFirstRow - 2) and is ~280px tall (~14 rows).
  // Push r past the chart bottom + 2-row buffer so the next section heading sits clear of it.
  r = Math.max(r, personDataFirstRow - 2 + 14) + 2;

  // ── Hours by task category ─────────────────────────────────────
  dbBand_(dashboard, r, 8, 'Hours by task category'); r++;
  dbHeaders_(dashboard, r, ['Category', 'Hours'], [2, 3]); r++;

  var catRows = Object.keys(categoryMap).map(function(cat) {
    return [cat, Math.round(categoryMap[cat]*100)/100];
  }).sort(function(a,b){ return b[1]-a[1]; });

  var catDataFirstRow = r;
  catRows.forEach(function(row) {
    dbDataRow_(dashboard, r, [
      {col:2, val:row[0], align:'left',  fmt:'',     bold:false},
      {col:3, val:row[1], align:'right', fmt:'0.00', bold:false},
    ]); r++;
  });
  // Category chart anchors at (catDataFirstRow - 2) and is ~280px tall (~14 rows).
  r = Math.max(r, catDataFirstRow - 2 + 14) + 2;

  // ── Top 10 issues ──────────────────────────────────────────────
  dbBand_(dashboard, r, 8, 'Top 10 issues by hours'); r++;
  dbHeaders_(dashboard, r, ['Issue', 'Hours', 'Type'], [2, 3, 4]); r++;

  var top10 = Object.keys(taskMap).map(function(k) {
    return [taskMap[k].name, Math.round(taskMap[k].hours*100)/100, taskMap[k].category];
  }).sort(function(a,b){ return b[1]-a[1]; }).slice(0,10);

  var top10DataFirstRow = r;
  top10.forEach(function(row) {
    dbDataRow_(dashboard, r, [
      {col:2, val:row[0], align:'left',   fmt:'',     bold:false},
      {col:3, val:row[1], align:'right',  fmt:'0.00', bold:false},
      {col:4, val:row[2], align:'center', fmt:'',     bold:false},
    ]); r++;
  });

  // Wrap text on the Type column (col 4) for all top10 rows so long category lists fit
  if (top10.length > 0) {
    dashboard.getRange(top10DataFirstRow, 4, top10.length, 1).setWrap(true);
  }

  // ── Charts ────────────────────────────────────────────────────
  dashboard.getCharts().forEach(function(c){ dashboard.removeChart(c); });
  if (personRows.length > 0) addPersonChart_(dashboard, personDataFirstRow, personRows.length);
  if (catRows.length > 0)    addCategoryChart_(dashboard, catDataFirstRow, catRows.length);
  if (top10.length > 0)      addTop10Chart_(dashboard, top10DataFirstRow, top10.length);

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

function listAllListsWithEntries() {
  var cfg   = readConfig();
  var range = resolveDateRange(cfg);
  SpreadsheetApp.getActive().toast('Scanning entries ' + range.label + '...', 'ClickUp');

  var memberIds = getTeamMemberIds(cfg.token, cfg.teamId);
  if (memberIds.length === 0) throw new Error('No team members found for this Team ID.');

  var entries = getTimeEntries(cfg.token, cfg.teamId, null, range.startMs, range.endMs, memberIds);
  var lists   = {};
  entries.forEach(function(e) {
    var loc    = e.task_location || {};
    var id     = loc.list_id  || (e.task && e.task.list && e.task.list.id)   || 'unknown';
    var name   = loc.list_name || (e.task && e.task.list && e.task.list.name) || '(unknown)';
    var folder = loc.folder_name || (e.task && e.task.folder && e.task.folder.name) || '';
    var space  = loc.space_name  || (e.task && e.task.space  && e.task.space.name)  || '';
    if (!lists[id]) lists[id] = { id: id, name: name, folder: folder, space: space, count: 0, hours: 0 };
    lists[id].count++;
    lists[id].hours += Number(e.duration || 0) / 3600000;
  });

  var rows = Object.keys(lists).map(function(k){ return lists[k]; }).sort(function(a, b){ return b.count - a.count; });

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LISTS_SHEET);
  if (!sheet) sheet = ss.insertSheet(LISTS_SHEET);
  sheet.clear();
  var header = ['List name', 'List ID', 'Folder', 'Space', '# entries', 'Total hours', 'Range', 'Display Label'];
  sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  if (rows.length > 0) {
    var data = rows.map(function(r){
      return [r.name, r.id, r.folder, r.space, r.count, Math.round(r.hours * 100) / 100, range.label, buildListLabel_(r)];
    });
    sheet.getRange(2, 1, data.length, header.length).setValues(data);
  }
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, header.length);
  applyListIdDropdown_(rows);
  SpreadsheetApp.getActive().toast('Found ' + rows.length + ' Lists. Config "List ID" dropdown updated.', 'ClickUp', 6);
}

function buildListLabel_(r) {
  var path = [];
  if (r.space)  path.push(r.space);
  if (r.folder) path.push(r.folder);
  return path.length > 0 ? r.name + ' (' + path.join(' > ') + ')' : r.name;
}

function applyListIdDropdown_(rows) {
  if (!rows || rows.length === 0) return;
  var labels      = rows.map(function(r){ return buildListLabel_(r); });
  var configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  if (!configSheet) return;
  configSheet.getRange('B4').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(labels, true).setAllowInvalid(false).build()
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
  if (EDITABLE_COLS.indexOf(col) === -1) return;
  recomputePendingForRow_(sheet, row);
}

function recomputePendingForRow_(sheet, row) {
  var rowValues  = sheet.getRange(row, 1, 1, COLUMNS.length).getValues()[0];
  var snapshotJson = rowValues[SNAPSHOT_COL - 1];
  if (!snapshotJson) { sheet.getRange(row, PENDING_COL).setValue(''); return; }
  var snap;
  try { snap = JSON.parse(snapshotJson); } catch (err) { snap = null; }
  if (!snap) { sheet.getRange(row, PENDING_COL).setValue('?'); return; }

  var currentDesc     = String(rowValues[DESCRIPTION_COL - 1] || '');
  var currentCategory = String(rowValues[LABELS_COL - 1]      || '');
  var currentBillable = rowValues[BILLABLE_COL - 1] === true;
  var diffs = [];
  if (currentDesc     !== String(snap.description || ''))           diffs.push('Desc');
  if (normalizeTagString_(currentCategory) !== normalizeTagString_(snap.tags || '')) diffs.push('Tags');
  if (currentBillable !== (snap.billable === true))                  diffs.push('Billable');
  sheet.getRange(row, PENDING_COL).setValue(diffs.join(', '));
}

function normalizeTagString_(s) {
  return String(s || '').split(',').map(function(t){ return t.trim(); }).filter(Boolean).sort().join(',');
}

function countPendingRows_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  var pending = sheet.getRange(2, PENDING_COL, sheet.getLastRow() - 1, 1).getValues();
  return pending.filter(function(r){ return r[0] && String(r[0]).length > 0; }).length;
}

// ---------- Sync ----------

function collectChanges_(requireConfirm) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DATA_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return { sheet: sheet, changes: [] };

  var lastRow = sheet.getLastRow();
  var data    = sheet.getRange(2, 1, lastRow - 1, COLUMNS.length).getValues();
  var changes = [];
  data.forEach(function(r, idx) {
    var pending = String(r[PENDING_COL - 1] || '').trim();
    if (!pending) return;
    if (requireConfirm && r[CONFIRM_COL - 1] !== true) return;
    var snap = null;
    try { snap = JSON.parse(r[SNAPSHOT_COL - 1] || '{}'); } catch (err) { snap = null; }
    changes.push({
      rowInSheet:  idx + 2,
      entryId:     String(r[ENTRY_ID_COL - 1] || '').trim(),
      taskId:      String(r[1] || ''),
      taskName:    String(r[2] || ''),
      pending:     pending,
      snap:        snap,
      newDesc:     String(r[DESCRIPTION_COL - 1] || ''),
      newTags:     String(r[LABELS_COL - 1]      || ''),
      newBillable: r[BILLABLE_COL - 1] === true,
    });
  });
  return { sheet: sheet, changes: changes };
}

function executeSyncChanges_(changes, sheet) {
  var cfg     = readConfig();
  var tagMaps = getTagMaps_();
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
      sheet.getRange(c.rowInSheet, SNAPSHOT_COL).setValue(newSnap);
      sheet.getRange(c.rowInSheet, PENDING_COL).setValue('');
      sheet.getRange(c.rowInSheet, CONFIRM_COL).setValue(false);
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
  if (result.changes.length === 0) { SpreadsheetApp.getActive().toast('No confirmed pending changes.', 'ClickUp'); return; }
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
  var resp = ui.alert('Sync ' + result.changes.length + ' change(s) to ClickUp?', lines.join('\n\n'), ui.ButtonSet.OK_CANCEL);
  if (resp !== ui.Button.OK) { SpreadsheetApp.getActive().toast('Sync cancelled.', 'ClickUp'); return; }
  var outcome = executeSyncChanges_(result.changes, result.sheet);
  var msg = 'Sync complete: ' + outcome.successCount + ' succeeded, ' + outcome.failCount + ' failed.';
  if (outcome.failCount > 0) msg += ' See "' + CHANGE_LOG_SHEET + '" tab.';
  SpreadsheetApp.getActive().toast(msg, 'ClickUp', 8);
}

function syncAndReload() {
  var result = collectChanges_(false);
  if (result.changes.length > 0) {
    SpreadsheetApp.getActive().toast('Syncing ' + result.changes.length + ' change(s)...', 'ClickUp');
    var outcome = executeSyncChanges_(result.changes, result.sheet);
    SpreadsheetApp.getActive().toast(outcome.successCount + ' synced, ' + outcome.failCount + ' failed. Refreshing...', 'ClickUp', 3);
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
  var lastRow  = sheet.getLastRow();
  var data     = sheet.getRange(2, 1, lastRow - 1, COLUMNS.length).getValues();
  var reverted = 0;
  data.forEach(function(r, idx) {
    var pending = String(r[PENDING_COL - 1] || '').trim();
    if (!pending) return;
    var snap;
    try { snap = JSON.parse(r[SNAPSHOT_COL - 1] || '{}'); } catch (err) { snap = null; }
    if (!snap) return;
    var rowNum = idx + 2;
    sheet.getRange(rowNum, DESCRIPTION_COL).setValue(snap.description || '');
    sheet.getRange(rowNum, LABELS_COL).setValue(snap.tags || '');
    sheet.getRange(rowNum, BILLABLE_COL).setValue(snap.billable === true);
    sheet.getRange(rowNum, PENDING_COL).setValue('');
    sheet.getRange(rowNum, CONFIRM_COL).setValue(false);
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
    var range = sheet.getRange(row, 1, 1, CONFIRM_COL);
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
