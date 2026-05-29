/**
 * ClickUp Time Entries → Google Sheet (with confirm-before-sync)
 *
 * Workflow:
 *   1. Refresh time entries → loads data, snapshots original values.
 *   2. Edit any of: Description, Labels, Billable → row gets a "Pending" status.
 *   3. Tick the Confirm checkbox on rows you want to send.
 *   4. Run "Sync pending changes" → confirmation dialog → API calls.
 *
 * Setup once:
 *   - "Setup config sheet"
 *   - Fill in token / Team ID / List ID
 *   - "Refresh tag list"
 *   - "Setup two-way sync" (installable onEdit trigger)
 */

// ---------- Constants ----------

const CONFIG_SHEET = 'Config';
const DATA_SHEET = 'Time Entries';
const LISTS_SHEET = 'Lists Found';
const TAGS_SHEET = 'Tags';
const CHANGE_LOG_SHEET = 'Change Log';
const CHANGE_LOG_MAX_ROWS = 5000;
const CLICKUP_BASE = 'https://api.clickup.com/api/v2';

const PRESETS = ['Current month', 'Previous month', 'Current quarter', 'Previous quarter', 'Custom'];
const BILLABLE_FILTERS = ['All', 'Billable only', 'Non-billable only'];

const COLUMNS = [
  'Date',            // 1
  'Task ID',         // 2
  'Task Name',       // 3
  'Description',     // 4   editable, syncable
  'Time (hours)',    // 5
  'User',            // 6
  'Labels (Tags)',   // 7   editable, syncable
  'Billable',        // 8   editable, syncable
  'Pending',         // 9   read-only status text
  'Confirm',         // 10  checkbox
  'Entry ID',        // 11  hidden
  'Snapshot',        // 12  hidden, JSON of original {description, tags, billable}
];

const COLUMN_WIDTHS = [100, 110, 280, 400, 90, 200, 220, 80, 130, 80, 120, 120];
const WRAP_COLUMNS = [3, 4]; // Task Name, Description

const DESCRIPTION_COL = 4;
const LABELS_COL = 7;
const BILLABLE_COL = 8;
const PENDING_COL = 9;
const CONFIRM_COL = 10;
const ENTRY_ID_COL = 11;
const SNAPSHOT_COL = 12;

// Editable columns that contribute to pending state
const EDITABLE_COLS = [DESCRIPTION_COL, LABELS_COL, BILLABLE_COL];

// Config row for "Last synced"
const LAST_SYNCED_ROW = 10;

// ---------- Menu ----------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ClickUp')
    .addItem('Refresh time entries', 'refreshTimeEntries')
    .addItem('Refresh tag list', 'refreshTagList')
    .addItem('List all Lists with time entries', 'listAllListsWithEntries')
    .addSeparator()
    .addItem('Sync pending changes', 'syncPendingChanges')
    .addItem('Sync & Reload', 'syncAndReload')
    .addItem('Discard pending changes', 'discardPendingChanges')
    .addSeparator()
    .addItem('Setup config sheet', 'setupConfigSheet')
    .addItem('Setup two-way sync', 'setupTwoWaySync')
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
    ['API Token', '', 'Your ClickUp personal API token (pk_...)'],
    ['Team ID', '', 'Workspace ID from app.clickup.com/{team_id}/...'],
    ['List ID', '', 'Run "List all Lists" first, then pick from dropdown'],
    ['Preset', 'Previous month', 'Pick from dropdown'],
    ['Custom start date', '', 'Only used if Preset = Custom (YYYY-MM-DD)'],
    ['Custom end date', '', 'Only used if Preset = Custom (YYYY-MM-DD), inclusive'],
    ['Include subtasks', 'Yes', 'Yes / No'],
    ['Billable filter', 'All', 'All / Billable only / Non-billable only'],
    ['Last synced', '', 'Auto-updated after a successful sync'],
  ];

  const existing = {};
  if (!isNew && sheet.getLastRow() > 0) {
    const data = sheet.getRange(1, 1, sheet.getLastRow(), 3).getValues();
    data.forEach((r, i) => {
      const key = String(r[0] || '').trim();
      if (key) existing[key] = { value: r[1] };
    });
  }

  const merged = SPEC.map((row, i) => {
    if (i === 0) return row;
    const key = row[0];
    if (existing[key] !== undefined) return [key, existing[key].value, row[2]];
    return row;
  });

  sheet.clear();
  sheet.getRange(1, 1, merged.length, 3).setValues(merged);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 240);
  sheet.setColumnWidth(3, 380);

  sheet.getRange('B5').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(PRESETS, true).build()
  );
  sheet.getRange('B8').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build()
  );
  sheet.getRange('B9').setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(BILLABLE_FILTERS, true).build()
  );

  const preserved = Object.keys(existing).filter(k => k !== 'Setting').length;
  SpreadsheetApp.getActive().toast(
    isNew ? 'Config sheet created. Fill in the values.' : 'Config refreshed. Preserved ' + preserved + ' existing value(s).',
    'ClickUp'
  );
}

function readConfig() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET);
  if (!sheet) throw new Error('No Config sheet. Run "Setup config sheet" first.');
  const values = sheet.getRange(2, 1, 9, 2).getValues();
  const map = {};
  values.forEach(([k, v]) => { map[k] = v; });

  var rawListId = String(map['List ID'] || '').trim();

  const cfg = {
    token: String(map['API Token'] || '').trim(),
    teamId: String(map['Team ID'] || '').trim(),
    listId: resolveListId_(rawListId),
    listLabel: rawListId,
    preset: String(map['Preset'] || '').trim(),
    customStart: map['Custom start date'],
    customEnd: map['Custom end date'],
    includeSubtasks: String(map['Include subtasks'] || 'Yes').trim().toLowerCase() === 'yes',
    billableFilter: String(map['Billable filter'] || 'All').trim(),
  };

  if (!cfg.token) throw new Error('Missing API Token in Config.');
  if (!cfg.teamId) throw new Error('Missing Team ID in Config.');
  if (!PRESETS.includes(cfg.preset)) throw new Error('Preset must be one of: ' + PRESETS.join(', '));
  if (!BILLABLE_FILTERS.includes(cfg.billableFilter)) throw new Error('Billable filter must be one of: ' + BILLABLE_FILTERS.join(', '));
  return cfg;
}

/**
 * Resolves the List ID value from Config B4.
 * If it looks like a display label from the dropdown, looks up the ID in Lists Found.
 * Returns the numeric List ID string, or empty string if blank.
 */
function resolveListId_(raw) {
  if (!raw) return '';
  // If it's purely numeric, treat as a raw List ID (backward compat)
  if (/^\d+$/.test(raw)) return raw;
  // Otherwise it's a display label — look it up in Lists Found
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var listsSheet = ss.getSheetByName(LISTS_SHEET);
  if (!listsSheet || listsSheet.getLastRow() < 2) {
    throw new Error('Lists Found sheet is empty. Run "List all Lists with time entries" first.');
  }
  var lastRow = listsSheet.getLastRow();
  var data = listsSheet.getRange(2, 1, lastRow - 1, 8).getValues(); // cols: name, id, folder, space, ..., ..., ..., label
  for (var i = 0; i < data.length; i++) {
    var label = String(data[i][7] || ''); // column 8 = Display Label
    if (label === raw) return String(data[i][1]); // column 2 = List ID
  }
  throw new Error('Could not find a matching List for "' + raw + '" in the Lists Found sheet. Try running "List all Lists with time entries" again.');
}

// ---------- Date range ----------

function resolveDateRange(cfg) {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const now = new Date();
  let start, end;
  switch (cfg.preset) {
    case 'Current month':
      start = new Date(now.getFullYear(), now.getMonth(), 1);
      end = new Date(now.getFullYear(), now.getMonth() + 1, 1);
      break;
    case 'Previous month':
      start = new Date(now.getFullYear(), now.getMonth() - 1, 1);
      end = new Date(now.getFullYear(), now.getMonth(), 1);
      break;
    case 'Current quarter': {
      const q = Math.floor(now.getMonth() / 3);
      start = new Date(now.getFullYear(), q * 3, 1);
      end = new Date(now.getFullYear(), q * 3 + 3, 1);
      break;
    }
    case 'Previous quarter': {
      const q = Math.floor(now.getMonth() / 3) - 1;
      const year = q < 0 ? now.getFullYear() - 1 : now.getFullYear();
      const qIdx = (q + 4) % 4;
      start = new Date(year, qIdx * 3, 1);
      end = new Date(year, qIdx * 3 + 3, 1);
      break;
    }
    case 'Custom':
      if (!cfg.customStart || !cfg.customEnd) throw new Error('Custom preset requires Custom start and end dates.');
      start = cfg.customStart instanceof Date ? cfg.customStart : new Date(cfg.customStart);
      const incEnd = cfg.customEnd instanceof Date ? cfg.customEnd : new Date(cfg.customEnd);
      end = new Date(incEnd.getFullYear(), incEnd.getMonth(), incEnd.getDate() + 1);
      break;
  }
  return {
    startMs: start.getTime(),
    endMs: end.getTime() - 1,
    label: Utilities.formatDate(start, tz, 'yyyy-MM-dd') + ' \u2192 ' + Utilities.formatDate(new Date(end.getTime() - 1), tz, 'yyyy-MM-dd'),
  };
}

// ---------- ClickUp API ----------

function cuFetch(path, token, params) {
  var qs = '';
  if (params) {
    var parts = [];
    Object.keys(params).forEach(function(k) {
      if (params[k] !== undefined && params[k] !== null && params[k] !== '') {
        parts.push(encodeURIComponent(k) + '=' + encodeURIComponent(params[k]));
      }
    });
    if (parts.length) qs = '?' + parts.join('&');
  }
  return cuRequest('get', path + qs, token);
}

function cuPut(path, token, payload) { return cuRequest('put', path, token, payload); }
function cuPost(path, token, payload) { return cuRequest('post', path, token, payload); }
function cuDelete(path, token, payload) { return cuRequest('delete', path, token, payload); }

function cuRequest(method, path, token, payload) {
  var opts = { method: method, headers: { Authorization: token }, muteHttpExceptions: true };
  if (payload !== undefined) {
    opts.contentType = 'application/json';
    opts.payload = JSON.stringify(payload);
  }
  var res = UrlFetchApp.fetch(CLICKUP_BASE + path, opts);
  var code = res.getResponseCode();
  var body = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('ClickUp API ' + code + ' on ' + method.toUpperCase() + ' ' + path + ': ' + body);
  }
  return body ? JSON.parse(body) : {};
}

function getTeamMemberIds(token, teamId) {
  var data = cuFetch('/team/' + teamId, token);
  var members = (data.team && data.team.members) || [];
  return members.map(function(m) { return m.user && m.user.id; }).filter(Boolean);
}

function getAllWorkspaceTags(token, teamId) {
  var data = cuFetch('/team/' + teamId + '/time_entries/tags', token);
  return Array.isArray(data.data) ? data.data : [];
}

function getTimeEntries(token, teamId, listId, startMs, endMs, assigneeIds) {
  var chunkSize = 100;
  var all = [];
  for (var i = 0; i < assigneeIds.length; i += chunkSize) {
    var chunk = assigneeIds.slice(i, i + chunkSize);
    var params = {
      start_date: startMs,
      end_date: endMs,
      assignee: chunk.join(','),
      include_task_tags: 'true',
      include_location_names: 'true',
    };
    if (listId) params.list_id = listId;
    var data = cuFetch('/team/' + teamId + '/time_entries', token, params);
    if (data.data) all = all.concat(data.data);
  }
  return all;
}

// ---------- Transform ----------

function entryToRow(e, tz) {
  var startDate = new Date(Number(e.start));
  var durHours = Number(e.duration || 0) / 3600000;
  var task = e.task || {};
  var tags = Array.isArray(e.tags) ? e.tags.map(function(t){ return t.name; }).join(', ') : '';
  var user = e.user && (e.user.username || e.user.email) || '';
  var displayId = task.custom_id || task.id || '';
  var billable = e.billable === true;
  var snapshot = JSON.stringify({
    description: e.description || '',
    tags: tags,
    billable: billable,
  });
  return [
    Utilities.formatDate(startDate, tz, 'yyyy-MM-dd'),
    displayId,
    task.name || '',
    e.description || '',
    Math.round(durHours * 100) / 100,
    user,
    tags,
    billable,
    '',           // Pending — empty
    false,        // Confirm — unchecked
    e.id || '',
    snapshot,
  ];
}

// ---------- Refresh ----------

function refreshTimeEntries(skipPendingCheck) {
  // Check for pending changes first (unless called from syncAndReload)
  if (!skipPendingCheck) {
    var pendingCount = countPendingRows_();
    if (pendingCount > 0) {
      var ui = SpreadsheetApp.getUi();
      var resp = ui.alert(
        'Pending changes',
        'You have ' + pendingCount + ' row(s) with pending changes. Refreshing will discard them.\n\n' +
        'YES = Sync first (run "Sync pending changes" before refreshing).\n' +
        'NO = Refresh anyway (discards pending edits).\n' +
        'CANCEL = Stop, do nothing.',
        ui.ButtonSet.YES_NO_CANCEL
      );
      if (resp === ui.Button.YES) { syncPendingChanges(); return; }
      if (resp !== ui.Button.NO) return;
    }
  }

  var cfg = readConfig();
  if (!cfg.listId) throw new Error('Missing List ID in Config. Use "List all Lists with time entries" to find one.');

  var range = resolveDateRange(cfg);
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  SpreadsheetApp.getActive().toast('Fetching ' + range.label + '...', 'ClickUp');

  var memberIds = getTeamMemberIds(cfg.token, cfg.teamId);
  if (memberIds.length === 0) throw new Error('No team members found for this Team ID.');

  var entries = getTimeEntries(cfg.token, cfg.teamId, cfg.listId, range.startMs, range.endMs, memberIds);

  var totalFetched = entries.length;
  if (cfg.billableFilter === 'Billable only') entries = entries.filter(function(e){ return e.billable === true; });
  else if (cfg.billableFilter === 'Non-billable only') entries = entries.filter(function(e){ return e.billable !== true; });

  entries.sort(function(a, b) { return Number(a.start) - Number(b.start); });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DATA_SHEET);
  if (!sheet) sheet = ss.insertSheet(DATA_SHEET);
  sheet.clear();
  sheet.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS]).setFontWeight('bold');

  if (entries.length > 0) {
    var rows = entries.map(function(e){ return entryToRow(e, tz); });
    sheet.getRange(2, 1, rows.length, COLUMNS.length).setValues(rows);
    sheet.getRange(2, BILLABLE_COL, rows.length, 1).insertCheckboxes();
    sheet.getRange(2, CONFIRM_COL, rows.length, 1).insertCheckboxes();
    applyLabelsDropdown(sheet, 2, rows.length);
  }

  sheet.setFrozenRows(1);
  for (var c = 0; c < COLUMN_WIDTHS.length; c++) sheet.setColumnWidth(c + 1, COLUMN_WIDTHS[c]);
  sheet.hideColumns(ENTRY_ID_COL);
  sheet.hideColumns(SNAPSHOT_COL);
  WRAP_COLUMNS.forEach(function(col) {
    sheet.getRange(1, col, Math.max(sheet.getMaxRows(), 1), 1).setWrap(true);
  });

  var filterNote = cfg.billableFilter !== 'All' ? ' [' + cfg.billableFilter + ': ' + entries.length + '/' + totalFetched + ']' : '';
  SpreadsheetApp.getActive().toast(entries.length + ' entries loaded (' + range.label + ')' + filterNote + '.', 'ClickUp', 5);
}

// ---------- Lists discovery ----------

function listAllListsWithEntries() {
  var cfg = readConfig();
  var range = resolveDateRange(cfg);
  SpreadsheetApp.getActive().toast('Scanning entries ' + range.label + '...', 'ClickUp');

  var memberIds = getTeamMemberIds(cfg.token, cfg.teamId);
  if (memberIds.length === 0) throw new Error('No team members found for this Team ID.');

  var entries = getTimeEntries(cfg.token, cfg.teamId, null, range.startMs, range.endMs, memberIds);

  var lists = {};
  entries.forEach(function(e) {
    var loc = e.task_location || {};
    var id = loc.list_id || (e.task && e.task.list && e.task.list.id) || 'unknown';
    var name = loc.list_name || (e.task && e.task.list && e.task.list.name) || '(unknown)';
    var folder = loc.folder_name || (e.task && e.task.folder && e.task.folder.name) || '';
    var space = loc.space_name || (e.task && e.task.space && e.task.space.name) || '';
    if (!lists[id]) lists[id] = { id: id, name: name, folder: folder, space: space, count: 0, hours: 0 };
    lists[id].count += 1;
    lists[id].hours += Number(e.duration || 0) / 3600000;
  });

  var rows = Object.keys(lists).map(function(k){ return lists[k]; }).sort(function(a, b){ return b.count - a.count; });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
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

  // Populate Config B4 dropdown with display labels
  applyListIdDropdown_(rows);

  SpreadsheetApp.getActive().toast('Found ' + rows.length + ' Lists. See "' + LISTS_SHEET + '" tab. Config "List ID" dropdown updated.', 'ClickUp', 6);
}

/**
 * Build a human-readable label for a list entry: "Name (Space > Folder)".
 */
function buildListLabel_(r) {
  var path = [];
  if (r.space) path.push(r.space);
  if (r.folder) path.push(r.folder);
  if (path.length > 0) return r.name + ' (' + path.join(' > ') + ')';
  return r.name;
}

/**
 * Apply the List ID dropdown on Config B4 using display labels from Lists Found rows.
 */
function applyListIdDropdown_(rows) {
  if (!rows || rows.length === 0) return;
  var labels = rows.map(function(r){ return buildListLabel_(r); });
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(CONFIG_SHEET);
  if (!configSheet) return;
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(labels, true)
    .setAllowInvalid(false) // dropdown only — reject manual input
    .build();
  configSheet.getRange('B4').setDataValidation(rule);
}

// ---------- Tag list ----------

function refreshTagList() {
  var cfg = readConfig();
  SpreadsheetApp.getActive().toast('Fetching workspace tags...', 'ClickUp');

  var tags = getAllWorkspaceTags(cfg.token, cfg.teamId);
  var names = tags.map(function(t){ return t.name; })
    .filter(function(n){ return n && n.length > 0; })
    .sort(function(a, b){ return a.toLowerCase().localeCompare(b.toLowerCase()); });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TAGS_SHEET);
  if (!sheet) sheet = ss.insertSheet(TAGS_SHEET);

  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p){ p.remove(); });
  sheet.clear();
  sheet.getRange(1, 1).setValue('Tag name').setFontWeight('bold');
  if (names.length > 0) sheet.getRange(2, 1, names.length, 1).setValues(names.map(function(n){ return [n]; }));
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 240);

  var protection = sheet.protect().setDescription('Tags list — managed by script');
  protection.setWarningOnly(false);
  var me = Session.getEffectiveUser();
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors().filter(function(u){ return u.getEmail() !== me.getEmail(); }));
  if (protection.canDomainEdit()) protection.setDomainEdit(false);

  SpreadsheetApp.getActive().toast('Loaded ' + names.length + ' tag(s). Re-run "Refresh time entries" to apply dropdowns.', 'ClickUp', 6);
}

function applyLabelsDropdown(dataSheet, firstRow, numRows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tagSheet = ss.getSheetByName(TAGS_SHEET);
  if (!tagSheet) return;
  var lastRow = tagSheet.getLastRow();
  if (lastRow < 2) return;
  var values = tagSheet.getRange(2, 1, lastRow - 1, 1).getValues().map(function(r){ return r[0]; }).filter(Boolean);
  if (values.length === 0) return;
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(true)
    .build();
  dataSheet.getRange(firstRow, LABELS_COL, numRows, 1).setDataValidation(rule);
}

// ---------- Edit handler: marks rows as pending ----------

function onClickUpEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== DATA_SHEET) return;
  var col = e.range.getColumn();
  var row = e.range.getRow();
  if (row < 2) return;
  if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;
  if (EDITABLE_COLS.indexOf(col) === -1) return;

  // Recompute pending state from snapshot vs current values
  recomputePendingForRow_(sheet, row);
}

function recomputePendingForRow_(sheet, row) {
  var rowValues = sheet.getRange(row, 1, 1, COLUMNS.length).getValues()[0];
  var snapshotJson = rowValues[SNAPSHOT_COL - 1];
  if (!snapshotJson) {
    sheet.getRange(row, PENDING_COL).setValue('');
    return;
  }
  var snap;
  try { snap = JSON.parse(snapshotJson); } catch (err) { snap = null; }
  if (!snap) {
    sheet.getRange(row, PENDING_COL).setValue('?');
    return;
  }

  var currentDesc = String(rowValues[DESCRIPTION_COL - 1] || '');
  var currentTags = String(rowValues[LABELS_COL - 1] || '');
  var currentBillable = rowValues[BILLABLE_COL - 1] === true;

  var diffs = [];
  if (currentDesc !== String(snap.description || '')) diffs.push('Desc');
  if (normalizeTagString_(currentTags) !== normalizeTagString_(snap.tags || '')) diffs.push('Tags');
  if (currentBillable !== (snap.billable === true)) diffs.push('Billable');

  sheet.getRange(row, PENDING_COL).setValue(diffs.join(', '));
}

function normalizeTagString_(s) {
  return String(s || '').split(',').map(function(t){ return t.trim(); }).filter(Boolean).sort().join(',');
}

function countPendingRows_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  var lastRow = sheet.getLastRow();
  var pending = sheet.getRange(2, PENDING_COL, lastRow - 1, 1).getValues();
  return pending.filter(function(r){ return r[0] && String(r[0]).length > 0; }).length;
}

// ---------- Sync ----------

/**
 * Collect pending changes from the Time Entries sheet.
 * @param {boolean} requireConfirm — if true, only rows with Confirm=true; if false, all pending rows.
 * @returns {{sheet: Sheet, changes: Object[]}}
 */
function collectChanges_(requireConfirm) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DATA_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return { sheet: sheet, changes: [] };

  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, COLUMNS.length).getValues();
  var changes = [];
  data.forEach(function(r, idx) {
    var pending = String(r[PENDING_COL - 1] || '').trim();
    if (!pending) return;
    if (requireConfirm && r[CONFIRM_COL - 1] !== true) return;
    var snap = null;
    try { snap = JSON.parse(r[SNAPSHOT_COL - 1] || '{}'); } catch (err) { snap = null; }
    changes.push({
      rowInSheet: idx + 2,
      entryId: String(r[ENTRY_ID_COL - 1] || '').trim(),
      taskId: String(r[1] || ''),
      taskName: String(r[2] || ''),
      pending: pending,
      snap: snap,
      newDesc: String(r[DESCRIPTION_COL - 1] || ''),
      newTags: String(r[LABELS_COL - 1] || ''),
      newBillable: r[BILLABLE_COL - 1] === true,
    });
  });
  return { sheet: sheet, changes: changes };
}

/**
 * Execute the actual API calls for a list of changes.
 * @returns {{successCount: number, failCount: number}}
 */
function executeSyncChanges_(changes, sheet) {
  var cfg = readConfig();
  var successCount = 0, failCount = 0;

  changes.forEach(function(c) {
    if (!c.entryId) {
      failCount++;
      logChange_('Failure', c.entryId, c.taskId, c.taskName, 'all', '', '(skipped)', 'Missing Entry ID');
      return;
    }

    var parts = c.pending.split(',').map(function(s){ return s.trim(); });
    var rowOK = true;

    var put = {};
    if (parts.indexOf('Billable') !== -1) put.billable = c.newBillable;
    if (parts.indexOf('Desc') !== -1) put.description = c.newDesc;
    if (Object.keys(put).length > 0) {
      var putErr = null;
      try {
        cuPut('/team/' + cfg.teamId + '/time_entries/' + c.entryId, cfg.token, put);
      } catch (err) {
        rowOK = false;
        putErr = err.message;
      }
      if (put.billable !== undefined) {
        logChange_(putErr ? 'Failure' : 'Success', c.entryId, c.taskId, c.taskName, 'Billable',
          c.snap ? (c.snap.billable ? 'Yes' : 'No') : '', put.billable ? 'Yes' : 'No', putErr || '');
      }
      if (put.description !== undefined) {
        logChange_(putErr ? 'Failure' : 'Success', c.entryId, c.taskId, c.taskName, 'Description',
          c.snap ? (c.snap.description || '') : '', put.description, putErr || '');
      }
    }

    if (parts.indexOf('Tags') !== -1) {
      var oldTags = parseTagList_(c.snap ? c.snap.tags : '');
      var newTags = parseTagList_(c.newTags);
      var oldSet = {}; oldTags.forEach(function(t){ oldSet[t] = true; });
      var newSet = {}; newTags.forEach(function(t){ newSet[t] = true; });
      var toAdd = newTags.filter(function(t){ return !oldSet[t]; });
      var toRemove = oldTags.filter(function(t){ return !newSet[t]; });

      toRemove.forEach(function(tag) {
        var tErr = null;
        try {
          cuDelete('/team/' + cfg.teamId + '/time_entries/tags', cfg.token, {
            time_entry_ids: [c.entryId], tags: [{ name: tag }],
          });
        } catch (err) { rowOK = false; tErr = err.message; }
        logChange_(tErr ? 'Failure' : 'Success', c.entryId, c.taskId, c.taskName, 'Labels (remove)', tag, '', tErr || '');
      });
      toAdd.forEach(function(tag) {
        var tErr = null;
        try {
          cuPost('/team/' + cfg.teamId + '/time_entries/tags', cfg.token, {
            time_entry_ids: [c.entryId], tags: [{ name: tag }],
          });
        } catch (err) { rowOK = false; tErr = err.message; }
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

/**
 * Sync with confirmation dialog. Only syncs rows that are Pending + Confirmed.
 */
function syncPendingChanges() {
  var result = collectChanges_(true);
  if (result.changes.length === 0) {
    SpreadsheetApp.getActive().toast('No confirmed pending changes.', 'ClickUp');
    return;
  }

  var ui = SpreadsheetApp.getUi();
  var lines = [];
  result.changes.slice(0, 10).forEach(function(c) {
    var parts = c.pending.split(',').map(function(s){ return s.trim(); });
    var summary = parts.map(function(p) {
      if (p === 'Desc') {
        var oldD = c.snap ? truncate_(c.snap.description, 40) : '(unknown)';
        return 'Desc: "' + oldD + '" \u2192 "' + truncate_(c.newDesc, 40) + '"';
      }
      if (p === 'Tags') {
        var oldT = c.snap ? (c.snap.tags || '(none)') : '(unknown)';
        return 'Tags: [' + oldT + '] \u2192 [' + (c.newTags || '(none)') + ']';
      }
      if (p === 'Billable') {
        var oldB = c.snap ? (c.snap.billable ? 'Yes' : 'No') : '(unknown)';
        return 'Billable: ' + oldB + ' \u2192 ' + (c.newBillable ? 'Yes' : 'No');
      }
      return p;
    }).join('; ');
    lines.push('Row ' + c.rowInSheet + ' (' + c.entryId + '): ' + summary);
  });
  if (result.changes.length > 10) lines.push('... and ' + (result.changes.length - 10) + ' more.');

  var resp = ui.alert(
    'Sync ' + result.changes.length + ' change(s) to ClickUp?',
    lines.join('\n\n'),
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) {
    SpreadsheetApp.getActive().toast('Sync cancelled.', 'ClickUp');
    return;
  }

  var outcome = executeSyncChanges_(result.changes, result.sheet);
  var msg = 'Sync complete: ' + outcome.successCount + ' succeeded, ' + outcome.failCount + ' failed.';
  if (outcome.failCount > 0) msg += ' See "' + CHANGE_LOG_SHEET + '" tab.';
  SpreadsheetApp.getActive().toast(msg, 'ClickUp', 8);
}

/**
 * Sync all pending rows (ignoring Confirm checkbox), skip dialog, then refresh.
 */
function syncAndReload() {
  var result = collectChanges_(false); // all pending, regardless of Confirm
  if (result.changes.length > 0) {
    SpreadsheetApp.getActive().toast('Syncing ' + result.changes.length + ' change(s)...', 'ClickUp');
    var outcome = executeSyncChanges_(result.changes, result.sheet);
    var msg = outcome.successCount + ' synced, ' + outcome.failCount + ' failed. Refreshing...';
    SpreadsheetApp.getActive().toast(msg, 'ClickUp', 3);
  }
  refreshTimeEntries(true); // skip pending check since we just synced
}

function discardPendingChanges() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DATA_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return;

  var pendingCount = countPendingRows_();
  if (pendingCount === 0) {
    SpreadsheetApp.getActive().toast('Nothing to discard.', 'ClickUp');
    return;
  }

  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'Discard ' + pendingCount + ' pending change(s)?',
    'Edits in the sheet will be reverted to the values currently in ClickUp (per snapshot). This cannot be undone.',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) return;

  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, COLUMNS.length).getValues();
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

function parseTagList_(raw) {
  if (raw == null || raw === '') return [];
  return String(raw).split(',').map(function(s){ return s.trim(); }).filter(function(s){ return s.length > 0; });
}

function truncate_(s, n) {
  s = String(s == null ? '' : s);
  return s.length > n ? s.substr(0, n - 1) + '\u2026' : s;
}

function flashRow_(sheet, row) {
  try {
    var range = sheet.getRange(row, 1, 1, CONFIRM_COL);
    var prev = range.getBackgrounds();
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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CHANGE_LOG_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(CHANGE_LOG_SHEET);
      sheet.getRange(1, 1, 1, 9).setValues([[
        'Timestamp', 'Status', 'Entry ID', 'Task ID', 'Task Name',
        'Field', 'Old value', 'New value', 'Error',
      ]]).setFontWeight('bold');
      sheet.setFrozenRows(1);
      // Sensible widths
      var widths = [150, 80, 110, 110, 260, 110, 280, 280, 280];
      for (var c = 0; c < widths.length; c++) sheet.setColumnWidth(c + 1, widths[c]);
    }
    var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    sheet.appendRow([
      Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss'),
      status,
      entryId || '',
      taskId || '',
      taskName || '',
      field || '',
      String(oldVal == null ? '' : oldVal),
      String(newVal == null ? '' : newVal),
      message || '',
    ]);
    // Cap retention: keep header + most recent CHANGE_LOG_MAX_ROWS rows
    var totalRows = sheet.getLastRow();
    var dataRows = totalRows - 1;
    if (dataRows > CHANGE_LOG_MAX_ROWS) {
      sheet.deleteRows(2, dataRows - CHANGE_LOG_MAX_ROWS);
    }
  } catch (err) { /* never let logging break a sync */ }
}

// ---------- Trigger setup ----------

function setupTwoWaySync() {
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'onClickUpEdit') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  ScriptApp.newTrigger('onClickUpEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  SpreadsheetApp.getActive().toast(
    'Trigger installed (removed ' + removed + ' old). Edits now mark rows as Pending.',
    'ClickUp', 6
  );
}
