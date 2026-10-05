/**
 * ClickUp Time Entries → Google Sheet (Report with confirm-before-sync)
 *
 * Version: 2.6.0
 *
 * Workflow:
 *   1. Refresh time entries → loads data into the Report sheet.
 *   2. Edit any of: Work Description, Task Category, Billable → row marked Pending.
 *   3. Tick the Confirm checkbox on rows you want to send.
 *   4. Run "Sync pending changes" → confirmation dialog → API calls.
 *
 * Multi-list: put several Lists in the Config "List ID" cell (smart chips /
 * comma-separated). When more than one List is selected, the Report gains a
 * leading "List" column so each entry shows which List it came from. A single
 * List produces the exact same layout as before (no List column).
 *
 * Setup once:
 *   - "Setup config sheet"  →  fill in token / Team ID / Rate
 *   - "List all Lists"       →  pick a List from the Config dropdown (all Lists, active or not)
 *   - "Refresh tag list"     →  fill in Display Name mappings on Tags sheet
 *   - "Setup two-way sync"   →  installable onEdit trigger
 *
 * Extra Users (2.6.0): ClickUp only returns time entries for the assignee IDs
 * you pass, and /team/{id} lists ACTIVE members only — so time logged by
 * deactivated users / guests is invisible unless their IDs are supplied.
 * Put those IDs in column A of the "Extra Users" tab (auto-created on first
 * Refresh / List all Lists). They are merged into the assignee list.
 */

// ---------- Constants ----------

const VERSION = '2.6.0';


const CONFIG_SHEET = 'Config';
const DATA_SHEET = 'Report';
const LISTS_SHEET = 'Lists Found';
const TAGS_SHEET = 'Tags';
const CHANGE_LOG_SHEET = 'Change Log';
const EXTRA_USERS_SHEET = 'Extra Users';
const CHANGE_LOG_MAX_ROWS = 5000;
const CLICKUP_BASE = 'https://api.clickup.com/api/v2';

const PRESETS = ['Current month', 'Previous month', 'Current quarter', 'Previous quarter', 'Custom'];
const BILLABLE_FILTERS = ['All', 'Billable only', 'Non-billable only'];

const COLUMNS = [
  'Date',              // 1
  'Issue Key',         // 2
  'Issue summary',     // 3
  'Work Description',  // 4   editable, syncable
  'Billed Hours',      // 5
  'Full name',         // 6
  'Task Category',     // 7   editable, syncable
  'Billable',          // 8   editable, syncable
  'Pending',           // 9   read-only status text
  'Confirm',           // 10  checkbox
  'Entry ID',          // 11  hidden
  'Snapshot',          // 12  hidden, JSON of original values
];

const COLUMN_WIDTHS = [100, 100, 280, 400, 110, 200, 200, 80, 130, 80, 120, 120];

// ----- Base (single-list) column layout -----
const DESCRIPTION_COL = 4;
const CATEGORY_COL = 7;
const BILLABLE_COL = 8;
const PENDING_COL = 9;
const CONFIRM_COL = 10;
const ENTRY_ID_COL = 11;
const SNAPSHOT_COL = 12;

const EDITABLE_COLS = [DESCRIPTION_COL, CATEGORY_COL, BILLABLE_COL];

// Header text for the conditional multi-list column (prepended as column 1).
const LIST_COL_HEADER = 'List';

/**
 * Resolve the active column layout for the Report sheet.
 * The Report gains a leading "List" column ONLY when multiple Lists are selected.
 * Single-list Reports are byte-identical to pre-2.5.0 layouts (no List column).
 *
 * Pass a boolean (building a fresh sheet) OR a sheet (reading an existing one:
 * detected from whether row-1 col-1 equals LIST_COL_HEADER). Every sheet-touching
 * function derives its indices from here, so there is a single source of truth and
 * no scattered "+1" offsets to keep in sync.
 *
 * Returns: { multi, off, hoursColLetter, columns, widths, and 1-based indices }.
 */
function getLayout_(multiOrSheet) {
  var multi;
  if (typeof multiOrSheet === 'boolean') {
    multi = multiOrSheet;
  } else if (multiOrSheet && typeof multiOrSheet.getRange === 'function') {
    multi = (multiOrSheet.getLastColumn() >= 1 && multiOrSheet.getLastRow() >= 1)
      ? String(multiOrSheet.getRange(1, 1).getValue()) === LIST_COL_HEADER
      : false;
  } else {
    multi = false;
  }
  var off = multi ? 1 : 0;
  var baseColumns = [
    'Date', 'Issue Key', 'Issue summary', 'Work Description', 'Billed Hours',
    'Full name', 'Task Category', 'Billable', 'Pending', 'Confirm', 'Entry ID', 'Snapshot',
  ];
  var baseWidths = [100, 100, 280, 400, 110, 200, 200, 80, 130, 80, 120, 120];
  var columns = multi ? [LIST_COL_HEADER].concat(baseColumns) : baseColumns.slice();
  var widths = multi ? [160].concat(baseWidths) : baseWidths.slice();
  var hoursColIndex = 5 + off; // "Billed Hours"
  return {
    multi: multi,
    off: off,
    columns: columns,
    widths: widths,
    hoursCol: hoursColIndex,
    hoursColLetter: columnToLetter_(hoursColIndex),
    labelCol: 4 + off, // summary label column (D, or E in multi mode)
    listCol: multi ? 1 : 0, // 0 = absent
    descriptionCol: DESCRIPTION_COL + off,
    categoryCol: CATEGORY_COL + off,
    billableCol: BILLABLE_COL + off,
    pendingCol: PENDING_COL + off,
    confirmCol: CONFIRM_COL + off,
    entryIdCol: ENTRY_ID_COL + off,
    snapshotCol: SNAPSHOT_COL + off,
    numCols: baseColumns.length + off,
  };
}

function columnToLetter_(col) {
  var s = '';
  while (col > 0) { var m = (col - 1) % 26; s = String.fromCharCode(65 + m) + s; col = (col - m - 1) / 26; }
  return s;
}

const LAST_SYNCED_ROW = 10;
const RATE_ROW = 11;

// Header style
const HEADER_BG = '#000000';
const HEADER_FG = '#ffffff';

// ---------- Menu ----------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ClickUp')
    .addItem('Refresh time entries', 'refreshTimeEntries')
    .addItem('Refresh tag list', 'refreshTagList')
    .addItem('List all Lists', 'listAllLists')
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
    ['Rate', '125', 'Hourly rate used in the Report summary block'],
    ['Target Contract Hours', '', 'Optional. If set, summary shows an overage block (billed hours over this target * Rate). Blank = standard Total Due block.'],
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
  const values = sheet.getRange(2, 1, 11, 2).getValues();
  const map = {};
  values.forEach(([k, v]) => { map[k] = v; });

  var rawListId = String(map['List ID'] || '').trim();
  var listSelection = resolveListSelection_(rawListId);

  const cfg = {
    token: String(map['API Token'] || '').trim(),
    teamId: String(map['Team ID'] || '').trim(),
    listId: listSelection.ids.join(','),
    listSelection: listSelection,
    listLabel: rawListId,
    preset: String(map['Preset'] || '').trim(),
    customStart: map['Custom start date'],
    customEnd: map['Custom end date'],
    includeSubtasks: String(map['Include subtasks'] || 'Yes').trim().toLowerCase() === 'yes',
    billableFilter: String(map['Billable filter'] || 'All').trim(),
    rate: parseFloat(map['Rate']) || 0,
    targetHours: (function() {
      var raw = String(map['Target Contract Hours'] == null ? '' : map['Target Contract Hours']).trim();
      if (raw === '') return null;
      var n = parseFloat(raw);
      return (isNaN(n) || n < 0) ? null : n;
    })(),
  };

  if (!cfg.token) throw new Error('Missing API Token in Config.');
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
  if (!listsSheet || listsSheet.getLastRow() < 2) {
    throw new Error('Lists Found sheet is empty. Run "List all Lists" first.');
  }
  var lastRow = listsSheet.getLastRow();
  var data = listsSheet.getRange(2, 1, lastRow - 1, 8).getValues();
  for (var i = 0; i < data.length; i++) {
    var label = String(data[i][7] || '');
    if (label === raw) return String(data[i][1]);
  }
  throw new Error('Could not find a matching List for "' + raw + '". Try running "List all Lists" again.');
}

/**
 * Split a possibly-multi List ID cell (chip mode gives comma-separated values)
 * into its parts, resolve each part to a numeric List ID, and build an
 * ID -> List-name map from the Lists Found sheet for the new "List" column.
 * Returns { ids: [..], count, idToName: {id: name} }.
 */
function resolveListSelection_(raw) {
  var parts = String(raw || '').split(',').map(function(s){ return s.trim(); }).filter(Boolean);
  var ids = parts.map(function(p){ return resolveListId_(p); }).filter(Boolean);
  // Dedupe while preserving order.
  var seen = {}, uniq = [];
  ids.forEach(function(id){ if (!seen[id]) { seen[id] = true; uniq.push(id); } });

  var idToName = {};
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var listsSheet = ss.getSheetByName(LISTS_SHEET);
  if (listsSheet && listsSheet.getLastRow() >= 2) {
    var data = listsSheet.getRange(2, 1, listsSheet.getLastRow() - 1, 8).getValues();
    data.forEach(function(r){ idToName[String(r[1])] = String(r[0] || ''); });
  }
  return { ids: uniq, count: uniq.length, idToName: idToName };
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

// ---------- Extra Users (deactivated / guest assignees) ----------

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
    .setFontWeight('bold').setBackground(HEADER_BG).setFontColor(HEADER_FG);
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
 * Returns an array of { id, name, folder, space } for every non-archived List.
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

// ---------- Tag mapping ----------

/**
 * Read the Tags sheet and build forward/reverse maps.
 * Forward: clickupTagName → displayName (only for tags with a display name set).
 * Reverse: displayName → clickupTagName.
 */
function getTagMaps_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TAGS_SHEET);
  var forward = {}; // clickup name → display name
  var reverse = {}; // display name → clickup name
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
 * Convert ClickUp tag names to display names using forward map.
 * Unmapped tags are hidden (excluded).
 */
function mapTagsForDisplay_(clickupTags, forwardMap) {
  if (!Array.isArray(clickupTags)) return '';
  var mapped = [];
  clickupTags.forEach(function(t) {
    var name = t.name || t;
    var display = forwardMap[name];
    if (display) mapped.push(display);
  });
  return mapped.join(', ');
}

/**
 * Convert display tag string back to ClickUp tag names using reverse map.
 */
function reverseMapTags_(displayString, reverseMap) {
  var displayNames = parseTagList_(displayString);
  return displayNames.map(function(d) {
    return reverseMap[d] || d; // fallback to display name if no reverse mapping
  });
}

// ---------- Transform ----------

function entryToRow(e, tz, tagForwardMap, layout, idToName) {
  var startDate = new Date(Number(e.start));
  var durHours = Number(e.duration || 0) / 3600000;
  var task = e.task || {};
  var user = e.user && (e.user.username || e.user.email) || '';
  var displayId = task.custom_id || task.id || '';
  var billable = e.billable === true;

  // Map tags: only include mapped tags
  var displayTags = mapTagsForDisplay_(e.tags || [], tagForwardMap);

  var snapshot = JSON.stringify({
    description: e.description || '',
    tags: displayTags,
    billable: billable,
  });
  var base = [
    Utilities.formatDate(startDate, tz, 'yyyy-MM-dd'),
    displayId,
    task.name || '',
    e.description || '',
    Math.round(durHours * 100) / 100,
    user,
    displayTags,
    billable,
    '',     // Pending
    false,  // Confirm
    e.id || '',
    snapshot,
  ];
  if (layout && layout.multi) {
    // Resolve this entry's List name for the leading "List" column.
    var loc = e.task_location || {};
    var lid = String(loc.list_id || (task.list && task.list.id) || '');
    var lname = (idToName && idToName[lid]) || loc.list_name || (task.list && task.list.name) || lid || '';
    return [lname].concat(base);
  }
  return base;
}

// ---------- Refresh ----------

function refreshTimeEntries(skipPendingCheck) {
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
  if (!cfg.listId) throw new Error('Missing List ID in Config. Use "List all Lists" to find one.');

  var range = resolveDateRange(cfg);
  var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  SpreadsheetApp.getActive().toast('Fetching ' + range.label + '...', 'ClickUp');

  var memberIds = getTeamMemberIds(cfg.token, cfg.teamId);
  if (memberIds.length === 0) throw new Error('No team members found for this Team ID.');
  var extra = getExtraUserIds_();
  var merged = mergeAssigneeIds_(memberIds, extra.ids);

  var entries = getTimeEntries(cfg.token, cfg.teamId, cfg.listId, range.startMs, range.endMs, merged.ids);

  var totalFetched = entries.length;
  if (cfg.billableFilter === 'Billable only') entries = entries.filter(function(e){ return e.billable === true; });
  else if (cfg.billableFilter === 'Non-billable only') entries = entries.filter(function(e){ return e.billable !== true; });

  entries.sort(function(a, b) { return Number(a.start) - Number(b.start); });

  // Conditional layout: the "List" column appears only when >1 List is selected.
  var layout = getLayout_(cfg.listSelection.count > 1);
  var idToName = cfg.listSelection.idToName;

  // Get tag mapping for display
  var tagMaps = getTagMaps_();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DATA_SHEET);
  if (!sheet) sheet = ss.insertSheet(DATA_SHEET);
  sheet.clear();
  // Wipe ALL data validations and borders across the entire sheet so leftover
  // checkboxes/dropdowns/borders from previous refreshes don't haunt empty rows.
  var fullRange = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  fullRange.clearDataValidations();
  fullRange.setBorder(false, false, false, false, false, false);
  // Apply Anek Tamil 11pt and text wrap to the entire sheet (header, data, summary)
  fullRange.setFontFamily('Anek Tamil').setFontSize(11).setWrap(true);

  // Header row: black background, white text, bold
  sheet.getRange(1, 1, 1, layout.columns.length).setValues([layout.columns])
    .setFontWeight('bold')
    .setBackground(HEADER_BG)
    .setFontColor(HEADER_FG);

  if (entries.length > 0) {
    var rows = entries.map(function(e){ return entryToRow(e, tz, tagMaps.forward, layout, idToName); });
    sheet.getRange(2, 1, rows.length, layout.columns.length).setValues(rows);
    sheet.getRange(2, layout.billableCol, rows.length, 1).insertCheckboxes();
    sheet.getRange(2, layout.confirmCol, rows.length, 1).insertCheckboxes();
    // Force 2-decimal display on Billed Hours column
    sheet.getRange(2, layout.hoursCol, rows.length, 1).setNumberFormat('0.00');
    applyCategoryDropdown_(sheet, 2, rows.length, layout);
  }

  sheet.setFrozenRows(1);
  for (var c = 0; c < layout.widths.length; c++) sheet.setColumnWidth(c + 1, layout.widths[c]);
  sheet.hideColumns(layout.entryIdCol);
  sheet.hideColumns(layout.snapshotCol);

  // Summary block — 3 blank rows of separation after last data row
  var dataRows = entries.length;
  var summaryStartRow = dataRows + 1 + 3 + 1; // header(1) + data + 3 blank rows + 1
  var hoursCol = layout.hoursCol;   // "Billed Hours" (E single-list, F multi-list)
  var labelCol = layout.labelCol;   // summary label column (D or E)
  var HL = layout.hoursColLetter;   // column letter for formulas
  var GREY_BG = '#999999'; // Google Sheets "Dark Gray 2"

  // Row 1 (both modes): Total Support Hours (normal weight, hours format 0.00)
  sheet.getRange(summaryStartRow, labelCol).setValue('Total Support Hours for the Month');
  if (dataRows > 0) {
    sheet.getRange(summaryStartRow, hoursCol)
      .setFormula('=ROUND(SUM(' + HL + '2:' + HL + (dataRows + 1) + '),2)')
      .setNumberFormat('0.00');
  } else {
    sheet.getRange(summaryStartRow, hoursCol).setValue(0).setNumberFormat('0.00');
  }

  var blockRows;

  if (cfg.targetHours == null) {
    // ----- Standard mode: Rate + Total Due -----
    blockRows = 3;

    // Row 2: Rate (normal weight, currency $0.00, Dark Gray 1 background)
    sheet.getRange(summaryStartRow + 1, labelCol).setValue('Rate');
    sheet.getRange(summaryStartRow + 1, hoursCol)
      .setValue(cfg.rate)
      .setNumberFormat('$0.00')
      .setBackground('#b7b7b7');

    // Row 3: Total Due (bold, black bg, white text, currency format $0.00)
    sheet.getRange(summaryStartRow + 2, labelCol).setValue('Total Due');
    sheet.getRange(summaryStartRow + 2, hoursCol)
      .setFormula('=' + HL + summaryStartRow + '*' + HL + (summaryStartRow + 1))
      .setNumberFormat('$0.00');
    sheet.getRange(summaryStartRow + 2, labelCol, 1, 2)
      .setFontWeight('bold')
      .setBackground(HEADER_BG)
      .setFontColor(HEADER_FG);

  } else {
    // ----- Overage mode: Target / Overage (hrs) / Overage ($rate/hr) -----
    blockRows = 4;
    var totalRow = summaryStartRow;
    var targetRow = summaryStartRow + 1;
    var overageHrsRow = summaryStartRow + 2;
    var overageDueRow = summaryStartRow + 3;
    var targetLabelNum = Math.round(cfg.targetHours); // whole number in label

    // Row 2: Target Contract Hours - {N} hrs
    sheet.getRange(targetRow, labelCol).setValue('Target Contract Hours - ' + targetLabelNum + ' hrs');
    sheet.getRange(targetRow, hoursCol).setValue(cfg.targetHours).setNumberFormat('0.00');

    // Row 3: Overage (hrs) = MAX(0, total - target), Dark Gray 2 value cell
    sheet.getRange(overageHrsRow, labelCol).setValue('Overage (hrs)');
    sheet.getRange(overageHrsRow, hoursCol)
      .setFormula('=MAX(0,' + HL + totalRow + '-' + HL + targetRow + ')')
      .setNumberFormat('0.00')
      .setBackground(GREY_BG)
      .setFontColor(HEADER_FG)
      .setFontWeight('bold');

    // Row 4: Overage (${rate}/hr) = overage hrs * rate, bold black/white
    sheet.getRange(overageDueRow, labelCol).setValue('Overage ($' + cfg.rate + '/hr)');
    sheet.getRange(overageDueRow, hoursCol)
      .setFormula('=' + HL + overageHrsRow + '*' + cfg.rate)
      .setNumberFormat('$0.00');
    sheet.getRange(overageDueRow, labelCol, 1, 2)
      .setFontWeight('bold')
      .setBackground(HEADER_BG)
      .setFontColor(HEADER_FG);
  }

  // Outer border only (no internal lines)
  sheet.getRange(summaryStartRow, labelCol, blockRows, 2)
    .setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  var filterNote = cfg.billableFilter !== 'All' ? ' [' + cfg.billableFilter + ': ' + entries.length + '/' + totalFetched + ']' : '';
  var extraNote = extraUsersNote_(extra, merged.addedCount);
  SpreadsheetApp.getActive().toast(entries.length + ' entries loaded (' + range.label + ')' + filterNote + '.' + extraNote,
    'ClickUp', extra.invalid.length > 0 ? 15 : 5);
}

// ---------- Lists discovery ----------

function listAllLists() {
  var cfg = readConfig();
  var range = resolveDateRange(cfg);
  SpreadsheetApp.getActive().toast('Fetching all Lists from the workspace...', 'ClickUp');

  // 1. Pull every non-archived List from the hierarchy.
  var hier = getAllListsHierarchy_(cfg.token, cfg.teamId);

  var lists = {};
  hier.forEach(function(l) {
    lists[l.id] = { id: l.id, name: l.name, folder: l.folder, space: l.space, count: 0, hours: 0 };
  });

  // 2. Second scan: fill entry counts / hours for the period (blank where none).
  SpreadsheetApp.getActive().toast('Scanning entries ' + range.label + ' for counts...', 'ClickUp');
  var memberIds = getTeamMemberIds(cfg.token, cfg.teamId);
  if (memberIds.length === 0) throw new Error('No team members found for this Team ID.');
  var extra = getExtraUserIds_();
  var merged = mergeAssigneeIds_(memberIds, extra.ids);
  var entries = getTimeEntries(cfg.token, cfg.teamId, null, range.startMs, range.endMs, merged.ids);
  entries.forEach(function(e) {
    var loc = e.task_location || {};
    var id = String(loc.list_id || (e.task && e.task.list && e.task.list.id) || '');
    if (!id || !lists[id]) return; // ignore entries whose list isn't in the (non-archived) hierarchy
    lists[id].count += 1;
    lists[id].hours += Number(e.duration || 0) / 3600000;
  });

  // Sort by activity (lists with entries first, desc), then alphabetically.
  var rows = Object.keys(lists).map(function(k){ return lists[k]; }).sort(function(a, b){
    if (b.count !== a.count) return b.count - a.count;
    return buildListLabel_(a).localeCompare(buildListLabel_(b));
  });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(LISTS_SHEET);
  if (!sheet) sheet = ss.insertSheet(LISTS_SHEET);
  sheet.clear();
  var header = ['List name', 'List ID', 'Folder', 'Space', '# entries', 'Total hours', 'Range', 'Display Label'];
  sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold')
    .setBackground(HEADER_BG).setFontColor(HEADER_FG);
  if (rows.length > 0) {
    var data = rows.map(function(r){
      var count = r.count > 0 ? r.count : '';
      var hours = r.count > 0 ? Math.round(r.hours * 100) / 100 : '';
      return [r.name, r.id, r.folder, r.space, count, hours, range.label, buildListLabel_(r)];
    });
    sheet.getRange(2, 1, data.length, header.length).setValues(data);
  }
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, header.length);
  protectSheet_(sheet, 'Lists Found — managed by script');
  applyListIdDropdown_(rows);

  var activeCount = rows.filter(function(r){ return r.count > 0; }).length;
  var extraNote = extraUsersNote_(extra, merged.addedCount);
  SpreadsheetApp.getActive().toast('Found ' + rows.length + ' Lists (' + activeCount + ' with entries ' + range.label + '). See "' + LISTS_SHEET + '" tab. Config dropdown updated.' + extraNote,
    'ClickUp', extra.invalid.length > 0 ? 15 : 6);
}

function buildListLabel_(r) {
  var path = [];
  if (r.space) path.push(r.space);
  if (r.folder) path.push(r.folder);
  if (path.length > 0) return r.name + ' (' + path.join(' > ') + ')';
  return r.name;
}

function applyListIdDropdown_(rows) {
  if (!rows || rows.length === 0) return;
  var labels = rows.map(function(r){ return buildListLabel_(r); });
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var configSheet = ss.getSheetByName(CONFIG_SHEET);
  if (!configSheet) return;
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(labels, true)
    .setAllowInvalid(false)
    .build();
  configSheet.getRange('B4').setDataValidation(rule);
}

// ---------- Sheet protection ----------

function protectSheet_(sheet, description) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p){ p.remove(); });
  var protection = sheet.protect().setDescription(description);
  protection.setWarningOnly(false);
  var me = Session.getEffectiveUser();
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors().filter(function(u){ return u.getEmail() !== me.getEmail(); }));
  if (protection.canDomainEdit()) protection.setDomainEdit(false);
}

// ---------- Tag list ----------

function refreshTagList() {
  var cfg = readConfig();
  SpreadsheetApp.getActive().toast('Fetching workspace tags...', 'ClickUp');

  var clickupTags = getAllWorkspaceTags(cfg.token, cfg.teamId);
  var newNames = clickupTags.map(function(t){ return t.name; })
    .filter(function(n){ return n && n.length > 0; })
    .sort(function(a, b){ return a.toLowerCase().localeCompare(b.toLowerCase()); });

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TAGS_SHEET);
  if (!sheet) sheet = ss.insertSheet(TAGS_SHEET);

  // Read existing mappings before clearing
  var existingMappings = {};
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function(p){ p.remove(); });
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function(p){ p.remove(); });
  if (sheet.getLastRow() >= 2) {
    var old = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
    old.forEach(function(r) {
      var tagName = String(r[0] || '').trim();
      var displayName = String(r[1] || '').trim();
      if (tagName && displayName) existingMappings[tagName] = displayName;
    });
  }

  sheet.clear();
  sheet.getRange(1, 1, 1, 2).setValues([['Tag name', 'Display Name']]).setFontWeight('bold')
    .setBackground(HEADER_BG).setFontColor(HEADER_FG);

  if (newNames.length > 0) {
    var rows = newNames.map(function(n){
      return [n, existingMappings[n] || ''];
    });
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 240);
  sheet.setColumnWidth(2, 240);

  // Protect column A only (tag names), leave column B editable
  var colARange = sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 1), 1);
  var protection = colARange.protect().setDescription('Tag names — managed by script');
  protection.setWarningOnly(false);
  var me = Session.getEffectiveUser();
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors().filter(function(u){ return u.getEmail() !== me.getEmail(); }));
  if (protection.canDomainEdit()) protection.setDomainEdit(false);

  var mapped = Object.keys(existingMappings).length;
  SpreadsheetApp.getActive().toast(
    'Loaded ' + newNames.length + ' tag(s), preserved ' + mapped + ' mapping(s). Fill in Display Name for tags you want in the dropdown.',
    'ClickUp', 8
  );
}

/**
 * Apply multi-select dropdown to the Task Category column using only mapped tags.
 */
function applyCategoryDropdown_(dataSheet, firstRow, numRows, layout) {
  if (!layout) layout = getLayout_(dataSheet);
  var tagMaps = getTagMaps_();
  var displayNames = [];
  // Get display names in order from the Tags sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tagSheet = ss.getSheetByName(TAGS_SHEET);
  if (!tagSheet || tagSheet.getLastRow() < 2) return;
  var data = tagSheet.getRange(2, 1, tagSheet.getLastRow() - 1, 2).getValues();
  data.forEach(function(r) {
    var display = String(r[1] || '').trim();
    if (display) displayNames.push(display);
  });
  if (displayNames.length === 0) return;
  // Sort alphabetically
  displayNames.sort(function(a, b){ return a.toLowerCase().localeCompare(b.toLowerCase()); });
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(displayNames, true)
    .setAllowInvalid(true)
    .build();
  dataSheet.getRange(firstRow, layout.categoryCol, numRows, 1).setDataValidation(rule);
}

// ---------- Edit handler ----------

function onClickUpEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== DATA_SHEET) return;
  var col = e.range.getColumn();
  var row = e.range.getRow();
  if (row < 2) return;
  if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;
  var layout = getLayout_(sheet);
  var editable = [layout.descriptionCol, layout.categoryCol, layout.billableCol];
  if (editable.indexOf(col) === -1) return;
  recomputePendingForRow_(sheet, row, layout);
}

function recomputePendingForRow_(sheet, row, layout) {
  if (!layout) layout = getLayout_(sheet);
  var rowValues = sheet.getRange(row, 1, 1, layout.numCols).getValues()[0];
  var snapshotJson = rowValues[layout.snapshotCol - 1];
  if (!snapshotJson) {
    sheet.getRange(row, layout.pendingCol).setValue('');
    return;
  }
  var snap;
  try { snap = JSON.parse(snapshotJson); } catch (err) { snap = null; }
  if (!snap) {
    sheet.getRange(row, layout.pendingCol).setValue('?');
    return;
  }

  var currentDesc = String(rowValues[layout.descriptionCol - 1] || '');
  var currentTags = String(rowValues[layout.categoryCol - 1] || '');
  var currentBillable = rowValues[layout.billableCol - 1] === true;

  var diffs = [];
  if (currentDesc !== String(snap.description || '')) diffs.push('Desc');
  if (normalizeTagString_(currentTags) !== normalizeTagString_(snap.tags || '')) diffs.push('Category');
  if (currentBillable !== (snap.billable === true)) diffs.push('Billable');

  sheet.getRange(row, layout.pendingCol).setValue(diffs.join(', '));
}

function normalizeTagString_(s) {
  return String(s || '').split(',').map(function(t){ return t.trim(); }).filter(Boolean).sort().join(',');
}

function countPendingRows_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  var layout = getLayout_(sheet);
  var lastRow = sheet.getLastRow();
  var pending = sheet.getRange(2, layout.pendingCol, lastRow - 1, 1).getValues();
  return pending.filter(function(r){ return r[0] && String(r[0]).length > 0; }).length;
}

// ---------- Sync ----------

function collectChanges_(requireConfirm) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DATA_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return { sheet: sheet, changes: [] };

  var layout = getLayout_(sheet);
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, layout.numCols).getValues();
  var changes = [];
  data.forEach(function(r, idx) {
    var pending = String(r[layout.pendingCol - 1] || '').trim();
    if (!pending) return;
    if (requireConfirm && r[layout.confirmCol - 1] !== true) return;
    var snap = null;
    try { snap = JSON.parse(r[layout.snapshotCol - 1] || '{}'); } catch (err) { snap = null; }
    changes.push({
      rowInSheet: idx + 2,
      entryId: String(r[layout.entryIdCol - 1] || '').trim(),
      taskId: String(r[1 + layout.off] || ''),   // Issue Key
      taskName: String(r[2 + layout.off] || ''), // Issue summary
      pending: pending,
      snap: snap,
      newDesc: String(r[layout.descriptionCol - 1] || ''),
      newTags: String(r[layout.categoryCol - 1] || ''),
      newBillable: r[layout.billableCol - 1] === true,
    });
  });
  return { sheet: sheet, changes: changes };
}

function executeSyncChanges_(changes, sheet) {
  var layout = getLayout_(sheet);
  var cfg = readConfig();
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

    var put = {};
    if (parts.indexOf('Billable') !== -1) put.billable = c.newBillable;
    if (parts.indexOf('Desc') !== -1) put.description = c.newDesc;
    if (Object.keys(put).length > 0) {
      var putErr = null;
      try {
        cuPut('/team/' + cfg.teamId + '/time_entries/' + c.entryId, cfg.token, put);
      } catch (err) { rowOK = false; putErr = err.message; }
      if (put.billable !== undefined) {
        logChange_(putErr ? 'Failure' : 'Success', c.entryId, c.taskId, c.taskName, 'Billable',
          c.snap ? (c.snap.billable ? 'Yes' : 'No') : '', put.billable ? 'Yes' : 'No', putErr || '');
      }
      if (put.description !== undefined) {
        logChange_(putErr ? 'Failure' : 'Success', c.entryId, c.taskId, c.taskName, 'Description',
          c.snap ? (c.snap.description || '') : '', put.description, putErr || '');
      }
    }

    if (parts.indexOf('Category') !== -1) {
      // Reverse-map display names → ClickUp tag names for API calls
      var oldDisplayTags = parseTagList_(c.snap ? c.snap.tags : '');
      var newDisplayTags = parseTagList_(c.newTags);
      var oldClickup = oldDisplayTags.map(function(d){ return tagMaps.reverse[d] || d; });
      var newClickup = newDisplayTags.map(function(d){ return tagMaps.reverse[d] || d; });

      var oldSet = {}; oldClickup.forEach(function(t){ oldSet[t] = true; });
      var newSet = {}; newClickup.forEach(function(t){ newSet[t] = true; });
      var toAdd = newClickup.filter(function(t){ return !oldSet[t]; });
      var toRemove = oldClickup.filter(function(t){ return !newSet[t]; });

      toRemove.forEach(function(tag) {
        var tErr = null;
        try {
          cuDelete('/team/' + cfg.teamId + '/time_entries/tags', cfg.token, {
            time_entry_ids: [c.entryId], tags: [{ name: tag }],
          });
        } catch (err) { rowOK = false; tErr = err.message; }
        logChange_(tErr ? 'Failure' : 'Success', c.entryId, c.taskId, c.taskName, 'Category (remove)', tag, '', tErr || '');
      });
      toAdd.forEach(function(tag) {
        var tErr = null;
        try {
          cuPost('/team/' + cfg.teamId + '/time_entries/tags', cfg.token, {
            time_entry_ids: [c.entryId], tags: [{ name: tag }],
          });
        } catch (err) { rowOK = false; tErr = err.message; }
        logChange_(tErr ? 'Failure' : 'Success', c.entryId, c.taskId, c.taskName, 'Category (add)', '', tag, tErr || '');
      });
    }

    if (rowOK) {
      successCount++;
      var newSnap = JSON.stringify({ description: c.newDesc, tags: c.newTags, billable: c.newBillable });
      sheet.getRange(c.rowInSheet, layout.snapshotCol).setValue(newSnap);
      sheet.getRange(c.rowInSheet, layout.pendingCol).setValue('');
      sheet.getRange(c.rowInSheet, layout.confirmCol).setValue(false);
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
      if (p === 'Category') {
        var oldT = c.snap ? (c.snap.tags || '(none)') : '(unknown)';
        return 'Category: [' + oldT + '] \u2192 [' + (c.newTags || '(none)') + ']';
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

function syncAndReload() {
  var result = collectChanges_(false);
  if (result.changes.length > 0) {
    SpreadsheetApp.getActive().toast('Syncing ' + result.changes.length + ' change(s)...', 'ClickUp');
    var outcome = executeSyncChanges_(result.changes, result.sheet);
    var msg = outcome.successCount + ' synced, ' + outcome.failCount + ' failed. Refreshing...';
    SpreadsheetApp.getActive().toast(msg, 'ClickUp', 3);
  }
  refreshTimeEntries(true);
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

  var layout = getLayout_(sheet);
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, layout.numCols).getValues();
  var reverted = 0;
  data.forEach(function(r, idx) {
    var pending = String(r[layout.pendingCol - 1] || '').trim();
    if (!pending) return;
    var snap;
    try { snap = JSON.parse(r[layout.snapshotCol - 1] || '{}'); } catch (err) { snap = null; }
    if (!snap) return;
    var rowNum = idx + 2;
    sheet.getRange(rowNum, layout.descriptionCol).setValue(snap.description || '');
    sheet.getRange(rowNum, layout.categoryCol).setValue(snap.tags || '');
    sheet.getRange(rowNum, layout.billableCol).setValue(snap.billable === true);
    sheet.getRange(rowNum, layout.pendingCol).setValue('');
    sheet.getRange(rowNum, layout.confirmCol).setValue(false);
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
  return s.length > n ? s.substr(0, n - 1) + '\u2026' : s;
}

function flashRow_(sheet, row) {
  try {
    var range = sheet.getRange(row, 1, 1, getLayout_(sheet).confirmCol);
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
      ]]).setFontWeight('bold').setBackground(HEADER_BG).setFontColor(HEADER_FG);
      sheet.setFrozenRows(1);
      var widths = [150, 80, 110, 110, 260, 110, 280, 280, 280];
      for (var c = 0; c < widths.length; c++) sheet.setColumnWidth(c + 1, widths[c]);
      protectSheet_(sheet, 'Change Log — managed by script');
    }
    var tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    sheet.appendRow([
      Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss'),
      status, entryId || '', taskId || '', taskName || '',
      field || '', String(oldVal == null ? '' : oldVal), String(newVal == null ? '' : newVal), message || '',
    ]);
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
