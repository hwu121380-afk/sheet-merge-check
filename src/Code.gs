/**
 * @OnlyCurrentDoc
 *
 * Sheet Merge Check — Google Sheets editor add-on prototype.
 *
 * Scopes are limited to the current spreadsheet and the add-on sidebar.
 * License-key checks are a local stub only. There is no payment call
 * and no UrlFetchApp, so no external_request scope is requested.
 */

var ADDON_NAME = 'Sheet Merge Check';
var MERGED_SHEET_NAME = 'Merged (Sheet Merge Check)';
var CHECK_SHEET_NAME = 'Check report (Sheet Merge Check)';
var LICENSE_PROPERTY = 'SMC_LICENSE_KEY';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(ADDON_NAME)
    .addItem('Open sidebar', 'showSidebar')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle(ADDON_NAME)
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSidebarState() {
  var license = MergeCheck.licenseStatus(_storedLicenseKey());
  return {
    sheets: listSourceSheets_(),
    license: license,
    freeMaxSources: MergeCheck.FREE_MAX_SOURCES,
    freeMaxRows: MergeCheck.FREE_MAX_ROWS,
    mergedSheetName: MERGED_SHEET_NAME,
    checkSheetName: CHECK_SHEET_NAME
  };
}

function saveLicenseKey(key) {
  var status = MergeCheck.licenseStatus(key);
  var props = PropertiesService.getUserProperties();
  if (status.accepted) {
    props.setProperty(LICENSE_PROPERTY, String(key).trim());
  } else if (!String(key || '').trim()) {
    props.deleteProperty(LICENSE_PROPERTY);
    status = MergeCheck.licenseStatus('');
  }
  return status;
}

/**
 * Run the merge and check. Does not write sheets unless `writeSheets` is true.
 * @param {{sheetNames: string[], keyColumn?: string, addSourceColumn?: boolean, writeSheets?: boolean}} payload
 */
function runMerge(payload) {
  payload = payload || {};
  var names = payload.sheetNames || [];
  if (!names.length) {
    throw new Error('Select at least one source sheet.');
  }
  var reserved = {};
  reserved[MERGED_SHEET_NAME] = true;
  reserved[CHECK_SHEET_NAME] = true;
  var sources = [];
  var ss = SpreadsheetApp.getActive();
  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    if (reserved[name]) {
      throw new Error('The sheet "' + name + '" is an output sheet. Pick a source tab instead.');
    }
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      throw new Error('Sheet not found: ' + name);
    }
    sources.push(_sheetToSource(sheet));
  }

  var result = MergeCheck.mergeAndCheck(sources, {
    addSourceColumn: !!payload.addSourceColumn,
    keyColumn: payload.keyColumn || '',
    licenseKey: _storedLicenseKey()
  });

  var written = { merged: false, check: false };
  if (payload.writeSheets) {
    if (!result.limits.allowed) {
      throw new Error(
        'Writing is blocked on the free tier. ' + result.limits.reasons.join(' ')
      );
    }
    _writeOutputSheets(ss, result);
    written.merged = true;
    written.check = true;
  }

  return {
    summary: MergeCheck.formatCheck(result),
    rowsIn: result.rowsIn,
    rowsOut: result.rowsOut,
    ok: result.ok,
    exactDuplicates: result.exactDuplicates,
    files: result.files,
    headers: result.headers,
    headerMismatches: result.headerMismatches,
    similarHeaders: result.similarHeaders,
    keyColumn: result.keyColumn,
    blankKeyCount: result.blankKeyCount,
    blankKeyFields: result.blankKeyFields.slice(0, 25),
    blankKeyFieldsTruncated: result.blankKeyFields.length > 25,
    duplicateKeys: result.duplicateKeys.slice(0, 25),
    duplicateKeysTruncated: result.duplicateKeys.length > 25,
    license: result.license,
    limits: result.limits,
    written: written,
    unionHeaders: result.dataHeaders
  };
}

function listSourceSheets_() {
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  var out = [];
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name === MERGED_SHEET_NAME || name === CHECK_SHEET_NAME) continue;
    var lastCol = Math.max(sheets[i].getLastColumn(), 1);
    var headers = sheets[i].getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    var clean = [];
    for (var c = 0; c < headers.length; c++) {
      if (String(headers[c]).trim() !== '') clean.push(String(headers[c]));
    }
    out.push({
      name: name,
      headers: clean,
      lastRow: sheets[i].getLastRow()
    });
  }
  return out;
}

function _sheetToSource(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    return { name: sheet.getName(), headers: [], rows: [] };
  }
  var values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  var headers = values[0].map(function (h) { return String(h == null ? '' : h); });
  var rows = [];
  for (var r = 1; r < values.length; r++) {
    rows.push(values[r]);
  }
  return { name: sheet.getName(), headers: headers, rows: rows };
}

function _storedLicenseKey() {
  return PropertiesService.getUserProperties().getProperty(LICENSE_PROPERTY) || '';
}

function _writeOutputSheets(ss, result) {
  var merged = _getOrCreateSheet_(ss, MERGED_SHEET_NAME);
  var check = _getOrCreateSheet_(ss, CHECK_SHEET_NAME);
  _writeGrid_(merged, _mergedGrid(result));
  _writeGrid_(check, MergeCheck.resultToCheckRows(result));
  check.activate();
}

function _mergedGrid(result) {
  var grid = [result.headers.slice()];
  for (var i = 0; i < result.rows.length; i++) {
    var row = [];
    for (var h = 0; h < result.headers.length; h++) {
      row.push(result.rows[i][result.headers[h]] || '');
    }
    grid.push(row);
  }
  return grid;
}

function _getOrCreateSheet_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function _writeGrid_(sheet, grid) {
  sheet.clearContents();
  if (!grid.length) return;
  var rows = grid.length;
  var cols = 1;
  for (var i = 0; i < grid.length; i++) {
    if (grid[i].length > cols) cols = grid[i].length;
  }
  var padded = [];
  for (var r = 0; r < grid.length; r++) {
    var line = grid[r].slice();
    while (line.length < cols) line.push('');
    padded.push(line);
  }
  sheet.getRange(1, 1, rows, cols).setValues(padded);
}
