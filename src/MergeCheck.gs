/**
 * Sheet Merge Check — core merge and check logic.
 *
 * This file is written so it can run in Google Apps Script and in Node.
 * It reuses the idea of the MIT-licensed csv-merge-check tool:
 * stack rows, do not rename columns, skip completely blank rows,
 * keep exact duplicates, and report a rows-in = rows-out check.
 *
 * Extra checks for this add-on (not in the original CLI):
 * header mismatches (including similar names such as date vs Date),
 * blank values in a chosen key column, and a free-tier limit stub.
 *
 * No network calls. No payment integration.
 */
var MergeCheck = (function () {
  var FREE_MAX_SOURCES = 3;
  var FREE_MAX_ROWS = 1000;
  var SOURCE_COLUMN_DEFAULT = 'source_sheet';
  // SAMPLE / DEV ONLY. Not a real paid license. Never sell this string.
  var SAMPLE_PRO_KEY = 'CMC-PRO-SAMPLE-UNLOCK';

  function isBlankRow(values) {
    for (var i = 0; i < values.length; i++) {
      if (String(values[i] == null ? '' : values[i]).trim() !== '') {
        return false;
      }
    }
    return true;
  }

  function unionHeaders(headerLists) {
    var seen = {};
    var headers = [];
    for (var i = 0; i < headerLists.length; i++) {
      var list = headerLists[i] || [];
      for (var j = 0; j < list.length; j++) {
        var name = list[j];
        if (name == null || name === '') continue;
        if (!Object.prototype.hasOwnProperty.call(seen, name)) {
          seen[name] = true;
          headers.push(String(name));
        }
      }
    }
    return headers;
  }

  function uniqueName(headers, base) {
    if (headers.indexOf(base) === -1) return base;
    var suffix = 2;
    while (headers.indexOf(base + '_' + suffix) !== -1) {
      suffix += 1;
    }
    return base + '_' + suffix;
  }

  function normalizeKey(value) {
    return String(value == null ? '' : value).toLowerCase().replace(/\s+/g, ' ').trim();
  }

  function similarHeaderPairs(headerLists) {
    var byNorm = {};
    for (var i = 0; i < headerLists.length; i++) {
      var list = headerLists[i] || [];
      for (var j = 0; j < list.length; j++) {
        var raw = String(list[j]);
        var norm = normalizeKey(raw);
        if (!norm) continue;
        if (!byNorm[norm]) byNorm[norm] = [];
        var already = false;
        for (var k = 0; k < byNorm[norm].length; k++) {
          if (byNorm[norm][k] === raw) already = true;
        }
        if (!already) byNorm[norm].push(raw);
      }
    }
    var pairs = [];
    Object.keys(byNorm).forEach(function (norm) {
      if (byNorm[norm].length > 1) {
        pairs.push({ normalized: norm, variants: byNorm[norm].slice() });
      }
    });
    return pairs;
  }

  function setMinus(left, rightSet) {
    var out = [];
    for (var i = 0; i < left.length; i++) {
      if (!Object.prototype.hasOwnProperty.call(rightSet, left[i])) {
        out.push(left[i]);
      }
    }
    return out;
  }

  function toSet(list) {
    var s = {};
    for (var i = 0; i < list.length; i++) {
      s[list[i]] = true;
    }
    return s;
  }

  function isValidLicenseKey(key) {
    return String(key == null ? '' : key).trim() === SAMPLE_PRO_KEY;
  }

  function licenseStatus(key) {
    var trimmed = String(key == null ? '' : key).trim();
    var isPro = isValidLicenseKey(trimmed);
    return {
      isPro: isPro,
      hasKey: trimmed.length > 0,
      accepted: isPro,
      sampleKey: SAMPLE_PRO_KEY,
      note: isPro
        ? 'SAMPLE license accepted (prototype stub only).'
        : trimmed
          ? 'License key not recognized. This prototype only accepts the sample stub key.'
          : 'No license key. Free-tier limits apply.'
    };
  }

  function checkLimits(sourceCount, rowCount, isPro) {
    var overSources = sourceCount > FREE_MAX_SOURCES;
    var overRows = rowCount > FREE_MAX_ROWS;
    var allowed = !!(isPro || (!overSources && !overRows));
    var reasons = [];
    if (!isPro && overSources) {
      reasons.push('Free tier allows up to ' + FREE_MAX_SOURCES + ' source tabs (this run has ' + sourceCount + ').');
    }
    if (!isPro && overRows) {
      reasons.push('Free tier allows up to ' + FREE_MAX_ROWS + ' data rows (this run has ' + rowCount + ').');
    }
    return {
      allowed: allowed,
      isPro: !!isPro,
      sourceCount: sourceCount,
      rowCount: rowCount,
      maxSources: FREE_MAX_SOURCES,
      maxRows: FREE_MAX_ROWS,
      reasons: reasons
    };
  }

  /**
   * Merge sources and build a check report.
   *
   * @param {Array<{name: string, headers: string[], rows: Array<Array<*>>}>} sources
   *   Each source is one sheet/tab. `rows` are data rows (no header), values
   *   aligned to that source's `headers`. Completely blank data rows are skipped.
   * @param {Object} [options]
   * @param {boolean} [options.addSourceColumn]
   * @param {string} [options.sourceColumn]
   * @param {string} [options.keyColumn]  Column used to flag blank keys.
   * @param {string} [options.licenseKey] SAMPLE stub only.
   * @return {Object} merge result plus checks and limit info
   */
  function mergeAndCheck(sources, options) {
    options = options || {};
    if (!sources || !sources.length) {
      throw new Error('Select at least one source sheet.');
    }

    var fileHeaders = [];
    var prepared = [];

    for (var s = 0; s < sources.length; s++) {
      var src = sources[s];
      var headers = (src.headers || []).map(function (h) { return String(h == null ? '' : h); });
      if (!headers.length || headers.every(function (h) { return h === ''; })) {
        throw new Error((src.name || 'Untitled') + ': no header row found');
      }
      var dataRows = [];
      var rawRows = src.rows || [];
      for (var r = 0; r < rawRows.length; r++) {
        var raw = rawRows[r] || [];
        var values = [];
        for (var c = 0; c < headers.length; c++) {
          var cell = c < raw.length ? raw[c] : '';
          values.push(cell == null ? '' : String(cell));
        }
        if (isBlankRow(values)) continue;
        var obj = {};
        for (var h = 0; h < headers.length; h++) {
          if (headers[h] === '') continue;
          obj[headers[h]] = values[h];
        }
        dataRows.push({ rowIndex: r + 2, values: obj }); // +2: header is row 1
      }
      fileHeaders.push(headers.filter(function (h) { return h !== ''; }));
      prepared.push({
        name: src.name || ('Sheet ' + (s + 1)),
        headers: fileHeaders[fileHeaders.length - 1],
        rows: dataRows
      });
    }

    var headers = unionHeaders(fileHeaders);
    var similar = similarHeaderPairs(fileHeaders);
    var firstSet = toSet(prepared[0].headers);
    var unionSet = toSet(headers);
    var mismatches = [];

    for (var i = 0; i < prepared.length; i++) {
      var srcSet = toSet(prepared[i].headers);
      var missingVsUnion = setMinus(headers, srcSet);
      var extraVsFirst = setMinus(prepared[i].headers, firstSet);
      var missingVsFirst = setMinus(prepared[0].headers, srcSet);
      var hasDiff = missingVsUnion.length > 0 || extraVsFirst.length > 0;
      if (hasDiff) {
        mismatches.push({
          source: prepared[i].name,
          missingVsUnion: missingVsUnion,
          extraVsFirst: prepared[i].name === prepared[0].name ? [] : extraVsFirst,
          missingVsFirst: prepared[i].name === prepared[0].name ? [] : missingVsFirst
        });
      }
    }

    var sourceColumn = '';
    var outputHeaders = headers.slice();
    if (options.addSourceColumn) {
      sourceColumn = uniqueName(outputHeaders, options.sourceColumn || SOURCE_COLUMN_DEFAULT);
      outputHeaders.push(sourceColumn);
    }

    var merged = [];
    var files = [];
    var blankKeys = [];
    var keyColumn = options.keyColumn ? String(options.keyColumn) : '';
    var keyCounts = {};

    for (var p = 0; p < prepared.length; p++) {
      var file = prepared[p];
      files.push({
        name: file.name,
        rowsIn: file.rows.length,
        headers: file.headers.slice()
      });
      for (var n = 0; n < file.rows.length; n++) {
        var item = file.rows[n];
        var out = {};
        for (var hi = 0; hi < headers.length; hi++) {
          var col = headers[hi];
          out[col] = Object.prototype.hasOwnProperty.call(item.values, col)
            ? item.values[col]
            : '';
        }
        if (sourceColumn) {
          out[sourceColumn] = file.name;
        }
        merged.push(out);
        if (keyColumn) {
          var keyVal = Object.prototype.hasOwnProperty.call(item.values, keyColumn)
            ? item.values[keyColumn]
            : out[keyColumn];
          var keyStr = keyVal == null ? '' : String(keyVal);
          if (keyStr.trim() === '') {
            blankKeys.push({
              source: file.name,
              rowIndex: item.rowIndex,
              keyColumn: keyColumn
            });
          } else {
            keyCounts[keyStr] = (keyCounts[keyStr] || 0) + 1;
          }
        }
      }
    }

    var exactDuplicates = 0;
    var seen = {};
    for (var m = 0; m < merged.length; m++) {
      var parts = [];
      for (var hc = 0; hc < headers.length; hc++) {
        parts.push(merged[m][headers[hc]] || '');
      }
      var sig = parts.join('\u0001');
      if (Object.prototype.hasOwnProperty.call(seen, sig)) {
        exactDuplicates += 1;
      } else {
        seen[sig] = true;
      }
    }

    var duplicateKeys = [];
    if (keyColumn) {
      Object.keys(keyCounts).forEach(function (k) {
        if (keyCounts[k] > 1) {
          duplicateKeys.push({ key: k, count: keyCounts[k] });
        }
      });
      duplicateKeys.sort(function (a, b) { return b.count - a.count; });
    }

    var rowsIn = 0;
    for (var f = 0; f < files.length; f++) rowsIn += files[f].rowsIn;
    var rowsOut = merged.length;
    var license = licenseStatus(options.licenseKey);
    var limits = checkLimits(prepared.length, rowsIn, license.isPro);

    return {
      files: files,
      headers: outputHeaders,
      dataHeaders: headers,
      rows: merged,
      rowsIn: rowsIn,
      rowsOut: rowsOut,
      ok: rowsIn === rowsOut,
      exactDuplicates: exactDuplicates,
      headerMismatches: mismatches,
      similarHeaders: similar,
      keyColumn: keyColumn,
      blankKeyFields: blankKeys,
      blankKeyCount: blankKeys.length,
      duplicateKeys: duplicateKeys,
      sourceColumn: sourceColumn,
      license: license,
      limits: limits,
      sampleLabel: 'Checks below may include SAMPLE – FICTIONAL DATA when run on the bundled samples.'
    };
  }

  function formatCheck(result) {
    var nameWidth = 4;
    for (var i = 0; i < result.files.length; i++) {
      if (result.files[i].name.length > nameWidth) nameWidth = result.files[i].name.length;
    }
    if (nameWidth < 16) nameWidth = 16;
    var countWidth = 8;
    function padRight(text, width) {
      var s = String(text);
      while (s.length < width) s += ' ';
      return s;
    }
    function padLeft(text, width) {
      var s = String(text);
      while (s.length < width) s = ' ' + s;
      return s;
    }
    var lines = [
      padRight('Source', nameWidth) + '  ' + padLeft('Rows in', countWidth),
      padRight(repeat('-', nameWidth), nameWidth) + '  ' + repeat('-', countWidth)
    ];
    for (var j = 0; j < result.files.length; j++) {
      lines.push(
        padRight(result.files[j].name, nameWidth) + '  ' +
        padLeft(String(result.files[j].rowsIn), countWidth)
      );
    }
    lines.push(padRight(repeat('-', nameWidth), nameWidth) + '  ' + repeat('-', countWidth));
    lines.push(padRight('TOTAL in', nameWidth) + '  ' + padLeft(String(result.rowsIn), countWidth));
    lines.push(padRight('TOTAL out', nameWidth) + '  ' + padLeft(String(result.rowsOut), countWidth));
    lines.push(
      padRight('Exact duplicates', nameWidth) + '  ' +
      padLeft(String(result.exactDuplicates), countWidth) +
      '  (kept; not removed)'
    );
    var status = result.ok ? 'OK' : 'MISMATCH';
    lines.push('Row-count check: ' + result.rowsIn + ' in = ' + result.rowsOut + ' out  ' + status);

    if (result.headerMismatches && result.headerMismatches.length) {
      lines.push('');
      lines.push('Header mismatches (columns are not renamed):');
      for (var h = 0; h < result.headerMismatches.length; h++) {
        var mm = result.headerMismatches[h];
        lines.push('  ' + mm.source);
        if (mm.missingVsUnion.length) {
          lines.push('    Missing vs union: ' + mm.missingVsUnion.join(', '));
        }
        if (mm.extraVsFirst.length) {
          lines.push('    Extra vs first source: ' + mm.extraVsFirst.join(', '));
        }
        if (mm.missingVsFirst.length) {
          lines.push('    Missing vs first source: ' + mm.missingVsFirst.join(', '));
        }
      }
    } else {
      lines.push('Header mismatches: none');
    }

    if (result.similarHeaders && result.similarHeaders.length) {
      lines.push('Similar header names (not merged automatically):');
      for (var sv = 0; sv < result.similarHeaders.length; sv++) {
        lines.push('  ' + result.similarHeaders[sv].variants.join('  |  '));
      }
    }

    if (result.keyColumn) {
      lines.push('Blank "' + result.keyColumn + '" fields: ' + result.blankKeyCount);
      lines.push('Duplicate "' + result.keyColumn + '" values (non-blank): ' + result.duplicateKeys.length + ' keys');
    }

    if (result.limits && !result.limits.allowed) {
      lines.push('');
      lines.push('Write blocked (free-tier limit): ' + result.limits.reasons.join(' '));
    }
    return lines.join('\n') + '\n';
  }

  function repeat(ch, n) {
    var s = '';
    for (var i = 0; i < n; i++) s += ch;
    return s;
  }

  function resultToCheckRows(result) {
    var rows = [['Check', 'Value', 'Notes']];
    rows.push(['Row-count status', result.ok ? 'OK' : 'MISMATCH', result.rowsIn + ' in = ' + result.rowsOut + ' out']);
    rows.push(['TOTAL in', String(result.rowsIn), 'Blank data rows were skipped and not counted']);
    rows.push(['TOTAL out', String(result.rowsOut), '']);
    rows.push(['Exact duplicate rows', String(result.exactDuplicates), 'Kept in the merged sheet; not removed']);
    for (var i = 0; i < result.files.length; i++) {
      rows.push(['Rows in: ' + result.files[i].name, String(result.files[i].rowsIn), result.files[i].headers.join(', ')]);
    }
    rows.push(['Header mismatches', String(result.headerMismatches.length), result.headerMismatches.length ? 'See rows below' : 'None']);
    for (var h = 0; h < result.headerMismatches.length; h++) {
      var mm = result.headerMismatches[h];
      rows.push([
        'Mismatch: ' + mm.source,
        '',
        'Missing vs union: ' + (mm.missingVsUnion.join(', ') || '—') +
        '; Extra vs first: ' + (mm.extraVsFirst.join(', ') || '—')
      ]);
    }
    if (result.similarHeaders.length) {
      for (var s = 0; s < result.similarHeaders.length; s++) {
        rows.push(['Similar headers', result.similarHeaders[s].variants.join(' | '), 'Not auto-renamed']);
      }
    }
    if (result.keyColumn) {
      rows.push(['Blank key fields (' + result.keyColumn + ')', String(result.blankKeyCount), '']);
      for (var b = 0; b < result.blankKeyFields.length; b++) {
        var bk = result.blankKeyFields[b];
        rows.push(['Blank key', bk.source + ' row ' + bk.rowIndex, bk.keyColumn]);
      }
      rows.push(['Duplicate key values', String(result.duplicateKeys.length), 'Non-blank ' + result.keyColumn]);
      for (var d = 0; d < result.duplicateKeys.length; d++) {
        rows.push(['Duplicate key', result.duplicateKeys[d].key, 'count ' + result.duplicateKeys[d].count]);
      }
    }
    rows.push(['License', result.license.isPro ? 'Pro (stub)' : 'Free', result.license.note]);
    rows.push(['Write allowed', result.limits.allowed ? 'Yes' : 'No', result.limits.reasons.join(' ')]);
    return rows;
  }

  return {
    FREE_MAX_SOURCES: FREE_MAX_SOURCES,
    FREE_MAX_ROWS: FREE_MAX_ROWS,
    SAMPLE_PRO_KEY: SAMPLE_PRO_KEY,
    mergeAndCheck: mergeAndCheck,
    formatCheck: formatCheck,
    licenseStatus: licenseStatus,
    isValidLicenseKey: isValidLicenseKey,
    checkLimits: checkLimits,
    resultToCheckRows: resultToCheckRows,
    isBlankRow: isBlankRow,
    unionHeaders: unionHeaders
  };
})();

if (typeof module !== 'undefined' && module.exports) {
  module.exports = MergeCheck;
}
