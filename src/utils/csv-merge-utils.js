/**
 * CSV merge utilities: header unification, column alignment, serialization, and blob download.
 */

var csvUtils = require('./csv-utils');

/**
 * Compute the unified header list (union, preserving first-seen order).
 *
 * @param {string[][]} fileHeadersArray - Array of header arrays, e.g. [["A","B"], ["B","C"]]
 * @returns {string[]} Unified header list, e.g. ["A","B","C"]
 */
function computeUnifiedHeaders(fileHeadersArray) {
  var seen = {};
  var unified = [];
  for (var f = 0; f < fileHeadersArray.length; f++) {
    var headers = fileHeadersArray[f];
    for (var h = 0; h < headers.length; h++) {
      var col = headers[h];
      if (!seen[col]) {
        seen[col] = true;
        unified.push(col);
      }
    }
  }
  return unified;
}

/**
 * Align multiple CSV datasets to a unified header list.
 *
 * @param {Array<{headers: string[], data: string[][]}>} filesData
 * @returns {{ headers: string[], data: string[][] }}
 */
function alignColumns(filesData) {
  // Build unified header list from all files
  var allHeaders = filesData.map(function (f) { return f.headers; });
  var headers = computeUnifiedHeaders(allHeaders);

  // Build column-name → index map for each file
  var fileMaps = filesData.map(function (f) {
    var m = {};
    for (var i = 0; i < f.headers.length; i++) {
      m[f.headers[i]] = i;
    }
    return m;
  });

  // Align every row from every file
  var data = [];
  for (var fi = 0; fi < filesData.length; fi++) {
    var rows = filesData[fi].data;
    var colMap = fileMaps[fi];
    for (var ri = 0; ri < rows.length; ri++) {
      var row = rows[ri];
      var aligned = [];
      for (var hi = 0; hi < headers.length; hi++) {
        var colName = headers[hi];
        var srcIdx = colMap[colName];
        aligned.push(srcIdx !== undefined ? String(row[srcIdx] !== undefined ? row[srcIdx] : '') : '');
      }
      data.push(aligned);
    }
  }

  return { headers: headers, data: data };
}

/**
 * Serialize a single CSV row according to RFC 4180.
 * Fields containing the delimiter, double-quotes, or newlines are wrapped
 * in double-quotes; embedded double-quotes are doubled.
 *
 * @param {string[]} row     - Array of field values
 * @param {string}   delimiter - Field separator (default ",")
 * @returns {string} CSV-formatted row string
 */
function serializeRow(row, delimiter) {
  var sep = delimiter || ',';
  var parts = [];
  for (var i = 0; i < row.length; i++) {
    var field = row[i] || '';
    var needsQuotes =
      field.indexOf(sep) !== -1 ||
      field.indexOf('"') !== -1 ||
      field.indexOf('\n') !== -1 ||
      field.indexOf('\r') !== -1;

    if (needsQuotes) {
      // Escape embedded double-quotes by doubling them, then wrap
      var escaped = field.replace(/"/g, '""');
      parts.push('"' + escaped + '"');
    } else {
      parts.push(field);
    }
  }
  return parts.join(sep);
}

/**
 * Build a CSV Blob and trigger browser download.
 *
 * @param {string[]} headers       - Header field names
 * @param {string[][]} data        - 2-D row data
 * @param {string}   encoding      - Text encoding (default "utf-8")
 * @param {string}   delimiter     - Field separator (default ",")
 * @param {string}   filename      - Download filename (default "export.csv")
 */
function buildCSVBlob(headers, data, encoding, delimiter, filename) {
  var enc = encoding || 'utf-8';
  var sep = delimiter || ',';
  var name = filename || 'export.csv';

  var lines = [];

  // Header row
  lines.push(serializeRow(headers, sep));

  // Data rows
  for (var r = 0; r < data.length; r++) {
    lines.push(serializeRow(data[r], sep));
  }

  var csvText = lines.join('\r\n');

  // Prepend UTF-8 BOM for Excel compatibility
  var isUTF = enc === 'utf-8' || enc === 'utf8';
  var blobContent = isUTF ? '﻿' + csvText : csvText;

  var blob, url;
  try {
    blob = new Blob([blobContent], { type: 'text/csv;charset=' + enc });
    url = URL.createObjectURL(blob);
  } catch (e) {
    console.error('Failed to create Blob or Object URL:', e);
    return;
  }

  var anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = name;
  anchor.style.display = 'none';
  document.body.appendChild(anchor);
  anchor.click();

  // Clean up
  document.body.removeChild(anchor);
  URL.revokeObjectURL(url);
}

/**
 * Fully read a CSV File into memory using the chunked parser.
 * Accumulates all batches into a single array before returning.
 *
 * @param {File}   file      - Browser File object
 * @param {string} delimiter - Field separator (default ",")
 * @param {string} encoding  - Text encoding (default "utf-8")
 * @returns {Promise<{headers: string[], data: string[][]}>}
 */
function readCSVFileFull(file, delimiter, encoding) {
  var delim = delimiter || ',';
  var enc = encoding || 'utf-8';

  return new Promise(function (resolve, reject) {
    var allRows = [];
    var headers = null;

    csvUtils.parseCSVFile(
      file,
      delim,
      0,          // batchSize 0 = flush only on finish
      enc,
      function (batch) {
        // First batch: use as headers; subsequent batches: data
        if (headers === null) {
          headers = batch;
          // No data yet — first batch is column headers
        } else {
          for (var i = 0; i < batch.length; i++) {
            allRows.push(batch[i]);
          }
        }
      },
      function () {
        resolve({ headers: headers, data: allRows });
      },
      function (err) {
        reject(err);
      }
    );
  });
}

module.exports = {
  computeUnifiedHeaders: computeUnifiedHeaders,
  alignColumns: alignColumns,
  serializeRow: serializeRow,
  buildCSVBlob: buildCSVBlob,
  readCSVFileFull: readCSVFileFull,
};
