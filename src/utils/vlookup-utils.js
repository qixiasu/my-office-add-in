var { getColumnLetter } = require("./concat-utils");

var SEPA = "!";
var ABS = "$";

function parseRangeAddress(address) {
  var sheet = "";
  var rangePart = address;

  var bangIdx = address.indexOf(SEPA);
  if (bangIdx !== -1) {
    sheet = address.substring(0, bangIdx);
    rangePart = address.substring(bangIdx + 1);
  }

  var parts = rangePart.split(":");
  var start = parseCellRef(parts[0]);
  var end = parseCellRef(parts[1] || parts[0]);

  return {
    sheet: sheet,
    startCol: start.col,
    startRow: start.row,
    endCol: end.col,
    endRow: end.row,
    colCount: end.col - start.col + 1,
    rowCount: end.row - start.row + 1,
  };
}

function parseCellRef(ref) {
  var col = 0;
  var i = 0;

  while (i < ref.length) {
    var ch = ref.charCodeAt(i);
    if (ch === 36) {
      i++;
    } else if (ch >= 65 && ch <= 90) {
      col = col * 26 + (ch - 64);
      i++;
    } else {
      break;
    }
  }

  // Skip any remaining $ between column and row
  while (i < ref.length && ref.charCodeAt(i) === 36) {
    i++;
  }

  var row = parseInt(ref.substring(i), 10) || 1;
  return { col: col - 1, row: row };
}

function buildColRange(parsed, colIndex) {
  var colLetter = getColumnLetter(parsed.startCol + colIndex);
  var prefix = parsed.sheet ? parsed.sheet + SEPA : "";
  var colAbs = ABS + colLetter + ABS;
  return prefix + colAbs + parsed.startRow + ":" + colAbs + parsed.endRow;
}

function buildIndexMatchFormula(lookupCellRef, lookupColRange, returnColRange, matchMode) {
  return (
    "=INDEX(" +
    returnColRange +
    ", MATCH(" +
    lookupCellRef +
    ", " +
    lookupColRange +
    ", " +
    matchMode +
    "))"
  );
}

function buildLookupIndex(lookupTable, matchColIndex) {
  var index = {};
  for (var r = 0; r < lookupTable.length; r++) {
    var key = lookupTable[r][matchColIndex];
    if (key === null || key === undefined) {
      key = "";
    }
    key = String(key);
    index[key] = r;
  }
  return index;
}

function staticLookup(
  lookupValues,
  lookupTable,
  matchColIndex,
  returnColIndices,
  matchMode,
  defaultValue
) {
  if (defaultValue === undefined || defaultValue === null) {
    defaultValue = "#N/A";
  }

  var results = [];

  var index = {};
  if (matchMode === 0) {
    for (var r = 0; r < lookupTable.length; r++) {
      var key = lookupTable[r][matchColIndex];
      if (key === null || key === undefined) {
        key = "";
      }
      key = String(key);
      index[key] = r;
    }
  }

  for (var i = 0; i < lookupValues.length; i++) {
    var row = [];
    var val = lookupValues[i];
    if (val === null || val === undefined) {
      val = "";
    }
    // Null/undefined in approximate mode → defaultValue (avoid Number("") → 0)
    if (matchMode !== 0 && (lookupValues[i] === null || lookupValues[i] === undefined)) {
      for (var q = 0; q < returnColIndices.length; q++) {
        row.push(defaultValue);
      }
      results.push(row);
      continue;
    }
    var valStr = String(val);

    if (matchMode === 0) {
      var matchedRow = index[valStr];
      if (matchedRow !== undefined) {
        for (var j = 0; j < returnColIndices.length; j++) {
          row.push(lookupTable[matchedRow][returnColIndices[j]]);
        }
      } else {
        for (var k = 0; k < returnColIndices.length; k++) {
          row.push(defaultValue);
        }
      }
    } else {
      // Approximate match: find largest value <= lookup
      // Assumes lookupTable is sorted ascending on match column (same as Excel VLOOKUP)
      var bestRow = -1;
      for (var m = 0; m < lookupTable.length; m++) {
        var tableVal = lookupTable[m][matchColIndex];
        if (tableVal === null || tableVal === undefined) continue;
        if (Number(tableVal) <= Number(val)) {
          bestRow = m;
        } else {
          break;
        }
      }
      if (bestRow >= 0) {
        for (var n = 0; n < returnColIndices.length; n++) {
          row.push(lookupTable[bestRow][returnColIndices[n]]);
        }
      } else {
        for (var p = 0; p < returnColIndices.length; p++) {
          row.push(defaultValue);
        }
      }
    }

    results.push(row);
  }

  return results;
}

module.exports = {
  parseRangeAddress: parseRangeAddress,
  parseCellRef: parseCellRef,
  buildColRange: buildColRange,
  buildIndexMatchFormula: buildIndexMatchFormula,
  staticLookup: staticLookup,
  buildLookupIndex: buildLookupIndex,
};
