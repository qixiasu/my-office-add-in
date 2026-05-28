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

  console.log("[DEBUG staticLookup] lookupValues:", JSON.stringify(lookupValues));
  console.log("[DEBUG staticLookup] lookupTable 行数:", lookupTable.length, "列数:", lookupTable[0] ? lookupTable[0].length : 0);
  console.log("[DEBUG staticLookup] matchColIndex:", matchColIndex, "returnColIndices:", JSON.stringify(returnColIndices), "matchMode:", matchMode);

  // 打印查找表第一列（match列）的实际内容
  if (lookupTable.length > 0) {
    var matchColValues = [];
    for (var mi = 0; mi < lookupTable.length; mi++) {
      matchColValues.push(lookupTable[mi][matchColIndex]);
    }
    console.log("[DEBUG staticLookup] 查找表 match 列内容:", JSON.stringify(matchColValues));
  }

  var index = {};
  if (matchMode === 0) {
    console.log("[DEBUG staticLookup] 构建 index 哈希表 (精确匹配模式)");
    for (var r = 0; r < lookupTable.length; r++) {
      var key = lookupTable[r][matchColIndex];
      if (key === null || key === undefined) {
        key = "";
      }
      key = String(key);
      console.log("[DEBUG staticLookup] index['" + key + "'] = " + r);
      index[key] = r;
    }
    console.log("[DEBUG staticLookup] index 哈希表构建完成:", JSON.stringify(index));
  }

  for (var i = 0; i < lookupValues.length; i++) {
    var row = [];
    var val = lookupValues[i];
    if (val === null || val === undefined) {
      val = "";
    }
    console.log("[DEBUG staticLookup] 查找第", i, "个值:", JSON.stringify(val), "类型:", typeof val);
    // Null/undefined in approximate mode → defaultValue (avoid Number("") → 0)
    if (matchMode !== 0 && (lookupValues[i] === null || lookupValues[i] === undefined)) {
      for (var q = 0; q < returnColIndices.length; q++) {
        row.push(defaultValue);
      }
      results.push(row);
      continue;
    }
    var valStr = String(val);
    console.log("[DEBUG staticLookup] valStr =", JSON.stringify(valStr));

    if (matchMode === 0) {
      var matchedRow = index[valStr];
      console.log("[DEBUG staticLookup] 精确匹配: index['" + valStr + "'] =", matchedRow);
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
};
