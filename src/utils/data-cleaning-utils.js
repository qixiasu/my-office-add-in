// src/utils/data-cleaning-utils.js

/**
 * 修剪空格
 * @param {Array<Array>} values - 二维数组
 * @param {string} mode - 'both' | 'all' | 'leading' | 'trailing'
 * @returns {Array<Array>} 新二维数组
 */
function trimSpaces(values, mode) {
  return values.map(function (row) {
    return row.map(function (cell) {
      if (typeof cell !== "string") return cell;
      switch (mode) {
        case "leading":
          return cell.replace(/^\s+/, "");
        case "trailing":
          return cell.replace(/\s+$/, "");
        case "all":
          return cell.replace(/\s+/g, " ").trim();
        case "both":
        default:
          return cell.trim();
      }
    });
  });
}

/**
 * 删除空行
 * @param {Array<Array>} values - 二维数组
 * @param {string} mode - 'all' | 'column' | 'ratio'
 * @param {number|null} columnIndex - 按列删除时指定的列索引
 * @param {number|null} ratioThreshold - 空值率阈值 (0-100)
 * @returns {Array<Array>} 新二维数组
 */
function removeEmptyRows(values, mode, columnIndex, ratioThreshold) {
  return values.filter(function (row) {
    if (mode === "all") {
      return row.some(function (cell) {
        return cell !== null && cell !== undefined && cell !== "";
      });
    }
    if (mode === "column") {
      var val = row[columnIndex];
      return val !== null && val !== undefined && val !== "";
    }
    if (mode === "ratio") {
      var emptyCount = 0;
      for (var i = 0; i < row.length; i++) {
        if (row[i] === null || row[i] === undefined || row[i] === "") {
          emptyCount++;
        }
      }
      var emptyRatio = (emptyCount / row.length) * 100;
      return emptyRatio < ratioThreshold;
    }
    return true;
  });
}

/**
 * 大小写转换
 * @param {Array<Array>} values - 二维数组
 * @param {string} mode - 'upper' | 'lower' | 'capitalize' | 'camel'
 * @returns {Array<Array>} 新二维数组
 */
function convertCase(values, mode) {
  return values.map(function (row) {
    return row.map(function (cell) {
      if (typeof cell !== "string") return cell;
      switch (mode) {
        case "upper":
          return cell.toUpperCase();
        case "lower":
          return cell.toLowerCase();
        case "capitalize":
          return cell.charAt(0).toUpperCase() + cell.slice(1);
        case "camel":
          return cell.replace(/[^a-zA-Z0-9]+(.)/g, function (_, chr) {
            return chr.toUpperCase();
          });
        default:
          return cell;
      }
    });
  });
}

/**
 * 清除不可见字符
 * @param {Array<Array>} values - 二维数组
 * @param {string} mode - 'control' | 'whitespace' | 'zero-width' | 'all'
 * @returns {Array<Array>} 新二维数组
 */
function removeInvisible(values, mode) {
  var patterns = {
    control: /[\x00-\x1F\x7F]/g,
    whitespace: /[\t\n\r\x0B\x0C\x1F]/g,
    "zero-width": /[​‌‍﻿⁠‎‏]/g,
    all: /[\x00-\x1F\x7F​‌‍﻿⁠‎‏]/g,
  };

  var regex = patterns[mode] || patterns.all;

  return values.map(function (row) {
    return row.map(function (cell) {
      if (typeof cell !== "string") return cell;
      return cell.replace(regex, "");
    });
  });
}

/**
 * 移除重复行
 * @param {Array<Array>} values - 二维数组
 * @param {Array<number>|null} keyColumns - 依据列索引数组，null=所有列
 * @param {string} keep - 'first' | 'last'
 * @returns {Array<Array>} 新二维数组
 */
function removeDuplicates(values, keyColumns, keep) {
  var seen = [];
  var result = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var key = keyColumns
      ? keyColumns.map(function (col) { return row[col]; }).join("|||")
      : row.join("|||");

    var foundIdx = -1;
    for (var j = 0; j < seen.length; j++) {
      if (seen[j] === key) {
        foundIdx = j;
        break;
      }
    }

    if (foundIdx === -1) {
      seen.push(key);
      result.push(row);
    } else if (keep === "last") {
      result.splice(foundIdx, 1);
      seen.splice(foundIdx, 1);
      seen.push(key);
      result.push(row);
    }
  }

  return result;
}

module.exports = {
  trimSpaces: trimSpaces,
  removeEmptyRows: removeEmptyRows,
  convertCase: convertCase,
  removeInvisible: removeInvisible,
  removeDuplicates: removeDuplicates,
};
