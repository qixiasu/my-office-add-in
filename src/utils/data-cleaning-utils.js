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
  // TODO: implement
  return values;
}

/**
 * 大小写转换
 * @param {Array<Array>} values - 二维数组
 * @param {string} mode - 'upper' | 'lower' | 'capitalize' | 'camel'
 * @returns {Array<Array>} 新二维数组
 */
function convertCase(values, mode) {
  // TODO: implement
  return values;
}

/**
 * 清除不可见字符
 * @param {Array<Array>} values - 二维数组
 * @param {string} mode - 'control' | 'whitespace' | 'zero-width' | 'all'
 * @returns {Array<Array>} 新二维数组
 */
function removeInvisible(values, mode) {
  // TODO: implement
  return values;
}

/**
 * 移除重复行
 * @param {Array<Array>} values - 二维数组
 * @param {Array<number>|null} keyColumns - 依据列索引数组，null=所有列
 * @param {string} keep - 'first' | 'last'
 * @returns {Array<Array>} 新二维数组
 */
function removeDuplicates(values, keyColumns, keep) {
  // TODO: implement
  return values;
}

module.exports = {
  trimSpaces: trimSpaces,
  removeEmptyRows: removeEmptyRows,
  convertCase: convertCase,
  removeInvisible: removeInvisible,
  removeDuplicates: removeDuplicates,
};
