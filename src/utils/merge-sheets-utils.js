/**
 * Excel 最大行数限制
 */
var EXCEL_MAX_ROWS = 1048576;

/**
 * 获取 Excel 列字母（A-Z, AA-ZZ, ...）
 * @param {number} colIndex - 0-based column index
 * @returns {string} Column letter(s)
 */
function getColumnLetter(colIndex) {
  var letter = "";
  var remaining = colIndex;
  do {
    letter = String.fromCharCode(65 + (remaining % 26)) + letter;
    remaining = Math.floor(remaining / 26) - 1;
  } while (remaining >= 0);
  return letter;
}

/**
 * 验证 Sheet 数据列数是否一致
 * @param {Array<{name: string, columnCount: number}>} sheets
 * @returns {{valid: boolean, expectedColumnCount: number|null, error: string|null}}
 */
function validateColumnConsistency(sheets) {
  if (!sheets || sheets.length < 2) {
    return {
      valid: false,
      expectedColumnCount: null,
      error: "请至少选择两个 Sheet"
    };
  }
  var expected = sheets[0].columnCount;
  for (var i = 1; i < sheets.length; i++) {
    if (sheets[i].columnCount !== expected) {
      return {
        valid: false,
        expectedColumnCount: expected,
        error: "所选 Sheet 列数不一致，请重新选择"
      };
    }
  }
  return { valid: true, expectedColumnCount: expected, error: null };
}

/**
 * 验证合并数据行数是否超出 Excel 最大限制
 * @param {number} rowCount - 合并后的数据行数
 * @returns {{valid: boolean, error: string|null}}
 */
function validateRowCount(rowCount) {
  if (rowCount > EXCEL_MAX_ROWS) {
    return {
      valid: false,
      error: "合并结果共 " + rowCount + " 行，超出 Excel 最大行数限制（" + EXCEL_MAX_ROWS + " 行），无法存放。请减少所选 Sheet 的数据量后重试。"
    };
  }
  return { valid: true, error: null };
}

/**
 * 生成唯一的 Sheet 名称
 * @param {string} baseName - 基础名称，如"合并结果"
 * @param {string[]} existingNames - 已存在的 Sheet 名称数组
 * @returns {string} 唯一可用名称
 */
function generateUniqueSheetName(baseName, existingNames) {
  if (existingNames.indexOf(baseName) === -1) {
    return baseName;
  }
  for (var i = 1; i <= 100; i++) {
    var name = baseName + "_" + i;
    if (existingNames.indexOf(name) === -1) {
      return name;
    }
  }
  return baseName + "_" + Date.now();
}

module.exports = {
  getColumnLetter: getColumnLetter,
  validateColumnConsistency: validateColumnConsistency,
  validateRowCount: validateRowCount,
  generateUniqueSheetName: generateUniqueSheetName,
  EXCEL_MAX_ROWS: EXCEL_MAX_ROWS
};