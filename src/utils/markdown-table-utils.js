/**
 * Markdown 表格转换工具
 * 将 Excel 区域数据转换为 GitHub 风格 Markdown 表格
 */

/**
 * 生成 Markdown 表格
 * @param {Array<Array>} values - 二维单元格值数组
 * @param {Object} options - 配置选项
 * @param {boolean} options.includeAlignment - 是否包含对齐标记（默认 true）
 * @param {boolean} options.preserveFormat - 是否保留显示格式（默认 true）
 * @returns {string} Markdown 表格字符串
 */
function generateMarkdownTable(values, options = {}) {
  const { includeAlignment = true, preserveFormat = true } = options;

  if (!values || !values.length) return '';

  // 过滤空行空列
  const cleaned = parseRange(values);
  if (!cleaned.length) return '';

  // 检测每列对齐方式
  const alignment = detectAlignment(cleaned, includeAlignment);

  // 构建表格字符串
  return buildTableString(cleaned, alignment);
}

/**
 * 过滤空行空列
 * @param {Array<Array>} values
 * @returns {Array<Array>}
 */
function parseRange(values) {
  // 过滤全空行
  const filteredRows = values.filter(row => row.some(cell => cell !== null && cell !== ''));

  if (!filteredRows.length) return [];

  // 过滤尾随空列
  const colCount = Math.max(...filteredRows.map(row => row.length));
  const nonEmptyCols = new Set();

  for (let col = 0; col < colCount; col++) {
    for (const row of filteredRows) {
      if (row[col] !== null && row[col] !== '') {
        nonEmptyCols.add(col);
        break;
      }
    }
  }

  return filteredRows.map(row =>
    Array.from(nonEmptyCols).sort((a, b) => a - b).map(col => row[col] ?? '')
  );
}

/**
 * 检测列对齐方式
 * @param {Array<Array>} rows
 * @param {boolean} includeAlignment
 * @returns {Array<string>}
 */
function detectAlignment(rows, includeAlignment) {
  if (!includeAlignment || rows.length < 2) {
    return rows[0].map(() => '---');
  }

  // 第一行为表头，不用于类型检测
  const dataRows = rows.slice(1);

  return rows[0].map((_, colIndex) => {
    const values = dataRows.map(row => row[colIndex]);
    const numericCount = values.filter(v => typeof v === 'number' || (!isNaN(parseFloat(v)) && isFinite(v))).length;
    const totalNonEmpty = values.filter(v => v !== null && v !== '').length;

    // 如果超过 50% 是数字，则右对齐
    if (totalNonEmpty > 0 && numericCount / totalNonEmpty > 0.5) {
      return '---:';
    }
    return ':---';
  });
}

/**
 * 构建 Markdown 表格字符串
 * @param {Array<Array>} rows
 * @param {Array<string>} alignment
 * @returns {string}
 */
function buildTableString(rows, alignment) {
  const lines = [];
  const colCount = rows[0].length;
  const separator = '|' + alignment.map(a => ` ${a} `).join('|') + '|';

  // 表头行
  lines.push('|' + rows[0].map(cell => ` ${escapeCell(cell)} `).join('|') + '|');
  // 分隔行
  lines.push(separator);
  // 数据行
  for (let i = 1; i < rows.length; i++) {
    lines.push('|' + rows[i].map(cell => ` ${escapeCell(cell)} `).join('|') + '|');
  }

  return lines.join('\n');
}

/**
 * 转义单元格内容中的特殊字符
 * @param {*} cell
 * @returns {string}
 */
function escapeCell(cell) {
  if (cell === null || cell === undefined) return '';
  const str = String(cell);
  return str.replace(/\|/g, '\\|').replace(/\n/g, '<br>');
}

export { generateMarkdownTable, parseRange, detectAlignment, buildTableString, escapeCell };