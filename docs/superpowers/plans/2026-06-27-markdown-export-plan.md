# Markdown 表格导出功能实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 用户选中 Excel 区域后，点击功能区按钮，弹窗预览 Markdown 表格并支持一键复制

**Architecture:** 转换逻辑封装在 `markdown-table-utils.js`，Taskpane 负责 UI 交互和 Excel API 调用

**Tech Stack:** Office JavaScript API, 原生 JavaScript (IE11 兼容)

---

## 文件结构

```
src/
├── utils/
│   ├── markdown-table-utils.js       # 新建：核心转换逻辑
│   └── markdown-table-utils.test.js  # 新建：单元测试
├── taskpane/
│   ├── markdown-export-taskpane.html # 新建：弹窗 HTML
│   ├── markdown-export-taskpane.js   # 新建：弹窗逻辑
│   └── markdown-export-taskpane.css  # 新建：弹窗样式
manifest.xml                           # 修改：添加功能区按钮
```

---

## Task 1: 创建 markdown-table-utils.js 基础版本

**Files:**
- Create: `src/utils/markdown-table-utils.js`

- [ ] **Step 1: 创建文件骨架**

```javascript
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
```

- [ ] **Step 2: 验证文件创建成功**

Run: `ls -la src/utils/markdown-table-utils.js`
Expected: 文件存在

- [ ] **Step 3: 提交**

```bash
git add src/utils/markdown-table-utils.js
git commit -m "feat(markdown-export): add markdown table utils core logic"
```

---

## Task 2: 创建单元测试

**Files:**
- Create: `src/utils/markdown-table-utils.test.js`

- [ ] **Step 1: 编写测试用例**

```javascript
import { generateMarkdownTable, parseRange, detectAlignment, escapeCell } from './markdown-table-utils';

describe('markdown-table-utils', () => {
  describe('parseRange', () => {
    it('should filter empty rows', () => {
      const input = [
        ['A', 'B'],
        [null, null],
        ['C', 'D']
      ];
      const result = parseRange(input);
      expect(result).toEqual([
        ['A', 'B'],
        ['C', 'D']
      ]);
    });

    it('should filter trailing empty columns', () => {
      const input = [
        ['A', 'B', null],
        ['C', 'D', null]
      ];
      const result = parseRange(input);
      expect(result).toEqual([
        ['A', 'B'],
        ['C', 'D']
      ]);
    });

    it('should return empty array for all empty input', () => {
      const input = [
        [null, null],
        [null, null]
      ];
      const result = parseRange(input);
      expect(result).toEqual([]);
    });
  });

  describe('detectAlignment', () => {
    it('should detect numeric columns as right-aligned', () => {
      const rows = [
        ['Name', 'Score'],
        ['Alice', 100],
        ['Bob', 95]
      ];
      const result = detectAlignment(rows, true);
      expect(result).toEqual([':---', '---:']);
    });

    it('should detect text columns as left-aligned', () => {
      const rows = [
        ['Name', 'City'],
        ['Alice', 'NYC'],
        ['Bob', 'LA']
      ];
      const result = detectAlignment(rows, true);
      expect(result).toEqual([':---', ':---']);
    });

    it('should return basic separator when includeAlignment is false', () => {
      const rows = [['A', 'B'], ['C', 'D']];
      const result = detectAlignment(rows, false);
      expect(result).toEqual(['---', '---']);
    });
  });

  describe('escapeCell', () => {
    it('should escape pipe characters', () => {
      expect(escapeCell('A|B')).toBe('A\\|B');
    });

    it('should replace newlines with br tag', () => {
      expect(escapeCell('A\nB')).toBe('A<br>B');
    });

    it('should handle null and undefined', () => {
      expect(escapeCell(null)).toBe('');
      expect(escapeCell(undefined)).toBe('');
    });
  });

  describe('generateMarkdownTable', () => {
    it('should generate basic markdown table', () => {
      const values = [
        ['Name', 'Age'],
        ['Alice', 25],
        ['Bob', 30]
      ];
      const result = generateMarkdownTable(values);
      expect(result).toContain('| Name | Age |');
      expect(result).toContain('| :--- | ---: |');
    });

    it('should return empty string for empty input', () => {
      expect(generateMarkdownTable([])).toBe('');
      expect(generateMarkdownTable(null)).toBe('');
    });
  });
});
```

- [ ] **Step 2: 运行测试验证**

Run: `npm test -- --testPathPattern=markdown-table-utils`
Expected: 所有测试通过

- [ ] **Step 3: 提交**

```bash
git add src/utils/markdown-table-utils.test.js
git commit -m "test(markdown-export): add markdown table utils unit tests"
```

---

## Task 3: 创建 markdown-export-taskpane 弹窗

**Files:**
- Create: `src/taskpane/markdown-export-taskpane.html`
- Create: `src/taskpane/markdown-export-taskpane.js`
- Create: `src/taskpane/markdown-export-taskpane.css`

- [ ] **Step 1: 创建 HTML**

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta http-equiv="Content-Security-Policy" content="default-src 'self' 'unsafe-inline' https://localhost:*;">
  <title>导出 Markdown</title>
</head>
<body>
  <div class="markdown-export-container">
    <div class="header">
      <h2>导出 Markdown 表格</h2>
      <button id="closeBtn" class="close-btn" aria-label="关闭">&times;</button>
    </div>
    
    <div id="warningBox" class="warning-box" style="display: none;">
      <span class="warning-icon">⚠️</span>
      <span id="warningText"></span>
    </div>

    <div class="preview-section">
      <div class="preview-header">
        <span>预览</span>
        <button id="copyBtn" class="copy-btn">复制到剪贴板</button>
      </div>
      <pre id="markdownPreview" class="markdown-preview"></pre>
    </div>

    <div id="toast" class="toast">已复制到剪贴板</div>
  </div>
</body>
</html>
```

- [ ] **Step 2: 创建 CSS**

```css
.markdown-export-container {
  display: flex;
  flex-direction: column;
  height: 100vh;
  padding: 16px;
  box-sizing: border-box;
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
  background: #fff;
  position: relative;
}

.header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 16px;
}

.header h2 {
  margin: 0;
  font-size: 18px;
  font-weight: 600;
  color: #333;
}

.close-btn {
  background: none;
  border: none;
  font-size: 24px;
  cursor: pointer;
  color: #666;
  padding: 0;
  line-height: 1;
}

.close-btn:hover {
  color: #333;
}

.warning-box {
  background: #fff3cd;
  border: 1px solid #ffc107;
  border-radius: 4px;
  padding: 10px 12px;
  margin-bottom: 16px;
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 13px;
  color: #856404;
}

.preview-section {
  flex: 1;
  display: flex;
  flex-direction: column;
  border: 1px solid #ddd;
  border-radius: 6px;
  overflow: hidden;
}

.preview-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 10px 12px;
  background: #f6f8fa;
  border-bottom: 1px solid #ddd;
  font-size: 13px;
  font-weight: 500;
  color: #333;
}

.copy-btn {
  background: #2ea44f;
  color: #fff;
  border: none;
  padding: 6px 12px;
  border-radius: 6px;
  font-size: 13px;
  cursor: pointer;
  transition: background-color 0.2s;
}

.copy-btn:hover {
  background: #2c974b;
}

.markdown-preview {
  flex: 1;
  margin: 0;
  padding: 12px;
  overflow: auto;
  font-family: 'Consolas', 'Monaco', monospace;
  font-size: 13px;
  line-height: 1.5;
  white-space: pre-wrap;
  word-break: break-word;
  background: #fff;
}

.toast {
  position: fixed;
  bottom: 20px;
  left: 50%;
  transform: translateX(-50%) translateY(100px);
  background: #333;
  color: #fff;
  padding: 10px 20px;
  border-radius: 6px;
  font-size: 14px;
  opacity: 0;
  transition: transform 0.3s, opacity 0.3s;
  z-index: 1000;
}

.toast.show {
  transform: translateX(-50%) translateY(0);
  opacity: 1;
}
```

- [ ] **Step 3: 创建 JS**

```javascript
import { generateMarkdownTable } from '../utils/markdown-table-utils';

let toastTimer = null;

Office.onReady(function(info) {
  if (info.host === Office.HostType.Excel) {
    initialize();
  }
});

function initialize() {
  document.getElementById('closeBtn').addEventListener('click', closePanel);
  document.getElementById('copyBtn').addEventListener('click', copyToClipboard);
  
  loadSelectedRange();
}

async function loadSelectedRange() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(['values', 'mergedRanges', 'numberFormat']);
      
      await context.sync();
      
      const values = range.values;
      const mergedRanges = range.mergedRanges;
      
      if (!values || !values.length || !values[0].length) {
        showError('请先选择数据区域');
        return;
      }
      
      // 检查数据量
      const rowCount = values.length;
      const colCount = values[0].length;
      if (rowCount > 10000) {
        showError('数据量过大（超过 10000 行），请缩小选区');
        return;
      }
      
      // 检查合并单元格
      if (mergedRanges && mergedRanges.length > 0) {
        showWarning('此表格含合并单元格，生成的 Markdown 可能不完美');
      }
      
      // 生成 Markdown
      const markdown = generateMarkdownTable(values, {
        includeAlignment: true,
        preserveFormat: true
      });
      
      document.getElementById('markdownPreview').textContent = markdown;
    });
  } catch (error) {
    showError('读取选区失败：' + error.message);
  }
}

async function copyToClipboard() {
  const markdown = document.getElementById('markdownPreview').textContent;
  
  try {
    await navigator.clipboard.writeText(markdown);
    showToast();
  } catch (error) {
    // Fallback for IE11
    const textarea = document.createElement('textarea');
    textarea.value = markdown;
    textarea.style.position = 'fixed';
    textarea.style.opacity = '0';
    document.body.appendChild(textarea);
    textarea.select();
    document.execCommand('copy');
    document.body.removeChild(textarea);
    showToast();
  }
}

function showToast() {
  const toast = document.getElementById('toast');
  toast.classList.add('show');
  
  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    toast.classList.remove('show');
  }, 2000);
}

function showError(message) {
  document.getElementById('warningBox').style.display = 'flex';
  document.getElementById('warningBox').style.background = '#f8d7da';
  document.getElementById('warningBox').style.borderColor = '#f5c6cb';
  document.getElementById('warningBox').style.color = '#721c24';
  document.getElementById('warningText').textContent = message;
  document.getElementById('markdownPreview').textContent = '';
}

function showWarning(message) {
  document.getElementById('warningBox').style.display = 'flex';
  document.getElementById('warningBox').style.background = '#fff3cd';
  document.getElementById('warningBox').style.borderColor = '#ffc107';
  document.getElementById('warningBox').style.color = '#856404';
  document.getElementById('warningText').textContent = message;
}

function closePanel() {
  Office.UI.close();
}
```

- [ ] **Step 4: 验证文件创建成功**

Run: `ls -la src/taskpane/markdown-export-taskpane.*`
Expected: 三个文件都存在

- [ ] **Step 5: 提交**

```bash
git add src/taskpane/markdown-export-taskpane.html src/taskpane/markdown-export-taskpane.js src/taskpane/markdown-export-taskpane.css
git commit -m "feat(markdown-export): add markdown export taskpane UI"
```

---

## Task 4: 更新 manifest.xml 添加功能区按钮

**Files:**
- Modify: `manifest.xml`

- [ ] **Step 1: 在现有功能组添加按钮，或新增功能组**

在 manifest.xml 中找到 "数据处理" 组（对应 "连接列" 等按钮），添加新按钮：

```xml
<control xsi:type="Button" id="MarkdownExportButton">
  <overwrites>
    <overwrite v="toString">导出 Markdown</overwrite>
  </overwrites>
  <icon>
    <bt:Image size="16" resid="icon16"/>
    <bt:Image size="32" resid="icon32"/>
    <bt:Image size="80" resid="icon80"/>
  </icon>
  <supertip>
    <title textresid="MarkdownExportButton.Supertip.Title"/>
    <description textresid="MarkdownExportButton.Supertip.Description"/>
  </supertip>
  <action type="ShowTaskpane">
    <sourceLocation resid="MarkdownExportTaskpane.Url"/>
  </action>
</control>
```

- [ ] **Step 2: 添加 Resources > URLs 中的 sourceLocation**

```xml
<bt:Urls>
  <bt:Url id="MarkdownExportTaskpane.Url" DefaultValue="https://localhost:3000/markdown-export-taskpane.html"/>
  <!-- 现有其他 URLs -->
</bt:Urls>
```

- [ ] **Step 3: 添加 Strings 中的文本资源**

```xml
<bt:Strings>
  <bt:String id="MarkdownExportButton.Supertip.Title" DefaultValue="导出 Markdown"/>
  <bt:String id="MarkdownExportButton.Supertip.Description" DefaultValue="将选中区域转换为 Markdown 表格"/>
  <!-- 现有其他 Strings -->
</bt:Strings>
```

- [ ] **Step 4: 验证 manifest.xml 语法**

Run: `npm run validate`
Expected: 验证通过

- [ ] **Step 5: 提交**

```bash
git add manifest.xml
git commit -m "feat(markdown-export): add ribbon button to manifest"
```

---

## Task 5: 更新 webpack.config.js 支持新入口

**Files:**
- Modify: `webpack.config.js`

- [ ] **Step 1: 添加新入口**

在 `entry` 配置中添加：

```javascript
'markdown-export-taskpane': './src/taskpane/markdown-export-taskpane.js',
```

- [ ] **Step 2: 验证构建**

Run: `npm run build:dev`
Expected: 构建成功，无错误

- [ ] **Step 3: 提交**

```bash
git add webpack.config.js
git commit -m "build(markdown-export): add markdown-export-taskpane entry"
```

---

## Task 6: 合并单元格支持（扩展功能）

**Files:**
- Modify: `src/utils/markdown-table-utils.js`
- Modify: `src/utils/markdown-table-utils.test.js`

- [ ] **Step 1: 更新 generateMarkdownTable 签名**

添加 `mergedRanges` 参数：

```javascript
/**
 * 生成 Markdown 表格（支持合并单元格）
 * @param {Array<Array>} values - 二维单元格值数组
 * @param {Object} options - 配置选项
 * @param {Array} options.mergedRanges - 合并区域信息
 * @param {boolean} options.includeAlignment - 是否包含对齐标记
 * @param {boolean} options.preserveFormat - 是否保留显示格式
 * @returns {string} Markdown 表格字符串
 */
function generateMarkdownTable(values, options = {}) {
  const { 
    mergedRanges = [], 
    includeAlignment = true, 
    preserveFormat = true 
  } = options;
  // ... 原有逻辑
}
```

- [ ] **Step 2: 实现 handleMergedCells 函数**

```javascript
/**
 * 处理合并单元格，标记需要 colspan/rowspan 的位置
 * @param {Array<Array>} values
 * @param {Array} mergedRanges - 合并区域信息
 * @returns {Map} 合并单元格位置映射
 */
function handleMergedCells(values, mergedRanges) {
  const mergedCells = new Map();
  
  for (const range of mergedRanges) {
    const { startRow, startColumn, rowCount, columnCount } = parseMergedRange(range);
    
    if (rowCount > 1 || columnCount > 1) {
      const key = `${startRow}-${startColumn}`;
      mergedCells.set(key, {
        rowspan: rowCount,
        colspan: columnCount,
        value: values[startRow]?.[startColumn] ?? ''
      });
    }
  }
  
  return mergedCells;
}

function parseMergedRange(range) {
  // Excel merged range format: "A1:C3"
  // 需要解析起始行列和行列数
  // ...
}
```

- [ ] **Step 3: 添加测试用例**

```javascript
it('should handle merged cells with rowspan and colspan', () => {
  // 测试代码
});
```

- [ ] **Step 4: 运行测试**

Run: `npm test -- --testPathPattern=markdown-table-utils`
Expected: 所有测试通过

- [ ] **Step 5: 提交**

```bash
git add src/utils/markdown-table-utils.js src/utils/markdown-table-utils.test.js
git commit -m "feat(markdown-export): add merged cells support"
```

---

## Task 7: 运行完整测试验证

- [ ] **Step 1: 运行所有测试**

Run: `npm test`
Expected: 所有测试通过

- [ ] **Step 2: 本地验证功能**

1. `npm run dev-server` 启动开发服务器
2. `npm run start` 启动 Excel 加载项
3. 选中包含数据的区域
4. 点击 "导出 Markdown" 按钮
5. 验证弹窗显示正确

- [ ] **Step 3: 提交所有更改**

```bash
git add -A
git commit -m "feat: complete markdown export feature"
```

---

## 实施检查清单

- [ ] Task 1: markdown-table-utils.js 基础版本
- [ ] Task 2: 单元测试
- [ ] Task 3: Taskpane 弹窗 UI
- [ ] Task 4: manifest.xml 功能区按钮
- [ ] Task 5: webpack.config.js 入口
- [ ] Task 6: 合并单元格支持
- [ ] Task 7: 完整测试验证