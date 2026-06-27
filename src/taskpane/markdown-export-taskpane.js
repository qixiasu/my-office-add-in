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
      range.load(['values', 'mergedRanges', 'numberFormat', 'rowCount', 'columnCount']);

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
  let success = false;

  try {
    await navigator.clipboard.writeText(markdown);
    success = true;
  } catch (error) {
    // Fallback for IE11
    try {
      const textarea = document.createElement('textarea');
      textarea.value = markdown;
      textarea.style.position = 'fixed';
      textarea.style.opacity = '0';
      document.body.appendChild(textarea);
      textarea.select();
      document.execCommand('copy');
      document.body.removeChild(textarea);
      success = true;
    } catch (e) {
      try { document.body.removeChild(textarea); } catch (e2) {}
    }
  }

  if (success) showToast();
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
  Office.context.ui.close();
}