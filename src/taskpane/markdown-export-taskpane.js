/* global Office, Excel */

import { generateMarkdownTable } from "../utils/markdown-table-utils";

let toastTimer = null;
let currentMarkdown = ""; // 缓存当前生成的 markdown

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    initialize();
  }
});

function initialize() {
  document.getElementById("refreshBtn").addEventListener("click", refreshSelection);
  document.getElementById("generateBtn").addEventListener("click", generateMarkdown);
  document.getElementById("copyBtn").addEventListener("click", copyToClipboard);

  // 初始化按钮状态
  updateButtonStates(false);
}

function updateButtonStates(hasPreview) {
  const copyBtn = document.getElementById("copyBtn");
  if (copyBtn) {
    copyBtn.disabled = !hasPreview;
  }
}

/**
 * 刷新选区信息（从 Excel 获取当前选区范围地址）
 */
async function refreshSelection() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["address"]);
      await context.sync();

      // 更新预览信息显示选区地址
      const address = range.address;
      document.getElementById("previewInfo").textContent = `已选: ${address}`;
    });
  } catch (error) {
    showError("读取选区失败：" + error.message);
  }
}

/**
 * 生成 Markdown 预览
 */
async function generateMarkdown() {
  try {
    // 清除之前的警告
    hideWarning();

    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      // eslint-disable-next-line office-addins/load-object-before-read -- mergedRanges needed to detect merged cells
      range.load(["values", "mergedRanges", "address", "rowCount", "columnCount"]);

      await context.sync();

      const values = range.values;
      const mergedRanges = range.mergedRanges;
      const address = range.address;

      if (!values || !values.length || !values[0].length) {
        showError("请先选择数据区域");
        updateButtonStates(false);
        return;
      }

      // 检查数据量
      const rowCount = values.length;
      const colCount = values[0].length;
      if (rowCount > 10000) {
        showError("数据量过大（超过 10000 行），请缩小选区");
        updateButtonStates(false);
        return;
      }

      // 检查合并单元格
      if (mergedRanges && mergedRanges.length > 0) {
        showWarning("此表格含合并单元格，生成的 Markdown 可能不完美");
      }

      // 生成 Markdown（只取前10行预览）
      const previewRows = Math.min(values.length, 10);
      const previewValues = values.slice(0, previewRows);
      const isTruncated = values.length > 10;

      currentMarkdown = generateMarkdownTable(values, {
        includeAlignment: true,
        preserveFormat: true,
      });

      // 显示预览（带截断提示）
      let previewMarkdown = generateMarkdownTable(previewValues, {
        includeAlignment: true,
        preserveFormat: true,
      });

      if (isTruncated) {
        previewMarkdown += "\n\n*... 共 " + values.length + " 行，仅显示前 10 行 ...*";
      }

      document.getElementById("markdownPreview").textContent = previewMarkdown;
      document.getElementById("previewInfo").textContent =
        `${address} | ${rowCount}行 × ${colCount}列`;

      updateButtonStates(true);
    });
  } catch (error) {
    showError("生成 Markdown 失败：" + error.message);
    updateButtonStates(false);
  }
}

/**
 * 复制到剪贴板
 */
async function copyToClipboard() {
  if (!currentMarkdown) {
    showError("请先生成 Markdown");
    return;
  }

  let success = false;
  let textarea = null;

  try {
    await navigator.clipboard.writeText(currentMarkdown);
    success = true;
  } catch {
    // Fallback for IE11
    textarea = document.createElement("textarea");
    textarea.value = currentMarkdown;
    textarea.style.position = "fixed";
    textarea.style.opacity = "0";
    document.body.appendChild(textarea);
    textarea.select();
    try {
      document.execCommand("copy");
      success = true;
    } catch {
      // Silent fail - IE11 fallback may not work in all contexts
    } finally {
      if (textarea.parentNode) {
        textarea.parentNode.removeChild(textarea);
      }
    }
  }

  if (success) {
    showToast();
  } else {
    showError("复制失败，请手动复制");
  }
}

function showToast() {
  const toast = document.getElementById("toast");
  toast.classList.add("show");

  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    toast.classList.remove("show");
  }, 2000);
}

function showError(message) {
  document.getElementById("warningBox").style.display = "flex";
  document.getElementById("warningBox").style.background = "#f8d7da";
  document.getElementById("warningBox").style.borderColor = "#f5c6cb";
  document.getElementById("warningBox").style.color = "#721c24";
  document.getElementById("warningText").textContent = message;
  document.getElementById("markdownPreview").textContent = "";
  currentMarkdown = "";
  updateButtonStates(false);
}

function showWarning(message) {
  document.getElementById("warningBox").style.display = "flex";
  document.getElementById("warningBox").style.background = "#fff3cd";
  document.getElementById("warningBox").style.borderColor = "#ffc107";
  document.getElementById("warningBox").style.color = "#856404";
  document.getElementById("warningText").textContent = message;
}

function hideWarning() {
  document.getElementById("warningBox").style.display = "none";
}
