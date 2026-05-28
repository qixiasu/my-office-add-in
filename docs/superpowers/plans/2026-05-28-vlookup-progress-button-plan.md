# VLOOKUP 执行按钮保护与进度显示实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在 VLOOKUP 执行过程中禁用按钮并显示视觉化进度条，防止用户重复点击

**Architecture:** 在 `performLookup` 函数入口禁用按钮，在 finally 中恢复；小数据模式使用 setInterval 模拟进度动画，大数据模式复用现有的分批进度更新

**Tech Stack:** Vanilla JS + CSS，Office Excel API

---

## 文件变更概览

| 文件 | 修改内容 |
|------|----------|
| `src/taskpane/vlookup-taskpane.html` | 进度条 DOM 插入到状态消息前 |
| `src/taskpane/vlookup-taskpane.css` | 进度条样式（进度条填充动画） |
| `src/taskpane/vlookup-taskpane.js` | 按钮禁用 + updateProgressUI + 进度集成 |

---

## Task 1: HTML 添加进度条 DOM

**Files:**
- Modify: `src/taskpane/vlookup-taskpane.html:116-118`

进度条 DOM 放在状态消息上方，按钮行下方。

- [ ] **Step 1: 修改 vlookup-taskpane.html**

找到第 116-118 行（状态消息 div）：

```html
        <!-- Status message -->
        <div id="statusMessage" class="status-message status-idle">状态：等待操作...</div>
```

在按钮行（第 112-114 行）后面、状态消息（第 117 行）前面，插入：

```html
        <!-- Progress bar (hidden by default) -->
        <div id="progressContainer" style="display:none;margin:12px 0">
            <div id="progressStatus" class="status-message status-loading" style="margin-bottom:6px;font-size:14px;font-weight:600">处理中...</div>
            <div style="height:8px;background:#e0e0e0;border-radius:4px;overflow:hidden">
                <div id="progressBarFill" style="width:0%;height:100%;background:linear-gradient(90deg,#0078d4,#00b4d8);border-radius:4px;transition:width 0.3s ease"></div>
            </div>
            <div id="progressDetail" style="color:#666;font-size:11px;margin-top:4px"></div>
        </div>

        <!-- Status message -->
        <div id="statusMessage" class="status-message status-idle">状态：等待操作...</div>
```

---

## Task 2: CSS 添加进度条样式

**Files:**
- Modify: `src/taskpane/vlookup-taskpane.css`

在文件末尾 `.status-loading` 后面添加：

- [ ] **Step 1: 添加进度条样式到 vlookup-taskpane.css**

在第 216 行后（`.status-loading` 块之后）添加：

```css
#progressContainer {
  margin: 12px 0;
}

#progressStatus {
  padding: 0;
  background: transparent;
  border: none;
  color: #0057b7;
}
```

---

## Task 3: JS 改造 performLookup — 按钮禁用 + 进度更新

**Files:**
- Modify: `src/taskpane/vlookup-taskpane.js:309-578`（performLookup 函数）

- [ ] **Step 1: 在 performLookup 开头添加按钮禁用逻辑**

找到 `performLookup` 函数（第 309 行），在第一行 `setStatus("处理中...", "info");` 之前插入：

```javascript
function performLookup(config) {
  var executeBtn = document.getElementById("executeBtn");
  var progressContainer = document.getElementById("progressContainer");
  var statusEl = document.getElementById("statusMessage");

  // 禁用按钮，显示进度条
  executeBtn.disabled = true;
  if (progressContainer) progressContainer.style.display = "block";
  statusEl.style.display = "none";
```

- [ ] **Step 2: 在 catch 之前添加 finally 块确保按钮恢复**

找到 `performLookup` 结尾的 `catch` 块（第 575-577 行）：

```javascript
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
  });
```

将其修改为：

```javascript
  }).catch(function (error) {
    setStatus("错误: " + error.message, "error");
  }).finally(function() {
    executeBtn.disabled = false;
    if (progressContainer) progressContainer.style.display = "none";
    statusEl.style.display = "block";
  });
```

- [ ] **Step 3: 添加 updateProgressUI 工具函数**

在 `setStatus` 函数（第 278 行）之后添加：

```javascript
function updateProgressUI(percent, completed, total) {
  var statusEl = document.getElementById("progressStatus");
  var detailEl = document.getElementById("progressDetail");
  var barEl = document.getElementById("progressBarFill");

  if (statusEl) {
    statusEl.textContent = "处理中... " + percent + "%";
    statusEl.className = "status-message status-loading";
  }

  if (barEl) {
    barEl.style.width = percent + "%";
  }

  if (detailEl && total > 0) {
    detailEl.textContent = "已完成 " + completed + " / " + total + " 行";
  }
}
```

- [ ] **Step 4: 修改小数据模式（if 分支）添加进度条**

找到第 389 行附近的小数据模式分支开头：

```javascript
    if (dataRowCount < LARGE_DATA_THRESHOLD) {
```

在这一行之后、`// Small data: single read...` 注释之前，添加模拟进度：

```javascript
      if (dataRowCount < LARGE_DATA_THRESHOLD) {
        var totalRows = dataRowCount;
        var progressStep = 0;
        var progressInterval = setInterval(function() {
          progressStep = Math.min(progressStep + 10, 90);
          var percent = progressStep;
          updateProgressUI(percent, Math.round(totalRows * percent / 100), totalRows);
        }, 200);
```

- [ ] **Step 5: 在小数据模式写入完成后清除定时器**

找到小数据模式写入完成后的位置（约第 507 行）：

```javascript
      setStatus(
        "完成! 已写入 " + results.length + " 行 x " + returnColCount + " 列静态值",
        "success"
      );
```

在这之后、`} else {`（大数据模式分支）之前，添加：

```javascript
        clearInterval(progressInterval);
        updateProgressUI(100, results.length, results.length);
      } else {
```

- [ ] **Step 6: 修改大数据模式（else 分支）使用 updateProgressUI**

找到大数据模式的进度更新行（第 570 行）：

```javascript
        setStatus("处理中... 已完成 " + processedRows + " / " + totalRows + " 行", "info");
```

替换为：

```javascript
        var percent = Math.round((processedRows / totalRows) * 100);
        updateProgressUI(percent, processedRows, totalRows);
```

找到大数据模式完成后的状态设置（约第 573 行）：

```javascript
      setStatus("完成! 已写入 " + totalRows + " 行 x " + returnColCount + " 列静态值", "success");
```

保持不变（updateProgressUI 在每批已调用，finally 阶段无需额外处理）。

---

## Task 4: 验证并测试

**Files:**
- Test: `src/taskpane/vlookup-taskpane.js`

- [ ] **Step 1: 启动开发服务器**

```bash
npm run dev-server
```

- [ ] **Step 2: 启动 Excel 调试**

```bash
npm run start
```

- [ ] **Step 3: 测试按钮禁用**

1. 在 Excel 中选择查找值区域
2. 选择查找表，点击"刷新选择"
3. 选择匹配列和返回列
4. 点击"执行查找"
5. 观察按钮立即变为灰色禁用状态
6. 进度条出现并更新
7. 完成后按钮恢复

- [ ] **Step 4: 测试小数据进度**

选择少量数据（<100000行），观察进度条动画是否正常显示。

---

## 验收标准

- [ ] 点击执行按钮后按钮立即禁用（灰色状态）
- [ ] 处理过程中显示进度条（蓝色渐变填充动画）
- [ ] 小数据模式（≤100000行）显示进度条
- [ ] 大数据模式（>100000行）显示进度条
- [ ] 处理完成后按钮恢复可用
- [ ] 错误发生时按钮也能恢复
- [ ] 进度条宽度平滑过渡（transition: width 0.3s）
