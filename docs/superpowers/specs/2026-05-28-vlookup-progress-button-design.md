# VLOOKUP 执行按钮保护与进度显示设计

## 背景与目标

用户反馈两个体验问题：
1. 查找功能执行时没有禁用按钮，导致用户可能多次点击
2. 查找过程中没有视觉进度反馈，用户以为程序卡死

## 设计方案（方案 A）

### 1. 执行按钮保护

**行为：**
- 用户点击「执行查找」按钮后，按钮立即被禁用（灰色 + cursor:not-allowed）
- 进入 `Excel.run` 处理阶段
- 处理完成（无论成功/失败）后恢复按钮可用
- 使用 `try-finally` 确保按钮状态恢复

**状态变化：**
```
空闲 → 执行查找（按钮点击）
       ↓
    按钮 disabled=true
    进度条开始动画
       ↓
    处理中...（不断更新进度）
       ↓
    完成/错误
       ↓
    按钮 disabled=false
```

### 2. 进度条显示

**适用场景：**
- 大数据模式（>100000行）：批处理 + 分批进度更新
- 小数据模式（≤100000行）：单次处理 + 模拟进度

**UI 结构：**
```
状态：处理中... 50%
[████████████████░░░░░░░] 50%
已完成 50000 / 100000 行
```

**组件说明：**
- `statusMessage`：显示百分比文字（如"处理中... 50%"）
- `progressBar`：视觉化进度条，蓝色渐变填充
- `statusDetail`：底部行数计数（如"已完成 X / Y 行"）

**进度条样式：**
- 高度：8px
- 背景：#e0e0e0（灰色）
- 填充：linear-gradient(#0057b7, #00b4d8)（蓝到青）
- 圆角：4px

### 3. 核心代码修改

**performLookup 函数改造：**

```javascript
async function performLookup(config) {
  var executeBtn = document.getElementById("executeBtn");
  
  // 禁用按钮
  executeBtn.disabled = true;
  
  try {
    await Excel.run(async (context) => {
      // ... 大数据模式 / 小数据模式逻辑 ...
      
      // 小数据模式：添加进度条
      if (dataRowCount < LARGE_DATA_THRESHOLD) {
        var progressStep = 0;
        var progressInterval = setInterval(function() {
          progressStep = Math.min(progressStep + 10, 90);
          var percent = progressStep;
          updateProgressUI(percent, 0, dataRowCount);
        }, 200);
        
        // ... 执行查找 ...
        
        clearInterval(progressInterval);
        updateProgressUI(100, dataRowCount, dataRowCount);
      } else {
        // 大数据模式：使用现有的分批进度
        while (processedRows < totalRows) {
          // ... 处理每批 ...
          var percent = Math.round((processedRows / totalRows) * 100);
          updateProgressUI(percent, processedRows, totalRows);
        }
      }
    });
  } finally {
    executeBtn.disabled = false;
  }
}

function updateProgressUI(percent, completed, total) {
  var statusEl = document.getElementById("statusMessage");
  var detailEl = document.getElementById("statusDetail");
  var barEl = document.getElementById("progressBarFill");
  
  statusEl.textContent = "处理中... " + percent + "%";
  statusEl.className = "status-message status-info";
  
  if (barEl) {
    barEl.style.width = percent + "%";
  }
  
  if (detailEl && total > 0) {
    detailEl.textContent = "已完成 " + completed + " / " + total + " 行";
  }
}
```

### 4. HTML 修改

在 `vlookup-taskpane.html` 中添加进度条组件：

```html
<div id="progressContainer" style="display:none;margin:12px 0">
  <div id="statusMessage" class="status-message status-info" 
       style="margin-bottom:6px;font-size:14px;font-weight:600">
    处理中...
  </div>
  <div style="height:8px;background:#e0e0e0;border-radius:4px;overflow:hidden">
    <div id="progressBarFill" style="width:0%;height:100%;
         background:linear-gradient(90deg,#0057b7,#00b4d8);border-radius:4px;
         transition:width 0.3s ease"></div>
  </div>
  <div id="statusDetail" style="color:#666;font-size:11px;margin-top:4px"></div>
</div>
```

### 5. 状态映射

| 状态 | 文字 | 样式 |
|------|------|------|
| idle | 状态：等待操作... | gray |
| loading | 处理中... 50% | blue + 进度条 |
| success | 完成! 已写入 X 行 | green |
| error | 错误: XXX | red |

## 实现步骤

1. 修改 `vlookup-taskpane.html`：添加进度条 DOM 结构
2. 修改 `vlookup-taskpane.js`：
   - 在 `performLookup` 开头禁用按钮
   - 添加 `updateProgressUI` 函数
   - 小数据模式使用 `setInterval` 模拟进度
   - 大数据模式复用现有的分批进度
   - 使用 `finally` 确保按钮恢复

## 验收标准

- [ ] 点击执行按钮后按钮立即禁用（灰色）
- [ ] 处理过程中显示进度条动画
- [ ] 小数据模式（≤100000行）也有进度条
- [ ] 处理完成后按钮恢复可用
- [ ] 错误发生时按钮也能恢复
