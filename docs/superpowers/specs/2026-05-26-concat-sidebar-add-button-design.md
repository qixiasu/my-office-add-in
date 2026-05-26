# 连接列侧边栏添加执行按钮设计

## 概述

在连接列侧边栏添加执行按钮，采用次要按钮（轮廓样式），移除 Enter 键触发，改用纯按钮点击执行。

## 设计决策

| 项目 | 选择 |
|------|------|
| 按钮样式 | 次要按钮（轮廓样式）- 蓝色边框 |
| 按钮位置 | 输入框下方 |
| 触发方式 | 仅按钮点击执行（移除 Enter 键监听） |

## 视觉设计

### 按钮样式

```css
.concat-button {
  width: 100%;
  padding: 10px 16px;
  background: transparent;
  color: #0078d4;
  border: 2px solid #0078d4;
  border-radius: 6px;
  font-size: 14px;
  font-weight: 600;
  cursor: pointer;
  transition: all 0.2s;
}

.concat-button:hover {
  background: #0078d4;
  color: white;
}

.concat-button:active {
  background: #005a9e;
  border-color: #005a9e;
}

.concat-button:disabled {
  background: #f5f5f5;
  color: #999;
  border-color: #999;
  cursor: not-allowed;
}
```

### 按钮位置

```
┌─────────────────────────────┐
│ 🔗 连接列工具                │
├─────────────────────────────┤
│ ┌─────────────────────────┐ │
│ │ 1. 选中要连接的两列      │ │
│ │ 2. 输入连接符，点击执行  │ │
│ └─────────────────────────┘ │
│                             │
│ 连接符 [________]           │
│                             │
│     [ 执行连接 ]            │  ← 按钮在输入框下方
│                             │
│ 状态: 等待操作...           │
└─────────────────────────────┘
```

## 交互设计

### 状态变化

| 状态 | 按钮样式 |
|------|----------|
| 等待 | 可点击，蓝色边框 |
| 处理中 | 禁用，不可点击 |
| 成功 | 可点击，蓝色边框 |
| 错误 | 可点击，蓝色边框 |

### 步骤引导更新

原：`2. 输入连接符，按 Enter 执行`  
改为：`2. 输入连接符，点击执行`

## 技术实现

### 文件变更

**修改文件：**
- `src/taskpane/concat-taskpane.html` — 添加按钮元素
- `src/taskpane/concat-taskpane.css` — 添加按钮样式
- `src/taskpane/concat-taskpane.js` — 移除 Enter 键监听，添加按钮点击事件

### concat-taskpane.html 变更

```html
<div class="concat-form-group">
  <label for="connector" class="concat-label">连接符</label>
  <input type="text" id="connector" class="concat-input" value="_" />
</div>

<button id="executeBtn" class="concat-button">执行连接</button>

<div id="status" class="status-message status-idle">状态：等待操作...</div>
```

### concat-taskpane.js 变更

1. 移除输入框的 `keydown` Enter 监听
2. 添加按钮的 `click` 事件监听
3. 按钮在处理中时设为禁用

```javascript
// 移除 Enter 键监听
// 原来的 keydown 监听代码删除

// 添加按钮点击事件
var executeBtn = document.getElementById("executeBtn");
executeBtn.addEventListener("click", function() {
  runConcat();
});

// 在 runConcat 开始时禁用按钮
executeBtn.disabled = true;
// 在 runConcat 结束时启用按钮
executeBtn.disabled = false;
```

## 错误处理

| 场景 | 按钮状态 | 状态消息 |
|------|----------|----------|
| 等待 | 可点击 | 等待操作... |
| 处理中 | 禁用 | 处理中... |
| 成功 | 可点击 | 完成! 已在第 C 列写入 X 行公式 |
| 错误 | 可点击 | 错误: [message] |

## 实施步骤

1. 修改 `concat-taskpane.html` 添加按钮，移除 autofocus
2. 在 `concat-taskpane.css` 添加按钮样式
3. 修改 `concat-taskpane.js` 移除 Enter 监听，添加按钮点击事件
4. 添加按钮禁用状态控制
5. 更新步骤引导文字
6. 测试完整流程
