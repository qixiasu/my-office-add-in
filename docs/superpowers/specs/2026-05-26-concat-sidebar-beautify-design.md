# 连接列侧边栏 - 友好引导风格设计

## 概述

对连接列侧边栏进行视觉美化，采用「友好引导」设计风格：蓝色强调色 + 信息卡片布局，适合新手用户。

## 设计方向

**选择：友好引导**
- 蓝色强调色 (#0078d4)
- 信息卡片布局
- 步骤引导，帮助新手操作

## 视觉设计

### 侧边栏 UI 结构

```
┌─────────────────────────────┐
│ 🔗 连接列工具                │
├─────────────────────────────┤
│ ┌─────────────────────────┐ │
│ │ 1. 选中要连接的两列      │ │
│ │ 2. 输入连接符，按 Enter  │ │
│ └─────────────────────────┘ │
│                             │
│ 连接符 [________]           │
│                             │
│ 状态: 等待操作...           │
└─────────────────────────────┘
```

### 视觉细节

| 元素 | 样式 |
|------|------|
| 标题 | 蓝色 #0078d4，20px，链接图标 |
| 步骤引导卡片 | 浅蓝背景 #f0f7ff，左侧蓝色边框 3px |
| 输入框 | 蓝色边框 2px #0078d4，圆角 6px，焦点时蓝色光晕 |
| 成功状态 | 绿色背景 #d4edda，绿色文字 #155724 |
| 错误状态 | 红色背景，红色文字 |

### 状态设计

**等待状态：**
- 背景：#f5f5f5
- 文字：#999

**成功状态：**
- 背景：#d4edda
- 边框：#c3e6cb
- 文字：#155724

**错误状态：**
- 背景：语义化红色
- 文字：语义化红色

## 技术实现

### 文件变更

**修改文件：**
- `src/taskpane/concat-taskpane.html` — 调整 HTML 结构
- `src/taskpane/concat-taskpane.css` — 添加样式

### CSS 样式

```css
/* 标题 */
.concat-title {
  font-size: 16px;
  font-weight: 700;
  color: #0078d4;
  display: flex;
  align-items: center;
  gap: 8px;
}

/* 步骤引导卡片 */
.guide-card {
  background: #f0f7ff;
  border-left: 3px solid #0078d4;
  padding: 12px;
  border-radius: 4px;
  margin-bottom: 16px;
}

/* 输入框 */
.concat-input {
  width: 100%;
  padding: 10px 12px;
  border: 2px solid #0078d4;
  border-radius: 6px;
  font-size: 15px;
  outline: none;
  transition: box-shadow 0.2s;
  box-sizing: border-box;
}

.concat-input:focus {
  box-shadow: 0 0 0 3px rgba(0,120,212,0.15);
}

/* 状态消息 */
.status-success {
  background: #d4edda;
  border: 1px solid #c3e6cb;
  color: #155724;
}

.status-error {
  background: #f8d7da;
  border: 1px solid #f5c6cb;
  color: #721c24;
}
```

## 错误处理

| 场景 | 显示 |
|------|------|
| 未选中两列 | 错误：请至少选择两列 |
| 无数据 | 错误：没有数据 |
| 数据量过大 | 错误：数据量过大（X行），单次最多支持 1050000 行 |
| 执行失败 | 错误：[error.message] |
| 成功 | 完成！已在第 C 列写入 X 行公式 |

## 实施步骤

1. 修改 `concat-taskpane.html` 添加步骤引导卡片和图标
2. 在 `concat-taskpane.css` 添加新样式
3. 更新输入框样式（聚焦光晕效果）
4. 更新状态显示逻辑（语义化颜色）
5. 测试完整流程
