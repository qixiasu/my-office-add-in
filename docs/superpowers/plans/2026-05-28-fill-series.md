# 填充序列工具 - 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox syntax for tracking.

**目标：** 在 Excel 侧边栏面板中实现填充序列功能，用户输入起始值和步长，对选中区域按线性序列填充数值，支持整数和小数。

**架构：** 新增独立面板 fill-series-taskpane（HTML/CSS/JS），新增 ribbon 按钮 FillSeriesButton，执行时读取选中区域，生成序列数组，写入 Excel 单元格的 2D 值。

**技术栈：** 纯 JavaScript（与现有 taskpane 一致），Office JavaScript API，Excel.run 模式。

---

## 文件结构

| 操作 | 文件 |
|------|------|
| Create | src/taskpane/fill-series-taskpane.html |
| Create | src/taskpane/fill-series-taskpane.js |
| Create | src/taskpane/fill-series-taskpane.css |
| Modify | manifest.xml（添加 FillSeriesButton） |

---

## 任务分解

### 任务 1：创建 fill-series-taskpane.html

**Files:**
- Create: src/taskpane/fill-series-taskpane.html

- [ ] **Step 1: 创建 HTML 文件**

参照 concat-taskpane.html 的结构，创建 fill-series-taskpane.html：
- 标题：填充序列工具
- 说明卡片：选中区域按起始值和步长填充
- 两个输入框（起始值、步长）水平排列
- 执行填充按钮
- 状态消息

- [ ] **Step 2: 提交**

git add src/taskpane/fill-series-taskpane.html
git commit -m "feat: add fill-series-taskpane HTML structure"

---

### 任务 2：创建 fill-series-taskpane.css

**Files:**
- Create: src/taskpane/fill-series-taskpane.css

- [ ] **Step 1: 创建样式文件**

创建 fill-series-taskpane.css：
- .fill-series-container 容器样式
- .fill-series-title 标题样式
- .guide-card 说明卡片样式
- .form-row 水平布局容器（display: flex; gap: 12px）
- .form-group 每个输入框的容器
- .fill-series-input 输入框样式（宽度100%，padding）
- .fill-series-button 执行按钮样式

- [ ] **Step 2: 提交**

git add src/taskpane/fill-series-taskpane.css
git commit -m "feat: add fill-series-taskpane styles"

---

### 任务 3：创建 fill-series-taskpane.js

**Files:**
- Create: src/taskpane/fill-series-taskpane.js

- [ ] **Step 1: 编写核心逻辑**

runFillSeries 函数核心逻辑：
- 读取 startValue 和 stepValue 输入
- 验证非空和数值格式
- Excel.run 获取选中区域
- 计算 rowCount, columnCount, total
- 生成 1D 序列数组：[start, start+step, ...]
- 转换为 2D 数组（列优先）
- range.values = result
- 显示成功状态

- [ ] **Step 2: 提交**

git add src/taskpane/fill-series-taskpane.js
git commit -m "feat: add fill-series-taskpane JavaScript logic"

---

### 任务 4：修改 manifest.xml 添加 FillSeriesButton

**Files:**
- Modify: manifest.xml

- [ ] **Step 1: 在 manifest.xml 中添加 FillSeriesButton**

在 HomeTab 的 CommandsGroup 中添加新的 Control：
- id: FillSeriesButton
- Label: 填充序列
- Action: ExecuteFunction，调用 showFillSeriesPanel

添加 Resources 字符串：
- FillSeriesLabel: 填充序列
- FillSeriesTooltip: 打开填充序列工具面板

添加 Commands 函数 showFillSeriesPanel：
- Office.addin.showAsTaskpane()
- event.completed()

- [ ] **Step 2: 提交**

git add manifest.xml
git commit -m "feat: add FillSeriesButton to ribbon"

---

### 任务 5：验证构建

- [ ] **Step 1: 运行开发服务器验证：**

npm run dev-server

预期：服务启动成功，fill-series-taskpane.html 可正常加载

---

## 验证计划

完成所有任务后，在 Excel 桌面版中测试：

1. 点击填充序列ribbon 按钮，打开侧边栏面板
2. 在 Excel 中选中一个区域（如 A1:C3）
3. 输入起始值（如 1）和步长（如 10）
4. 点击执行填充按钮
5. 验证选中区域填充了正确的序列值（按列优先）

预期结果：
- A1=1, A2=11, A3=21
- B1=31, B2=41, B3=51
- C1=61, C2=71, C3=81
