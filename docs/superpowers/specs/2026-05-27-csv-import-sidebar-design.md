# CSV 导入侧边栏设计

**日期**: 2026-05-27
**状态**: 已批准
**需求**: 将 CSV 导入功能从弹窗操作改为侧边栏面板，采用向导式流程，按连接列风格美化界面

---

## 1. 功能行为

- 用户通过 Ribbon 按钮打开独立的侧边栏面板（与连接列面板平级）
- 向导式 3 步骤流程，引导用户依次完成操作
- 导入数据固定从 **A1** 单元格开始写入
- 支持 .csv/.tsv/.txt 文件，分隔符可自定义

---

## 2. 向导步骤

| 步骤 | 名称 | 内容 |
|------|------|------|
| **Step 1** | 选择文件 | 文件选择器，显示已选文件名 |
| **Step 2** | 设置分隔符 | 分隔符输入框（默认 `,`），显示分隔符说明 |
| **Step 3** | 确认导入 | 文件信息摘要、行列数预览、「上一步」「导入」按钮 |

### 2.1 步骤指示器

- 顶部显示 3 步圆点指示器（Step 1 / Step 2 / Step 3）
- 当前步骤高亮（蓝色），已完成步骤显示绿色勾
- 连接线颜色随步骤状态变化

### 2.2 步骤卡片

每个步骤独占一个内容卡片，包裹在 `step-content` 容器内。

### 2.3 状态流转

```
Step 1 → 选择文件 → 自动跳转 Step 2
Step 2 → 设置分隔符 → 自动跳转 Step 3
Step 3 → 点击「导入」→ 显示处理中 → 完成/错误
Step 3 → 点击「上一步」→ 返回 Step 2
Step 2 → 点击「上一步」→ 返回 Step 1
```

---

## 3. UI 样式

### 3.1 视觉风格

与连接列面板保持一致：
- 主色调：`#0078d4`（蓝色）
- 标题：16px，Bold，蓝色，带 SVG 图标
- 引导卡片：浅蓝背景 `#f0f7ff`，左侧 3px 蓝色边框
- 输入框：2px 蓝色边框，6px 圆角，聚焦时蓝色阴影
- 按钮：透明背景，蓝色边框，hover 时蓝色填充白色文字
- 状态消息：圆角卡片，颜色区分成功/错误/加载中

### 3.2 步骤指示器样式

```css
.step-circle { /* 步骤圆点 */ }
.step.active .step-circle { background: #0078d4; }
.step.done .step-circle { background: #28a745; }
.step::after { /* 连接线 */ }
```

### 3.3 按钮样式

- 主要操作：「下一步」「导入」等，使用蓝色填充按钮 `.btn-primary`
- 次要操作：「上一步」「取消」等，使用透明边框按钮 `.btn`
- 按钮宽度根据内容自适应，不强迫撑满整行

---

## 4. 文件结构

```
src/
├── taskpane/
│   ├── csv-import-taskpane.html  # 新建
│   ├── csv-import-taskpane.css  # 新建
│   └── csv-import-taskpane.js   # 新建
└── utils/
    └── csv-utils.js             # 复用（parseCSV）
```

---

## 5. 核心交互

### 5.1 文件选择

- `input[type="file"]` 接受 `.csv,.tsv,.txt`
- 选择文件后自动显示文件名（绿色字体）
- 自动跳转 Step 2

### 5.2 分隔符设置

- 输入框预填充逗号
- 支持任意单字符分隔符
- 跳转到 Step 3 后锁定分隔符输入（防止误改）

### 5.3 确认导入

- 显示摘要：`文件名.csv，共 N 行 × M 列`
- 点击「导入」后禁用按钮，显示处理状态
- 完成后显示成功消息

### 5.4 Excel 数据写入

- 数据从 A1 单元格开始写入
- 使用 `Excel.run()` 批量写入 values
- 进度状态显示处理中行数
- 错误时显示友好错误消息

---

## 6. 错误处理

| 场景 | 处理 |
|------|------|
| 文件为空 | 报错「文件为空」 |
| 数据超过 MAX_ROWS (1050000) | 报错「数据量过大（X 行）」 |
| 文件读取失败 | 报错「文件读取失败」 |
| Excel API 错误 | 报错「错误: X」 |

---

## 7. manifest.xml 改动

将 `ImportCsvButton` 从 `ExecuteFunction` 改为 `ShowTaskpane`，新增独立 `TaskpaneId` 和 `SourceLocation`：

```xml
<Control xsi:type="Button" id="ImportCsvButton">
  ...
  <Action xsi:type="ShowTaskpane">
    <TaskpaneId>CSVImportTaskpaneId</TaskpaneId>
    <SourceLocation resid="CSVImportTaskpane.Url"/>
  </Action>
</Control>

<bt:Url id="CSVImportTaskpane.Url" DefaultValue="https://localhost:3000/csv-import-taskpane.html"/>
```

---

## 8. 测试场景

| 场景 | 预期行为 |
|------|---------|
| 选择 CSV 文件 | 文件名显示，自动进入 Step 2 |
| 输入分隔符 | 自动进入 Step 3，显示摘要 |
| Step 3 点击「上一步」 | 返回 Step 2，分隔符可改 |
| Step 2 点击「上一步」 | 返回 Step 1，可重新选文件 |
| Step 3 点击「导入」 | 数据写入 A1，显示成功状态 |
| 空文件 | 报错「文件为空」 |
| 大数据量（> MAX_ROWS） | 报错「数据量过大」 |