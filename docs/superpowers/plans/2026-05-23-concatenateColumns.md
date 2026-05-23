# Concatenate Columns Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 用户选中两列后点击功能区按钮，在第二列后插入新列并用指定连接符合并两列数据

**Architecture:** 使用 Office.js Excel JavaScript API 实现。功能区按钮通过 manifest.xml 注册，处理逻辑写入 commands.js。核心流程：获取选中区域 → 弹出对话框获取连接符 → 读取单元格值 → 插入列 → 写入连接结果

**Tech Stack:** Office.js Excel JavaScript API, manifest.xml

---

## File Structure

| 文件 | 作用 |
|------|------|
| `manifest.xml` | 功能区按钮注册（替换原 demo 按钮） |
| `src/commands/commands.js` | 功能区按钮处理逻辑（替换原 demo 代码） |
| `src/taskpane/taskpane.js` | 任务窗格代码（可保留为空或简化） |

---

## Task 1: 清理原 demo 代码

**Files:**
- Modify: `manifest.xml` (删除原 TaskpaneButton，只保留 ConcatenateButton)
- Modify: `src/commands/commands.js` (删除原 action 函数，只保留 concatenateColumns)

- [ ] **Step 1: 在 manifest.xml 中删除原 TaskpaneButton 按钮**

删除整个 `TaskpaneButton` Control 元素，只保留 `ConcatenateButton`

- [ ] **Step 2: 在 manifest.xml 中删除原 ShortStrings 和 LongStrings**

删除 `GetStarted.Title`, `GetStarted.Description`, `TaskpaneButton.Label`, `TaskpaneButton.Tooltip` 等残留资源

- [ ] **Step 3: 在 commands.js 中删除原 action 函数**

删除整个 `action` 函数和相关代码，只保留新的 `concatenateColumns` 函数

- [ ] **Step 4: 验证 manifest.xml 仍然有效**

运行：`npm run validate`
Expected: 验证通过

---

## Task 2: manifest.xml 添加功能区按钮

**Files:**
- Modify: `manifest.xml:59-61` (在 `</ExtensionPoint>` 前插入新按钮)
- Modify: `manifest.xml:74-78` (在 `<bt:ShortStrings>` 添加按钮文字资源)

- [ ] **Step 1: 在 manifest.xml 的 ShortStrings 中添加按钮标签**

在 `TaskpaneButton.Label` 后面添加：
```xml
<bt:String id="ConcatenateButton.Label" DefaultValue="连接列"/>
```

在 `CommandsGroup.Label` 后面添加：
```xml
<bt:String id="ConcatenateButton.Label" DefaultValue="连接列"/>
```

- [ ] **Step 2: 在 manifest.xml 的 LongStrings 中添加按钮 Tooltip**

在 `TaskpaneButton.Tooltip` 后面添加：
```xml
<bt:String id="ConcatenateButton.Tooltip" DefaultValue="选中两列后点击，将数据用连接符合并到新列"/>
```

- [ ] **Step 3: 在 manifest.xml 的 Images 中添加按钮图标**

在 `Icon.80x80` 后面添加（可以复用现有图标）：
```xml
<bt:Image id="ConcatenateIcon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
<bt:Image id="ConcatenateIcon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
```

- [ ] **Step 4: 在 manifest.xml 的 ExtensionPoint 中添加新按钮**

在 `TaskpaneButton` Control 后面添加：
```xml
<Control xsi:type="Button" id="ConcatenateButton">
  <Label resid="ConcatenateButton.Label"/>
  <Supertip>
    <Title resid="ConcatenateButton.Label"/>
    <Description resid="ConcatenateButton.Tooltip"/>
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="ConcatenateIcon.16x16"/>
    <bt:Image size="32" resid="ConcatenateIcon.32x32"/>
  </Icon>
  <Action xsi:type="ExecuteFunction">
    <FunctionName>concatenateColumns"/>
  </Action>
</Control>
```

- [ ] **Step 5: 验证 manifest.xml 语法正确**

运行：`npm run validate`
Expected: 验证通过

---

## Task 3: commands.js 实现连接列核心逻辑

**Files:**
- Modify: `src/commands/commands.js` (添加 concatenateColumns 函数)
- Test: 手动测试

- [ ] **Step 1: 查阅 Excel JS API 文档中 Range 和 Column 相关 API**

查看 PDF 中以下章节确认 API 用法：
- Range 的 `columnCount`, `columnIndex` 属性
- Range 的 `getColumn(index)` 方法
- Worksheet 的 `insert` 方法插入列
- Cell 值的读取和写入方法

- [ ] **Step 2: 在 commands.js 中添加 promptConnector 函数（弹出输入框）**

在 `action` 函数后面添加：
```javascript
function promptConnector() {
  return new Promise((resolve) => {
    const connector = prompt("请输入连接符（默认 _）:", "_");
    resolve(connector === "" ? "_" : connector);
  });
}
```

- [ ] **Step 3: 在 commands.js 中添加 concatenateColumns 函数**

```javascript
async function concatenateColumns(event) {
  try {
    await Excel.run(async (context) => {
      // 1. 获取选中区域
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.load("columnCount");
      await context.sync();

      // 2. 验证至少选中两列
      if (range.columnCount < 2) {
        Office.context.ui.messageBox("请至少选择两列");
        event.completed();
        return;
      }

      // 3. 获取连接符
      const connector = await promptConnector();

      // 4. 获取第一列和第二列的数据
      const firstColumn = range.getColumn(0);
      const secondColumn = range.getColumn(1);
      firstColumn.load("values");
      secondColumn.load("values");
      await context.sync();

      // 5. 在第二列后面插入新列
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const secondColumnIndex = range.columnIndex + 1;
      worksheet.getRange(`:${secondColumnIndex + 1}`).insert(Excel.InsertShiftDirection.shiftRight);
      await context.sync();

      // 6. 写入连接后的数据
      const rowCount = firstColumn.values.length;
      const resultRange = worksheet.getRange(`:${secondColumnIndex}`, `${secondColumnIndex + rowCount - 1}`);
      const resultValues = [];
      for (let i = 0; i < rowCount; i++) {
        const val1 = firstColumn.values[i][0] || "";
        const val2 = secondColumn.values[i][0] || "";
        resultValues.push([[val1 + connector + val2]]);
      }
      resultRange.values = resultValues;
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
  event.completed();
}
```

- [ ] **Step 4: 在 commands.js 中注册新函数**

将 `Office.actions.associate("action", action);` 后面添加：
```javascript
Office.actions.associate("concatenateColumns", concatenateColumns);
```

- [ ] **Step 5: 手动测试**

1. 启动 dev server: `npm run dev-server`
2. 在 Excel 中选中两列（如 A1:B3，包含 "hello"/"world" 等数据）
3. 点击功能区"连接列"按钮
4. 在弹出框输入 `_`（或默认直接点确定）
5. 检查 C 列是否正确插入并写入连接后的数据

Expected: C1="hello_world", C2=第二行数据连接结果, ...

---

## Task 4: 错误处理增强

**Files:**
- Modify: `src/commands/commands.js` (增强错误处理)

- [ ] **Step 1: 添加错误处理覆盖用户取消 prompt 的情况**

修改 `concatenateColumns` 中的 prompt 调用：
```javascript
const connector = await promptConnector();
if (connector === null) {
  event.completed();
  return;
}
```

- [ ] **Step 2: 验证空工作表处理**

测试在空工作表中点击按钮，确认不会崩溃

---

## Task 5: Commit

- [ ] **Step 1: Commit 所有修改**

```bash
git add manifest.xml src/commands/commands.js
git commit -m "feat: add concatenate columns functionality

- Add ribbon button '连接列' in manifest
- Implement concatenateColumns in commands.js
- User selects two columns and clicks button
- Prompt for connector (default _)
- Insert new column after second column
- Write concatenated values as plain text"
```

---

## 自检清单

- [ ] manifest.xml 验证通过
- [ ] 功能区按钮显示并可点击
- [ ] 选中两列后点击能正确插入新列
- [ ] 连接符正确使用
- [ ] 空值处理正确（保留连接符）
- [ ] 用户取消 prompt 时不执行
- [ ] 已 commit