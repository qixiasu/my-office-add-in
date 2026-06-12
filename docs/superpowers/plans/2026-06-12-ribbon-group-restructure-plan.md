# 功能区分组重构 — 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 将 DataProcessingGroup（5 个按钮）拆分为 DataProcessingGroup（3 个按钮）+ QuickSelectGroup（2 个按钮），确保所有按钮以大图标显示。

**Architecture:** 只修改 `manifest.xml`。将 SelectToEndButton 和 SelectToRightButton 从 DataProcessingGroup 移动到新建的 QuickSelectGroup 分组，其他所有资源引用保持不变。

**Tech Stack:** Office Add-in manifest XML

---

### Task 1: 从 DataProcessingGroup 移除两个选区按钮

**Files:**
- Modify: `manifest.xml`

- [ ] **Step 1: 移除 SelectToEndButton 控件块**

找到 DataProcessingGroup 中的 `SelectToEndButton` 控件块（从 `<Control xsi:type="Button" id="SelectToEndButton">` 到其对应的 `</Control>`），将其**完整删除**。

当前 DataProcessingGroup 中 SelectToEndButton 位于 SplitSheetButton 和 SelectToRightButton 之间。删除后 SelectToRightButton 会成为最后一个控件。

- [ ] **Step 2: 移除 SelectToRightButton 控件块**

找到 DataProcessingGroup 中剩余的 `SelectToRightButton` 控件块（从 `<Control xsi:type="Button" id="SelectToRightButton">` 到其对应的 `</Control>`），将其**完整删除**。

- [ ] **Step 3: 验证 manifest.xml**

运行：

```bash
npm run validate
```

期望：通过。

- [ ] **Step 4: 提交**

```bash
git add manifest.xml
git commit -m "refactor: remove SelectToEndButton and SelectToRightButton from DataProcessingGroup"
```

---

### Task 2: 新增 QuickSelectGroup 分组和资源引用

**Files:**
- Modify: `manifest.xml`

- [ ] **Step 1: 在 DataConversionGroup 和 LookupImportGroup 之间插入 QuickSelectGroup**

找到 DataConversionGroup 的 `</Group>` 闭合标签之后、LookupImportGroup 的 `<Group id="LookupImportGroup">` 开始标签之前，插入以下新分组定义：

```xml
              <Group id="QuickSelectGroup">
                <Label resid="QuickSelectGroup.Label"/>
                <Icon>
                  <bt:Image size="16" resid="ToolsIcon.16x16"/>
                  <bt:Image size="32" resid="ToolsIcon.32x32"/>
                  <bt:Image size="80" resid="ToolsIcon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="SelectToEndButton">
                  <Label resid="SelectToEndButton.Label"/>
                  <Supertip>
                    <Title resid="SelectToEndButton.Label"/>
                    <Description resid="SelectToEndButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="SelectToEndIcon.16x16"/>
                    <bt:Image size="32" resid="SelectToEndIcon.32x32"/>
                    <bt:Image size="80" resid="SelectToEndIcon.80x80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>selectToEnd</FunctionName>
                  </Action>
                </Control>
                <Control xsi:type="Button" id="SelectToRightButton">
                  <Label resid="SelectToRightButton.Label"/>
                  <Supertip>
                    <Title resid="SelectToRightButton.Label"/>
                    <Description resid="SelectToRightButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="SelectToRightIcon.16x16"/>
                    <bt:Image size="32" resid="SelectToRightIcon.32x32"/>
                    <bt:Image size="80" resid="SelectToRightIcon.80x80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>selectToRight</FunctionName>
                  </Action>
                </Control>
              </Group>
```

> 注意缩进：使用 14 个空格（与相邻的 DataConversionGroup 和 LookupImportGroup 一致）。

- [ ] **Step 2: 在 ShortStrings 中新增分组标签**

找到 `<bt:ShortStrings>` 块末尾，在最后一个 `<bt:String>` 之后、`</bt:ShortStrings>` 之前添加：

```xml
        <bt:String id="QuickSelectGroup.Label" DefaultValue="快速选区"/>
```

- [ ] **Step 3: 验证 manifest.xml**

运行：

```bash
npm run validate
```

期望：通过。

- [ ] **Step 4: 提交**

```bash
git add manifest.xml
git commit -m "feat: add QuickSelectGroup with SelectToEnd and SelectToRight buttons"
```

---

### Task 3: 运行测试并验证

**Files:**
- 无需修改代码

- [ ] **Step 1: 运行 lint**

```bash
npm run lint
```

期望：无新错误。

- [ ] **Step 2: 运行测试**

```bash
npm test
```

期望：115 passed, 全部通过（本次变更不涉及代码逻辑，测试仍应全部通过）。

- [ ] **Step 3: 验证整体 manifest**

再次运行验证确保最终 manifest.xml 无问题：

```bash
npm run validate
```

期望：通过。
