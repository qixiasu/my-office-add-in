# 功能区分组重构 — 设计文档

**日期**: 2026-06-12
**状态**: 已批准

## 概述

将"数据处理"分组中的 5 个按钮拆分为两个分组，确保每个分组按钮数 ≤ 3，使 Excel 功能区以大图标显示所有按钮。

## 问题

`DataProcessingGroup` 包含 5 个按钮（连接列、字段计数、按列拆分、向下选区、向右选区），超出 Office Ribbon 单组大图标显示的上限，部分按钮被降级为小图标（16px）或折叠进溢出菜单。

## 方案

按功能类型拆分为两组：

| 新分组 | 按钮 | 数量 |
|--------|------|------|
| **数据处理** (DataProcessingGroup) | 连接列、字段计数、按列拆分 | 3个 |
| **快速选区** (QuickSelectGroup) **新增** | 向下选区、向右选区 | 2个 |

每个分组 ≤ 3 个按钮，确保所有按钮以大图标（32px）完整显示。

## 涉及文件

| 文件 | 变更 |
|------|------|
| `manifest.xml` | 从 DataProcessingGroup 移除 SelectToEndButton 和 SelectToRightButton；新增 QuickSelectGroup；新增 QuickSelectGroup.Label 字符串资源 |

## 不受影响的文件

- `src/commands/commands.html` — 函数注册不变
- `assets/*` — 图标文件不变
- `webpack.config.js` — 构建配置不变

## manifest.xml 变更详情

### 1. DataProcessingGroup 缩减

移除 `SelectToEndButton` 和 `SelectToRightButton` 两个 `<Control>` 块，保留 `ConcatenateButton`、`CountValuesButton`、`SplitSheetButton`。

### 2. 新增 QuickSelectGroup

在 `DataConversionGroup` 和 `LookupImportGroup` 之间插入：

```xml
<Group id="QuickSelectGroup">
  <Label resid="QuickSelectGroup.Label"/>
  <Icon>
    <bt:Image size="16" resid="ToolsIcon.16x16"/>
    <bt:Image size="32" resid="ToolsIcon.32x32"/>
    <bt:Image size="80" resid="ToolsIcon.80x80"/>
  </Icon>
  <!-- 向下选区 -->
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
  <!-- 向右选区 -->
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

### 3. 新增 ShortStrings

```xml
<bt:String id="QuickSelectGroup.Label" DefaultValue="快速选区"/>
```

其他现有资源（按钮标签、提示、图标）保持不变。

## 验证方法

1. 运行 `npm run validate` 确保 manifest.xml 语法正确
2. 运行 `npm run lint` 确保无 lint 问题
3. 启动 Excel 加载插件，观察分区是否正确、图标是否全部大图标显示
