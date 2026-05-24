# 修复 CSV 中文乱码 — 设计文档

**日期**: 2026-05-24
**状态**: 待实施

## 问题

Excel 打开 UTF-8 编码（无 BOM）的 CSV 文件时，使用系统默认编码（中文 Windows 下为 GBK/CP936）解读字节，导致中文显示为乱码。

## 方案

使用 `iconv-lite` 纯 JS 库对已打开工作表中的乱码文本做编码逆转：乱码文本 → GBK 编码为字节 → UTF-8 解码为正确中文。

## 架构

在现有项目结构上新增：

```
src/
  utils/
    encoding-utils.js      # 新增：编码逆转工具函数
  commands/
    commands.html          # 修改：注册 fixGarbledText 命令
manifest.xml               # 修改：新增"修复乱码" ribbon 按钮
package.json               # 修改：新增 iconv-lite 依赖
```

## 功能规格

### 触发方式

"我的工具"选项卡新增 **"修复乱码"** 按钮（ExecuteFunction），点击后直接执行，无需弹窗。

### 处理逻辑

1. 获取当前活动工作表的 `getUsedRange()`
2. 读取所有单元格的值到二维数组
3. 对每个文本值尝试编码逆转：
   - `iconv.encode(value, 'gbk')` → 得到字节
   - `iconv.decode(bytes, 'utf-8')` → 得到修复后文本
4. 对比修复前后的文本：
   - **不同** → 写入修复后的值
   - **相同** → 保持原值（说明不需要修复，如英文/数字/已正确的文本）
5. 一次性将修复后的数组写回工作表
6. 弹出消息框告知结果（"修复完成，共处理 N 个单元格，修复 M 个"）

### 边界情况

- **大数据量**：以 1050000 行为上限（与现有 `MAX_ROWS` 一致），超出时提示用户
- **空工作表**：提示"没有数据"
- **公式单元格**：`range.values` 返回的是计算结果值，不影响公式；但如果是公式产生的文本，修复后会被替换为静态文本
- **非文本值**：数字、布尔值不做编码逆转处理

### 性能考量

- 一次性读写全部数据（先读 → 内存处理 → 一次写回），避免逐个单元格的 API 调用
- `iconv-lite` 是纯 JS 实现，GBK 编码表约 200KB

## 文件变更清单

| 文件 | 操作 | 说明 |
|------|------|------|
| `package.json` | 修改 | 新增 `iconv-lite` 依赖 |
| `src/utils/encoding-utils.js` | 新增 | `fixGarbledText(value)` 工具函数 |
| `src/commands/commands.html` | 修改 | 注册 `fixGarbledText` 命令，实现处理逻辑 |
| `manifest.xml` | 修改 | 新增"修复乱码"按钮配置 |

## 测试

- 单元测试：`encoding-utils.js` 的编码逆转函数（输入乱码字符串，验证输出正确中文）
- 手动测试：准备 UTF-8 无 BOM 中文 CSV 文件，用 Excel 打开后点击按钮验证修复效果
