# VLOOKUP 性能优化设计

## 背景

当前的增强查找功能在数据行数较多时查找速度缓慢。通过代码分析，发现以下性能瓶颈：

1. **调试日志过多**：`staticLookup` 函数内每条记录查找都打印多条 console.log，严重拖慢执行
2. **哈希索引重复构建**：每次调用 `staticLookup` 时，精确匹配模式都会重新遍历整个查找表构建哈希索引
3. **近似匹配线性扫描**：非精确匹配时每条记录线性扫描整个查找表

## 优化目标

- 减少调试日志对执行性能的影响
- 避免哈希索引重复构建
- 提升大数据量场景下的查找速度

## 优化方案

### 优化一：移除调试日志

从 `staticLookup` 函数中删除所有循环内的 `console.log` 调试语句。

**涉及位置**：`src/utils/vlookup-utils.js` 第 91-128 行附近

### 优化二：哈希索引缓存

在 `staticLookup` 中增加可选的 `indexCache` 参数，将哈希索引的构建与查找分离。

**接口变更**：

```javascript
// 新增：独立索引构建函数
function buildLookupIndex(lookupTable, matchColIndex) { ... }

// staticLookup 签名增加可选参数
function staticLookup(
  lookupValues,
  lookupTable,
  matchColIndex,
  returnColIndices,
  matchMode,
  defaultValue,
  indexCache  // 新增：可选的预构建索引
) { ... }
```

**调用方式变更**：

```javascript
// taskpane.js 批次查找流程
// 1. 首次查找前构建索引
var indexCache = buildLookupIndex(g_lookupTableData, config.matchColIndex);

// 2. 后续批次传递索引
var batchResults = staticLookup(
  batchLookupValues,
  g_lookupTableData,
  config.matchColIndex,
  config.returnColIndices,
  config.matchMode,
  config.defaultValue,
  indexCache  // 传入缓存的索引
);
```

### 优化三：近似匹配二分查找（可选增强）

将近似匹配的线性扫描 O(m × n) 优化为二分查找 O(m × log n)。

**前提条件**：查找表必须按匹配列升序排列（与 Excel VLOOKUP 要求一致）

## 改动范围

| 文件 | 改动内容 |
|------|----------|
| `src/utils/vlookup-utils.js` | 1. 删除 staticLookup 内所有 console.log<br>2. 新增 `buildLookupIndex` 函数<br>3. staticLookup 增加 `indexCache` 参数 |
| `src/taskpane/vlookup-taskpane.js` | 1. 首次调用前构建索引并缓存<br>2. 后续所有 staticLookup 调用传入索引缓存 |

## 性能预估

| 场景 | 优化前 | 优化后 |
|------|--------|--------|
| 精确匹配，100k 行查找表，单次 10k 查找值 | 索引重复构建 10 次 | 索引仅构建 1 次 |
| 精确匹配，100k 行查找表，10 批次 | 索引重复构建 10 次 | 索引仅构建 1 次 |

## 测试验证

1. 精确匹配：验证索引缓存后查找结果与原来一致
2. 近似匹配：验证二分查找结果与线性扫描一致
3. 性能对比：对比优化前后执行时间
