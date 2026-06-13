# 导出大数据量查询结果优化设计

## 概述

当用户将 SQL 查询结果写入新工作表时，如果数据量很大（几十万行），在预览模式下需要重新执行全量查询。由于 sql.js 是同步执行，此查询会阻塞浏览器主线程，导致 UI 完全冻结若干秒到几十秒。

本设计通过**分页流式查询**替代单次全量查询，将大数据量导出过程拆分为多个小批次查询 + 逐批写入 Excel 的循环，每批之间 yield 给浏览器更新界面，彻底消除 UI 冻结。

## 背景

### 当前导出流程

```
用户执行 SELECT 查询
  → sql-query-taskpane.js 自动追加 LIMIT 200（预览模式）
  → 快速返回 200 行预览结果
用户点击「📝 写入新工作表」
  → showSheetNameDialog() 弹出工作表名称对话框
  → 用户输入名称，点击「确定」
  → dbManager.exec(currentOriginalSQL) ← 同步全量查询，UI 冻结！
  → 全量结果返回后，开始分块写入 Excel（有进度条）
  → 写入完成
```

### 问题根因

`dbManager.exec()` 是同步调用，底层由 sql.js（WebAssembly SQLite）执行。对于几十万行数据的查询，执行时间可达 10-30+ 秒，期间浏览器主线程被完全占用，无法处理 UI 更新、按钮点击等操作。

## 设计

### 核心思路

用 `SELECT * FROM (原始SQL) LIMIT ? OFFSET ?` 循环分批查询，每次只取一小批（5000 行），查完后立即写入 Excel，然后通过 `setTimeout(0)` yield 给浏览器更新 UI。

### 改进后流程

```
用户点击「写入新工作表」
  → 弹出工作表名称对话框
  → 用户输入名称，点击「确定」
  → 按钮变为 ⏳ 正在导出数据...
  → 进度条出现，显示 "已写入 0 行"
  → ── 循环开始（每轮 = 查询 5000 行 + 写入 Excel）──
      → 执行 SELECT * FROM (原始SQL) LIMIT 5000 OFFSET N   ← 毫秒级返回
      → 写入 Excel（当前批）
      → 更新进度 "已写入 M 行"
      → yield 给浏览器（setTimeout(0)）
      → N += 5000，继续下一轮
  → ── 某批返回行数 < 5000，循环结束 ──
  → 恢复按钮，显示 "已将 N 行结果写入新工作表「名称」"
```

### 关键变化

| 方面 | 原来 | 改进后 |
|------|------|--------|
| 查询方式 | 单次 `dbManager.exec(原始SQL)` | 循环 `LIMIT/OFFSET` 分批查询 |
| UI 响应 | 冻结 10-30+ 秒 | 每批之间更新 UI，无冻结 |
| 进度范围 | 仅写入阶段 | 查询 + 写入全程 |
| 进度文本 | "已写入 N/M 行" | "已写入 N 行"（无 COUNT） |
| 数据流 | 全量数据在内存 → 逐批写入 | 逐批查询 → 逐批写入 |

## 技术方案

### 分页查询函数

```javascript
/**
 * 为原始 SQL 构建分页查询语句
 * 将原始 SQL 包裹为子查询以支持 LIMIT/OFFSET
 * @param {string} originalSQL - 原始 SQL
 * @param {number} limit - 每批行数
 * @param {number} offset - 偏移量
 * @returns {string}
 */
function buildPaginationQuery(originalSQL, limit, offset) {
  var sql = originalSQL.replace(/[;\s]*$/, '');
  return 'SELECT * FROM (' + sql + ') LIMIT ' + limit + ' OFFSET ' + offset;
}
```

### 导出核心循环

在 `writeResultToSheet()` 中，预览模式的重新查询部分改为分页循环：

```javascript
function writeResultToSheet() {
  if (!currentQueryResult) return;

  showSheetNameDialog(function (sheetName) {
    if (isPreviewResult && currentOriginalSQL) {
      // C2 方案：分页流式查询 + 逐批写入
      var CHUNK_SIZE = 5000;
      var offset = 0;
      var totalWritten = 0;
      var columns = null;
      var sheetCreated = false;
      var finalSheetName = sheetName;

      // 设置 loading 状态
      setWriteButtonLoading(true);

      function processNextChunk() {
        // yield 给浏览器更新 UI
        setTimeout(function () {
          // 查询当前批
          var paginatedSQL = buildPaginationQuery(currentOriginalSQL, CHUNK_SIZE, offset);
          var result = dbManager.exec(paginatedSQL);

          if (result.type === "error") {
            handleExportError("导出失败: " + result.message);
            return;
          }

          if (result.rowCount === 0) {
            // 没有更多数据了
            handleExportComplete(totalWritten, finalSheetName);
            return;
          }

          if (!columns && result.columns) {
            columns = result.columns;
          }

          var chunkRows = result.rows;
          var batchStartRow = totalWritten;

          // 写入 Excel
          Excel.run(function (context) {
            if (!sheetCreated) {
              // 第一批：创建表 + 写表头 + 写数据
              var sheetCollection = context.workbook.worksheets;
              sheetCollection.load("items/name");
              return context.sync().then(function () {
                finalSheetName = generateUniqueSheetName(sheetCollection, sheetName);
                var newSheet = sheetCollection.add(finalSheetName);
                newSheet.position = 0;
                sheetCreated = true;

                var rangeRows = chunkRows.length + 1;
                var range = newSheet.getRangeByIndexes(0, 0, rangeRows, columns.length);
                var values = [columns];
                for (var r = 0; r < chunkRows.length; r++) {
                  values.push(chunkRows[r]);
                }
                range.values = values;
                return context.sync();
              });
            } else {
              // 后续批次：追加数据
              var sheet = context.workbook.worksheets.getItem(finalSheetName);
              var range = sheet.getRangeByIndexes(batchStartRow + 1, 0, chunkRows.length, columns.length);
              range.values = chunkRows;
              return context.sync();
            }
          }).then(function () {
            totalWritten += chunkRows.length;
            offset += CHUNK_SIZE;
            updateWriteProgress(totalWritten);
            processNextChunk(); // 继续下一批
          }).catch(function (error) {
            handleExportError("写入失败: " + (error.message || error));
          });
        }, 0);
      }

      processNextChunk();
    } else {
      // 非预览模式：数据已在内存中，保持原有逐批写入逻辑
      // ...
    }
  });
}
```

### 辅助函数

```javascript
/**
 * 设置导出按钮的 loading 状态
 */
function setWriteButtonLoading(isLoading) {
  var btn = document.getElementById("writeSheetBtn");
  if (isLoading) {
    btn.disabled = true;
    btn.textContent = "⏳ 正在导出数据...";
    btn.classList.add("sql-button-loading");
    document.getElementById("writeProgress").style.display = "flex";
    document.getElementById("writeProgressText").textContent = "正在导出数据...";
  } else {
    btn.disabled = false;
    btn.textContent = "📝 写入新工作表";
    btn.classList.remove("sql-button-loading");
    document.getElementById("writeProgress").style.display = "none";
  }
}

/**
 * 更新写入进度
 * @param {number} written - 已写入行数
 */
function updateWriteProgress(written) {
  var fillEl = document.getElementById("writeProgressFill");
  var textEl = document.getElementById("writeProgressText");
  textEl.textContent = "已写入 " + written + " 行";
}
```

## UI 改动

### 改动清单

| 元素 | 改动 |
|------|------|
| `#writeSheetBtn` | 点击「确定」后立即显示 `⏳ 正在导出数据...` |
| `#writeProgress` | 复用现有进度条容器，无改动 |
| `#writeProgressFill` | 不再显示百分比，仅作为装饰性进度指示 |
| `#writeProgressText` | 文本改为 "已写入 N 行"（仅绝对行数） |
| `#queryStatus` | 最终状态提示保持现有风格 |

### 不需改动的元素

- 工作表名称对话框 (`showSheetNameDialog`)
- 结果表格渲染
- 按钮样式和 CSS
- HTML 模板

## 错误处理

| 场景 | 处理方式 |
|------|----------|
| 某批查询返回错误 | 停止流程，恢复按钮，显示 "导出失败: [错误]" |
| 某批 Excel 写入失败 | 停止流程，恢复按钮，已写入数据保留，显示 "写入失败: [错误]" |
| 某批返回行数 < 每批大小 | 正常终止，显示完成消息 |
| 非 SELECT 查询进入导出 | 已有 `if (!currentQueryResult) return;` 守卫 |

## SQL 兼容性

`SELECT * FROM (原始SQL) LIMIT N OFFSET M` 在 SQLite 中已验证兼容：

- `SELECT ... WHERE ...`
- `SELECT ... JOIN ... ON ...`
- `SELECT ... GROUP BY ... HAVING ...`
- `SELECT ... ORDER BY ...`
- `SELECT ... UNION ...`
- `SELECT ... LIMIT ...`（用户显式 LIMIT 时走非预览路径，不受影响）

唯一需要处理的是原始 SQL 末尾的 `;` 和空白字符，已在 `buildPaginationQuery` 中 strip 处理。

## 性能分析

对于 N 行数据、每批 K 行的场景：

- **总查询复杂度：** 与单次全量查询相同（SQLite 仍会扫描全表）
- **总内存使用：** 从 O(N) 降为 O(K)（每次只保留 K 行在内存中）
- **UI 响应性：** 从完全冻结提升为可交互（每次 sql.js 执行仅为 K 行的查询时间）
- **总耗时：** 与原来基本持平（查询 + 写入总工作量不变）
