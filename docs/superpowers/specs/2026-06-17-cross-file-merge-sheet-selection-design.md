# 跨文件合并：支持手动选择每个文件的 Sheet

## 背景

当前跨文件合并功能使用每个 Excel 文件的激活 sheet（`SheetNames[0]`），用户无法指定要合并哪个 sheet。当 Excel 文件包含多个 sheet 时，无法选择要合并的目标 sheet。

## 目标

在文件选中后、合并执行前，允许用户为每个文件单独选择要合并的 sheet。默认选中第一个 sheet，用户可按需更改。

## 用户交互流程

1. 用户点击"选择 Excel 文件"，选中多个文件
2. 文件列表显示每个文件，旁边显示该文件所有 sheet 的下拉选择框（默认选第一个）
3. 用户可调整每个文件对应的 sheet
4. 用户设置表头行号，点击"执行合并"
5. 合并结果新增"SrcSheet"列，记录每行数据来源于源文件的哪个 sheet（结合之前新增的 Sheet 名列功能）

## 数据流变化

```
旧: selectedFiles = [{ file, name }]
新: selectedFiles = [{ file, name, sheetName }]
```

## 详细设计

### 1. 新增 getSheetNames 函数

**文件**: `src/utils/cross-file-merge-utils.js`

```javascript
/**
 * 获取 Excel 文件的所有 sheet 名称
 * @param {File} file - Browser File 对象
 * @returns {Promise<string[]>} sheet 名称数组
 */
function getSheetNames(file) {
  return new Promise(function (resolve, reject) {
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, { type: "array" });
        resolve(workbook.SheetNames);
      } catch (err) {
        reject(new Error("读取文件 sheet 失败: " + err.message));
      }
    };
    reader.onerror = function () {
      reject(new Error("文件读取失败"));
    };
    reader.readAsArrayBuffer(file);
  });
}
```

### 2. 修改 parseExcelFile 函数

**文件**: `src/utils/cross-file-merge-utils.js`

`parseExcelFile(file)` 增加可选参数 `sheetName`，如果不传则使用第一个 sheet：

```javascript
/**
 * 使用 SheetJS 解析 Excel 文件的指定 sheet
 * @param {File} file - Browser File 对象
 * @param {string} [sheetName] - 可选，指定 sheet 名称，不传则使用第一个
 * @returns {Promise<{data: Array, sheetName: string}>}
 */
function parseExcelFile(file, sheetName) {
  return new Promise(function (resolve, reject) {
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, { type: "array" });
        var targetSheetName = sheetName || workbook.SheetNames[0];
        var sheet = workbook.Sheets[targetSheetName];
        if (!sheet) {
          return reject(new Error("工作表 '" + targetSheetName + "' 不存在或为空"));
        }
        var jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        resolve({ data: jsonData, sheetName: targetSheetName });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = function () {
      reject(new Error("文件读取失败"));
    };
    reader.readAsArrayBuffer(file);
  });
}
```

### 3. 修改文件列表渲染（cross-file-merge-taskpane.js）

**文件**: `src/taskpane/cross-file-merge-taskpane.js`

#### 3.1 修改 selectedFiles 数据结构

`selectedFiles` 每项从 `{ file, name }` 变为 `{ file, name, sheetName, sheetNames, loadingSheets }`：

```javascript
selectedFiles = Array.from(files).map(function(file) {
  return {
    file: file,
    name: file.name,
    sheetName: null,        // 当前选中的 sheet（默认在加载后设置）
    sheetNames: [],         // 该文件所有 sheet 名称列表
    loadingSheets: true     // 是否正在加载 sheet 列表
  };
});
```

#### 3.2 选中文件后异步获取 sheet 列表

```javascript
function onFileSelected(e) {
  var files = e.target.files;
  if (!files || files.length === 0) return;

  selectedFiles = Array.from(files).map(function(file) {
    return {
      file: file,
      name: file.name,
      sheetName: null,
      sheetNames: [],
      loadingSheets: true
    };
  });

  renderFileList();
  updateExecuteButton();

  // 异步获取每个文件的 sheet 列表
  selectedFiles.forEach(function(item, index) {
    crossFileMergeUtils.getSheetNames(item.file)
      .then(function(names) {
        item.sheetNames = names;
        item.sheetName = names[0]; // 默认选第一个
        item.loadingSheets = false;
        renderFileList(); // 重新渲染以显示下拉框
      })
      .catch(function(err) {
        item.sheetNames = [];
        item.sheetName = null;
        item.loadingSheets = false;
        item.loadError = err.message;
        renderFileList();
      });
  });
}
```

#### 3.3 修改 renderFileList 渲染下拉框

每个文件项渲染 sheet 选择下拉框：

```javascript
function renderFileList() {
  var listEl = document.getElementById("fileList");

  if (selectedFiles.length === 0) {
    listEl.innerHTML = '<div class="cfm-empty">未选择文件</div>';
    return;
  }

  var html = "";
  selectedFiles.forEach(function(item, index) {
    var sheetSelect = "";
    if (item.loadingSheets) {
      sheetSelect = '<span class="cfm-sheet-loading">加载中...</span>';
    } else if (item.loadError) {
      sheetSelect = '<span class="cfm-sheet-error" title="' + escapeHtml(item.loadError) + '">加载失败</span>';
    } else if (item.sheetNames.length === 0) {
      sheetSelect = '<span class="cfm-sheet-error">无 sheet</span>';
    } else {
      var options = item.sheetNames.map(function(name) {
        var selected = name === item.sheetName ? ' selected' : '';
        return '<option value="' + escapeHtml(name) + '"' + selected + '>' + escapeHtml(name) + '</option>';
      }).join('');
      sheetSelect = '<select class="cfm-sheet-select" data-index="' + index + '">' + options + '</select>';
    }

    html +=
      '<div class="cfm-file-item" data-index="' + index + '">' +
        '<span class="cfm-file-name" title="' + escapeHtml(item.name) + '">' + escapeHtml(item.name) + '</span>' +
        '<span class="cfm-file-sheet">' + sheetSelect + '</span>' +
        '<button class="cfm-file-remove" data-index="' + index + '">移除</button>' +
      '</div>';
  });

  listEl.innerHTML = html;

  // 绑定下拉框 change 事件
  var sheetSelects = listEl.querySelectorAll(".cfm-sheet-select");
  sheetSelects.forEach(function(select) {
    select.onchange = function() {
      var idx = parseInt(select.dataset.index, 10);
      selectedFiles[idx].sheetName = select.value;
    };
  });

  // 绑定移除按钮
  var removeButtons = listEl.querySelectorAll(".cfm-file-remove");
  removeButtons.forEach(function(btn) {
    btn.onclick = function() {
      var idx = parseInt(btn.dataset.index, 10);
      selectedFiles.splice(idx, 1);
      renderFileList();
      updateExecuteButton();
    };
  });
}
```

#### 3.4 修改 executeMerge 使用用户选择的 sheet

```javascript
var parsePromises = selectedFiles.map(function(item) {
  return crossFileMergeUtils.parseExcelFile(item.file, item.sheetName)
    .then(function(result) {
      fileDataList.push({
        data: result.data,
        name: item.name,
        sheetName: result.sheetName, // 使用用户选择的 sheet 名称
        columnCount: result.data.length > 0 ? result.data[0].length : 0
      });
      setStatus("已解析 " + fileDataList.length + "/" + selectedFiles.length + " 个文件...", "loading");
    })
    .catch(function(err) {
      throw new Error("解析文件 '" + item.name + "' 的 Sheet '" + item.sheetName + "' 失败: " + err.message);
    });
});
```

### 4. CSS 样式

**文件**: `src/taskpane/cross-file-merge-taskpane.css`

```css
.cfm-file-item {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 6px 8px;
  border-bottom: 1px solid #f0f0f0;
}

.cfm-file-name {
  flex: 1;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.cfm-file-sheet {
  flex-shrink: 0;
}

.cfm-sheet-select {
  padding: 2px 6px;
  border: 1px solid #d0d0d0;
  border-radius: 4px;
  font-size: 12px;
  min-width: 80px;
  max-width: 150px;
}

.cfm-sheet-loading {
  font-size: 12px;
  color: #888;
}

.cfm-sheet-error {
  font-size: 12px;
  color: #d83b01;
}
```

### 5. 单元测试更新

**文件**: `src/utils/cross-file-merge-utils.test.js`

- 新增 `getSheetNames` 测试用例
- `parseExcelFile` 增加 `sheetName` 参数的测试用例
- `executeMerge` 相关测试需要更新（如果测试中有 mock 文件选中逻辑）

## 边界情况

| 情况 | 处理方式 |
|------|---------|
| 文件 sheet 加载中 | 显示"加载中..."，执行按钮仍可用（依赖用户等待） |
| 文件 sheet 加载失败 | 显示"加载失败"，执行时该文件报错 |
| 文件无任何 sheet | 显示"无 sheet"，执行时该文件报错 |
| 文件仅有一个 sheet | 下拉框只有一个选项，无需用户选择 |
| 用户移除了部分文件 | 正常处理，数组 splice |
| 合并时某个 sheet 不存在 | 报错指向具体文件和 sheet 名 |

## 兼容性

- 不影响现有合并逻辑，仅扩展 sheet 选择能力
- 旧行为（无 sheet 选择）已被新流程覆盖
