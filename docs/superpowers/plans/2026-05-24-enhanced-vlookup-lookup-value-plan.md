# Enhanced VLOOKUP - lookup_value Parameter UI Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add lookup_value input field to vlookup dialog with refresh-target toggle so users can re-select either the lookup value source or the lookup table.

**Architecture:** Add a read-only "查找值区域" input to the dialog with a toggle (刷新查找值 / 刷新查找表) that controls which field the refresh-selection button updates. commands.html message handler is extended to distinguish between the two refresh targets.

**Tech Stack:** Vanilla JS dialog, Office.js, plain HTML/CSS

---

## File Map

```
src/commands/vlookup-dialog.html   — UI changes: toggle + lookup value input
src/commands/commands.html         — Message handler: differentiate refresh target
```

---

## Task 1: Add toggle and lookup value input to dialog UI

**Files:**
- Modify: `src/commands/vlookup-dialog.html`

- [ ] **Step 1: Add CSS for toggle tabs**

Find the `<style>` block around line 8. Add after existing `.section-title`:

```css
        .toggle-tabs { margin-bottom: 16px; }
        .toggle-tabs button {
            padding: 6px 16px;
            cursor: pointer;
            font-size: 13px;
            border: 1px solid #aaa;
            background: #f5f5f5;
            margin-right: 4px;
        }
        .toggle-tabs button.active {
            background: #0078d4;
            color: #fff;
            border-color: #0078d4;
        }
        .lookup-value-row { margin-bottom: 14px; }
        .lookup-value-row input { background: #f0f0f0; }
```

- [ ] **Step 2: Add toggle HTML after `<h3>增强查找</h3>`**

Insert at line ~39, after the h3 closing tag:

```html
    <div class="toggle-tabs">
        <button id="tabLookupValue" class="active">刷新查找值</button>
        <button id="tabLookupTable">刷新查找表</button>
    </div>

    <div class="row lookup-value-row">
        <label for="lookupValue">查找值区域：</label>
        <input type="text" id="lookupValue" readonly style="background:#f0f0f0;" />
        <div class="info">自动取自当前选区第一列</div>
    </div>
```

- [ ] **Step 3: Add refresh-target state variable**

Find `var MAX_ROWS = 1050000;` around line 96. Add after it:

```javascript
        var refreshTarget = "lookupValue"; // "lookupValue" or "lookupTable"
```

- [ ] **Step 4: Add tab switching logic**

Find `Office.onReady(function () {` around line 419. Add before it (or inside onReady before existing handlers):

```javascript
        function setRefreshTarget(target) {
            refreshTarget = target;
            document.getElementById("tabLookupValue").classList.toggle("active", target === "lookupValue");
            document.getElementById("tabLookupTable").classList.toggle("active", target === "lookupTable");
        }

        document.getElementById("tabLookupValue").onclick = function() { setRefreshTarget("lookupValue"); };
        document.getElementById("tabLookupTable").onclick = function() { setRefreshTarget("lookupTable"); };
```

- [ ] **Step 5: Update refreshSelection to send target + update lookupValue input display**

Find `function refreshSelection()` around line 370. Replace the entire function:

```javascript
        function refreshSelection() {
            document.getElementById("status").textContent = "正在更新选区...";
            document.getElementById("status").style.color = "green";

            var cbs = document.querySelectorAll("#returnCols input[type='checkbox']");
            var checkedCols = [];
            for (var i = 0; i < cbs.length; i++) {
                if (cbs[i].checked) { checkedCols.push(cbs[i].value); }
            }

            Office.context.ui.messageParent(JSON.stringify({
                type: "refreshSelection",
                target: refreshTarget,
                state: {
                    headerRow: document.getElementById("headerRow").value,
                    matchCol: document.getElementById("matchCol").value,
                    returnCols: checkedCols.join(","),
                    matchMode: document.querySelector("input[name='matchMode']:checked").value,
                    outputType: document.querySelector("input[name='outputType']:checked").value,
                    lookupValue: document.getElementById("lookupValue").value
                }
            }));
        }
```

- [ ] **Step 6: Update restoreState to handle lookupValue and refreshTarget URL params**

Find `function restoreState()` around line 395. Add after the existing param checks:

```javascript
            if (params.get("refreshTarget")) {
                setRefreshTarget(params.get("refreshTarget"));
            }
            if (params.get("lookupValue")) {
                document.getElementById("lookupValue").value = params.get("lookupValue");
            }
```

- [ ] **Step 7: Update onReady to auto-fill lookupValue from preselected range**

Find the preselected fill block around line 421:

```javascript
            var preselected = getPreselectedRange();
            if (preselected) {
                document.getElementById("lookupRange").value = preselected;
                // Extract first column from preselected range as lookup value
                document.getElementById("lookupValue").value = extractFirstColAddress(preselected);
                refreshColumns();
                restoreState();
            }
```

Add helper function before `Office.onReady`:

```javascript
        function extractFirstColAddress(rangeStr) {
            try {
                var parsed = parseRangeAddress(rangeStr);
                var firstColLetter = getColumnLetter(parsed.startCol);
                var prefix = parsed.sheet ? parsed.sheet + "!" : "";
                return prefix + "$" + firstColLetter + "$" + parsed.startRow + ":$" + firstColLetter + "$" + parsed.endRow;
            } catch (e) {
                return rangeStr;
            }
        }
```

- [ ] **Step 8: Commit**

```bash
git add src/commands/vlookup-dialog.html
git commit -m "feat: add lookup_value field and refresh-target toggle to vlookup dialog"
```

---

## Task 2: Update commands.html to handle two refresh targets

**Files:**
- Modify: `src/commands/commands.html`

- [ ] **Step 1: Update the refreshSelection handler to distinguish two targets**

Find the `if (msg.type === "refreshSelection")` block around line 257. Replace the URL-building block inside `context.sync().then`:

For `lookupValue` target:
```javascript
if (savedState.lookupValue) {
    reopenUrl += "&lookupValue=" + encodeURIComponent(savedState.lookupValue);
}
```

Add `refreshTarget` to URL:
```javascript
if (savedState.refreshTarget) {
    reopenUrl += "&refreshTarget=" + encodeURIComponent(savedState.refreshTarget);
}
```

The full handler should preserve all existing params plus the two new ones.

- [ ] **Step 2: Commit**

```bash
git add src/commands/commands.html
git commit -m "feat: handle two refresh targets in commands.html message handler"
```

---

## Verification

1. Open Excel, select a range (e.g., B2:D50 on Sheet1), click Enhanced VLOOKUP
2. Verify "查找值区域" shows the first column address (e.g., B2:B50)
3. Verify "查找表区域" is empty (user fills)
4. Click "刷新查找表" tab, click "刷新选择" — should refresh the lookup table area
5. Click "刷新查找值" tab, click "刷新选择" — should refresh the lookup value area
6. Enter lookup table range, read headers, select match/return cols, execute — should work correctly
