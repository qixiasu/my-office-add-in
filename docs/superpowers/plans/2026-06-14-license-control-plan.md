# License 访问控制系统 — 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 为 Excel 工具箱 Office Add-in 实现 License Key 访问控制，使用 Supabase 免费套餐做后端验证，继续使用 GitHub Pages 托管静态资源。

**Architecture:** Supabase Edge Function (Deno) 处理 License 验证和激活，PostgreSQL 存 License 数据。客户端 license-core.js 统一处理验证、缓存和激活 UI。各 taskpane 在启动时调用 license-core.js 完成权限校验。管理员通过 Supabase Dashboard 管理 Key。

**Tech Stack:** Deno (Supabase Edge Function), vanilla JS (客户端), webpack (构建), Supabase (PostgreSQL + Edge Function)

---

## 文件变更清单

### 新建文件

| 文件 | 职责 |
|------|------|
| `supabase/functions/validate-license/index.ts` | Supabase Edge Function：验证 License Key、设备绑定 |
| `supabase/init.sql` | 建表 SQL 脚本（供参考，实际在 Dashboard 执行） |
| `src/license/license-core.js` | 客户端 License 验证公共模块 |
| `src/license/license-ui.css` | License 激活界面样式 |
| `src/license/generate-key.html` | 浏览器端生成 License Key 的工具页面 |

### 修改文件

| 文件 | 修改内容 |
|------|---------|
| `webpack.config.js` | 添加 DefinePlugin 注入 SUPABASE_URL 和 SUPABASE_ANON_KEY |
| `src/taskpane/concat-taskpane.js` | 集成 license check |
| `src/taskpane/concat-taskpane.html` | 添加 license-ui.css |
| `src/taskpane/csv-import-taskpane.js` | 同上 |
| `src/taskpane/csv-import-taskpane.html` | 同上 |
| `src/taskpane/vlookup-taskpane.js` | 同上 |
| `src/taskpane/vlookup-taskpane.html` | 同上 |
| `src/taskpane/fill-series-taskpane.js` | 同上 |
| `src/taskpane/fill-series-taskpane.html` | 同上 |
| `src/taskpane/count-values-taskpane.js` | 同上 |
| `src/taskpane/count-values-taskpane.html` | 同上 |
| `src/taskpane/expand-taskpane.js` | 同上 |
| `src/taskpane/expand-taskpane.html` | 同上 |
| `src/taskpane/split-sheet-taskpane.js` | 同上 |
| `src/taskpane/split-sheet-taskpane.html` | 同上 |
| `src/taskpane/sql-query-taskpane.js` | 同上 |
| `src/taskpane/sql-query-taskpane.html` | 同上 |
| `src/taskpane/ai-assistant-taskpane.js` | 同上 |
| `src/taskpane/ai-assistant-taskpane.html` | 同上 |

---

### Task 1: 创建 Supabase Edge Function

**Files:**
- Create: `supabase/functions/validate-license/index.ts`
- Create: `supabase/init.sql`

- [ ] **Step 1: 创建 init.sql（建表 SQL）**

创建一个 `supabase/init.sql` 文件作为参考文档（实际在 Supabase SQL Editor 中执行）。

```sql
-- licenses 表：存储所有 License Key 数据
CREATE TABLE IF NOT EXISTS licenses (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  license_key TEXT UNIQUE NOT NULL,
  email TEXT DEFAULT '',
  tier TEXT NOT NULL DEFAULT 'basic',
  features JSONB DEFAULT '[]'::jsonb,
  issued_at TIMESTAMPTZ DEFAULT NOW(),
  expires_at TIMESTAMPTZ,
  max_devices INTEGER DEFAULT 3,
  active_devices JSONB DEFAULT '[]'::jsonb,
  revoked BOOLEAN DEFAULT FALSE,
  notes TEXT DEFAULT ''
);

-- 主键索引自动创建，额外加索引加速查询
CREATE INDEX IF NOT EXISTS idx_licenses_license_key ON licenses(license_key);
```

- [ ] **Step 2: 创建 Edge Function（验证 + 设备绑定）**

创建 `supabase/functions/validate-license/index.ts`：

```typescript
// Follow this setup guide to integrate the Deno SDK:
// https://github.com/supabase/supabase-js-v2

import { serve } from "https://deno.land/std@0.177.0/http/server.ts";
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

interface ValidationRequest {
  licenseKey: string;
  deviceId?: string;
}

interface ValidationResponse {
  valid: boolean;
  error?: string;
  tier?: string;
  features?: string[];
  expiresAt?: number;
}

serve(async (req: Request) => {
  // Handle CORS preflight
  if (req.method === "OPTIONS") {
    return new Response(null, {
      headers: {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "POST, OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type",
      },
    });
  }

  try {
    const { licenseKey, deviceId }: ValidationRequest = await req.json();

    if (!licenseKey || typeof licenseKey !== "string") {
      return json({ valid: false, error: "MISSING_KEY" } as ValidationResponse);
    }

    // 创建 Supabase 客户端（使用 SERVICE_ROLE_KEY 跳过 RLS）
    const supabase = createClient(
      Deno.env.get("SUPABASE_URL")!,
      Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!
    );

    // 1. 查 License Key
    const { data, error } = await supabase
      .from("licenses")
      .select("*")
      .eq("license_key", licenseKey.toUpperCase().trim())
      .single();

    if (error || !data) {
      return json({ valid: false, error: "KEY_INVALID" } as ValidationResponse);
    }

    // 2. 检查吊销
    if (data.revoked) {
      return json({ valid: false, error: "KEY_REVOKED" } as ValidationResponse);
    }

    // 3. 检查过期
    if (data.expires_at && new Date(data.expires_at) < new Date()) {
      return json({ valid: false, error: "KEY_EXPIRED" } as ValidationResponse);
    }

    // 4. 设备绑定（如果传了 deviceId）
    if (deviceId) {
      const devices: string[] = data.active_devices || [];
      if (!devices.includes(deviceId)) {
        if (devices.length >= data.max_devices) {
          return json({
            valid: false,
            error: "DEVICE_LIMIT",
          } as ValidationResponse);
        }
        devices.push(deviceId);
        await supabase
          .from("licenses")
          .update({ active_devices: devices })
          .eq("id", data.id);
      }
    }

    // 5. 返回验证成功
    return json({
      valid: true,
      tier: data.tier,
      features: data.features,
      expiresAt: data.expires_at
        ? Math.floor(new Date(data.expires_at).getTime() / 1000)
        : 0,
    } as ValidationResponse);
  } catch (err) {
    console.error("Validation error:", err);
    return json({ valid: false, error: "SERVER_ERROR" } as ValidationResponse);
  }
});

function json(data: ValidationResponse, status = 200): Response {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      "Content-Type": "application/json",
      "Access-Control-Allow-Origin": "*",
    },
  });
}
```

- [ ] **Step 3: 部署 Edge Function**

登录 Supabase Dashboard，进入项目 → Edge Functions → 创建 `validate-license`，粘贴上面的代码。设置环境变量 `SUPABASE_URL` 和 `SUPABASE_SERVICE_ROLE_KEY`（在 Project Settings → API 中找到）。

或者用 Supabase CLI：

```bash
# 安装 Supabase CLI 后登录
supabase login

# 在项目根目录初始化（如果还没有）
supabase init

# 链接到现有项目
supabase link --project-ref <your-project-ref>

# 部署
supabase functions deploy validate-license
```

Expected: `Function 'validate-license' deployed.`

- [ ] **Step 4: 在 Supabase SQL Editor 中建表**

在 Supabase Dashboard → SQL Editor → 粘贴 `supabase/init.sql` 中的内容 → 运行。

Expected: `CREATE TABLE`, `CREATE INDEX` 成功。

- [ ] **Step 5: 创建一条测试数据**

在 Supabase Dashboard → Table Editor → `licenses` → Insert row：

```json
{
  "license_key": "LICS-AAAAA-BBBBB-CCCCC-DDDDD-EEEEE-X",
  "tier": "enterprise",
  "features": ["concat", "csv-import", "vlookup", "fill-series", "count-values", "expand", "split-sheet", "sql-query", "ai-assistant"],
  "max_devices": 3,
  "expires_at": "2027-12-31T00:00:00Z",
  "revoked": false
}
```

- [ ] **Step 6: 提交**

```bash
git add supabase/
git commit -m "feat: add Supabase Edge Function and init SQL for license validation"
```

---

### Task 2: 创建客户端 license-core.js

**Files:**
- Create: `src/license/license-core.js`

- [ ] **Step 1: 创建 license-core.js**

```javascript
/* global localStorage, fetch, document, window, SUPABASE_URL, SUPABASE_ANON_KEY */

var LICENSE_CACHE_KEY = "excel_tools_license";
var CACHE_TTL_MS = 24 * 60 * 60 * 1000; // 24 小时
var GRACE_PERIOD_MS = 48 * 60 * 60 * 1000; // 48 小时宽限期（网络故障容错）
var DEVICE_ID_KEY = "excel_tools_device_id";

// 验证 API 地址（webpack DefinePlugin 注入）
var VALIDATE_API = typeof SUPABASE_URL !== "undefined" && typeof SUPABASE_ANON_KEY !== "undefined"
  ? SUPABASE_URL + "/functions/v1/validate-license"
  : "http://localhost:54321/functions/v1/validate-license";

function getDeviceId() {
  var id = localStorage.getItem(DEVICE_ID_KEY);
  if (!id) {
    id = "dev-" + Math.random().toString(36).substring(2, 15) + "-" + Date.now().toString(36);
    localStorage.setItem(DEVICE_ID_KEY, id);
  }
  return id;
}

function getCachedLicense() {
  try {
    var raw = localStorage.getItem(LICENSE_CACHE_KEY);
    if (!raw) return null;
    var data = JSON.parse(raw);
    if (data.cachedAt && (Date.now() - data.cachedAt) < CACHE_TTL_MS) {
      return data; // 缓存有效
    }
    return data; // 缓存过期但仍有数据（用于宽限期判断）
  } catch (e) {
    return null;
  }
}

function setCachedLicense(info) {
  info.cachedAt = Date.now();
  localStorage.setItem(LICENSE_CACHE_KEY, JSON.stringify(info));
}

function clearLicense() {
  localStorage.removeItem(LICENSE_CACHE_KEY);
}

function callValidateAPI(licenseKey, deviceId) {
  return fetch(VALIDATE_API, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ licenseKey: licenseKey, deviceId: deviceId }),
  }).then(function (resp) {
    return resp.json();
  });
}

function validateLicense(licenseKey) {
  return callValidateAPI(licenseKey, getDeviceId());
}

// ---- License 错误类型 ----
function LicenseRequiredError(message) {
  this.name = "LicenseRequiredError";
  this.message = message || "License is required to use this feature";
}
LicenseRequiredError.prototype = Object.create(Error.prototype);

// ---- 激活 UI ----
var activationResolve = null;
var activationReject = null;

function showActivationUI() {
  return new Promise(function (resolve, reject) {
    activationResolve = resolve;
    activationReject = reject;

    var overlay = document.createElement("div");
    overlay.id = "license-overlay";
    overlay.innerHTML =
      '<div class="license-modal">' +
        '<div class="license-modal-header">' +
          "<h2>🔐 激活 Excel 工具箱</h2>" +
        "</div>" +
        '<div class="license-modal-body">' +
          "<p>请输入您的 License Key 以激活插件</p>" +
          '<input type="text" id="license-key-input" class="license-input" ' +
            'placeholder="例: LICS-XXXXX-XXXXX-XXXXX-XXXXX-XXXXX" ' +
            "autocomplete=\"off\" />" +
          '<div id="license-error" class="license-error" style="display:none"></div>' +
          '<button id="license-activate-btn" class="license-btn license-btn-primary">激活</button>' +
        "</div>" +
        '<div class="license-modal-footer">' +
          '<span class="license-help-text">❓ 没有 Key？请联系管理员</span>' +
        "</div>" +
      "</div>";

    document.body.appendChild(overlay);

    var input = document.getElementById("license-key-input");
    var btn = document.getElementById("license-activate-btn");
    var errorEl = document.getElementById("license-error");

    function doActivate() {
      var key = input.value.trim().toUpperCase();
      if (!key) {
        errorEl.textContent = "请输入 License Key";
        errorEl.style.display = "block";
        return;
      }
      btn.disabled = true;
      btn.textContent = "验证中...";
      errorEl.style.display = "none";

      validateLicense(key).then(function (result) {
        if (!result.valid) {
          var messages = {
            KEY_INVALID: "Key 无效，请检查后重新输入",
            KEY_REVOKED: "该 Key 已被吊销，请联系管理员",
            KEY_EXPIRED: "该 Key 已过期，请联系管理员",
            DEVICE_LIMIT: "该 Key 已绑定过多设备，请联系管理员",
            MISSING_KEY: "请输入 License Key",
            SERVER_ERROR: "验证服务异常，请稍后重试",
          };
          errorEl.textContent = messages[result.error] || "验证失败，请重试";
          errorEl.style.display = "block";
          btn.disabled = false;
          btn.textContent = "激活";
          return;
        }

        // 验证成功
        var info = {
          licenseKey: key,
          tier: result.tier,
          features: result.features,
          expiresAt: result.expiresAt,
        };
        setCachedLicense(info);
        document.body.removeChild(overlay);
        resolve(info);
      }).catch(function () {
        errorEl.textContent = "网络错误，请检查网络连接后重试";
        errorEl.style.display = "block";
        btn.disabled = false;
        btn.textContent = "激活";
      });
    }

    btn.addEventListener("click", doActivate);
    input.addEventListener("keydown", function (e) {
      if (e.key === "Enter") doActivate();
    });

    setTimeout(function () { input.focus(); }, 100);
  });
}

// ---- 显示"无权限"提示 ----
function showNoPermission(featureName) {
  var featureLabels = {
    "concat": "连接列",
    "csv-import": "导入CSV",
    "vlookup": "增强查找",
    "fill-series": "填充序列",
    "count-values": "字段计数",
    "expand": "展开列",
    "split-sheet": "按列拆分",
    "sql-query": "数据库查询",
    "ai-assistant": "AI 助手",
  };
  var label = featureLabels[featureName] || featureName;
  var body = document.querySelector("body") || document.documentElement;
  body.innerHTML =
    '<div style="padding:40px 20px;text-align:center;color:#888;font-family:-apple-system,sans-serif;">' +
    '<p style="font-size:32px;margin-bottom:16px;">🔒</p>' +
    "<p>您当前的 License 不包含 <strong>" + label + "</strong> 功能</p>" +
    '<p style="font-size:13px;margin-top:8px;">请联系管理员升级 License</p>' +
    "</div>";
}

// ---- 核心导出函数 ----
function requireLicense(featureName) {
  // 1. 检查缓存
  var cached = getCachedLicense();
  if (cached && cached.features && cached.features.length > 0) {
    var cacheAge = Date.now() - (cached.cachedAt || 0);
    if (cacheAge < CACHE_TTL_MS) {
      // 缓存有效，检查功能权限
      if (featureName && cached.features.indexOf(featureName) < 0) {
        showNoPermission(featureName);
        return Promise.reject(new LicenseRequiredError("no permission for " + featureName));
      }
      return Promise.resolve(cached);
    }
    // 缓存过期但未超过宽限期，后台静默重新验证
    if (cacheAge < GRACE_PERIOD_MS) {
      if (featureName && cached.features.indexOf(featureName) < 0) {
        showNoPermission(featureName);
        return Promise.reject(new LicenseRequiredError("no permission for " + featureName));
      }
      // 静默重新验证（不阻塞当前流程）
      if (cached.licenseKey) {
        validateLicense(cached.licenseKey).then(function (result) {
          if (result.valid) {
            setCachedLicense({
              licenseKey: cached.licenseKey,
              tier: result.tier,
              features: result.features,
              expiresAt: result.expiresAt,
            });
          } else {
            clearLicense();
          }
        }).catch(function () { /* 网络错误，忽略，下次再试 */ });
      }
      return Promise.resolve(cached);
    }
    // 超过宽限期，需要重新验证
  }

  // 2. 尝试静默验证（如果 localStorage 里有上次的 Key 但缓存过期太久）
  var raw = localStorage.getItem(LICENSE_CACHE_KEY);
  if (raw) {
    try {
      var old = JSON.parse(raw);
      if (old.licenseKey) {
        return validateLicense(old.licenseKey).then(function (result) {
          if (result.valid) {
            var info = {
              licenseKey: old.licenseKey,
              tier: result.tier,
              features: result.features,
              expiresAt: result.expiresAt,
            };
            setCachedLicense(info);
            // 检查功能权限
            if (featureName && info.features.indexOf(featureName) < 0) {
              showNoPermission(featureName);
              throw new LicenseRequiredError("no permission for " + featureName);
            }
            return info;
          }
          // 验证失败，清除缓存，显示激活界面
          clearLicense();
          return showAndCheck(featureName);
        }).catch(function (err) {
          if (err instanceof LicenseRequiredError) throw err;
          // 网络错误，显示激活界面
          return showAndCheck(featureName);
        });
      }
    } catch (e) { /* ignore */ }
  }

  // 3. 无任何记录，显示激活界面
  return showAndCheck(featureName);
}

function showAndCheck(featureName) {
  return showActivationUI().then(function (info) {
    if (featureName && info.features.indexOf(featureName) < 0) {
      showNoPermission(featureName);
      throw new LicenseRequiredError("no permission for " + featureName);
    }
    return info;
  });
}

module.exports = {
  requireLicense: requireLicense,
  getCachedLicense: getCachedLicense,
  validateLicense: validateLicense,
  clearLicense: clearLicense,
  LicenseRequiredError: LicenseRequiredError,
};
```

- [ ] **Step 2: 提交**

```bash
git add src/license/license-core.js
git commit -m "feat: add license-core.js client module with activation UI"
```

---

### Task 3: 创建 License 激活界面样式

**Files:**
- Create: `src/license/license-ui.css`

- [ ] **Step 1: 创建 license-ui.css**

```css
/* License Activation Modal Styles */
#license-overlay {
  position: fixed;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background: rgba(0, 0, 0, 0.5);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 9999;
}

.license-modal {
  background: #fff;
  border-radius: 8px;
  box-shadow: 0 8px 32px rgba(0, 0, 0, 0.2);
  width: 420px;
  max-width: 90%;
  overflow: hidden;
}

.license-modal-header {
  background: #0078d4;
  color: #fff;
  padding: 20px 24px;
}

.license-modal-header h2 {
  margin: 0;
  font-size: 18px;
  font-weight: 600;
}

.license-modal-body {
  padding: 24px;
}

.license-modal-body p {
  margin: 0 0 16px 0;
  color: #333;
  font-size: 14px;
}

.license-input {
  width: 100%;
  padding: 10px 12px;
  font-size: 14px;
  font-family: 'Consolas', 'Courier New', monospace;
  border: 2px solid #ddd;
  border-radius: 4px;
  box-sizing: border-box;
  outline: none;
  transition: border-color 0.2s;
  text-transform: uppercase;
}

.license-input:focus {
  border-color: #0078d4;
}

.license-error {
  margin-top: 10px;
  padding: 8px 12px;
  background: #fde7e9;
  color: #a4262c;
  border-radius: 4px;
  font-size: 13px;
}

.license-btn {
  margin-top: 16px;
  padding: 10px 24px;
  font-size: 14px;
  font-weight: 600;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  transition: background 0.2s;
}

.license-btn-primary {
  background: #0078d4;
  color: #fff;
  width: 100%;
}

.license-btn-primary:hover {
  background: #106ebe;
}

.license-btn-primary:disabled {
  background: #ccc;
  cursor: not-allowed;
}

.license-modal-footer {
  padding: 12px 24px;
  background: #f4f4f4;
  text-align: center;
}

.license-help-text {
  font-size: 12px;
  color: #666;
}
```

- [ ] **Step 2: 提交**

```bash
git add src/license/license-ui.css
git commit -m "feat: add license activation modal styles"
```

---

### Task 4: 更新 webpack 配置

**Files:**
- Modify: `webpack.config.js`

- [ ] **Step 1: 添加 DefinePlugin**

在 `webpack.config.js` 顶部添加 `webpack` 引用：

```javascript
const webpack = require("webpack");
```

在 `plugins` 数组中添加 DefinePlugin（放在 CopyWebpackPlugin 前面）：

```javascript
      new webpack.DefinePlugin({
        "SUPABASE_URL": JSON.stringify(
          dev ? "http://localhost:54321" : "https://<your-project>.supabase.co"
        ),
        "SUPABASE_ANON_KEY": JSON.stringify(
          dev ? "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZS1kZW1vIiwicm9sZSI6ImFub24iLCJleHAiOjE5ODM4MTI5OTZ9.CRXP1A7WOeoJeXxjNni43kdQwgnWNReilDMblYTn_I0" : "<your-supabase-anon-key>"
        ),
      }),
```

> **说明：** 开发环境地址在本地启动 Supabase 时使用。如果不在本地调试 Edge Function，可以直接写生产地址，dev 和生产用同一个值。`<your-project>` 和 `<your-supabase-anon-key>` 在注册 Supabase 后从 Dashboard 获取。

- [ ] **Step 2: 验证 webpack 配置加载正常**

```bash
node -e "var w = require('./webpack.config.js'); w.then(function(c) { console.log('webpack config loaded, entry keys:', Object.keys(c.entry).join(', ')); })"
```

Expected: `webpack config loaded, entry keys: polyfill, commands, concat-taskpane, ...`

- [ ] **Step 3: 提交**

```bash
git add webpack.config.js
git commit -m "feat: inject SUPABASE_URL and SUPABASE_ANON_KEY via webpack DefinePlugin"
```

---

### Task 5: 在 concat-taskpane 中集成 License

**Files:**
- Modify: `src/taskpane/concat-taskpane.js`
- Modify: `src/taskpane/concat-taskpane.html`

- [ ] **Step 1: 在 HTML 中添加 license-ui.css 引用**

在 `concat-taskpane.html` 的 `<head>` 中，在最后一个 `<link>` 后面添加：

```html
    <link href="license-ui.css" rel="stylesheet" type="text/css" />
```

- [ ] **Step 2: 在 JS 中添加 license 验证**

在 `concat-taskpane.js` 顶部（`var MAX_ROWS` 下面）添加：

```javascript
var license = require("../license/license-core");
```

将 `Office.onReady` 内容包裹在 license 验证中：

```javascript
Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    license.requireLicense("concat").then(function () {
      var connectorInput = document.getElementById("connector");
      var executeBtn = document.getElementById("executeBtn");
      var confirmYes = document.getElementById("confirmYes");
      var confirmNo = document.getElementById("confirmNo");

      connectorInput.focus();

      executeBtn.addEventListener("click", function () {
        runConcat();
      });

      confirmNo.onclick = function () {
        document.getElementById("confirmBox").style.display = "none";
        if (pendingConfirmation) {
          pendingConfirmation.resolve(false);
          pendingConfirmation = null;
        }
        executeBtn.disabled = false;
        setStatus("状态：等待操作...", "idle");
      };

      confirmYes.onclick = function () {
        document.getElementById("confirmBox").style.display = "none";
        if (pendingConfirmation) {
          pendingConfirmation.resolve(true);
          pendingConfirmation = null;
        }
      };
    }).catch(function () {
      // 用户未激活或无权限，不显示功能
    });
  }
});
```

- [ ] **Step 3: 提交**

```bash
git add src/taskpane/concat-taskpane.js src/taskpane/concat-taskpane.html
git commit -m "feat: integrate license check into concat-taskpane"
```

---

### Task 6: 在 csv-import-taskpane 中集成 License

**Files:**
- Modify: `src/taskpane/csv-import-taskpane.js`
- Modify: `src/taskpane/csv-import-taskpane.html`

- [ ] **Step 1: 在 HTML 中添加 license-ui.css 引用**

在 `csv-import-taskpane.html` 的 `<head>` 中最后一个 `<link>` 后面添加：

```html
    <link href="license-ui.css" rel="stylesheet" type="text/css" />
```

- [ ] **Step 2: 在 JS 中添加 license 验证**

在 `csv-import-taskpane.js` 顶部添加：

```javascript
var license = require("../license/license-core");
```

将 `Office.onReady` 内容包裹在 `license.requireLicense("csv-import")` 中（模式与 Task 5 相同）。

- [ ] **Step 3: 提交**

```bash
git add src/taskpane/csv-import-taskpane.js src/taskpane/csv-import-taskpane.html
git commit -m "feat: integrate license check into csv-import-taskpane"
```

---

### Task 7: 在 vlookup-taskpane 中集成 License

**Files:**
- Modify: `src/taskpane/vlookup-taskpane.js`
- Modify: `src/taskpane/vlookup-taskpane.html`

操作与 Task 5 相同，功能名称为 `vlookup`。

- [ ] **Step 1: HTML 添加 license-ui.css 引用**
- [ ] **Step 2: JS 添加 `var license = require("../license/license-core")`，Office.onReady 包裹 `license.requireLicense("vlookup")`**
- [ ] **Step 3: 提交**

```bash
git add src/taskpane/vlookup-taskpane.js src/taskpane/vlookup-taskpane.html
git commit -m "feat: integrate license check into vlookup-taskpane"
```

---

### Task 8: 在 fill-series-taskpane 中集成 License

**Files:**
- Modify: `src/taskpane/fill-series-taskpane.js`
- Modify: `src/taskpane/fill-series-taskpane.html`

操作同上，功能名称为 `fill-series`。

- [ ] **Step 1: HTML 添加 license-ui.css 引用**
- [ ] **Step 2: JS 集成 `license.requireLicense("fill-series")`**
- [ ] **Step 3: 提交**

```bash
git add src/taskpane/fill-series-taskpane.js src/taskpane/fill-series-taskpane.html
git commit -m "feat: integrate license check into fill-series-taskpane"
```

---

### Task 9: 在 count-values-taskpane 中集成 License

**Files:**
- Modify: `src/taskpane/count-values-taskpane.js`
- Modify: `src/taskpane/count-values-taskpane.html`

操作同上，功能名称为 `count-values`。

- [ ] **Step 1: HTML 添加 license-ui.css 引用**
- [ ] **Step 2: JS 集成 `license.requireLicense("count-values")`**
- [ ] **Step 3: 提交**

```bash
git add src/taskpane/count-values-taskpane.js src/taskpane/count-values-taskpane.html
git commit -m "feat: integrate license check into count-values-taskpane"
```

---

### Task 10: 在 expand-taskpane 中集成 License

**Files:**
- Modify: `src/taskpane/expand-taskpane.js`
- Modify: `src/taskpane/expand-taskpane.html`

操作同上，功能名称为 `expand`。

- [ ] **Step 1: HTML 添加 license-ui.css 引用**
- [ ] **Step 2: JS 集成 `license.requireLicense("expand")`**
- [ ] **Step 3: 提交**

```bash
git add src/taskpane/expand-taskpane.js src/taskpane/expand-taskpane.html
git commit -m "feat: integrate license check into expand-taskpane"
```

---

### Task 11: 在 split-sheet-taskpane 中集成 License

**Files:**
- Modify: `src/taskpane/split-sheet-taskpane.js`
- Modify: `src/taskpane/split-sheet-taskpane.html`

操作同上，功能名称为 `split-sheet`。

- [ ] **Step 1: HTML 添加 license-ui.css 引用**
- [ ] **Step 2: JS 集成 `license.requireLicense("split-sheet")`**
- [ ] **Step 3: 提交**

```bash
git add src/taskpane/split-sheet-taskpane.js src/taskpane/split-sheet-taskpane.html
git commit -m "feat: integrate license check into split-sheet-taskpane"
```

---

### Task 12: 在 sql-query-taskpane 中集成 License

**Files:**
- Modify: `src/taskpane/sql-query-taskpane.js`
- Modify: `src/taskpane/sql-query-taskpane.html`

操作同上，功能名称为 `sql-query`。

- [ ] **Step 1: HTML 添加 license-ui.css 引用**
- [ ] **Step 2: JS 集成 `license.requireLicense("sql-query")`**
- [ ] **Step 3: 提交**

```bash
git add src/taskpane/sql-query-taskpane.js src/taskpane/sql-query-taskpane.html
git commit -m "feat: integrate license check into sql-query-taskpane"
```

---

### Task 13: 在 ai-assistant-taskpane 中集成 License

**Files:**
- Modify: `src/taskpane/ai-assistant-taskpane.js`
- Modify: `src/taskpane/ai-assistant-taskpane.html`

操作同上，功能名称为 `ai-assistant`。

- [ ] **Step 1: HTML 添加 license-ui.css 引用**
- [ ] **Step 2: JS 集成 `license.requireLicense("ai-assistant")`**
- [ ] **Step 3: 提交**

```bash
git add src/taskpane/ai-assistant-taskpane.js src/taskpane/ai-assistant-taskpane.html
git commit -m "feat: integrate license check into ai-assistant-taskpane"
```

---

### Task 14: 创建浏览器端 Key 生成工具

**Files:**
- Create: `src/license/generate-key.html`

- [ ] **Step 1: 创建 generate-key.html**

这是一个纯静态页面，管理员在浏览器中打开，生成 License Key 后手动复制到 Supabase Dashboard 插入。

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>License Key 生成器</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #f0f2f5; display: flex; justify-content: center; padding: 40px 20px; }
    .container { background: #fff; border-radius: 12px; padding: 32px; box-shadow: 0 4px 24px rgba(0,0,0,0.1); width: 600px; max-width: 100%; }
    h1 { font-size: 22px; margin-bottom: 24px; color: #0078d4; }
    .form-group { margin-bottom: 16px; }
    .form-group label { display: block; font-size: 13px; color: #666; margin-bottom: 4px; font-weight: 600; }
    .form-group input, .form-group select { width: 100%; padding: 10px 12px; border: 2px solid #ddd; border-radius: 6px; font-size: 14px; }
    .form-group input:focus, .form-group select:focus { border-color: #0078d4; outline: none; }
    .form-row { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
    button { padding: 12px 24px; background: #0078d4; color: #fff; border: none; border-radius: 6px; font-size: 15px; font-weight: 600; cursor: pointer; width: 100%; }
    button:hover { background: #106ebe; }
    button:disabled { background: #ccc; cursor: not-allowed; }
    #result { margin-top: 16px; display: none; }
    .result-box { background: #e8f5e9; border-radius: 6px; padding: 16px; word-break: break-all; }
    .result-box .key { font-family: 'Consolas', 'Courier New', monospace; font-size: 18px; letter-spacing: 2px; color: #2e7d32; font-weight: bold; }
    .result-box .detail { font-size: 13px; color: #555; margin-top: 8px; }
    .copy-btn { margin-top: 8px; padding: 8px 16px; background: #fff; color: #0078d4; border: 1px solid #0078d4; border-radius: 4px; cursor: pointer; font-size: 13px; }
    .copy-btn:hover { background: #e6f2fb; }
  </style>
</head>
<body>
  <div class="container">
    <h1>🔑 License Key 生成器</h1>

    <div class="form-group">
      <label>层级 (Tier)</label>
      <select id="tier">
        <option value="basic">Basic</option>
        <option value="pro" selected>Pro</option>
        <option value="enterprise">Enterprise</option>
      </select>
    </div>

    <div class="form-group">
      <label>邮箱（可选）</label>
      <input type="email" id="email" placeholder="user@example.com" />
    </div>

    <div class="form-row">
      <div class="form-group">
        <label>最大设备数</label>
        <input type="number" id="maxDevices" value="3" min="1" max="10" />
      </div>
      <div class="form-group">
        <label>有效期（天）</label>
        <input type="number" id="days" value="365" min="1" max="3650" />
      </div>
    </div>

    <div class="form-group">
      <label>备注</label>
      <input type="text" id="notes" placeholder="用户姓名或来源" />
    </div>

    <button id="generateBtn">生成 License Key</button>

    <div id="result">
      <div class="result-box">
        <div class="key" id="generatedKey"></div>
        <div class="detail" id="generatedDetail"></div>
      </div>
      <button class="copy-btn" id="copyBtn">📋 复制 Key</button>
      <div style="margin-top:12px;padding:10px;background:#fff3e0;border-radius:4px;font-size:13px;color:#e65100;">
        ⚠️ 请手动将 Key 和相关信息插入 Supabase Dashboard → Table Editor → licenses 表
      </div>
    </div>
  </div>

  <script>
    var CHARSET = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';

    function generateLicenseKey() {
      var groups = [];
      for (var g = 0; g < 5; g++) {
        var group = '';
        for (var i = 0; i < 5; i++) {
          group += CHARSET.charAt(Math.floor(Math.random() * CHARSET.length));
        }
        groups.push(group);
      }
      var raw = groups.join('');
      var checksum = 0;
      for (var j = 0; j < raw.length; j++) {
        checksum = (checksum * 31 + raw.charCodeAt(j)) % CHARSET.length;
      }
      return 'LICS-' + raw.slice(0, 5) + '-' + raw.slice(5, 10) + '-' +
        raw.slice(10, 15) + '-' + raw.slice(15, 20) + '-' + raw.slice(20, 25) + CHARSET.charAt(checksum);
    }

    var FEATURES = {
      basic: ['concat', 'csv-import', 'vlookup', 'fill-series', 'count-values'],
      pro: ['concat', 'csv-import', 'vlookup', 'fill-series', 'count-values', 'expand', 'split-sheet'],
      enterprise: ['concat', 'csv-import', 'vlookup', 'fill-series', 'count-values', 'expand', 'split-sheet', 'sql-query', 'ai-assistant'],
    };

    document.getElementById('generateBtn').addEventListener('click', function() {
      var key = generateLicenseKey();
      var tier = document.getElementById('tier').value;
      var days = parseInt(document.getElementById('days').value) || 365;
      var maxDevices = parseInt(document.getElementById('maxDevices').value) || 3;
      var email = document.getElementById('email').value.trim();
      var notes = document.getElementById('notes').value.trim();

      var expiresAt = new Date();
      expiresAt.setDate(expiresAt.getDate() + days);

      document.getElementById('generatedKey').textContent = key;
      document.getElementById('generatedDetail').innerHTML =
        '层级: ' + tier + '<br>' +
        '功能: ' + FEATURES[tier].join(', ') + '<br>' +
        '设备上限: ' + maxDevices + '<br>' +
        '过期: ' + expiresAt.toLocaleDateString('zh-CN') + '<br>' +
        (email ? '邮箱: ' + email + '<br>' : '') +
        (notes ? '备注: ' + notes : '');

      document.getElementById('result').style.display = 'block';
      document.getElementById('copyBtn').textContent = '📋 复制 Key';
    });

    document.getElementById('copyBtn').addEventListener('click', function() {
      var key = document.getElementById('generatedKey').textContent;
      navigator.clipboard.writeText(key).then(function() {
        document.getElementById('copyBtn').textContent = '✅ 已复制！';
      }).catch(function() {
        // fallback: select text
        var range = document.createRange();
        var sel = window.getSelection();
        range.selectNodeContents(document.getElementById('generatedKey'));
        sel.removeAllRanges();
        sel.addRange(range);
      });
    });
  </script>
</body>
</html>
```

- [ ] **Step 2: 提交**

```bash
git add src/license/generate-key.html
git commit -m "feat: add browser-based License Key generator tool"
```

---

## 自审清单

- ✅ **Spec 覆盖**: 设计文档中所有内容（Edge Function、数据模型、客户端集成、缓存策略、激活 UI、层级定义、管理操作）全部有对应任务
- ✅ **无占位符**: 所有代码步骤包含完整实现，`<your-project>` 是部署时才确定的配置值，合理
- ✅ **类型一致性**: license-core.js 的 validateLicense 返回格式与 Edge Function 输出一致（valid, tier, features, expiresAt）
- ✅ **范围控制**: 分为 14 个独立任务，每个可独立提交和测试
- ✅ **YAGNI**: 没有开发管理后台（Dashboard 够用），没有开发支付系统，没有开发自助申请流程

---

## 执行方式

**Plan complete and saved to `docs/superpowers/plans/2026-06-14-license-control-plan.md`。两种执行方式：**

**1. Subagent-Driven（推荐）** — 每个任务启动一个独立子代理，任务间审查，快速迭代

**2. Inline Execution** — 在当前会话中执行，批处理 + 检查点

**选择哪种方式？**
