# Office Add-in License 访问控制系统设计

## 概述

Excel 工具箱（Office Add-in）目前通过 GitHub Pages 公开分发，任何获取 manifest.xml 的人均可使用，管理员无法控制用户范围或功能权限。本文档设计了一套面向中国境内用户的 License 访问控制系统，使用 Supabase（免费套餐）提供后端验证服务，继续使用 GitHub Pages 托管静态资源。

## 目标

- 通过 License Key 实现用户白名单控制
- 支持功能粒度权限（不同层级用户可见不同功能）
- 支持 Key 吊销和设备数限制，有效防滥用
- 管理员可随时生成/吊销 Key，无需修改代码
- **零服务器运维**，全托管方案
- **继续使用 GitHub Pages**，不迁移静态资源

## 架构

```
┌─ GitHub Pages (静态托管，保持不变) ────────┐
│  /my-office-add-in/                        │
│    ├── taskpane.html / concat-taskpane.html │
│    ├── *.js (webpack 打包)                  │
│    └── manifest.xml                         │
└───────────────────────────────────────────┘
         │  HTTPS 请求
         ▼
┌─ Supabase (免费套餐) ─────────────────────┐
│                                           │
│  Edge Function                            │
│  POST /validate-license                   │
│       ↓                                   │
│  PostgreSQL: licenses 表                   │
│                                           │
│  Supabase Dashboard: 管理员管理 Key        │
└───────────────────────────────────────────┘

┌─ 用户 Excel ─────────────────────────────┐
│  manifest.xml → GitHub Pages 加载资源       │
│    ↓                                      │
│  license-core.js → 调 Supabase 验证       │
│    ↓                                      │
│  features 控制 → 显示/隐藏各功能面板        │
└───────────────────────────────────────────┘
```

## 技术选型

| 组件 | 选择 | 原因 |
|------|------|------|
| 静态托管 | GitHub Pages（不变） | 无需迁移，国内可访问 |
| 后端服务 | Supabase (免费套餐) | 全托管，零运维，有 Dashboard |
| 验证逻辑 | Supabase Edge Function (Deno) | 与数据库同集群，延迟低 |
| 数据库 | Supabase PostgreSQL | 免费 500MB |
| License Key 算法 | 生成+校验码 | 简单防错 |

## License Key 系统

### Key 格式

```
LICS-XXXXX-XXXXX-XXXXX-XXXXX-XXXXX
  │     │      │      │      │
 前缀   5组大写字母+数字（排除 0/O/1/I/L）
```

- 前缀 `LICS` 固定标识
- 每组 5 字符，5 组 + 1 个校验字符 = 26 字符
- 排除易混淆字符: `0`、`O`、`1`、`I`、`L`
- 最后一位为校验码（前 25 字符的 CRC）

### Key 生成流程

管理员通过 `generate-key.html` 网页生成 Key，写入 Supabase 数据库。

### 数据模型

PostgreSQL 表 `licenses`：

```sql
CREATE TABLE licenses (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  license_key TEXT UNIQUE NOT NULL,
  email TEXT DEFAULT '',
  tier TEXT NOT NULL DEFAULT 'basic',
  features JSONB,              -- 功能列表
  issued_at TIMESTAMPTZ DEFAULT NOW(),
  expires_at TIMESTAMPTZ,      -- 过期时间
  max_devices INTEGER DEFAULT 3,
  active_devices JSONB DEFAULT '[]'::jsonb,
  revoked BOOLEAN DEFAULT FALSE,
  notes TEXT DEFAULT ''
);
```

### 层级定义

| 层级 | 功能 | 价格概念 |
|------|------|---------|
| `basic` | 数据处理（连接列、导入CSV、增强查找、填充序列、字段计数） | 基础版 |
| `pro` | basic + 数据转换（展开列、按列拆分） | 进阶版 |
| `enterprise` | pro + 数据库查询 + AI 助手 | 完整版 |

## Edge Function 接口

### POST /validate-license

请求:
```json
{ "licenseKey": "LICS-XXXXX-XXXXX-XXXXX-XXXXX-XXXXX", "deviceId": "dev-abc123" }
```

响应：
```json
// 验证成功
{ "valid": true, "tier": "pro", "features": ["concat", "csv-import", ...], "expiresAt": 1784092800 }
// 验证失败
{ "valid": false, "error": "KEY_INVALID" | "KEY_REVOKED" | "KEY_EXPIRED" | "DEVICE_LIMIT" }
// 服务端错误
{ "valid": false, "error": "SERVER_ERROR" }
```

逻辑流程：
1. 从 Postgres 查询 `license_key` 匹配记录
2. 检查 `revoked` 标记
3. 检查 `expires_at` 是否过期
4. 如果传了 `deviceId`：检查是否已绑定 → 未超限则加入 `active_devices`
5. 返回验证结果

## 客户端集成设计

### 文件结构

```
src/
├── license/
│   ├── license-core.js       # 验证、缓存、激活逻辑
│   └── license-ui.css        # 激活界面样式
├── taskpane/
│   ├── xxx-taskpane.js       # 各面板代码，集成 license check
│   └── ...
```

### license-core.js API

```javascript
async function requireLicense(featureName)
// 检查 License，如果无效则显示激活界面等待用户输入
// 返回 { tier, features }
// 如果有 featureName 参数且用户无此功能权限，显示"无权限"提示
// 如果用户放弃激活，抛出 LicenseRequiredError

function getCachedLicense()
// 获取本地缓存的 License 信息（不触网）

async function validateLicense(key, deviceId)
// 手动触发验证

function clearLicense()
// 清除本地 License（退出登录）
```

### 各 Taskpane 集成方式

每个 taskpane js 文件只需在 `Office.onReady` 中加：

```javascript
var license = require("../license/license-core");

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    license.requireLicense("concat").then(function () {
      initPanel(); // 有权限，正常初始化
    }).catch(function () {
      // 无权限或用户取消，不显示功能
    });
  }
});
```

### 缓存策略

- 首次验证成功后，将 Key + features + 时间戳存入 `localStorage`
- 后续加载先读缓存，如果缓存时间 < 24 小时则静默使用
- 缓存过期后，调用 Supabase 静默重新验证（不阻塞 UI）
- 重新验证失败但缓存仍在 48h 宽限期内 → 允许继续使用（防网络故障）
- 超过宽限期 → 清除缓存，下次加载回到激活界面

### 激活界面行为

```
用户打开任意面板
  → 检测无有效 License → 显示 License 激活弹窗
    → 输入 Key → 调 Supabase Edge Function
      → 成功 → 自动绑定设备 → 进入功能面板
      → 失败 → 显示具体错误原因
  → 已激活用户再次打开 → 无缝进入功能（静默验证）
  → Key 被吊销后 24h 内 → 弹出 "License 已失效，请联系管理员"
```

## 管理员操作流程

**原则：不开发管理后台，初期直接用 Supabase Dashboard。**

| 操作 | 步骤 |
|------|------|
| 生成 Key | Supabase Dashboard → Table Editor → 点击 Insert row → 手动输入 license_key、tier、expires_at 等 |
| 批量生成 | 提供一个 `generate-key.html` 页面，在浏览器本地生成 Key 并显示，管理员手动复制到 Supabase 插入 |
| 吊销 | 找到对应行 → 勾选 `revoked = true` |
| 查看状态 | Dashboard 直接浏览所有记录 |
| 查看设备绑定 | 看 `active_devices` 数组长度 |

> **未来可选项：** 当 Key 数量多了以后，可以开发一个简单的管理页面（纯静态 HTML + 直接调 Supabase REST API），部署在 GitHub Pages 上，需要管理员 Token 才能访问。

## 防滥用机制

| 层 | 措施 | 说明 |
|----|------|------|
| Key | 设备绑定 | 一个 Key 最多绑定 N 台设备（管理员配置） |
| Key | 吊销能力 | 管理员可随时吊销任何 Key |
| Key | 过期机制 | Key 设置有效期，到期自动失效 |
| 代码 | 混淆 | JS 在构建时做基本混淆 |
| 传输 | HTTPS 加密 | 全链路 HTTPS（Supabase 默认） |
| Edge Function | 请求频率限制 | Supabase 层面可配置 Rate Limiting |

## 项目实施

### 阶段 1: 后端搭建（~30 分钟）

1. 注册 Supabase 账号（可用 GitHub 登录）
2. 创建项目，选免费套餐
3. SQL Editor 中执行 `CREATE TABLE licenses` 建表
4. 创建并部署 Edge Function `validate-license`
5. 记录项目 URL 和 Anon Key

### 阶段 2: 客户端开发（~2 小时）

1. 开发 `license-core.js` 公共模块
2. 开发 License 激活 UI 样式
3. 在每个 taskpane 中集成 `requireLicense()`
4. 配置 webpack DefinePlugin 注入 Supabase URL

### 阶段 3: 管理工具（~1 小时）

1. 创建 `generate-key.html` 浏览器端 Key 生成工具
2. 编写使用说明文档

### 阶段 4: 测试验证（~1 小时）

1. 启动 dev server 测试验证流程
2. 测试各种错误情况（无效 Key、吊销、过期、设备超限）
3. 测试缓存机制

## 开放问题

- 是否需要开发者提供 Supabase 服务账号，还是你自己注册？
- 初期只接受你手动生成 Key 分发，还是未来需要自助购买/申请流程？
- 是否需要开发 Web 管理页面？
