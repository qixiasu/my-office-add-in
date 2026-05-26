# Office Add-in 本地 HTTP 服务

## 概述

将 Excel Add-in 的前端 build 结果通过一个轻量级 Go HTTP 服务托管，提供系统托盘支持和用户可控的开机自启功能。实现单文件分发给同事时，无需安装 Node.js 或 Python。

## 背景

**现状：**
- Excel Add-in 前端 build 后输出到 `dist` 文件夹
- 需要 Node.js 的 http-server 或 Python 的 http.server 托管
- 分享给同事时，对方需要安装 Node.js 或 Python

**目标：**
- 编译成单个 Go 二进制文件（Windows/macOS/Linux）
- 双击即用，无需安装运行时
- 内置系统托盘和开机自启功能

## 架构

```
┌─────────────────────────────────────────┐
│  office-addin-server (Go 二进制)        │
│                                          │
│  ┌─────────────────────────────────┐    │
│  │  HTTP Server (静态文件服务)       │    │
│  │  - 端口: 3000 (可配置)            │    │
│  │  - 目录: ./dist (可配置)          │    │
│  └─────────────────────────────────┘    │
│                                          │
│  ┌─────────────────────────────────┐    │
│  │  System Tray                    │    │
│  │  - 右键菜单                       │    │
│  │  - 打开/关闭开机自启               │    │
│  │  - 退出程序                       │    │
│  └─────────────────────────────────┘    │
│                                          │
│  ┌─────────────────────────────────┐    │
│  │  Config (config.json)            │    │
│  │  - port: 3000                    │    │
│  │  - dir: ./dist                    │    │
│  │  - autostart: false              │    │
│  └─────────────────────────────────┘    │
└─────────────────────────────────────────┘
```

## 功能清单

### 1. HTTP 静态文件服务
- 托管 `dist` 文件夹（或自定义目录）
- 默认端口 `3000`
- 支持目录可配置（命令行参数或配置文件）

### 2. 系统托盘
- 启动后最小化到系统托盘
- 托盘图标右键菜单：
  - **打开管理页面** — 在浏览器打开 `http://localhost:<port>`
  - **开启开机自启** — 勾选后写入系统自启
  - **关闭开机自启** — 取消自启
  - **退出** — 关闭程序

### 3. 开机自启
- 用户通过托盘菜单开关控制
- 状态保存到 `config.json`
- Windows：通过注册表 `HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run`
- macOS：通过 `~/Library/LaunchAgents` plist
- Linux：通过 `~/.config/autostart` desktop 文件

### 4. 配置文件
- 文件名：`config.json`
- 位置：程序同目录下
- 格式：
```json
{
  "port": 3000,
  "dir": "./dist",
  "autostart": false
}
```

## 使用方式

```bash
# 默认配置（端口3000，托管dist目录）
./office-addin-server

# 自定义端口
./office-addin-server -port 8080

# 自定义目录和端口
./office-addin-server -dir ./dist -port 3000

# 显示帮助
./office-addin-server -h
```

## 项目结构

```
office-addin-server/
├── main.go              # 主程序
├── go.mod               # Go 模块
├── config.go            # 配置管理
├── tray.go              # 系统托盘
├── autostart.go         # 开机自启
├── static.go           # 静态文件服务
└── build.sh            # 跨平台编译脚本
```

## 技术选型

| 组件 | 库 | 说明 |
|------|-----|------|
| HTTP 服务 | Go 标准库 `net/http` | 无需第三方库 |
| 系统托盘 | `github.com/getlantern/systray` | 跨平台托盘支持 |
| 命令行参数 | 标准库 `flag` | 轻量解析 |
| 配置文件 | Go 标准库 `encoding/json` | 无需第三方库 |

## 错误处理

| 场景 | 处理 |
|------|------|
| 端口被占用 | 提示用户端口已被占用，建议换端口 |
| dist 目录不存在 | 创建目录并提示"请将 Add-in build 结果放入 dist 目录" |
| 配置文件损坏 | 使用默认配置，重写配置文件 |

## 打包发布

提供 `build.sh` 编译脚本，支持：
- Windows: `office-addin-server.exe`
- macOS: `office-addin-server`
- Linux: `office-addin-server`

编译产物单个文件，无依赖。