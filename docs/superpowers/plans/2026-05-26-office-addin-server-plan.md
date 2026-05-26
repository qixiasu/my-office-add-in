# Office Add-in HTTP Server Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Go-based HTTP server with system tray and autostart for distributing the Excel Add-in without Node.js/Python

**Architecture:** Single Go binary that serves static files from dist folder, with system tray for quick access and autostart toggle

**Tech Stack:** Go standard library (net/http, encoding/json, flag), github.com/getlantern/systray for cross-platform tray

---

## File Structure

```
office-addin-server/
├── main.go       # Entry point, coordinates all components
├── config.go     # Config loading/saving, port/dir/autostart
├── http.go       # Static file HTTP server
├── tray.go       # System tray setup and menu handling
├── autostart.go  # Windows autostart via registry
└── build.sh      # Cross-platform build script
```

---

## Task 1: 项目初始化

**Files:**
- Create: `office-addin-server/go.mod`

Run:
```bash
cd office-addin-server && go mod init office-addin-server
```

Install dependencies:
```bash
cd office-addin-server && go get github.com/getlantern/systray@latest
```

---

## Task 2: config.go - 配置管理

**Files:**
- Create: `office-addin-server/config.go`

```go
package main

import (
    "encoding/json"
    "os"
    "path/filepath"
)

type Config struct {
    Port      int    `json:"port"`
    Dir       string `json:"dir"`
    Autostart bool   `json:"autostart"`
}

func DefaultConfig() *Config {
    return &Config{
        Port:      3000,
        Dir:       "./dist",
        Autostart: false,
    }
}

func LoadConfig() (*Config, error) {
    cfg := DefaultConfig()

    exePath, err := os.Executable()
    if err != nil {
        return cfg, err
    }
    cfgDir := filepath.Dir(exePath)
    cfgFile := filepath.Join(cfgDir, "config.json")

    data, err := os.ReadFile(cfgFile)
    if err != nil {
        if os.IsNotExist(err) {
            return cfg, nil
        }
        return nil, err
    }

    if err := json.Unmarshal(data, cfg); err != nil {
        cfg = DefaultConfig()
        return cfg, nil
    }

    return cfg, nil
}

func SaveConfig(cfg *Config) error {
    exePath, err := os.Executable()
    if err != nil {
        return err
    }
    cfgDir := filepath.Dir(exePath)
    cfgFile := filepath.Join(cfgDir, "config.json")

    data, err := json.MarshalIndent(cfg, "", "  ")
    if err != nil {
        return err
    }

    return os.WriteFile(cfgFile, data, 0644)
}
```

---

## Task 3: http.go - HTTP 静态文件服务

**Files:**
- Create: `office-addin-server/http.go`

```go
package main

import (
    "fmt"
    "net/http"
    "os"
    "path/filepath"
)

type HTTPServer struct {
    port int
    dir  string
    server *http.Server
}

func NewHTTPServer(port int, dir string) *HTTPServer {
    return &HTTPServer{
        port: port,
        dir:  dir,
    }
}

func (h *HTTPServer) Start() error {
    absDir, err := filepath.Abs(h.dir)
    if err != nil {
        return fmt.Errorf("failed to get absolute path: %w", err)
    }

    if _, err := os.Stat(absDir); os.IsNotExist(err) {
        if err := os.MkdirAll(absDir, 0755); err != nil {
            return fmt.Errorf("failed to create directory %s: %w", absDir, err)
        }
    }

    mux := http.NewServeMux()
    mux.Handle("/", http.FileServer(http.Dir(absDir)))

    h.server = &http.Server{
        Addr:    fmt.Sprintf(":%d", h.port),
        Handler: mux,
    }

    return h.server.ListenAndServe()
}

func (h *HTTPServer) Stop() error {
    if h.server != nil {
        return h.server.Close()
    }
    return nil
}
```

---

## Task 4: autostart.go - 开机自启

**Files:**
- Create: `office-addin-server/autostart.go`

```go
package main

import (
    "fmt"
    "os"
    "path/filepath"
    "syscall"
    "unsafe"
)

const (
    regKeyPath = `Software\Microsoft\Windows\CurrentVersion\Run`
)

var (
    advapi32 = syscall.NewLazyDLL("advapi32.dll")
    regOpenKeyExW  = advapi32.NewProc("RegOpenKeyExW")
    regSetValueExW = advapi32.NewProc("RegSetValueExW")
    regCloseKey    = advapi32.NewProc("RegCloseKey")
)

func SetAutostart(name string, exePath string) error {
    var key syscall.Handle
    ret, _, _ := regOpenKeyExW.Call(
        syscall.HKEY_CURRENT_USER,
        syscall.StringToUTF16Ptr(regKeyPath),
        0,
        syscall.KEY_WRITE,
        uintptr(unsafe.Pointer(&key)),
    )
    if ret != 0 {
        return fmt.Errorf("failed to open registry key: %d", ret)
    }
    defer regCloseKey.Call(uintptr(key))

    exePathPtr := syscall.StringToUTF16Ptr(exePath)
    ret, _, _ = regSetValueExW.Call(
        uintptr(key),
        syscall.StringToUTF16Ptr(name),
        0,
        syscall.REG_SZ,
        uintptr(unsafe.Pointer(exePathPtr)),
        uintptr(len(exePath)*2),
    )
    if ret != 0 {
        return fmt.Errorf("failed to set registry value: %d", ret)
    }

    return nil
}

func RemoveAutostart(name string) error {
    var key syscall.Handle
    ret, _, _ := regOpenKeyExW.Call(
        syscall.HKEY_CURRENT_USER,
        syscall.StringToUTF16Ptr(regKeyPath),
        0,
        syscall.KEY_WRITE,
        uintptr(unsafe.Pointer(&key)),
    )
    if ret != 0 {
        return fmt.Errorf("failed to open registry key: %d", ret)
    }
    defer regCloseKey.Call(uintptr(key))

    ret, _, _ = syscall.NewLazyDLL("advapi32.dll").NewProc("RegDeleteValueW").Call(
        uintptr(key),
        syscall.StringToUTF16Ptr(name),
    )
    if ret != 0 && ret != 2 {
        return fmt.Errorf("failed to delete registry value: %d", ret)
    }

    return nil
}
```

---

## Task 5: tray.go - 系统托盘

**Files:**
- Create: `office-addin-server/tray.go`

```go
package main

import (
    "fmt"
    "os"
    "os/exec"
    "path/filepath"

    "github.com/getlantern/systray"
)

type TrayManager struct {
    port       int
    autostart  bool
    onToggle   func(bool)
}

func NewTrayManager(port int, autostart bool, onToggle func(bool)) *TrayManager {
    return &TrayManager{
        port:      port,
        autostart: autostart,
        onToggle:  onToggle,
    }
}

func (t *TrayManager) Setup() {
    systray.SetTemplateIcon(nil, nil)
    systray.SetTitle("Office Add-in Server")
    systray.SetTooltip(fmt.Sprintf("Serving on port %d", t.port))

    mOpen := systray.AddMenuItem("Open", "Open in browser")
    mToggleAuto := systray.AddMenuItem("Enable Autostart", "Toggle autostart")
    if t.autostart {
        mToggleAuto.Checked = true
    }
    mQuit := systray.AddMenuItem("Quit", "Exit")

    go func() {
        for {
            select {
            case <-mOpen.ClickedCh:
                t.openBrowser()
            case <-mToggleAuto.ClickedCh:
                t.autostart = !t.autostart
                mToggleAuto.Checked = t.autostart
                if t.onToggle != nil {
                    t.onToggle(t.autostart)
                }
            case <-mQuit.ClickedCh:
                systray.Quit()
                os.Exit(0)
            }
        }
    }()
}

func (t *TrayManager) openBrowser() {
    url := fmt.Sprintf("http://localhost:%d", t.port)
    exec.Command("cmd", "/c", "start", url).Start()
}

func GetExecutablePath() (string, error) {
    exe, err := os.Executable()
    if err != nil {
        return "", err
    }
    return filepath.Abs(exe)
}
```

---

## Task 6: main.go - 主程序入口

**Files:**
- Create: `office-addin-server/main.go`

```go
package main

import (
    "flag"
    "fmt"
    "log"
    "os"
    "os/signal"
    "syscall"

    "github.com/getlantern/systray"
)

var (
    flagPort    = flag.Int("port", 3000, "HTTP server port")
    flagDir     = flag.String("dir", "./dist", "Static files directory")
    flagHelp    = flag.Bool("h", false, "Show help")
)

func main() {
    flag.Parse()

    if *flagHelp {
        flag.Usage()
        os.Exit(0)
    }

    cfg, err := LoadConfig()
    if err != nil {
        log.Printf("Warning: failed to load config: %v, using defaults", err)
        cfg = DefaultConfig()
    }

    port := *flagPort
    dir := *flagDir
    if dir == "./dist" && cfg.Dir != "" {
        dir = cfg.Dir
    }
    if port == 3000 && cfg.Port != 0 {
        port = cfg.Port
    }

    exePath, err := GetExecutablePath()
    if err != nil {
        log.Fatalf("Failed to get executable path: %v", err)
    }

    httpServer := NewHTTPServer(port, dir)

    sigCh := make(chan os.Signal, 1)
    signal.Notify(sigCh, syscall.SIGINT, syscall.SIGTERM)

    go func() {
        <-sigCh
        httpServer.Stop()
        os.Exit(0)
    }()

    systray.Run(func() {
        tray := NewTrayManager(port, cfg.Autostart, func(enabled bool) {
            cfg.Autostart = enabled
            if err := SaveConfig(cfg); err != nil {
                log.Printf("Failed to save config: %v", err)
            }

            if enabled {
                if err := SetAutostart("OfficeAddinServer", exePath); err != nil {
                    log.Printf("Failed to enable autostart: %v", err)
                }
            } else {
                if err := RemoveAutostart("OfficeAddinServer"); err != nil {
                    log.Printf("Failed to disable autostart: %v", err)
                }
            }
        })
        tray.Setup()

        go func() {
            if err := httpServer.Start(); err != nil {
                log.Printf("HTTP server error: %v", err)
            }
        }()
    }, func() {})
}
```

---

## Task 7: build.sh - 跨平台编译脚本

**Files:**
- Create: `office-addin-server/build.sh`

```bash
#!/bin/bash

APP_NAME="office-addin-server"

echo "Building $APP_NAME..."

# Clean previous builds
rm -rf dist

mkdir -p dist

# Build for current platform
if [[ "$OSTYPE" == "darwin"* ]]; then
    GOOS=darwin GOARCH=amd64 go build -o "dist/$APP_NAME" .
    GOOS=darwin GOARCH=arm64 go build -o "dist/${APP_NAME}-arm64" .
elif [[ "$OSTYPE" == "linux-gnu"* ]]; then
    GOOS=linux GOARCH=amd64 go build -o "dist/$APP_NAME" .
else
    GOOS=windows GOARCH=amd64 go build -o "dist/${APP_NAME}.exe" .
fi

echo "Build complete!"
ls -la dist/
```

---

## Task 8: 测试流程

1. **编译 Go 程序**
```bash
cd office-addin-server && go build -o office-addin-server.exe .
```

2. **创建测试 dist 目录**
```bash
mkdir -p dist && echo "test" > dist/test.txt
```

3. **运行程序**
```bash
./office-addin-server.exe
```

4. **验证**
- 访问 `http://localhost:3000/test.txt` 确认能看到 "test"
- 检查系统托盘图标是否出现
- 测试托盘菜单功能

---

## Task 9: 完整分发包

**创建分发目录结构：**
```
release/
├── office-addin-server.exe   # Go 编译的可执行文件
└── dist/                       # Excel Add-in build 结果
    ├── taskpane.html
    ├── commands.html
    └── ...
```

用户只需：
1. 解压 release 包
2. 运行 `office-addin-server.exe`
3. 在 Excel 中加载 Add-in

---

## 实施检查清单

- [ ] Task 1: 项目初始化 (go.mod)
- [ ] Task 2: config.go 配置管理
- [ ] Task 3: http.go HTTP 服务
- [ ] Task 4: autostart.go 开机自启
- [ ] Task 5: tray.go 系统托盘
- [ ] Task 6: main.go 主程序
- [ ] Task 7: build.sh 编译脚本
- [ ] Task 8: 测试流程
- [ ] Task 9: 分发包结构