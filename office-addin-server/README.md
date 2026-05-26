# Office Add-in Server

本地 HTTP 服务器，用于托管 Excel Add-in 静态文件。

## 使用方法

1. 运行 `build.sh` 编译二进制文件
2. 双击运行 `office-addin-server.exe`
3. 系统托盘会出现图标，右键菜单可：
   - Open - 在浏览器打开
   - Quit - 退出程序

## 配置

首次运行会在同目录生成 `config.json`：

```json
{
  "port": 3000,
  "dir": "dist",
  "autostart": false
}
```

- `port`: HTTP 服务端口
- `dir`: 静态文件目录
- `autostart`: 开机自启（Windows Registry）

## 开发

```bash
go run .
```