# My Office Add-in

一个用于 Excel 的 Microsoft Office 加载项，基于 Office JavaScript API 构建。

## 功能

在 Excel 功能区的「我的工具」选项卡下提供以下工具：

### 数据处理
| 按钮 | 功能 |
|------|------|
| 连接列 | 选中两列后点击，将数据用连接符合并到新列 |
| 字段计数 | 统计选中列每个字段值出现的次数 |
| 按列拆分 | 将数据按某一列的值拆分为多个新工作表 |

### 数据转换
| 按钮 | 功能 |
|------|------|
| 展开列 | 将宽表数据转换为长表格式，第一列作为键，其他列的值展开为独立行 |
| 填充序列 | 打开填充序列工具面板 |

### 查找与导入
| 按钮 | 功能 |
|------|------|
| 增强查找 | 在查找表中匹配一个值，返回该行指定列的数据，支持精确匹配和模糊匹配 |
| 导入CSV | 用 UTF-8 编码正确导入 CSV/文本文件，支持自定义分隔符 |

## 技术栈

- **框架**: Office JavaScript API
- **构建工具**: Webpack + Babel
- **兼容**: 目标 IE11

## 开发

### 前置条件

- Node.js
- Microsoft Excel 桌面版

### 安装依赖

```bash
npm install
```

### 启动开发服务器

```bash
npm run dev-server
```

### 运行加载项

```bash
npm start
```

### 其他命令

```bash
npm run build        # 生产构建
npm run build:dev    # 开发构建
npm run lint         # 检查 lint 问题
npm run lint:fix     # 自动修复 lint 问题
npm run validate     # 验证 manifest.xml
npm run test         # 运行测试
npm run watch        # 监听模式开发
```

## 生产部署

已配置 GitHub Actions 自动部署到 GitHub Pages：

- 部署地址: https://qixiasu.github.io/my-office-add-in/
- manifest.xml 中的资源 URL 在构建时自动替换为生产地址

## 项目结构

```
src/
  taskpane/           # 各工具面板 UI
    concat-taskpane.*    # 连接列
    count-values-taskpane.*  # 字段计数
    split-sheet-taskpane.*    # 按列拆分
    expand-taskpane.*    # 展开列
    fill-series-taskpane.*   # 填充序列
    vlookup-taskpane.*   # 增强查找
    csv-import-taskpane.*    # 导入CSV
  commands/           # 命令处理器
  utils/              # 工具函数
manifest.xml         # 加载项清单配置
webpack.config.js     # Webpack 配置
```

## 许可证

MIT