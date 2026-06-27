# GitHub Pages 部署改造设计

**日期：** 2026-05-31

## 目标

将项目从本地 Go HTTP 服务器迁移到 GitHub Pages，利用 GitHub Actions 实现自动构建和部署。

## 现状分析

- **前端构建**：webpack 打包，输出到 `dist/` 目录
- **本地服务**：Go 服务器（`office-addin-server/`）提供 HTTPS 静态文件服务
- **Manifest 配置**：
  - 开发环境：`https://localhost:3000/`
  - 生产环境：`https://www.contoso.com/`
- **部署方式**：无自动部署

## 目标状态

- **托管服务**：GitHub Pages（默认域名 `https://qixiasu.github.io/my-office-add-in/`）
- **自动化**：推送 main 分支 → 自动构建 → 部署到 `gh-pages` 分支
- **删除**：Go 服务器（不再需要）

## 实施方案

### 1. 删除 Go 服务器

删除 `office-addin-server/` 目录。

### 2. 修改 webpack.config.js

将生产 URL 从 `https://www.contoso.com/` 改为：

```javascript
const urlProd = "https://qixiasu.github.io/my-office-add-in/";
```

### 3. 修改 manifest.xml

- `<SupportUrl>` 改为 `https://qixiasu.github.io/my-office-add-in/`
- `<AppDomain>` 配置根据需要调整

### 4. 新增 GitHub Actions workflow

创建 `.github/workflows/deploy.yml`：

```yaml
name: Deploy to GitHub Pages

on:
  push:
    branches:
      - main

jobs:
  deploy:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4

      - name: Setup Node.js
        uses: actions/setup-node@v4
        with:
          node-version: '20'
          cache: 'npm'

      - name: Install dependencies
        run: npm ci

      - name: Build
        run: npm run build

      - name: Deploy to GitHub Pages
        uses: peaceiris/actions-gh-pages@v3
        with:
          github_token: ${{ secrets.GITHUB_TOKEN }}
          publish_dir: ./dist
          publish_branch: gh-pages
```

### 5. 额外配置

创建 `.nojekyll` 文件（空白文件），防止 GitHub Pages 使用 Jekyll 处理目录。

## 文件变更清单

| 文件/目录 | 操作 |
|-----------|------|
| `office-addin-server/` | 删除 |
| `.github/workflows/deploy.yml` | 新增 |
| `.nojekyll` | 新增 |
| `webpack.config.js` | 修改 urlProd |
| `manifest.xml` | 修改 SupportUrl 和 AppDomain |

## 部署流程

1. 开发者推送代码到 `main` 分支
2. GitHub Actions 自动触发：
   - 安装 Node.js 依赖
   - 执行 `npm run build`（生产构建，webpack 会把 manifest 中的 localhost 替换为 `https://qixiasu.github.io/my-office-add-in/`）
   - 将 `dist/` 目录内容 force push 到 `gh-pages` 分支
3. GitHub Pages 自动从 `gh-pages` 分支提供服务

## 注意事项

- GitHub Pages 默认启用 HTTPS
- Office Add-in 要求 HTTPS，GitHub Pages 满足此要求
- `dist/` 目录会被 GitHub Actions force push 到 `gh-pages`，本地无需手动管理该分支