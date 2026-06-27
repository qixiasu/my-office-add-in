# GitHub Pages 部署改造设计

**日期：** 2026-05-31

**更新：** 2026-06-27（迁移至 Gitee Pages）

## 目标

~~将项目从本地 Go HTTP 服务器迁移到 GitHub Pages~~ → 已完成 GitHub Pages 部署

**当前目标：** 将托管服务从 GitHub Pages 迁移到 Gitee Pages，利用 Gitee Pages 实现自动构建和部署（国内访问更稳定）。

## 现状分析

- **前端构建**：webpack 打包，输出到 `dist/` 目录
- **托管服务**：GitHub Pages（`https://qixiasu.github.io/my-office-add-in/`）
- **Manifest 配置**：
  - 开发环境：`https://localhost:3000/`
  - 生产环境：`https://qixiasu.github.io/my-office-add-in/`
- **部署方式**：GitHub Actions 自动部署

## 目标状态

- **托管服务**：Gitee Pages（`https://qinxuan1989.gitee.io/my-office-add-in/`）
- **自动化**：推送 main 分支 → 自动构建 → 部署到 Gitee Pages
- **删除**：GitHub Actions workflow（不再需要）

## 实施方案

### 1. 修改 webpack.config.js

将生产 URL 改为：

```javascript
const urlProd = "https://qinxuan1989.gitee.io/my-office-add-in/";
```

### 2. 修改 manifest.xml

- `<SupportUrl>` 改为 `https://qinxuan1989.gitee.io/my-office-add-in/`
- `<AppDomain>` 改为 `https://qinxuan1989.gitee.io`

### 3. 删除 GitHub Actions workflow

删除 `.github/workflows/deploy.yml`。

### 4. Gitee Pages 配置

在 Gitee 仓库设置中启用 Gitee Pages：
- 仓库 → 设置 → Gitee Pages
- 选择分支（通常为 `gh-pages` 或 `master`）
- 勾选"强制使用 HTTPS"
- 点击启动

### 5. 更新文档

- README.md：更新部署地址
- CLAUDE.md：更新生产 URL 注释

## 文件变更清单

| 文件/目录 | 操作 |
|-----------|------|
| `.github/workflows/deploy.yml` | 删除 |
| `webpack.config.js` | 修改 urlProd |
| `manifest.xml` | 修改 SupportUrl 和 AppDomain |
| `README.md` | 更新部署地址 |
| `CLAUDE.md` | 更新生产 URL 注释 |

## 部署流程

1. 开发者推送代码到 `main` 分支
2. 手动执行 `npm run build` 生成本地 `dist/` 目录
3. 将 `dist/` 目录内容推送到 Gitee 仓库的 `gh-pages` 分支（或 master 分支，取决于 Gitee Pages 配置）
4. Gitee Pages 自动从对应分支提供服务

**注意：** Gitee 不支持像 GitHub 那样的 Actions 自动部署，需要手动构建推送或使用第三方 CI（如 Jenkins）。

## 注意事项

- Gitee Pages 默认启用 HTTPS
- Office Add-in 要求 HTTPS，Gitee Pages 满足此要求
- Gitee Pages 需要手动部署，不支持 GitHub Actions
- 考虑后续使用 Gitee Go 或其他 CI 工具实现自动化