# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Microsoft Office Add-in for Excel, built with the Office JavaScript API. Provides a taskpane UI and ribbon button commands that operate on Excel worksheet data. Pure JavaScript project with webpack bundling and Babel transpilation (targeting IE11).

## Commands

```bash
# Development
npm run dev-server    # Start webpack dev server (https://localhost:3000)
npm run start         # Launch Excel desktop and sideload the add-in
npm run stop          # Stop debugging session
npm run watch         # Webpack watch mode for development

# Build
npm run build         # Production build (replaces localhost URLs with production domain)
npm run build:dev     # Development build

# Lint & Format
npm run lint          # Check for lint issues
npm run lint:fix      # Auto-fix lint issues
npm run prettier      # Run prettier formatting

# Validate
npm run validate      # Validate manifest.xml
```

To debug in VS Code: F5 → select "Excel Desktop (Edge Chromium)". This runs the "Debug: Excel Desktop" pre-launch task which starts Excel with the add-in sideloaded, then attaches the Edge debugger on port 9229.

## Architecture

```
manifest.xml          # Add-in manifest — defines ribbon buttons, permissions, resource URLs
webpack.config.js     # Two entry points: taskpane + commands
src/
  taskpane/
    taskpane.html     # Taskpane UI (Fluent UI styled panel)
    taskpane.js       # Taskpane logic — Office.onReady → binds button → Excel.run
    taskpane.css      # Styles (Microsoft Fabric design)
  commands/
    commands.html     # Commands host page (empty shell, just loads office.js)
    commands.js       # ExecuteFunction handlers registered via Office.actions.associate
assets/               # Icons referenced in manifest
```

**Two execution paths:**

1. **Taskpane** (`ShowTaskpane` action): Opens a side panel in Excel. The HTML page loads `office.js`, `taskpane.js` runs on `Office.onReady`, then binds click handlers. All Excel operations happen inside `Excel.run` callbacks which manage the request context.

2. **Commands** (`ExecuteFunction` action): Ribbon button triggers a function registered via `Office.actions.associate("name", handler)`. These run in a hidden browser process — `console.log` output goes to VS Code Debug Console. The handler receives an `event` object and **must** call `event.completed()` when done, otherwise the add-in hangs.

**Key webpack details:**
- `HtmlWebpackPlugin` generates two HTML files: `taskpane.html` and `commands.html`, each with their respective chunks
- `CopyWebpackPlugin` copies `assets/*` and `manifest*.xml` to output, replacing `localhost:3000` URLs with production URLs in production mode
- Dev server uses HTTPS with self-signed certs (required by Office)
- Polyfill chunk (core-js + regenerator-runtime) is included in both pages for IE11 support

**manifest.xml key points:**
- Host: `Workbook` (Excel only, not Word/PowerPoint/Outlook)
- Permissions: `ReadWriteDocument`
- Two ribbon buttons in the Home tab under "Commands Group":
  - `TaskpaneButton` → `ShowTaskpane` → opens `taskpane.html`
  - `ConcatenateButton` → `ExecuteFunction` → calls `concatenateColumns` function
- All resource URLs point to `https://localhost:3000/` in dev; replaced at production build time

## Reference

The official Excel JavaScript API reference is at `javascript-api-office-js-docs-reference-excel-js-preview.pdf` in the project root.

## Key APIs in Use

- `Office.onReady()` — initialization entry point
- `Office.actions.associate()` — register ribbon button command handlers
- `Excel.run()` — required wrapper for all Excel object model operations
- `context.workbook.getSelectedRange()` — get user's current selection
- `worksheet.getRange().insert()` — insert cells/columns with shift direction
- `worksheet.getRange().values` — read/write range values as 2D arrays
