// src/utils/cm6-sql-editor.js

var { EditorState, Compartment } = require("@codemirror/state");
var { EditorView, placeholder, keymap } = require("@codemirror/view");
var { syntaxHighlighting, HighlightStyle } = require("@codemirror/language");
var { sql } = require("@codemirror/lang-sql");
var { tags } = require("@lezer/highlight");

/**
 * Office 主题：浅色风格，匹配 Fabric UI
 */
var sqlHighlightStyle = HighlightStyle.define([
  { tag: tags.keyword, color: "#0078d4", fontWeight: "600" },
  { tag: tags.string, color: "#d32f2f" },
  { tag: tags.number, color: "#098658" },
  { tag: tags.comment, color: "#808080", fontStyle: "italic" },
  { tag: tags.operator, color: "#333333" },
  { tag: tags.variableName, color: "#333333" },
]);

var sqlTheme = EditorView.theme({
  "&": {
    backgroundColor: "#ffffff",
    fontSize: "12px",
    fontFamily: "'Consolas', 'Courier New', monospace",
  },
  ".cm-cursor": {
    borderLeftColor: "#0078d4",
  },
  "&.cm-focused .cm-selectionBackground, .cm-selectionBackground": {
    backgroundColor: "#0078d420",
  },
  ".cm-activeLine": {
    backgroundColor: "#f5f5f5",
  },
  ".cm-placeholder": {
    color: "#aaa",
    fontSize: "12px",
    fontStyle: "normal",
  },
  ".cm-gutters": {
    backgroundColor: "#fafafa",
    borderRight: "1px solid #e8e8e8",
    color: "#aaa",
  },
  ".cm-foldPlaceholder": {
    backgroundColor: "#f0f0f0",
    border: "none",
    color: "#999",
  },
});

/**
 * 创建 SQL 编辑器实例
 *
 * @param {HTMLElement} container - 挂载的 DOM 元素
 * @param {object} options
 * @param {function(): object} options.getSchema - 返回 { tableName: ["col1", "col2"] } 格式的 schema
 * @returns {object} 编辑器控制器
 *   - getValue(): string
 *   - setValue(sql: string): void
 *   - updateSchema(): void
 *   - onExecute(fn: function): void
 *   - destroy(): void
 */
function createSqlEditor(container, options) {
  if (!container || typeof container.appendChild !== "function") {
    throw new Error("createSqlEditor: container must be a valid HTMLElement");
  }
  options = options || {};
  var getSchema =
    options.getSchema ||
    function () {
      return {};
    };

  var schemaCompartment = new Compartment();
  var onExecute = null;

  // Ctrl/Cmd + Enter 执行查询
  var executeKeybinding = keymap.of([
    {
      key: "Ctrl-Enter",
      mac: "Cmd-Enter",
      run: function () {
        if (onExecute) {
          onExecute();
          return true;
        }
        return false;
      },
    },
  ]);

  var extensions = [
    // 基础功能集（括号匹配、折叠、历史、行号、选中等）
    require("codemirror").basicSetup,
    // Compartment 占位，用于热替换 schema
    schemaCompartment.of(sql({ schema: getSchema() })),
    // 语法高亮主题
    syntaxHighlighting(sqlHighlightStyle),
    // 编辑器主题
    sqlTheme,
    // 快捷键
    executeKeybinding,
    // Placeholder
    placeholder("输入 SQL 语句\n\n例如: SELECT * FROM 表名\n或: SELECT COUNT(*) FROM 表名"),
  ];

  var view = new EditorView({
    state: EditorState.create({
      doc: "",
      extensions: extensions,
    }),
    parent: container,
  });

  return {
    /** 获取当前编辑器中的 SQL */
    getValue: function () {
      return view.state.doc.toString();
    },

    /** 设置编辑器内容（自动高亮） */
    setValue: function (sql) {
      var doc = view.state.doc;
      view.dispatch({
        changes: { from: 0, to: doc.length, insert: sql || "" },
      });
    },

    /** 热替换自动补全 Schema（导入新表后调用） */
    updateSchema: function () {
      view.dispatch({
        effects: schemaCompartment.reconfigure(sql({ schema: getSchema() })),
      });
    },

    /** 注册 Ctrl+Enter 回调 */
    onExecute: function (fn) {
      onExecute = fn;
    },

    /** 销毁编辑器，释放资源 */
    destroy: function () {
      view.destroy();
    },
  };
}

module.exports = { createSqlEditor };
