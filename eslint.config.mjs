import officeAddins from "eslint-plugin-office-addins";
import tsParser from "@typescript-eslint/parser";
import globals from "globals";

export default [
  ...officeAddins.configs.recommended,
  {
    languageOptions: {
      parser: tsParser,
      globals: {
        ...globals.browser,
        ...globals.commonjs,
        Office: "writable",
        Excel: "writable",
        Word: "writable",
      },
    },
    rules: {
      // 代码中用 /* global */ 注释声明，与 config globals 可能重复
      "no-redeclare": "off",
      // console.log 用于调试输出
      "no-console": "off",
    },
  },
  {
    files: ["**/*.test.js", "**/*.spec.js", "**/test/**"],
    rules: {
      "@typescript-eslint/no-unused-vars": "off",
      "no-undef": "off",
    },
  },
];
