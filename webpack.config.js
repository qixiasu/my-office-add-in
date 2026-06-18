/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
const urlProd = "https://qixiasu.github.io/my-office-add-in/";

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      commands: "./src/commands/commands.js",
      "concat-taskpane": ["./src/taskpane/concat-taskpane.js", "./src/taskpane/concat-taskpane.html"],
      "csv-import-taskpane": ["./src/taskpane/csv-import-taskpane.js", "./src/taskpane/csv-import-taskpane.html"],
      "csv-merge-taskpane": ["./src/taskpane/csv-merge-taskpane.js", "./src/taskpane/csv-merge-taskpane.html"],
      "vlookup-taskpane": ["./src/taskpane/vlookup-taskpane.js", "./src/taskpane/vlookup-taskpane.html"],
      "fill-series-taskpane": ["./src/taskpane/fill-series-taskpane.js", "./src/taskpane/fill-series-taskpane.html"],
      "count-values-taskpane": ["./src/taskpane/count-values-taskpane.js", "./src/taskpane/count-values-taskpane.html"],
      "expand-taskpane": ["./src/taskpane/expand-taskpane.js", "./src/taskpane/expand-taskpane.html"],
      "split-sheet-taskpane": ["./src/taskpane/split-sheet-taskpane.js", "./src/taskpane/split-sheet-taskpane.html"],
      "sql-query-taskpane": [
        "./src/taskpane/sql-query-taskpane.js",
        "./src/taskpane/sql-query-taskpane.html",
      ],
      "ai-assistant-taskpane": [
        "./src/taskpane/ai-assistant-taskpane.js",
        "./src/taskpane/ai-assistant-taskpane.html",
      ],
      "data-cleaning-taskpane": [
        "./src/taskpane/data-cleaning-taskpane.js",
        "./src/taskpane/data-cleaning-taskpane.html",
      ],
      "merge-sheets-taskpane": [
        "./src/taskpane/merge-sheets-taskpane.js",
        "./src/taskpane/merge-sheets-taskpane.html",
      ],
      "cross-file-merge-taskpane": [
        "./src/taskpane/cross-file-merge-taskpane.js",
        "./src/taskpane/cross-file-merge-taskpane.html",
      ],
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
          {
            from: "src/commands/csv-import-dialog.html",
            to: "csv-import-dialog.html",
          },
          {
            from: "node_modules/sql.js/dist/sql-wasm-browser.wasm",
            to: "assets/sql-wasm-browser.wasm",
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
        scriptLoading: "blocking",
      }),
      new HtmlWebpackPlugin({
        filename: "concat-taskpane.html",
        template: "./src/taskpane/concat-taskpane.html",
        chunks: ["polyfill", "concat-taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "csv-import-taskpane.html",
        template: "./src/taskpane/csv-import-taskpane.html",
        chunks: ["polyfill", "csv-import-taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "csv-merge-taskpane.html",
        template: "./src/taskpane/csv-merge-taskpane.html",
        chunks: ["polyfill", "csv-merge-taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "vlookup-taskpane.html",
        template: "./src/taskpane/vlookup-taskpane.html",
        chunks: ["polyfill", "vlookup-taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "fill-series-taskpane.html",
        template: "./src/taskpane/fill-series-taskpane.html",
        chunks: ["polyfill", "fill-series-taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "count-values-taskpane.html",
        template: "./src/taskpane/count-values-taskpane.html",
        chunks: ["polyfill", "count-values-taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "expand-taskpane.html",
        template: "./src/taskpane/expand-taskpane.html",
        chunks: ["polyfill", "expand-taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "split-sheet-taskpane.html",
        template: "./src/taskpane/split-sheet-taskpane.html",
        chunks: ["polyfill", "split-sheet-taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "sql-query-taskpane.html",
        template: "./src/taskpane/sql-query-taskpane.html",
        chunks: ["polyfill", "sql-query-taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "ai-assistant-taskpane.html",
        template: "./src/taskpane/ai-assistant-taskpane.html",
        chunks: ["polyfill", "ai-assistant-taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "data-cleaning-taskpane.html",
        template: "./src/taskpane/data-cleaning-taskpane.html",
        chunks: ["polyfill", "data-cleaning-taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "merge-sheets-taskpane.html",
        template: "./src/taskpane/merge-sheets-taskpane.html",
        chunks: ["polyfill", "merge-sheets-taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "cross-file-merge-taskpane.html",
        template: "./src/taskpane/cross-file-merge-taskpane.html",
        chunks: ["polyfill", "cross-file-merge-taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "index.html",
        template: "./src/index.html",
        chunks: [],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
      static: [
        {
          directory: require("path").resolve(__dirname, "src/commands"),
          publicPath: "/",
          serveIndex: false,
        },
      ],
    },
  };

  return config;
};
