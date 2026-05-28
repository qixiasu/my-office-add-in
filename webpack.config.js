/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
const urlProd = "https://www.contoso.com/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

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
      "vlookup-taskpane": ["./src/taskpane/vlookup-taskpane.js", "./src/taskpane/vlookup-taskpane.html"],
      "fill-series-taskpane": ["./src/taskpane/fill-series-taskpane.js", "./src/taskpane/fill-series-taskpane.html"],
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
        filename: "vlookup-taskpane.html",
        template: "./src/taskpane/vlookup-taskpane.html",
        chunks: ["polyfill", "vlookup-taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "fill-series-taskpane.html",
        template: "./src/taskpane/fill-series-taskpane.html",
        chunks: ["polyfill", "fill-series-taskpane"],
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
