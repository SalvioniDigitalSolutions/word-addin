const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

async function getHttpsOptions() {
  const certs = await devCerts.getHttpsServerOptions();
  return {
    key: certs.key,
    cert: certs.cert,
    ca: certs.ca,
  };
}

module.exports = async (env, argv) => {
  const isProd = argv.mode === "production";
  const httpsOptions = isProd ? false : await getHttpsOptions();

  return {
    entry: {
      taskpane: "./src/taskpane/taskpane.js",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true,
    },
    resolve: {
      extensions: [".js"],
    },
    module: {
      rules: [
        {
          test: /\.css$/i,
          use: ["style-loader", "css-loader"],
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: "./src/taskpane/taskpane.html",
        filename: "taskpane.html",
        chunks: ["taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [{ from: "assets", to: "assets", noErrorOnMissing: true }],
      }),
    ],
    devServer: {
      static: path.join(__dirname, "dist"),
      compress: true,
      port: 3000,
      server: {
        type: "https",
        options: httpsOptions,
      },
      headers: { "Access-Control-Allow-Origin": "*" },
      hot: true,
      /* Full-screen overlay shows handleError@Script error. in Word's webview with no useful stack. */
      client: {
        overlay: { errors: false, warnings: false, runtimeErrors: false },
      },
      proxy: [
        {
          context: ["/api"],
          target: "http://127.0.0.1:3548",
          changeOrigin: true,
        },
      ],
    },
  };
};
