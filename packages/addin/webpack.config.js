const path = require("path");
const webpack = require("webpack");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

module.exports = async (env, argv) => {
  const isProd = argv.mode === "production";
  const httpsOptions = isProd ? undefined : await devCerts.getHttpsServerOptions();

  return {
    devtool: isProd ? false : "source-map",
    entry: {
      taskpane: "./src/taskpane/taskpane.ts",
      commands: "./src/commands/commands.ts",
    },
    output: {
      filename: "[name].js",
      path: path.resolve(__dirname, "dist"),
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js"],
      extensionAlias: { ".js": [".ts", ".js"] },
    },
    module: {
      rules: [
        { test: /\.tsx?$/, loader: "ts-loader", exclude: /node_modules/ },
        { test: /\.css$/, use: ["style-loader", "css-loader"] },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["commands"],
      }),
      new webpack.DefinePlugin({
        // Fallback MUST be one of the WELL_KNOWN_DEV_CLIENT_IDS in src/auth.ts
        // so the dev-mode banner fires. A private app id here would silently
        // bypass the entire pre-release safety net.
        ADDIN_CLIENT_ID: JSON.stringify(
          process.env.DATAVERSE_LOAD_CLIENT_ID || "2ad88395-b77d-4561-9441-d0e40824f9bc"
        ),
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "manifest.xml", to: "manifest.xml" },
          { from: "assets", to: "assets", noErrorOnMissing: true },
        ],
      }),
    ],
    devServer: {
      static: { directory: path.resolve(__dirname, "dist") },
      server: { type: "https", options: httpsOptions },
      port: 3000,
      hot: true,
    },
  };
};
