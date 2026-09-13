// Build for the Power Platform ToolBox (PPTB) tool.
//
// Same UI, different wrapper: the entry re-exports the add-in's taskpane
// (host.ts detects the ToolBox bridge at runtime), and the page template is
// the add-in's taskpane.html with the office.js CDN <script> stripped —
// PPTB's CSP (script-src 'self') would block it, and there is no Office host
// to wait for anyway.
//
// The output in dist/ is the complete publishable tool package: index.html,
// the bundle, the icon, and a package.json carrying the PPTB manifest fields
// (displayName, icon, main, features.minAPI) from tool.package.json. Publish
// by running `npm publish --access public` FROM dist/ — see README.md.
const path = require("path");
const fs = require("fs");
const webpack = require("webpack");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

const ADDIN = path.resolve(__dirname, "../addin");

/** The add-in's page, minus the office.js CDN script tag. */
function pptbTemplate() {
  const html = fs.readFileSync(path.join(ADDIN, "src/taskpane/taskpane.html"), "utf8");
  const stripped = html.replace(/^\s*<script[^>]*appsforoffice[^>]*><\/script>\s*$/m, "");
  if (stripped === html) {
    throw new Error(
      "Expected to strip the office.js <script> tag from taskpane.html and didn't — check the template."
    );
  }
  return stripped.replace(/<title>[^<]*<\/title>/, "<title>dvload — Excel to Dataverse</title>");
}

module.exports = (env, argv) => {
  const isProd = argv.mode === "production";

  return {
    devtool: isProd ? false : "source-map",
    entry: { taskpane: "./src/index.ts" },
    output: {
      filename: "[name].js",
      path: path.resolve(__dirname, "dist"),
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js"],
      extensionAlias: { ".js": [".ts", ".js"] },
      // Same stub as the add-in build: core's readTableFromFile lazily
      // imports node:fs/promises, which a browser bundle never reaches.
      fallback: { "fs/promises": false },
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          loader: "ts-loader",
          exclude: /node_modules/,
          // The entry pulls sources from ../addin, so point ts-loader at one
          // tsconfig that covers both trees.
          options: { configFile: path.resolve(__dirname, "tsconfig.json") },
        },
        { test: /\.css$/, use: ["style-loader", "css-loader"] },
      ],
    },
    plugins: [
      new webpack.NormalModuleReplacementPlugin(/^node:/, (resource) => {
        resource.request = resource.request.replace(/^node:/, "");
      }),
      new HtmlWebpackPlugin({
        filename: "index.html",
        templateContent: pptbTemplate(),
        chunks: ["taskpane"],
      }),
      new webpack.DefinePlugin({
        // Telemetry stays off in the ToolBox build: PPTB's CSP has no
        // exception for the ingestion endpoint, so a baked-in connection
        // string would only produce console noise.
        ADDIN_AI_CONNECTION: JSON.stringify(""),
      }),
      new CopyWebpackPlugin({
        patterns: [
          // Icons resolve under dist/, which is where `icon` in the manifest
          // points. The manifest and README are NOT copied here: they belong
          // beside dist/ in the assembled package, not inside it — see
          // scripts/assemble-package.mjs.
          { from: "assets/icon.svg", to: "icons/icon.svg" },
        ],
      }),
    ],
  };
};
