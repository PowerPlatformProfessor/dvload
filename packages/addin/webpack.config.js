const path = require("path");
const webpack = require("webpack");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

// There is no client id in this build any more.
//
// The bundle used to bake one in for @azure/msal-browser, along with a guard
// that failed the production build if it was a borrowed Microsoft app id.
// Both are gone: the UI no longer authenticates. It asks the local sidecar
// (`dvload serve`) for tokens over a same-origin fetch, and the sidecar uses
// the CLI's auth stack — which *can* borrow Microsoft's pre-consented
// Dataverse client, because a native client redirects to http://localhost
// and never meets CORS. See packages/addin/src/auth.ts.
//
// Consequence for packaging: this output is served from loopback by the
// CLI, not from a public origin, so there is nothing host-specific to
// configure at build time.
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
      // @dvload/core's readTableFromFile dynamically imports node:fs/promises
      // for CSV sources. The add-in only ever uses the buffer-based reader,
      // so stub the Node module out (false → empty module).
      fallback: { "fs/promises": false },
    },
    module: {
      rules: [
        { test: /\.tsx?$/, loader: "ts-loader", exclude: /node_modules/ },
        { test: /\.css$/, use: ["style-loader", "css-loader"] },
      ],
    },
    plugins: [
      // Webpack 5 doesn't resolve the "node:" scheme; strip the prefix so
      // the resolve.fallback above can take over.
      new webpack.NormalModuleReplacementPlugin(/^node:/, (resource) => {
        resource.request = resource.request.replace(/^node:/, "");
      }),
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
        // Application Insights connection string; empty = telemetry fully off.
        ADDIN_AI_CONNECTION: JSON.stringify(process.env.DVLOAD_AI_CONNECTION_STRING || ""),
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "manifest.xml", to: "manifest.xml" },
          { from: "assets", to: "assets", noErrorOnMissing: true },
        ],
      }),
    ],
    // Kept for UI-only work (HMR while restyling). Note that sign-in and
    // anything touching Dataverse will NOT work here: /api/* is served by
    // `dvload serve`, not by this dev server. For a working end-to-end loop
    // run `webpack --watch` and point `dvload serve` at dist/ instead.
    devServer: {
      static: { directory: path.resolve(__dirname, "dist") },
      server: { type: "https", options: httpsOptions },
      port: 3000,
      hot: true,
    },
  };
};
