const path = require("path");
const webpack = require("webpack");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

// Keep in sync with WELL_KNOWN_DEV_CLIENT_IDS in src/auth.ts.
// Note: unlike the CLI, the add-in cannot actually authenticate with these
// — a browser flow needs a `spa` redirect URI on the app registration for
// both the redirect and the token endpoint's CORS header, and you can't
// add one to a Microsoft-owned app. They're listed to catch a misconfigured
// build, not as a usable fallback.
const WELL_KNOWN_DEV_CLIENT_IDS = [
  "2ad88395-b77d-4561-9441-d0e40824f9bc", // Microsoft PowerApps
  "51f81489-12ee-4a9e-aaae-a2591f45987d", // Microsoft Dynamics CRM (XRM Tooling)
];

module.exports = async (env, argv) => {
  const isProd = argv.mode === "production";
  const httpsOptions = isProd ? undefined : await devCerts.getHttpsServerOptions();

  const clientId = process.env.DATAVERSE_LOAD_CLIENT_ID || "e6828b0f-9fde-43f8-85d0-602660d498bb";

  // A production bundle must never ship with a borrowed Microsoft client id
  // (the consent screen would show "Microsoft PowerApps", and it's against
  // Microsoft's terms). Fail the build rather than relying on the runtime
  // banner. DVLOAD_ALLOW_DEV_CLIENT=1 is an explicit local escape hatch.
  if (isProd && WELL_KNOWN_DEV_CLIENT_IDS.includes(clientId) && process.env.DVLOAD_ALLOW_DEV_CLIENT !== "1") {
    throw new Error(
      "Refusing production build: DATAVERSE_LOAD_CLIENT_ID is unset or points at a " +
        "well-known Microsoft client id. Register your own Entra ID app " +
        "(see PRE-RELEASE-CHECKLIST.md) and set DATAVERSE_LOAD_CLIENT_ID, or set " +
        "DVLOAD_ALLOW_DEV_CLIENT=1 to build a local dev bundle anyway."
    );
  }

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
        // Defaults to the registered "dataverse-load" multi-tenant app.
        // Override with DATAVERSE_LOAD_CLIENT_ID at build time to ship
        // against a different registration.
        ADDIN_CLIENT_ID: JSON.stringify(clientId),
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
    devServer: {
      static: { directory: path.resolve(__dirname, "dist") },
      server: { type: "https", options: httpsOptions },
      port: 3000,
      hot: true,
    },
  };
};
