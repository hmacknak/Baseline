/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const { DefinePlugin } = require("webpack");

const { execSync } = require("child_process");

const urlDev = "https://localhost:3000/";

/**
 * Short commit id of this build. It stamps the download page links, is baked into the add-in, and is
 * published as version.json so a running add-in can tell when a newer build is live.
 */
function buildVersion() {
  let sha = (process.env.GITHUB_SHA || "").slice(0, 7);
  if (!sha) {
    try {
      sha = execSync("git rev-parse --short HEAD", { stdio: ["ignore", "pipe", "ignore"] }).toString().trim();
    } catch {
      sha = "dev";
    }
  }
  return { sha, label: `${new Date().toISOString().slice(0, 10)} (${sha})` };
}
const urlProd = "https://hmacknak.github.io/Baseline/"; // GitHub Pages, published by .github/workflows/pages.yml

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const version = dev ? { sha: "dev", label: "development" } : buildVersion();
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
    },
    output: {
      clean: true,
      // Content-hashed names so a published update can never be served from a stale cache.
      filename: dev ? "[name].js" : "[name].[contenthash:8].js",
    },
    resolve: {
      extensions: [".ts", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader"
          },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.css$/,
          type: "asset/resource",
          generator: {
            filename: dev ? "[name][ext]" : "[name].[contenthash:8][ext]",
          },
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
      new DefinePlugin({
        __BUILD_VERSION__: JSON.stringify(version.sha),
        __BUILD_LABEL__: JSON.stringify(version.label),
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "site/index.html",
            to: "index.html",
            transform(content) {
              return content
                .toString()
                .replace(/__VERSION_LABEL__/g, version.label)
                .replace(/__VERSION__/g, version.sha);
            },
          },
          {
            from: "demo/Baseline-demo.xlsx",
            to: "Baseline-demo.xlsx",
          },
          {
            from: "demo/Baseline-demo-video.mp4",
            to: "Baseline-demo-video.mp4",
          },
          {
            from: "site/version.json",
            to: "version.json",
            transform(content) {
              return content
                .toString()
                .replace(/__VERSION_LABEL__/g, version.label)
                .replace(/__VERSION__/g, version.sha);
            },
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
        ],
      }),
    ],
    performance: {
      // The demo video and workbook are downloads, not part of the add-in bundle.
      assetFilter: name => !/\.(mp4|xlsx|map)$/.test(name),
    },
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
