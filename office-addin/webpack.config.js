/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");

const urlDev = "https://localhost:3000/";
/** Office on the web / SharePoint / Microsoft 365 hosts allowed credentialed CORS to the dev server. */
const M365_ORIGIN = /^https:\/\/([\w-]+\.)*(officeapps\.live\.com|office\.com|office\.net|sharepoint\.com|cloud\.microsoft)$/;
// Hosting origin for production builds; set by scripts/Deploy-AdvAddinWeb.ps1.
const urlProd = process.env.ADV_ADDIN_URL_PROD || "https://www.contoso.com/";
// Dev sideload and the centrally deployed add-in must not share an Id, or Office treats them as one add-in.
const MANIFEST_ID_DEV = "34ca07a9-c204-475a-a213-694399a38430";
const MANIFEST_ID_PROD = "852d67be-2456-421f-a280-251aaa88bb63";
const DISPLAY_NAME_DEV = "Global Advocate (ADVTest dev)";
const DISPLAY_NAME_PROD = "Global Advocate";

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
      react: ["react", "react-dom"],
      taskpane: {
        import: ["./src/taskpane/index.tsx", "./src/taskpane/taskpane.html"],
        dependOn: "react",
      },
    },
    output: {
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader",
          },
        },
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: ["ts-loader"],
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        {
          test: /\.(png|jpg|jpeg|ttf|woff|woff2|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "react"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.{json,xml}",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                // Without trailing slash so AppDomain entries are rewritten too.
                return content
                  .toString()
                  .replaceAll(urlDev.replace(/\/$/, ""), urlProd.replace(/\/$/, ""))
                  .replaceAll(MANIFEST_ID_DEV, MANIFEST_ID_PROD)
                  .replaceAll(DISPLAY_NAME_DEV, DISPLAY_NAME_PROD);
              }
            },
          },
        ],
      }),
      new webpack.ProvidePlugin({
        Promise: ["es6-promise", "Promise"],
      }),
    ],
    // The manifest* copy glob watches the project root; E2E output/session writes would rebuild + reload the pane mid-test.
    watchOptions: {
      ignored: ["**/node_modules/**", "**/e2e/**", "**/.auth/**", "**/test-results/**", "**/dist/**", "**/scripts/**"],
    },
    devServer: {
      hot: true,
      // Explicit, else the cross-origin check 403s the ribbon's no-cors <img> icon loads from officeapps.live.com.
      allowedHosts: ["localhost"],
      // Office on the web fetches ribbon icons with credentials, which rejects "*": echo Microsoft 365 origins only.
      headers: (req) => {
        const origin = req.headers.origin;
        return origin && M365_ORIGIN.test(origin)
          ? { "Access-Control-Allow-Origin": origin, "Access-Control-Allow-Credentials": "true", Vary: "Origin" }
          : { "Access-Control-Allow-Origin": "*" };
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
