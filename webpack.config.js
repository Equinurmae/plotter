const devCerts = require("office-addin-dev-certs");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const fs = require("fs");
const webpack = require("webpack");

const urlDev="https://localhost:3000/";
const urlProd="https://www.contoso.com/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const buildType = dev ? "dev" : "prod";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: "@babel/polyfill",
      metrics: "./src/metrics/metrics.js",
      commands: "./src/commands/commands.js",
      pacing: "./src/pacing/pacing.js",
      structure: "./src/structure/structure.js",
      pos: "./src/pos/pos.js",
      pov: "./src/pov/pov.js",
      sentiment: "./src/sentiment/sentiment.js"
    },
    resolve: {
      extensions: [".ts", ".tsx", ".html", ".js"]
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader", 
            options: {
              presets: ["@babel/preset-env"]
            }
          }
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader"
        },
        {
          test: /\.(png|jpg|jpeg|gif)$/,
          loader: "file-loader",
          options: {
            name: '[path][name].[ext]',          
          }
        }
      ]
    },
    plugins: [
      new CleanWebpackPlugin(),
      new HtmlWebpackPlugin({
        filename: "metrics.html",
        template: "./src/metrics/metrics.html",
        chunks: ["polyfill", "metrics"]
      }),
      new CopyWebpackPlugin({
        patterns: [
        {
          to: "metrics.css",
          from: "./src/metrics/metrics.css"
        },
        {
          to: "metrics_worker.js",
          from: "./src/metrics/metrics_worker.js"
        },
        {
          to: "[name]." + buildType + ".[ext]",
          from: "manifest*.xml",
          transform(content) {
            if (dev) {
              return content;
            } else {
              return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
            }
          }
        }
      ]}),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      }),
      new HtmlWebpackPlugin({
        filename: "pacing.html",
        template: "./src/pacing/pacing.html",
        chunks: ["polyfill", "pacing"]
      }),
      new CopyWebpackPlugin({
        patterns: [
        {
          to: "pacing.css",
          from: "./src/pacing/pacing.css"
        },
        {
          to: "pacing_worker.js",
          from: "./src/pacing/pacing_worker.js"
        },
        {
          to: "[name]." + buildType + ".[ext]",
          from: "manifest*.xml",
          transform(content) {
            if (dev) {
              return content;
            } else {
              return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
            }
          }
        }
      ]}),
      new HtmlWebpackPlugin({
        filename: "structure.html",
        template: "./src/structure/structure.html",
        chunks: ["polyfill", "structure"]
      }),
      new CopyWebpackPlugin({
        patterns: [
        {
          to: "structure.css",
          from: "./src/structure/structure.css"
        },
        {
          to: "structure_worker.js",
          from: "./src/structure/structure_worker.js"
        },
        {
          to: "[name]." + buildType + ".[ext]",
          from: "manifest*.xml",
          transform(content) {
            if (dev) {
              return content;
            } else {
              return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
            }
          }
        }
      ]}),
      new HtmlWebpackPlugin({
        filename: "pos.html",
        template: "./src/pos/pos.html",
        chunks: ["polyfill", "pos"]
      }),
      new CopyWebpackPlugin({
        patterns: [
        {
          to: "pos.css",
          from: "./src/pos/pos.css"
        },
        {
          to: "pos_worker.js",
          from: "./src/pos/pos_worker.js"
        },
        {
          to: "[name]." + buildType + ".[ext]",
          from: "manifest*.xml",
          transform(content) {
            if (dev) {
              return content;
            } else {
              return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
            }
          }
        }
      ]}),
      new HtmlWebpackPlugin({
        filename: "pov.html",
        template: "./src/pov/pov.html",
        chunks: ["polyfill", "pov"]
      }),
      new CopyWebpackPlugin({
        patterns: [
        {
          to: "pov.css",
          from: "./src/pov/pov.css"
        },
        {
          to: "pov_worker.js",
          from: "./src/pov/pov_worker.js"
        },
        {
          to: "[name]." + buildType + ".[ext]",
          from: "manifest*.xml",
          transform(content) {
            if (dev) {
              return content;
            } else {
              return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
            }
          }
        }
      ]}),
      new HtmlWebpackPlugin({
        filename: "sentiment.html",
        template: "./src/sentiment/sentiment.html",
        chunks: ["polyfill", "sentiment"]
      }),
      new CopyWebpackPlugin({
        patterns: [
        {
          to: "sentiment.css",
          from: "./src/sentiment/sentiment.css"
        },
        {
          to: "sentiment_worker.js",
          from: "./src/sentiment/sentiment_worker.js"
        },
        {
          to: "[name]." + buildType + ".[ext]",
          from: "manifest*.xml",
          transform(content) {
            if (dev) {
              return content;
            } else {
              return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
            }
          }
        }
      ]}),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*"
      },      
      https: (options.https !== undefined) ? options.https : await devCerts.getHttpsServerOptions(),
      port: process.env.npm_package_config_dev_server_port || 3000
    }
  };

  return config;
};
