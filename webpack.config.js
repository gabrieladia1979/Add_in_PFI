const path = require("path");
const webpack = require("webpack");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
module.exports = async (_env, argv) => {
  const isDev = argv.mode === "development";

  return {
    entry: {
      taskpane: path.resolve(__dirname, "src/taskpane/taskpane.tsx"),
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      clean: true,
      filename: "[name].bundle.js",
      publicPath: "/",
    },
    devtool: isDev ? "source-map" : false,
    resolve: {
      extensions: [".ts", ".tsx", ".js"],
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: "ts-loader",
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader", "postcss-loader"],
        },
        {
          test: /\.(png|jpg|jpeg|gif|svg)$/i,
          type: "asset/resource",
        },
      ],
    },
    plugins: [
      new webpack.DefinePlugin({
        "process.env.API_BASE_URL": JSON.stringify(
          process.env.API_BASE_URL || "http://127.0.0.1:8001"
        ),
        "process.env.API_KEY": JSON.stringify(
          process.env.API_KEY || "PHISHARG_DEV_KEY_123"
        ),
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: path.resolve(__dirname, "src/taskpane/taskpane.html"),
        chunks: ["taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "manifest.xml", to: "manifest.xml" },
          { from: "assets", to: "assets" },
        ],
      }),
    ],
    devServer: {
      port: 3000,
      hot: true,
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      https: true,
      proxy: [
        {
          context: ["/api"],
          target: "http://127.0.0.1:8001",
          changeOrigin: true,
        },
      ],
    },
  };
};
