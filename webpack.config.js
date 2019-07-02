"use strict";
const path = require("path");
const TerserPlugin = require("terser-webpack-plugin");
const BundleAnalyzerPlugin = require('webpack-bundle-analyzer').BundleAnalyzerPlugin;
const { allTokens } = require("fast-formula-parser/grammar/parsing");

// extract the names of the TokenTypes to avoid name mangling them.
const allTokenNames = allTokens.map(tokenType => tokenType.name);

module.exports = {
    mode: "production",
    entry: "./lib/XlsxPopulate.js",
    output: {
        path: path.resolve(__dirname, "./browser/"),
        filename: "xlsx-populate.min.js",
        library: "XlsxPopulate",
        libraryTarget: "umd",

        // https://github.com/webpack/webpack/issues/6784#issuecomment-375941431
        globalObject: "typeof self !== 'undefined' ? self : this"
    },
    optimization: {
        minimizer: [
            new TerserPlugin({
                sourceMap: true,
                parallel: true,
                terserOptions: {
                    compress: true,
                    mangle: {
                        // Avoid mangling TokenType names.
                        reserved: allTokenNames
                    }
                }
            })
        ]
    },
    plugins: [
        new BundleAnalyzerPlugin({
            analyzerMode: 'static',
            defaultSizes: 'parsed',
            reportFilename: 'webpack-report.html'
        })
    ],
    target: 'web',
    node: {
        fs: 'empty'
    },
    devtool: 'source-map'
};
