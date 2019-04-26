const HtmlWebpackPlugin = require('html-webpack-plugin');
const path = require('path');

module.exports = {
    entry: {
        polyfill: 'babel-polyfill',
        app: './src/index.js',
        'function-file': './function-file/function-file.js'
    },
    output: {
        path: path.resolve(__dirname, "..")
    },
    module: {
        rules: [
            {
                test: /\.js$/,
                exclude: /node_modules/,
                use: 'babel-loader'
            },
            {
                test: /\.html$/,
                exclude: /node_modules/,
                use: 'html-loader'
            },
            {
                test: /\.(png|jpg|jpeg|gif)$/,
                use: 'file-loader'
            }
        ]
    },
    plugins: [
        new HtmlWebpackPlugin({
            template: './index.html',
            chunks: ['polyfill', 'app']
        }),
        new HtmlWebpackPlugin({
            template: './function-file/function-file.html',
            filename: 'function-file/function-file.html',
            chunks: ['function-file']
        })
    ]
};
