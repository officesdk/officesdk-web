const path = require('path');
const HtmlWebpackPlugin = require('html-webpack-plugin');
const TsconfigPathsPlugin = require('tsconfig-paths-webpack-plugin');

module.exports = {
  mode: 'development',
  entry: {
    index: './pages/index/index.ts',

    // Test connection cases.
    rpc: './pages/rpc/rpc.ts',
    rpcServer: './pages/rpc/server.ts',

    // Test document cases.
    document: './pages/document/document.ts',
    documentServer: './pages/document/server.ts',

    // Test spreadsheet cases.
    spreadsheet: './pages/spreadsheet/spreadsheet.ts',
    spreadsheetServer: './pages/spreadsheet/server.ts',

    // Test pdf cases.
    pdf: './pages/pdf/pdf.ts',
    pdfServer: './pages/pdf/server.ts',

    // Test presentation cases.
    presentation: './pages/presentation/presentation.ts',
    presentationServer: './pages/presentation/server.ts',
  },
  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: '[name].bundle.js',
    chunkFilename: '[name].chunk.js',
    clean: true,
  },
  module: {
    rules: [
      {
        test: /\.js$/,
        exclude: /node_modules/,
        use: {
          loader: 'babel-loader',
          options: { cacheDirectory: true, cacheCompression: false },
        },
      },
      {
        test: /\.tsx?$/,
        use: { loader: 'babel-loader', options: { cacheDirectory: true, cacheCompression: false } },
      },
      {
        test: /\.css$/,
        use: ['style-loader', 'css-loader'],
      },
    ],
  },
  plugins: [
    new HtmlWebpackPlugin({
      chunks: ['index'],
      template: './pages/index/index.html',
    }),
    new HtmlWebpackPlugin({
      filename: 'rpc.html',
      chunks: ['rpc'],
      template: './pages/rpc/rpc.html',
    }),
    new HtmlWebpackPlugin({
      filename: 'template.html',
      chunks: [],
      template: './pages/shared/template.html',
    }),
    new HtmlWebpackPlugin({
      filename: 'document.html',
      chunks: ['document'],
      template: './pages/document/document.html',
    }),
    new HtmlWebpackPlugin({
      filename: 'spreadsheet.html',
      chunks: ['spreadsheet'],
      template: './pages/spreadsheet/spreadsheet.html',
    }),
    new HtmlWebpackPlugin({
      filename: 'pdf.html',
      chunks: ['pdf'],
      template: './pages/pdf/pdf.html',
    }),
    new HtmlWebpackPlugin({
      filename: 'presentation.html',
      chunks: ['presentation'],
      template: './pages/presentation/presentation.html',
    }),
  ],
  devServer: {
    static: [
      {
        directory: path.join(__dirname, 'dist'),
      },
      {
        directory: path.join(__dirname, 'static'),
        publicPath: '/static',
      },
    ],
    compress: true,
    port: 3344,
    hot: 'only',
    open: true,
  },
  resolve: {
    extensions: ['.js', '.ts', '.json'],
    plugins: [new TsconfigPathsPlugin()],
  },
  devtool: 'source-map',
};
