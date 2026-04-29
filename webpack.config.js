const fs = require('fs')
const path = require('path')
const webpack = require('webpack')
const HtmlWebpackPlugin = require('html-webpack-plugin')
const MiniCssExtractPlugin = require('mini-css-extract-plugin')
const CopyWebpackPlugin = require('copy-webpack-plugin')

// Locally-generated trusted dev cert (mkcert), produced by `npm run dev:setup`.
// Both files exist together or not at all. If absent, webpack-dev-server falls back to its
// default self-signed cert (only valid for localhost).
function readDevCert() {
  const certPath = path.resolve(__dirname, '.certs/dev.pem')
  const keyPath = path.resolve(__dirname, '.certs/dev-key.pem')
  if (!fs.existsSync(certPath) || !fs.existsSync(keyPath)) return null
  return {
    cert: fs.readFileSync(certPath),
    key: fs.readFileSync(keyPath),
  }
}

// In dev mode, REQUIRE `manifest.dev.xml` (produced by `npm run dev:setup` — the single
// source of truth for dev URL substitution). Fail loudly if it's missing so the build
// can't silently ship a manifest pointing at outlook.retyc.com to a local dev server.
// In prod mode, always use `manifest.xml` so the Docker entrypoint can substitute BASE_URL.
function pickManifestSource(mode) {
  if (mode === 'production') return 'manifest.xml'
  const devManifest = path.resolve(__dirname, 'manifest.dev.xml')
  if (!fs.existsSync(devManifest)) {
    throw new Error(
      'manifest.dev.xml is missing. Run `npm run dev:setup` first ' +
      '(override DEV_HOST=... if needed; defaults to 192.168.122.1).',
    )
  }
  return 'manifest.dev.xml'
}

/** @type {import('webpack').Configuration} */
module.exports = (env, argv) => {
  const devCert = readDevCert()
  const manifestSource = pickManifestSource(argv.mode)
  return {
  devtool: argv.mode === 'production' ? false : 'source-map',
  entry: {
    taskpane: './src/taskpane/taskpane.ts',
  },
  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: '[name].js',
    clean: true,
  },
  resolve: {
    extensions: ['.ts', '.js'],
    fallback: {
      buffer: require.resolve('buffer/'),
      stream: require.resolve('stream-browserify'),
      path: require.resolve('path-browserify'),
      fs: false,
      crypto: false,
      util: false,
      assert: false,
    },
    alias: {
      // The SDK's package.json declares index.js (ESM) and index.cjs — use CJS for stable browser bundling.
      '@retyc/sdk': path.resolve(__dirname, 'node_modules/@retyc/sdk/dist/index.cjs'),
      'node:buffer': require.resolve('buffer/'),
      'node:stream': require.resolve('stream-browserify'),
      'node:path': require.resolve('path-browserify'),
      'node:fs': false,
      'node:crypto': false,
      'node:util': false,
    },
  },
  module: {
    rules: [
      {
        test: /\.ts$/,
        use: [{ loader: 'ts-loader', options: { transpileOnly: true } }],
        exclude: /node_modules/,
      },
      {
        test: /\.css$/,
        use: [MiniCssExtractPlugin.loader, 'css-loader'],
      },
    ],
  },
  plugins: [
    new webpack.ProvidePlugin({
      Buffer: ['buffer', 'Buffer'],
      process: 'process/browser',
    }),
    new MiniCssExtractPlugin({ filename: '[name].css' }),
    new HtmlWebpackPlugin({
      template: './src/taskpane/taskpane.html',
      filename: 'taskpane.html',
      chunks: ['taskpane'],
      inject: 'body',
    }),
    new CopyWebpackPlugin({
      patterns: [
        { from: 'assets', to: 'assets' },
        // Source picked by `pickManifestSource()`:
        //   - dev mode + `manifest.dev.xml` present (after `npm run dev:setup`) → that one
        //   - otherwise → canonical `manifest.xml` (prod-ready, https://outlook.retyc.com)
        // Docker entrypoint substitutes BASE_URL into the prod manifest at container start.
        { from: manifestSource, to: 'manifest.xml' },
      ],
    }),
  ],
  devServer: {
    port: 3000,
    host: '0.0.0.0',
    allowedHosts: 'all',
    server: devCert ? { type: 'https', options: devCert } : 'https',
    headers: { 'Access-Control-Allow-Origin': '*' },
    historyApiFallback: {
      rewrites: [{ from: /^\/$/, to: '/taskpane.html' }],
    },
  },
  ignoreWarnings: [
    /Critical dependency/,
    /Module not found.*node:/,
  ],
  }
}
