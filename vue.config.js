//头部引用
const NodePolyfillPlugin = require('node-polyfill-webpack-plugin')
module.exports = {
  // publicPath: '/admin',
  // outputDir: 'dist/admin',
  // assetsDir: 'static',
  lintOnSave: false,
  configureWebpack: {
    externals: {
      fs: require('fs')
    },
    plugins: [new NodePolyfillPlugin()]
  }
}

