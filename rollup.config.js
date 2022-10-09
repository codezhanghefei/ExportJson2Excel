export default {
  // 入口文件
  input: 'src/main.js',
  output: {
    // 打包名称
    name: 'index',
    // 打包的文件
    file: 'dist/index.js',
    // 打包的格式，umd 支持 commonjs/amd/life 三种方式
    format: 'umd',
  },
}