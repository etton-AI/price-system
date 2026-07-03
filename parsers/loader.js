/**
 * 解析器加载器 —— 放在 parsers/ 目录避开 webpack 打包
 * 返回原始 Node.js module.createRequire，确保 xlsx 等依赖正确解析
 */
const path = require("path");
const { createRequire } = require("module");

const parsersDir = __dirname;
const parsersRequire = createRequire(path.join(parsersDir, "_index.js"));

/**
 * 加载并执行指定解析器
 * @param {string} fileName - 解析器文件名 (如 "tiantu_us.js")
 * @param {string} exportName - 导出函数名 (如 "parseTiantu")
 * @param {string} filePath - Excel 文件路径
 * @returns {Array} 解析结果
 */
function loadAndParse(fileName, exportName, filePath) {
  const mod = parsersRequire("./" + fileName);
  if (typeof mod[exportName] !== "function") {
    throw new Error(`${fileName} 未导出 ${exportName}`);
  }
  return mod[exportName](filePath);
}

module.exports = { loadAndParse, parsersRequire };
