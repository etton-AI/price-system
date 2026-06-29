#!/usr/bin/env node
/**
 * build_db.js — 统一构建入口
 *
 * 自动识别比价工具文件夹中所有 Excel 文件，调用对应解析器，
 * 输出统一结构的 prices.json 到 data/ 目录。
 *
 * 用法: node price_db/parsers/build_db.js
 */

const path = require("path");
const fs = require("fs");
const { parseETTON } = require("./etton_us");
const { parseTiantu } = require("./tiantu_us");
const { parseTiantuUK } = require("./tiantu_uk");
const { parseYingmei } = require("./yingmei_us");
const { parseHaohui } = require("./haohui_us");
const { parseHaopeng } = require("./haopeng_us");
const { parseXinglian } = require("./xinglian_us");
const { parseXinyi } = require("./xinyi_eu");
const { parseHangle } = require("./hangle");
const { parseXinyun } = require("./fengyun");
const { parseHuaweier } = require("./huaweier");
const { parseTiantuAir } = require("./tiantu_air");
const { parseKaixin } = require("./kaixin");
const { parseXinsheng } = require("./xinsheng");

// ── 供应商识别规则 ──
function identifySupplier(fileName) {
  const n = fileName.toLowerCase();
  // 跳过非价格表文件（船期表、出运计划等）
  if (n.includes("出运计划") || n.includes("船期") || n.includes("schedule")) return "skip";
  // 天图英国
  if ((n.includes("天图") || n.includes("tiantu")) && n.includes("英国") && !n.includes("美")) return "tiantu_uk";
  // 天图空运 (英美空运文件 — 必须在普通tiantu之前)
  if ((n.includes("天图") || n.includes("tiantu")) && (n.includes("空运") || n.includes("air"))) return "tiantu_air";
  if (n.includes("皓辉") || n.includes("haohui")) return "haohui";
  if (n.includes("皓鹏") || n.includes("haopeng")) return "haopeng";
  if (n.includes("星链") || n.includes("xinglian")) return "xinglian";
  if (n.includes("心一") || n.includes("xinyi")) return "xinyi";
  if (n.includes("航乐") || n.includes("hangle") || n.includes("yue")) return "hangle";
  if (n.includes("etton") || n.includes("易通")) return "etton";
  if (n.includes("天图") || n.includes("tiantu")) return "tiantu";
  if (n.includes("英美") || n.includes("yingmei")) return "yingmei";
  if (n.includes("丰运") || n.includes("fengyun")) return "fengyun";
  if (n.includes("华威尔") || n.includes("huaweier")) return "huaweier";
  if (n.includes("凯鑫") || n.includes("kaixin")) return "kaixin";
  if (n.includes("新胜") || n.includes("xinsheng")) return "xinsheng";
  return null;
}

// ── 主流程 ──
function main() {
  // excels 目录（parsers 的上一级的 excels/）
  const baseDir = path.resolve(__dirname, "..", "excels");
  console.log("📂 扫描目录:", baseDir);

  const files = fs.readdirSync(baseDir).filter((f) => (f.endsWith(".xlsx") || f.endsWith(".xls")) && !f.startsWith("~$"));
  console.log(`📋 发现 ${files.length} 个 Excel 文件\n`);

  const allPrices = [];
  const stats = {};

  for (const file of files) {
    const supplier = identifySupplier(file);
    if (!supplier) {
      console.log(`⏭ 跳过 (无法识别供应商): ${file}`);
      continue;
    }
    if (supplier === "skip") {
      console.log(`⏭ 跳过 (非价格表文件): ${file}`);
      continue;
    }

    const filePath = path.join(baseDir, file);
    console.log(`\n🔍 解析: ${file} → 供应商: ${supplier}`);

    let results = [];
    try {
      switch (supplier) {
        case "etton":
          results = parseETTON(filePath);
          break;
        case "tiantu":
          results = parseTiantu(filePath);
          break;
        case "tiantu_uk":
          results = parseTiantuUK(filePath);
          break;
        case "yingmei":
          results = parseYingmei(filePath);
          break;
        case "haohui":
          results = parseHaohui(filePath);
          break;
        case "haopeng":
          results = parseHaopeng(filePath);
          break;
        case "xinglian":
          results = parseXinglian(filePath);
          break;
        case "xinyi":
          results = parseXinyi(filePath);
          break;
        case "hangle":
          results = parseHangle(filePath);
          break;
        case "tiantu_air":
          results = parseTiantuAir(filePath);
          break;
        case "fengyun":
          results = parseXinyun(filePath);
          break;
        case "huaweier":
          results = parseHuaweier(filePath);
          break;
        case "kaixin":
          results = parseKaixin(filePath);
          break;
        case "xinsheng":
          results = parseXinsheng(filePath);
          break;
      }
    } catch (err) {
      console.error(`  ❌ 解析失败: ${err.message}`);
      console.error(err.stack);
      continue;
    }

    // 计算生效日期（从文件名提取）
    // 支持格式: 2026-06-23, 2026年6月23日, 6月24日, 6.25, 06.22
    let effectiveDate = "";
    const dateMatch1 = file.match(/(\d{4})[年.\-]?(\d{1,2})[月.\-](\d{1,2})/);
    if (dateMatch1) {
      effectiveDate = `${dateMatch1[1]}-${String(dateMatch1[2]).padStart(2, "0")}-${String(dateMatch1[3]).padStart(2, "0")}`;
    } else {
      // Try "M月D日" without year prefix
      const dateMatch2 = file.match(/(\d{1,2})[月](\d{1,2})[日]/);
      if (dateMatch2) {
        const now = new Date();
        effectiveDate = `${now.getFullYear()}-${String(parseInt(dateMatch2[1])).padStart(2, "0")}-${String(parseInt(dateMatch2[2])).padStart(2, "0")}`;
      } else {
        const dateMatch3 = file.match(/(\d{1,2})[.·](\d{1,2})/);
        if (dateMatch3) {
          const now = new Date();
          effectiveDate = `${now.getFullYear()}-${String(parseInt(dateMatch3[1])).padStart(2, "0")}-${String(parseInt(dateMatch3[2])).padStart(2, "0")}`;
        }
      }
    }

    // 设置生效日期和来源文件
    for (const r of results) {
      r.effective_date = effectiveDate;
      r.source_file = file;
    }

    stats[supplier] = (stats[supplier] || 0) + results.length;
    allPrices.push(...results);
    console.log(`  ✅ 导入 ${results.length} 条记录 (生效日期: ${effectiveDate})`);
  }

  // ── 输出去重 ──
  // 唯一键: supplier + channel_name + destination_code + origin_region + billing_type + min_quantity
  const seen = new Set();
  const deduped = [];
  for (const r of allPrices) {
    const key = `${r.supplier}|${r.channel_name}|${r.destination_code}|${r.origin_region}|${r.billing_type}|${r.min_quantity}`;
    if (!seen.has(key)) {
      seen.add(key);
      deduped.push(r);
    }
  }
  const dupCount = allPrices.length - deduped.length;
  if (dupCount > 0) {
    console.log(`\n⚠ 去重: 移除 ${dupCount} 条重复记录`);
  }

  // ── 输出 JSON (压缩格式减小体积) ──
  const outDir = path.resolve(__dirname, "..", "data");
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true });

  const outPath = path.join(outDir, "prices.json");
  const output = {
    generated_at: new Date().toISOString(),
    total_records: deduped.length,
    stats,
    data: deduped,
  };

  // 压缩格式写入 (无缩进，节省 ~40% 空间)
  const jsonStr = JSON.stringify(output);
  fs.writeFileSync(outPath, jsonStr, "utf-8");
  console.log(`\n💾 已保存: ${outPath} (${(Buffer.byteLength(jsonStr)/1024/1024).toFixed(1)} MB)`);

  // 同时复制到 public/data/ (供 Web 部署)
  const publicDir = path.resolve(__dirname, "..", "public", "data");
  if (!fs.existsSync(publicDir)) fs.mkdirSync(publicDir, { recursive: true });
  const publicPath = path.join(publicDir, "prices.json");
  fs.writeFileSync(publicPath, jsonStr, "utf-8");
  console.log(`🌐 已复制到: ${publicPath}`);
  console.log(`📊 总计: ${deduped.length} 条价格记录`);
  console.log(`   易通ETTON: ${stats.etton || 0} 条`);
  console.log(`   天图通逊: ${stats.tiantu || 0} 条`);
  console.log(`   天图通逊(英国): ${stats.tiantu_uk || 0} 条`);
  console.log(`   英美跨境: ${stats.yingmei || 0} 条`);
  console.log(`   皓辉国际: ${stats.haohui || 0} 条`);
  console.log(`   皓鹏国际: ${stats.haopeng || 0} 条`);
  console.log(`   星链专线: ${stats.xinglian || 0} 条`);
  console.log(`   心一供应链: ${stats.xinyi || 0} 条`);
  console.log(`   航乐国际: ${stats.hangle || 0} 条`);
  console.log(`   天图空运(英美): ${stats.tiantu_air || 0} 条`);
  console.log(`   丰运跨境: ${stats.fengyun || 0} 条`);
  console.log(`   华威尔: ${stats.huaweier || 0} 条`);
  console.log(`   凯鑫科技: ${stats.kaixin || 0} 条`);
  console.log(`   新胜供应链: ${stats.xinsheng || 0} 条`);

  // ── 简要数据质量报告 ──
  const uniqueWarehouses = new Set(deduped.map((r) => r.destination_code));
  const uniqueChannels = new Set(deduped.map((r) => r.channel_name));
  console.log(`\n📈 数据覆盖: ${uniqueWarehouses.size} 个目的仓, ${uniqueChannels.size} 个渠道`);
}

main();
