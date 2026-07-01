/**
 * 上传供应商最新价格表 Excel
 * POST /api/price-query/upload
 *
 * 接收 Excel 文件 → 自动识别供应商 → 解析 → 合并入库 → 刷新缓存
 */

import { NextRequest, NextResponse } from "next/server";
import { refreshCache, getDataPath, type PriceEntry } from "@/lib/price-store";
import fs from "fs";
import path from "path";
import os from "os";
// eslint-disable-next-line @typescript-eslint/no-require-imports
import { createRequire } from "module";

const nodeRequire = createRequire(import.meta.url);

/** 与 build_db.js 一致的供应商识别 */
function identifySupplier(fileName: string): string | null {
  const n = fileName.toLowerCase();
  // 跳过非价格表文件（船期表、出运计划等）
  if (n.includes("出运计划") || n.includes("船期") || n.includes("schedule")) return "skip";
  // 天图英国（必须不含"美"）
  if ((n.includes("天图") || n.includes("tiantu")) && n.includes("英国") && !n.includes("美")) return "tiantu_uk";
  // 天图空运（必须在普通 tiantu 之前）
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

/** 解析器文件名 → 导出函数名映射 */
const PARSER_REGISTRY: Record<string, { file: string; exportName: string }> = {
  etton:      { file: "etton_us.js",    exportName: "parseETTON" },
  tiantu:     { file: "tiantu_us.js",   exportName: "parseTiantu" },
  tiantu_uk:  { file: "tiantu_uk.js",   exportName: "parseTiantuUK" },
  tiantu_air: { file: "tiantu_air.js",  exportName: "parseTiantuAir" },
  yingmei:    { file: "yingmei_us.js",  exportName: "parseYingmei" },
  haohui:     { file: "haohui_us.js",   exportName: "parseHaohui" },
  haopeng:    { file: "haopeng_us.js",  exportName: "parseHaopeng" },
  xinglian:   { file: "xinglian_us.js", exportName: "parseXinglian" },
  xinyi:      { file: "xinyi_eu.js",    exportName: "parseXinyi" },
  hangle:     { file: "hangle.js",      exportName: "parseHangle" },
  fengyun:    { file: "fengyun.js",     exportName: "parseXinyun" },
  huaweier:   { file: "huaweier.js",    exportName: "parseHuaweier" },
  kaixin:     { file: "kaixin.js",      exportName: "parseKaixin" },
  xinsheng:   { file: "xinsheng.js",    exportName: "parseXinsheng" },
};

/** 多线路供应商：同一 Excel 可能包含多个国家的 Sheet，需依次尝试所有子解析器 */
const SUPPLIER_PARSER_GROUP: Record<string, string[]> = {
  tiantu: ["tiantu", "tiantu_uk", "tiantu_air"],
};

/** 根据识别出的供应商 key 返回需要尝试的全部解析器 */
function getParserKeys(supplier: string): string[] {
  const base = supplier.replace(/_(uk|air|us)$/, "");
  return SUPPLIER_PARSER_GROUP[base] || [supplier];
}

function parseWithNode(filePath: string, supplier: string): PriceEntry[] {
  const fileName = path.basename(filePath);

  // 自动识别供应商
  if (!supplier) {
    const identified = identifySupplier(fileName);
    if (!identified) {
      throw new Error(
        `无法识别供应商，文件名需包含供应商标识（如 ETTON/易通、天图/tiantu、英美/yingmei、皓辉/haohui、皓鹏/haopeng、星链/xinglian 等）`
      );
    }
    if (identified === "skip") {
      throw new Error(`非价格表文件（船期/出运计划），已跳过: ${fileName}`);
    }
    supplier = identified;
  }

  // 提取生效日期
  let effectiveDate = "";
  const dm1 = fileName.match(/(\d{4})[年.-]?(\d{1,2})[月.-]?(\d{1,2})/);
  if (dm1) {
    effectiveDate = `${dm1[1]}-${String(dm1[2]).padStart(2, "0")}-${String(dm1[3]).padStart(2, "0")}`;
  } else {
    const dm2 = fileName.match(/(\d{1,2})[.·](\d{1,2})/);
    if (dm2) {
      effectiveDate = `${new Date().getFullYear()}-${String(parseInt(dm2[1])).padStart(2, "0")}-${String(parseInt(dm2[2])).padStart(2, "0")}`;
    }
  }

  // 多线路供应商：依次尝试所有子解析器，各取对应 Sheet 的数据
  const parsersDir = path.join(process.cwd(), "parsers");
  const parserKeys = getParserKeys(supplier);
  const allResults: PriceEntry[] = [];
  const parsedLines: string[] = [];

  for (const key of parserKeys) {
    const entry = PARSER_REGISTRY[key];
    if (!entry) continue;

    try {
      const mod = nodeRequire(path.join(parsersDir, entry.file));
      const parseFn = mod[entry.exportName];
      if (typeof parseFn !== "function") {
        console.warn(`[upload] ⚠ ${entry.file} 未导出 ${entry.exportName}，跳过`);
        continue;
      }

      const results: PriceEntry[] = parseFn(filePath);
      if (results.length > 0) {
        // 标记数据来源文件和生效日期
        for (const r of results) {
          r.source_file = fileName;
          r.effective_date = effectiveDate;
        }
        allResults.push(...results);
        parsedLines.push(`${key}(${results.length}条)`);
        console.log(`[upload]   ✅ ${key}: ${results.length} 条`);
      }
    } catch (err) {
      // 某个子解析器失败不影响其他解析器（该线路可能不存在于此文件）
      console.log(`[upload]   ⏭ ${key}: 无匹配数据 (${(err as Error).message.slice(0, 60)})`);
    }
  }

  if (allResults.length === 0) {
    throw new Error(
      `文件 "${fileName}" 未能解析出任何价格数据。尝试的解析器: ${parserKeys.join(", ")}`
    );
  }

  console.log(`[upload] 📊 总计: ${allResults.length} 条 (线路: ${parsedLines.join(" + ")})`);
  return allResults;
}

export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();
    const files = formData.getAll("files") as File[];

    if (files.length === 0) {
      return NextResponse.json({ success: false, error: "请上传至少一个 Excel 文件" }, { status: 400 });
    }

    const results: { file: string; supplier: string; count: number; effectiveDate: string }[] = [];
    const allNewRecords: PriceEntry[] = [];
    const suppliersUpdated = new Set<string>();

    // 逐个处理上传的文件
    for (const file of files) {
      if (!file.name.endsWith(".xlsx") && !file.name.endsWith(".xls")) {
        continue;
      }

      console.log(`[upload] 处理: ${file.name} (${(file.size / 1024).toFixed(0)} KB)`);

      // 保存到临时文件
      const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "price-upload-"));
      const tmpFile = path.join(tmpDir, file.name);
      const bytes = await file.arrayBuffer();
      fs.writeFileSync(tmpFile, Buffer.from(bytes));

      try {
        const records = parseWithNode(tmpFile, "");
        allNewRecords.push(...records);

        if (records.length > 0) {
          suppliersUpdated.add(records[0].supplier);
        }

        const effectiveDate = records.length > 0 ? records[0].effective_date : "";
        results.push({
          file: file.name,
          supplier: records.length > 0 ? records[0].supplier : "未知",
          count: records.length,
          effectiveDate,
        });

        console.log(`[upload] ✅ ${file.name}: ${records.length} 条 (${records[0]?.supplier})`);
      } finally {
        // 清理临时文件
        try { fs.unlinkSync(tmpFile); fs.rmdirSync(tmpDir); } catch {}
      }
    }

    if (allNewRecords.length === 0) {
      return NextResponse.json({ success: false, error: "未能从上传文件中解析到任何价格数据" }, { status: 400 });
    }

    // ── 合并到现有数据库 ──
    // 读取现有数据
    const dataPath = getDataPath();
    let existingData: PriceEntry[] = [];
    if (fs.existsSync(dataPath)) {
      const raw = JSON.parse(fs.readFileSync(dataPath, "utf-8"));
      existingData = raw.data || [];
    }

    // 去掉被更新供应商的旧数据
    const updatedSuppliers = Array.from(suppliersUpdated);
    const preserved = existingData.filter(
      (r: PriceEntry) => !updatedSuppliers.some((s) => r.supplier.includes(s) || s.includes(r.supplier))
    );

    console.log(`[upload] 保留其他供应商: ${preserved.length} 条, 新增: ${allNewRecords.length} 条`);

    // 合并且去重
    const merged = [...preserved, ...allNewRecords];
    const seen = new Set<string>();
    const deduped: PriceEntry[] = [];
    for (const r of merged) {
      const key = `${r.supplier}|${r.channel_name}|${r.destination_code}|${r.origin_region}|${r.billing_type}|${r.min_quantity}`;
      if (!seen.has(key)) {
        seen.add(key);
        deduped.push(r);
      }
    }

    // 统计（与 build_db.js 一致的供应商映射）
    const supplierStatsMap: Record<string, string> = {
      "易通": "etton", "天图通逊": "tiantu", "英美": "yingmei",
      "皓辉": "haohui", "皓鹏": "haopeng", "星链": "xinglian",
      "心一": "xinyi", "航乐": "hangle", "丰运": "fengyun",
      "华威尔": "huaweier", "凯鑫": "kaixin", "新胜": "xinsheng",
    };
    // 天图细分：有 country 字段后，检查是否包含英国/空运标记
    const stats: Record<string, number> = {};
    for (const r of deduped) {
      let key = "other";
      for (const [name, slug] of Object.entries(supplierStatsMap)) {
        if (r.supplier.includes(name)) { key = slug; break; }
      }
      stats[key] = (stats[key] || 0) + 1;
    }

    // 写入文件
    const output = {
      generated_at: new Date().toISOString(),
      total_records: deduped.length,
      stats,
      data: deduped,
    };

    // 确保目录存在
    const dir = path.dirname(dataPath);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

    fs.writeFileSync(dataPath, JSON.stringify(output), "utf-8");

    // 刷新内存缓存
    refreshCache();

    const dupRemoved = allNewRecords.length + preserved.length - deduped.length;

    return NextResponse.json({
      success: true,
      message: `导入完成：${results.map((r) => `${r.file} (${r.supplier} ${r.count}条)`).join(", ")}`,
      files: results,
      totals: {
        new: allNewRecords.length,
        preserved,
        deduped,
        dupRemoved,
      },
      stats,
    });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : "上传处理失败";
    console.error("[upload] 错误:", message);
    return NextResponse.json({ success: false, error: message }, { status: 500 });
  }
}
