/**
 * 上传供应商最新价格表 Excel
 * POST /api/price-query/upload
 *
 * 接收 Excel 文件 → 自动识别供应商 → 解析 → 合并入库 → 刷新缓存
 * ⚠ 仅 admin 角色可调用，需 JWT Bearer token
 */

import { NextRequest, NextResponse } from "next/server";
import { refreshCache, getDataPath, type PriceEntry } from "@/lib/price-store";
import { extractBearerToken, verifyToken, logUpload } from "@/lib/auth";
import fs from "fs";
import path from "path";
import os from "os";

/** 最大上传文件大小: 15MB */
const MAX_FILE_SIZE = 15 * 1024 * 1024;

interface PriceEntryWithCountry extends PriceEntry {
  country?: string;
}

// 通过 require("module").createRequire 创建原生 Node.js require
// 锚定到 parsers/_index.js，所有通过 parsersRequire 加载的模块
// 都使用原生 Node.js 模块解析（绕过 webpack 打包）
const parsersDir = path.join(process.cwd(), "parsers");
console.log(`[upload] parsersDir=${parsersDir}, exists=${fs.existsSync(parsersDir)}`);

// eslint-disable-next-line @typescript-eslint/no-require-imports
const parsersRequire = require("module").createRequire(path.join(parsersDir, "_index.js"));

// 通过 parsersRequire 加载 loader.js（原生 require 环境）
// loader.js 内部的 require("xlsx") 等调用均使用原生 Node.js 解析
const { parseFile } = parsersRequire("./loader.js");
console.log(`[upload] loader 加载成功, parseFile=${typeof parseFile}`);

export async function POST(request: NextRequest) {
  try {
    // ── 鉴权: 仅 admin 可上传 ──
    const token = extractBearerToken(request.headers.get("authorization"));
    if (!token) {
      return NextResponse.json({ success: false, error: "未提供认证令牌，仅管理员可上传" }, { status: 401 });
    }
    let jwtUser: { username: string; role: string };
    try {
      const payload = await verifyToken(token);
      jwtUser = { username: payload.sub, role: payload.role };
      if (payload.role !== "admin") {
        return NextResponse.json({ success: false, error: "仅管理员可上传报价表" }, { status: 403 });
      }
    } catch {
      return NextResponse.json({ success: false, error: "令牌无效或已过期" }, { status: 401 });
    }

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
      // ── 格式校验: 仅 .xlsx ──
      if (!file.name.endsWith(".xlsx")) {
        console.log(`[upload] ⏭ 跳过非 xlsx 文件: ${file.name}`);
        continue;
      }

      // ── 大小校验: ≤15MB ──
      if (file.size > MAX_FILE_SIZE) {
        console.log(`[upload] ⏭ 文件过大 (${(file.size / 1024 / 1024).toFixed(1)}MB): ${file.name}`);
        results.push({
          file: file.name,
          supplier: "跳过",
          count: 0,
          effectiveDate: "",
        });
        continue;
      }

      console.log(`[upload] 处理: ${file.name} (${(file.size / 1024).toFixed(0)} KB)`);

      // 保存到临时文件
      const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), "price-upload-"));
      const tmpFile = path.join(tmpDir, file.name);
      const bytes = await file.arrayBuffer();
      fs.writeFileSync(tmpFile, Buffer.from(bytes));

      try {
        // 调用 loader.js 的 parseFile（原生 Node.js require 环境，可直接 require("xlsx")）
        const records = parseFile(tmpFile, "") as PriceEntry[];
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

    // ── 合并到现有数据库（内存优化：单遍构建，避免多份中间数组） ──
    const dataPath = getDataPath();
    const updatedCountries = new Set(allNewRecords.map((r) => (r as PriceEntryWithCountry).country || "美国"));
    const updatedSuppliers = Array.from(suppliersUpdated);

    console.log(`[upload] 新增: ${allNewRecords.length} 条 (国家: ${[...updatedCountries].join(",")})`);

    // 构建 merged 数组：先保留未更新的数据，再加新数据
    const seen = new Set<string>();
    const finalData: PriceEntry[] = [];
    let preservedCount = 0;

    if (fs.existsSync(dataPath)) {
      const raw = JSON.parse(fs.readFileSync(dataPath, "utf-8"));
      const existingData: PriceEntry[] = raw.data || [];

      // 第一遍：保留不属于被更新供应商+国家的数据
      for (const r of existingData) {
        const rc = (r as PriceEntryWithCountry).country || "美国";
        const sameSupplier = updatedSuppliers.some((s) => r.supplier.includes(s) || s.includes(r.supplier));
        const sameCountry = updatedCountries.has(rc);
        if (!(sameSupplier && sameCountry)) {
          const key = `${r.supplier}|${r.channel_name}|${r.destination_code}|${r.origin_region}|${r.billing_type}|${r.min_quantity}`;
          if (!seen.has(key)) {
            seen.add(key);
            finalData.push(r);
            preservedCount++;
          }
        }
      }
      // 释放原始 JSON 对象引用，帮助 GC
      (raw as unknown as Record<string, unknown>).data = null;
    }

    console.log(`[upload] 保留其他数据: ${preservedCount} 条`);

    // 第二遍：加入新数据（去重）
    for (const r of allNewRecords) {
      const key = `${r.supplier}|${r.channel_name}|${r.destination_code}|${r.origin_region}|${r.billing_type}|${r.min_quantity}`;
      if (!seen.has(key)) {
        seen.add(key);
        finalData.push(r);
      }
    }

    // 统计（与 build_db.js 一致的供应商映射）
    const supplierStatsMap: Record<string, string> = {
      "易通": "etton", "天图通逊": "tiantu", "英美": "yingmei",
      "皓辉": "haohui", "皓鹏": "haopeng", "星链": "xinglian",
      "心一": "xinyi", "航乐": "hangle", "丰运": "fengyun",
      "华威尔": "huaweier", "凯鑫": "kaixin", "新胜": "xinsheng", "美琦": "meiqi",
    };
    const stats: Record<string, number> = {};
    for (const r of finalData) {
      let key = "other";
      for (const [name, slug] of Object.entries(supplierStatsMap)) {
        if (r.supplier.includes(name)) { key = slug; break; }
      }
      stats[key] = (stats[key] || 0) + 1;
    }

    // 写入文件
    const output = {
      generated_at: new Date().toISOString(),
      total_records: finalData.length,
      stats,
      data: finalData,
    };

    // 确保目录存在
    const dir = path.dirname(dataPath);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

    fs.writeFileSync(dataPath, JSON.stringify(output), "utf-8");
    console.log(`[upload] 💾 已写入 ${finalData.length} 条记录`);

    // ── 记录操作日志 ──
    for (const r of results) {
      if (r.count > 0) {
        logUpload({
          timestamp: new Date().toISOString(),
          username: jwtUser.username,
          fileName: r.file,
          fileSize: files.find((f) => f.name === r.file)?.size || 0,
          recordCount: r.count,
          supplier: r.supplier,
          result: "success",
        });
      }
    }

    // 刷新内存缓存（惰性加载，不立即读取文件）
    refreshCache();

    const dupRemoved = allNewRecords.length + preservedCount - finalData.length;

    return NextResponse.json({
      success: true,
      message: `导入完成：${results.map((r) => `${r.file} (${r.supplier} ${r.count}条)`).join(", ")}`,
      files: results,
      totals: {
        new: allNewRecords.length,
        preserved: preservedCount,
        deduped: finalData.length,
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
