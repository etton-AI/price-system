/**
 * 上传供应商最新价格表 Excel
 * POST /api/price-query/upload
 *
 * 接收 Excel 文件 → 自动识别供应商 → 解析 → 合并入库 → 刷新缓存
 * ⚠ 仅 admin 角色可调用，需 JWT Bearer token
 *
 * ⚠ 性能优化（2026-07-04）:
 * - 使用异步 I/O 避免阻塞事件循环（prices.json 已超 50MB）
 * - 合并逻辑使用"标记删除 + 追加"策略，避免构建庞大的中间数组
 * - 及时释放大对象引用，防止 OOM 导致容器重启
 */

import { NextRequest, NextResponse } from "next/server";
import {
  refreshCache,
  getDataPath,
  getBackupDataPath,
  readDataAsync,
  writeDataAsync,
  type PriceEntry,
} from "@/lib/price-store";
import { extractBearerToken, verifyToken, logUpload } from "@/lib/auth";
import fs from "fs";
import fsPromises from "fs/promises";
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

    // ── 合并到现有数据库（异步 I/O + 内存优化） ──
    const updatedCountries = new Set(allNewRecords.map((r) => (r as PriceEntryWithCountry).country || "美国"));
    const updatedSuppliers = Array.from(suppliersUpdated);

    console.log(`[upload] 新增: ${allNewRecords.length} 条 (国家: ${[...updatedCountries].join(",")})`);

    // 阶段 1: 异步读取现有数据
    const existing = await readDataAsync();
    const existingData: PriceEntry[] = existing.data || [];

    console.log(`[upload] 现有数据: ${existingData.length} 条, 待合并`);

    // 阶段 2: 合并（使用 Set 去重，比构建中间数组更省内存）
    const keyToEntry = new Map<string, PriceEntry>();

    // 2a: 先插入所有新记录
    for (const r of allNewRecords) {
      const key = `${r.supplier}|${r.channel_name}|${r.destination_code}|${r.origin_region}|${r.billing_type}|${r.min_quantity}`;
      keyToEntry.set(key, r);
    }
    const newKeySet = new Set(keyToEntry.keys());

    // 2b: 保留未被替换的旧记录
    let preservedCount = 0;
    for (const r of existingData) {
      const rc = (r as PriceEntryWithCountry).country || "美国";
      const sameSupplier = updatedSuppliers.some((s) => r.supplier.includes(s) || s.includes(r.supplier));
      const sameCountry = updatedCountries.has(rc);
      if (!(sameSupplier && sameCountry)) {
        const key = `${r.supplier}|${r.channel_name}|${r.destination_code}|${r.origin_region}|${r.billing_type}|${r.min_quantity}`;
        if (!keyToEntry.has(key)) {
          keyToEntry.set(key, r);
          preservedCount++;
        }
      }
    }

    console.log(`[upload] 保留其他数据: ${preservedCount} 条`);

    // 合并为最终数组
    const finalData = Array.from(keyToEntry.values());

    // 释放不再需要的大对象引用（帮助 GC）
    (existing as unknown as { data: null }).data = null;

    // 统计
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

    // 构建输出
    const output = {
      generated_at: new Date().toISOString(),
      total_records: finalData.length,
      stats,
      data: finalData,
    };

    // 阶段 3: 异步写入（关键！不阻塞事件循环）
    await writeDataAsync(output);

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

    const dupCount = allNewRecords.length + preservedCount - finalData.length;

    return NextResponse.json({
      success: true,
      message: `导入完成：${results.map((r) => `${r.file} (${r.supplier} ${r.count}条)`).join(", ")}`,
      files: results,
      totals: {
        new: allNewRecords.length,
        preserved: preservedCount,
        deduped: finalData.length,
        dupRemoved: Math.max(0, dupCount),
      },
      stats,
    });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : "上传处理失败";
    console.error("[upload] 错误:", message);
    if (err instanceof Error && err.stack) {
      console.error("[upload] 堆栈:", err.stack.split("\n").slice(0, 5).join("\n"));
    }
    return NextResponse.json({ success: false, error: message }, { status: 500 });
  }
}
