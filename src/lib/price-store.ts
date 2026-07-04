/**
 * 价格数据共享模块
 * 查询 API 和上传 API 共用，支持缓存和刷新
 *
 * ⚠ 性能注意：
 * - prices.json 可能超过 50MB，必须使用异步 I/O 避免阻塞事件循环
 * - 上传后仅刷新缓存（设为 null），下次查询时惰性加载，避免双倍内存峰值
 */

import fs from "fs";
import fsPromises from "fs/promises";
import path from "path";

export interface PriceEntry {
  supplier: string;
  channel_name: string;
  vessel_config: string;
  delivery_method: string;
  destination_code: string;
  destination_type: string;
  destination_region: string;
  origin_region: string;
  origin_cities: string[];
  billing_type: string;
  min_quantity: string;
  min_quantity_value: number;
  unit_price: number;
  price_unit: string;
  transit_time_min: number | null;
  transit_time_max: number | null;
  transit_time_desc: string;
  claim_rule: string;
  effective_date: string;
  source_file: string;
}

export interface PriceData {
  generated_at: string;
  total_records: number;
  stats: Record<string, number>;
  data: PriceEntry[];
}

let cache: PriceData | null = null;

/** 主数据文件路径（运行时读写） */
const DATA_PATH = path.join(process.cwd(), "public", "data", "prices.json");

/** 备份数据路径（构建时生成，作为 fallback） */
const BACKUP_DATA_PATH = path.join(process.cwd(), "data", "prices.json");

/**
 * 同步获取数据（查询 API 使用，依赖内存缓存）
 * 首次调用时同步读取文件，后续命中缓存
 */
export function getData(): PriceData {
  if (cache) return cache;

  // 优先读主数据文件，不存在则 fallback 到备份路径
  const readPath = fs.existsSync(DATA_PATH) ? DATA_PATH : BACKUP_DATA_PATH;

  if (!fs.existsSync(readPath)) {
    console.warn("[price-store] 数据文件不存在，返回空");
    return { generated_at: "", total_records: 0, stats: {}, data: [] };
  }

  const raw = fs.readFileSync(readPath, "utf-8");
  cache = JSON.parse(raw) as PriceData;
  console.log(`[price-store] 数据已加载: ${cache.total_records} 条 (${readPath})`);
  return cache;
}

/**
 * 异步读取数据 —— 用于上传 API 合并逻辑
 * 绕过缓存，直接从文件读取，避免污染查询缓存
 */
export async function readDataAsync(): Promise<PriceData> {
  const readPath = fs.existsSync(DATA_PATH) ? DATA_PATH : BACKUP_DATA_PATH;

  if (!fs.existsSync(readPath)) {
    console.warn("[price-store] 异步读取: 数据文件不存在，返回空");
    return { generated_at: "", total_records: 0, stats: {}, data: [] };
  }

  const raw = await fsPromises.readFile(readPath, "utf-8");
  const data = JSON.parse(raw) as PriceData;
  console.log(`[price-store] 异步读取: ${data.total_records} 条 (${readPath})`);
  return data;
}

/**
 * 异步写入数据 —— 用于上传 API
 * 同时写入主路径和备份路径，确保数据冗余
 */
export async function writeDataAsync(data: PriceData): Promise<void> {
  const json = JSON.stringify(data);
  console.log(`[price-store] 异步写入: ${data.total_records} 条, ${(Buffer.byteLength(json) / 1024 / 1024).toFixed(1)} MB`);

  // 确保目录存在
  for (const p of [DATA_PATH, BACKUP_DATA_PATH]) {
    const dir = path.dirname(p);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  }

  // 并行写入两个路径
  await Promise.all([
    fsPromises.writeFile(DATA_PATH, json, "utf-8"),
    fsPromises.writeFile(BACKUP_DATA_PATH, json, "utf-8"),
  ]);

  console.log(`[price-store] 异步写入完成: ${DATA_PATH}`);
}

/** 刷新缓存（上传新数据后调用，延迟加载避免双倍内存） */
export function refreshCache(): void {
  cache = null;
  // 不立即重新加载 —— 下次查询时惰性加载，避免上传请求内存峰值翻倍
}

/** 获取数据文件路径（上传 API 使用） */
export function getDataPath(): string {
  return DATA_PATH;
}

/** 获取备份数据路径 */
export function getBackupDataPath(): string {
  return BACKUP_DATA_PATH;
}
