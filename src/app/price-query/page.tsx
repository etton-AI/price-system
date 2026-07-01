"use client";

import { useState, useEffect } from "react";
import Link from "next/link";

interface PriceEntry {
  supplier: string;
  channel_name: string;
  vessel_config: string;
  delivery_method: string;
  destination_code: string;
  destination_region: string;
  origin_region: string;
  origin_cities: string[];
  billing_type: string;
  min_quantity: string;
  unit_price: number;
  price_unit: string;
  transit_time_min: number | null;
  transit_time_max: number | null;
  transit_time_desc: string;
  claim_rule: string;
  effective_date: string;
  source_file: string;
  transport_mode?: string;
}

interface QueryResult {
  success: boolean;
  results: PriceEntry[];
  total: number;
  best: PriceEntry | null;
  stats: { total: number; generated_at: string } | null;
  error?: string;
}

// ── 线路配置 ──
const LINE_CONFIG: Record<string, {
  label: string;
  icon: string;
  transportModes: string[];
  warehouses: string[];
  supplierDesc: string;
}> = {
  "美国": {
    label: "美线",
    icon: "🇺🇸",
    transportModes: ["全部", "海运", "空运"],
    warehouses: ["ONT8", "LGB8", "LAX9", "SBD1", "SMF3", "SCK4", "LAS1", "FTW1", "DFW6", "IAH3", "MEM1", "MDW2", "IND9"],
    supplierDesc: "ETTON · 天图 · 英美 · 皓辉 · 皓鹏 · 星链",
  },
  "英国": {
    label: "英国线",
    icon: "🇬🇧",
    transportModes: ["全部", "海运", "卡航/专车", "铁路", "空运"],
    warehouses: ["BHX4", "LBA4", "BHX8", "BHX7", "LBA2", "MAN4", "MAN8", "LTN7", "LPL2"],
    supplierDesc: "英美 · 天图 · 航乐",
  },
  "欧洲": {
    label: "欧洲线",
    icon: "🇪🇺",
    transportModes: ["全部", "海运", "卡航/专车", "铁路", "空运"],
    warehouses: ["DTM2", "WRO5", "HAJ1", "DUS2", "CDG7", "LYS1", "MXP5", "ZAZ1", "PRG1", "AMS1", "BHX4"],
    supplierDesc: "英美 · 航乐 · 心一",
  },
  "加拿大": {
    label: "加拿大线",
    icon: "🇨🇦",
    transportModes: ["全部", "海运", "空运"],
    warehouses: ["YYZ1", "YYZ3", "YYZ4", "YYZ7", "YYZ9", "YVR1", "YVR2", "YVR3", "YXX2", "YOW1", "YOW3", "YEG1", "YEG2"],
    supplierDesc: "美琦 · ETTON",
  },
  "墨西哥": {
    label: "墨西哥线",
    icon: "🇲🇽",
    transportModes: ["全部", "海运", "空运"],
    warehouses: ["MEX1", "MEX2", "MEX3", "MTY1", "MTY2", "GDL1", "QRO1"],
    supplierDesc: "美琦",
  },
  "巴西": {
    label: "巴西线",
    icon: "🇧🇷",
    transportModes: ["全部", "海运", "空运"],
    warehouses: ["GRU1", "GRU2", "GRU3", "VCP1", "VCP2", "REC1"],
    supplierDesc: "皓鹏",
  },
  "澳大利亚": {
    label: "澳洲线",
    icon: "🇦🇺",
    transportModes: ["全部", "海运", "空运"],
    warehouses: ["SYD1", "SYD3", "MEL1", "MEL5", "BNE1", "PER1", "ADL1"],
    supplierDesc: "星链 · 皓鹏 · ETTON",
  },
};

const ALL_SUPPLIERS = ["ETTON易通", "天图通逊", "英美跨境", "皓辉国际", "皓鹏国际", "星链专线", "心一供应链", "航乐国际", "丰运跨境", "华威尔", "美琦国际"];

// ── 常用城市 ──
const POPULAR_CITIES = ["深圳", "东莞", "广州", "义乌", "上海", "宁波", "厦门", "泉州", "武汉"];

// ── 送货方式 ──
const DELIVERY_METHODS = ["", "卡派", "海派", "快递派", "整柜直送", "直送", "自提"];

// ── 供应商颜色标签 ──
function supplierBadge(s: string) {
  if (s.includes("易通") || s.includes("ETTON")) return { bg: "bg-blue-100", text: "text-blue-700", border: "border-blue-300" };
  if (s.includes("天图")) return { bg: "bg-green-100", text: "text-green-700", border: "border-green-300" };
  if (s.includes("英美")) return { bg: "bg-purple-100", text: "text-purple-700", border: "border-purple-300" };
  if (s.includes("皓辉")) return { bg: "bg-orange-100", text: "text-orange-700", border: "border-orange-300" };
  if (s.includes("皓鹏")) return { bg: "bg-cyan-100", text: "text-cyan-700", border: "border-cyan-300" };
  if (s.includes("星链")) return { bg: "bg-pink-100", text: "text-pink-700", border: "border-pink-300" };
  if (s.includes("心一")) return { bg: "bg-rose-100", text: "text-rose-700", border: "border-rose-300" };
  if (s.includes("航乐")) return { bg: "bg-teal-100", text: "text-teal-700", border: "border-teal-300" };
  if (s.includes("丰运")) return { bg: "bg-amber-100", text: "text-amber-700", border: "border-amber-300" };
  if (s.includes("华威尔")) return { bg: "bg-lime-100", text: "text-lime-700", border: "border-lime-300" };
  if (s.includes("凯鑫")) return { bg: "bg-violet-100", text: "text-violet-700", border: "border-violet-300" };
  if (s.includes("新胜")) return { bg: "bg-sky-100", text: "text-sky-700", border: "border-sky-300" };
  if (s.includes("美琦")) return { bg: "bg-indigo-100", text: "text-indigo-700", border: "border-indigo-300" };
  return { bg: "bg-gray-100", text: "text-gray-700", border: "border-gray-300" };
}

export default function PriceQueryPage() {
  const [country, setCountry] = useState("美国");
  const [transportMode, setTransportMode] = useState("全部");
  const [dest, setDest] = useState("");
  const [origin, setOrigin] = useState("");
  const [weight, setWeight] = useState("");
  const [vessel, setVessel] = useState("");
  const [method, setMethod] = useState("");
  const [selectedSuppliers, setSelectedSuppliers] = useState<string[]>([]);
  const [limit, setLimit] = useState("30");

  const [loading, setLoading] = useState(false);
  const [results, setResults] = useState<PriceEntry[]>([]);
  const [best, setBest] = useState<PriceEntry | null>(null);
  const [total, setTotal] = useState(0);
  const [error, setError] = useState("");
  const [stats, setStats] = useState<{ total: number; generated_at: string } | null>(null);
  const [showExport, setShowExport] = useState(false);
  const [exportData, setExportData] = useState("");
  const [showUpload, setShowUpload] = useState(false);
  const [uploadFiles, setUploadFiles] = useState<FileList | null>(null);
  const [uploading, setUploading] = useState(false);
  const [uploadMsg, setUploadMsg] = useState<{ type: "success" | "error"; text: string } | null>(null);
  const [showSupplierDropdown, setShowSupplierDropdown] = useState(false);
  // 供应商元数据: { name: { countries: [...], latestDate: "2026-06-29" } }
  const [supplierMeta, setSupplierMeta] = useState<Record<string, { countries: string[]; latestDate: string }>>({});

  const lineConfig = LINE_CONFIG[country];

  // 首次加载时获取数据统计和供应商元数据
  useEffect(() => {
    fetch("/api/price-query?meta=1")
      .then((r) => r.json())
      .then((d) => {
        if (d.stats) setStats(d.stats);
        if (d.meta) setSupplierMeta(d.meta);
      })
      .catch(() => {});
  }, []);

  // 当前线路可用的供应商列表（已过滤+含日期）
  const availableSuppliers = ALL_SUPPLIERS.filter((s) => {
    if (!supplierMeta || Object.keys(supplierMeta).length === 0) return true;
    return Object.keys(supplierMeta).some((key) => s.includes(key.replace(/\s/g, "").slice(0, 4)) || key.includes(s.slice(0, 4)));
  }).filter((s) => {
    if (Object.keys(supplierMeta).length === 0) return true;
    // Filter by country
    for (const [key, meta] of Object.entries(supplierMeta)) {
      if ((s.includes(key.replace(/\s/g, "").slice(0, 4)) || key.includes(s.slice(0, 4))) && meta.countries) {
        if (country === "美国") return meta.countries.includes("美国");
        if (country === "英国") return meta.countries.includes("英国");
        if (country === "欧洲") return meta.countries.some((c: string) => !["美国", "英国", "加拿大", "墨西哥", "巴西", "澳大利亚"].includes(c));
        if (country === "加拿大") return meta.countries.includes("加拿大");
        if (country === "墨西哥") return meta.countries.includes("墨西哥");
        if (country === "巴西") return meta.countries.includes("巴西");
        if (country === "澳大利亚") return meta.countries.includes("澳大利亚");
      }
    }
    return false;
  });

  // 供应商日期格式化
  const supplierDateLabel = (supplier: string) => {
    for (const [key, meta] of Object.entries(supplierMeta)) {
      if (supplier.includes(key.slice(0, 4)) || key.includes(supplier.slice(0, 4))) {
        if (meta.latestDate) {
          const d = meta.latestDate.slice(5); // "06-29"
          return ` (${d.replace("-", "/")})`;
        }
      }
    }
    return "";
  };

  // 切换线路时重置运输方式和目的仓
  const handleCountryChange = (c: string) => {
    setCountry(c);
    setTransportMode("全部");
    setDest("");
    setVessel("");
  };

  const toggleSupplier = (supplier: string) => {
    setSelectedSuppliers((prev) =>
      prev.includes(supplier) ? prev.filter((s) => s !== supplier) : [...prev, supplier]
    );
  };

  const handleSearch = async () => {
    if (!dest.trim() && !origin.trim() && !vessel.trim()) {
      setError("请至少输入目的仓、发货城市或渠道关键词");
      return;
    }
    setError("");
    setLoading(true);
    try {
      const params = new URLSearchParams();
      params.set("country", country);
      if (transportMode !== "全部") params.set("transport_mode", transportMode);
      if (dest.trim()) params.set("dest", dest.trim());
      if (origin.trim()) params.set("origin", origin.trim());
      if (weight.trim()) params.set("weight", weight.trim());
      if (vessel.trim()) params.set("vessel", vessel.trim());
      if (method) params.set("method", method);
      if (selectedSuppliers.length > 0) params.set("supplier", selectedSuppliers.join(","));
      params.set("top", limit);

      const resp = await fetch(`/api/price-query?${params.toString()}`);
      const data: QueryResult = await resp.json();
      if (data.success) {
        setResults(data.results);
        setBest(data.best);
        setTotal(data.total);
        if (data.stats) setStats(data.stats);
      } else {
        setError(data.error || "查询失败");
        setResults([]);
        setBest(null);
        setTotal(0);
      }
    } catch (e: unknown) {
      setError("网络错误: " + (e instanceof Error ? e.message : ""));
    }
    setLoading(false);
  };

  const handleExportCSV = () => {
    if (results.length === 0) return;
    const headers = ["供应商", "渠道名", "运输方式", "船配置", "送仓方式", "目的仓", "发货仓", "计费方式", "起收量", "单价", "价格单位", "时效", "赔付规则", "生效日期"];
    const keys: (keyof PriceEntry)[] = ["supplier", "channel_name", "transport_mode", "vessel_config", "delivery_method", "destination_code", "origin_region", "billing_type", "min_quantity", "unit_price", "price_unit", "transit_time_desc", "claim_rule", "effective_date"];
    const lines = [headers.join(",")];
    for (const r of results) {
      lines.push(keys.map((k) => `"${String(r[k] ?? "").replace(/"/g, '""')}"`).join(","));
    }
    const csv = lines.join("\n");
    const blob = new Blob(["﻿" + csv], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `比价查询_${country}_${dest || "all"}_${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleUpload = async () => {
    if (!uploadFiles || uploadFiles.length === 0) return;
    setUploading(true);
    setUploadMsg(null);
    try {
      const form = new FormData();
      for (let i = 0; i < uploadFiles.length; i++) {
        form.append("files", uploadFiles[i]);
      }
      const resp = await fetch("/api/price-query/upload", { method: "POST", body: form });
      const data = await resp.json();
      if (data.success) {
        setUploadMsg({ type: "success", text: data.message });
        setUploadFiles(null);
        setStats(data.stats ? { total: data.totals.deduped, generated_at: new Date().toISOString() } : stats);
      } else {
        setUploadMsg({ type: "error", text: data.error || "上传失败" });
      }
    } catch {
      setUploadMsg({ type: "error", text: "网络错误" });
    }
    setUploading(false);
  };

  const handleExportJSON = () => {
    const json = JSON.stringify({ query: { country, transportMode, dest, origin, weight, vessel, method, suppliers: selectedSuppliers }, results, best, total }, null, 2);
    setExportData(json);
    setShowExport(true);
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === "Enter") handleSearch();
  };

  // ── 快捷预设 ──
  const quickPresets = (() => {
    if (country === "美国") {
      return [
        { label: "ONT8+深圳+EXX", action: () => { setDest("ONT8"); setOrigin("深圳"); setWeight("100"); setVessel("EXX"); setMethod("卡派"); setTransportMode("海运"); } },
        { label: "ONT8+美森卡派", action: () => { setDest("ONT8"); setOrigin("深圳"); setVessel("美森"); setMethod("卡派"); setWeight(""); setTransportMode("海运"); } },
        { label: "LAX9+义乌+美森", action: () => { setDest("LAX9"); setOrigin("义乌"); setVessel("美森"); setMethod("卡派"); setWeight(""); setTransportMode("海运"); } },
        { label: "ONT8+空运", action: () => { setDest("ONT8"); setOrigin("深圳"); setVessel(""); setMethod(""); setWeight("50"); setTransportMode("空运"); } },
      ];
    } else if (country === "英国") {
      return [
        { label: "BHX4+深圳+卡航", action: () => { setDest("BHX4"); setOrigin("深圳"); setVessel("卡航"); setTransportMode("卡航"); setWeight(""); } },
        { label: "BHX4+深圳+海运", action: () => { setDest("BHX4"); setOrigin("深圳"); setVessel("海运"); setTransportMode("海运"); setWeight(""); } },
        { label: "LBA4+义乌+铁运", action: () => { setDest("LBA4"); setOrigin("义乌"); setVessel("铁运"); setTransportMode("铁路"); setWeight(""); } },
        { label: "BHX4+空运", action: () => { setDest("BHX4"); setOrigin("深圳"); setVessel(""); setTransportMode("空运"); setWeight("50"); } },
      ];
    } else {
      return [
        { label: "DTM2+深圳+海运", action: () => { setDest("DTM2"); setOrigin("深圳"); setVessel("海运"); setTransportMode("海运"); setWeight(""); } },
        { label: "WRO5+深圳+铁路", action: () => { setDest("WRO5"); setOrigin("深圳"); setVessel("铁路"); setTransportMode("铁路"); setWeight(""); } },
        { label: "CDG7+深圳+卡航", action: () => { setDest("CDG7"); setOrigin("深圳"); setVessel("卡航"); setTransportMode("卡航"); setWeight(""); } },
        { label: "德国+空运", action: () => { setDest("DTM2"); setOrigin("深圳"); setVessel(""); setTransportMode("空运"); setWeight("50"); } },
      ];
    }
  })();

  return (
    <div className="min-h-screen bg-gray-50">
      {/* 顶部横幅 */}
      <div className="bg-gradient-to-r from-blue-700 to-blue-900 text-white">
        <div className="max-w-7xl mx-auto px-4 py-6">
          <div className="flex items-center justify-between">
            <div>
              <h1 className="text-2xl font-bold">
                {lineConfig.icon} {lineConfig.label}FBA比价查询
              </h1>
              <p className="text-blue-200 text-sm mt-1">
                覆盖 {lineConfig.supplierDesc}
                {stats && (
                  <span className="ml-2 text-blue-300">
                    | {stats.total.toLocaleString()} 条价格数据 | 更新于 {stats.generated_at?.slice(0, 10)}
                  </span>
                )}
              </p>
            </div>
            <Link href="/" className="text-blue-200 hover:text-white text-sm underline">← 返回首页</Link>
          </div>
        </div>
      </div>

      <div className="max-w-7xl mx-auto px-4 py-6">
        {/* 查询表单 */}
        <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-6 mb-6">
          {/* 步骤1: 线路选择 */}
          <div className="flex flex-wrap items-center gap-2 mb-4 pb-4 border-b border-gray-100">
            <span className="text-sm font-medium text-gray-700 mr-2">① 线路：</span>
            {Object.entries(LINE_CONFIG).map(([key, cfg]) => (
              <button
                key={key}
                onClick={() => handleCountryChange(key)}
                className={`px-4 py-1.5 rounded-full text-sm font-medium transition-colors ${
                  country === key
                    ? "bg-blue-600 text-white shadow-sm"
                    : "bg-gray-100 text-gray-600 hover:bg-gray-200"
                }`}
              >
                {cfg.icon} {cfg.label}
              </button>
            ))}
          </div>

          {/* 步骤2: 运输方式 */}
          <div className="flex flex-wrap items-center gap-2 mb-4 pb-4 border-b border-gray-100">
            <span className="text-sm font-medium text-gray-700 mr-2">② 运输方式：</span>
            {lineConfig.transportModes.map((mode) => (
              <button
                key={mode}
                onClick={() => setTransportMode(mode)}
                className={`px-3 py-1.5 rounded-full text-sm font-medium transition-colors ${
                  transportMode === mode
                    ? "bg-indigo-600 text-white shadow-sm"
                    : "bg-gray-100 text-gray-600 hover:bg-gray-200"
                }`}
              >
                {mode === "全部" ? "🔀 全部" : mode}
              </button>
            ))}
          </div>

          {/* 步骤3: 供应商多选 */}
          <div className="flex flex-wrap items-center gap-2 mb-4 pb-4 border-b border-gray-100">
            <span className="text-sm font-medium text-gray-700 mr-2">③ 供应商：</span>
            <div className="relative">
              <button
                onClick={() => setShowSupplierDropdown(!showSupplierDropdown)}
                className="px-3 py-1.5 rounded-full text-sm font-medium bg-gray-100 text-gray-600 hover:bg-gray-200 transition-colors"
              >
                {selectedSuppliers.length === 0
                  ? "🏢 全部供应商"
                  : `🏢 已选 ${selectedSuppliers.length} 家`}
              </button>
              {showSupplierDropdown && (
                <div className="absolute top-full left-0 mt-1 bg-white border border-gray-200 rounded-lg shadow-lg z-20 p-2 min-w-[220px]">
                  {availableSuppliers.map((s) => (
                    <label
                      key={s}
                      className="flex items-center gap-2 px-3 py-2 hover:bg-gray-50 rounded cursor-pointer text-sm"
                    >
                      <input
                        type="checkbox"
                        checked={selectedSuppliers.includes(s)}
                        onChange={() => toggleSupplier(s)}
                        className="rounded"
                      />
                      <span className={`inline-block px-2 py-0.5 rounded text-xs font-medium ${supplierBadge(s).bg} ${supplierBadge(s).text}`}>
                        {s}{supplierDateLabel(s)}
                      </span>
                    </label>
                  ))}
                  {availableSuppliers.length === 0 && (
                    <span className="text-xs text-gray-400 px-3 py-2">加载中...</span>
                  )}
                  {selectedSuppliers.length > 0 && (
                    <button
                      onClick={() => { setSelectedSuppliers([]); }}
                      className="w-full text-center text-xs text-gray-400 hover:text-gray-600 mt-1 pt-1 border-t border-gray-100"
                    >
                      清除选择
                    </button>
                  )}
                </div>
              )}
            </div>
            {/* 已选供应商标签 */}
            {selectedSuppliers.map((s) => {
              const badge = supplierBadge(s);
              return (
                <span
                  key={s}
                  className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-xs font-medium ${badge.bg} ${badge.text} border ${badge.border}`}
                >
                  {s}
                  <button onClick={() => toggleSupplier(s)} className="ml-0.5 hover:opacity-70">&times;</button>
                </span>
              );
            })}
          </div>

          {/* 步骤4: 查询条件 */}
          <div className="grid grid-cols-2 md:grid-cols-4 lg:grid-cols-8 gap-3">
            {/* 目的仓 */}
            <div className="col-span-2 lg:col-span-1">
              <label className="block text-xs font-medium text-gray-600 mb-1">目的仓 *</label>
              <input
                type="text"
                value={dest}
                onChange={(e) => setDest(e.target.value.toUpperCase())}
                onKeyDown={handleKeyDown}
                placeholder={lineConfig.warehouses[0]}
                list="warehouse-list"
                className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none"
              />
              <datalist id="warehouse-list">
                {lineConfig.warehouses.map((w) => (<option key={w} value={w} />))}
              </datalist>
            </div>

            {/* 发货城市 */}
            <div className="col-span-2 lg:col-span-1">
              <label className="block text-xs font-medium text-gray-600 mb-1">发货城市</label>
              <input
                type="text"
                value={origin}
                onChange={(e) => setOrigin(e.target.value)}
                onKeyDown={handleKeyDown}
                placeholder="深圳"
                list="city-list"
                className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none"
              />
              <datalist id="city-list">
                {POPULAR_CITIES.map((c) => (<option key={c} value={c} />))}
              </datalist>
            </div>

            {/* 重量 */}
            <div className="lg:col-span-1">
              <label className="block text-xs font-medium text-gray-600 mb-1">重量(KG)</label>
              <input
                type="number"
                value={weight}
                onChange={(e) => setWeight(e.target.value)}
                onKeyDown={handleKeyDown}
                placeholder="100"
                className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none"
              />
            </div>

            {/* 渠道关键词 (formerly 船司) */}
            <div className="col-span-2 lg:col-span-1">
              <label className="block text-xs font-medium text-gray-600 mb-1">渠道/船司</label>
              <input
                type="text"
                value={vessel}
                onChange={(e) => setVessel(e.target.value)}
                onKeyDown={handleKeyDown}
                placeholder="EXX/美森/卡航"
                className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none"
              />
            </div>

            {/* 送仓方式 */}
            <div className="lg:col-span-1">
              <label className="block text-xs font-medium text-gray-600 mb-1">送仓方式</label>
              <select
                value={method}
                onChange={(e) => setMethod(e.target.value)}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none bg-white"
              >
                {DELIVERY_METHODS.map((m) => (<option key={m} value={m}>{m || "全部"}</option>))}
              </select>
            </div>

            {/* 显示条数 + 搜索按钮 */}
            <div className="lg:col-span-1">
              <label className="block text-xs font-medium text-gray-600 mb-1">显示条数</label>
              <select
                value={limit}
                onChange={(e) => setLimit(e.target.value)}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:border-blue-500 outline-none bg-white"
              >
                <option value="10">10条</option>
                <option value="30">30条</option>
                <option value="50">50条</option>
                <option value="100">100条</option>
                <option value="0">全部</option>
              </select>
            </div>
          </div>

          {/* 操作按钮行 */}
          <div className="flex items-center gap-3 mt-4">
            <button
              onClick={handleSearch}
              disabled={loading}
              className="px-6 py-2.5 bg-blue-600 text-white rounded-lg font-medium hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed text-sm transition-colors"
            >
              {loading ? "🔍 查询中..." : "🔍 查询"}
            </button>
            {results.length > 0 && (
              <>
                <button onClick={handleExportCSV} className="px-4 py-2 text-sm text-gray-600 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-colors">
                  📥 导出CSV
                </button>
                <button onClick={handleExportJSON} className="px-4 py-2 text-sm text-gray-600 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-colors">
                  📋 导出JSON
                </button>
              </>
            )}
            {error && <span className="text-red-500 text-sm">{error}</span>}
          </div>

          {/* 快捷预设 */}
          <div className="mt-3 flex flex-wrap gap-2">
            <span className="text-xs text-gray-400 pt-1">快捷:</span>
            {quickPresets.map((preset, i) => (
              <button
                key={i}
                onClick={preset.action}
                className="text-xs px-2 py-1 bg-gray-100 hover:bg-gray-200 rounded text-gray-600 transition-colors"
              >
                {preset.label}
              </button>
            ))}
            <button onClick={() => { setDest(""); setOrigin("深圳"); setMethod(""); setVessel(""); setWeight(""); }} className="text-xs px-2 py-1 bg-gray-100 hover:bg-gray-200 rounded text-gray-600 transition-colors">深圳所有渠道</button>
          </div>
        </div>

        {/* ── 上传区域 ── */}
        <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-5 mb-6">
          <button
            onClick={() => setShowUpload(!showUpload)}
            className="flex items-center gap-2 text-sm font-medium text-gray-600 hover:text-gray-800 transition-colors"
          >
            <span className={`transform transition-transform ${showUpload ? "rotate-90" : ""}`}>▶</span>
            📤 上传供应商最新报价表
            {stats && (
              <span className="text-xs text-gray-400 ml-2">
                ({stats.total.toLocaleString()} 条数据 | {stats.generated_at?.slice(0, 10)})
              </span>
            )}
          </button>

          {showUpload && (
            <div className="mt-4 pt-4 border-t border-gray-100">
              <p className="text-xs text-gray-500 mb-3">
                上传供应商 Excel 报价表（.xlsx），系统自动识别供应商并更新价格库。支持一次上传多个文件。
              </p>

              {/* 文件拖拽区 */}
              <div
                className={`border-2 border-dashed rounded-lg p-6 text-center transition-colors ${
                  uploadFiles && uploadFiles.length > 0
                    ? "border-green-400 bg-green-50"
                    : "border-gray-300 hover:border-blue-400 bg-gray-50"
                }`}
                onDragOver={(e) => e.preventDefault()}
                onDrop={(e) => {
                  e.preventDefault();
                  setUploadFiles(e.dataTransfer.files);
                  setUploadMsg(null);
                }}
              >
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  multiple
                  className="hidden"
                  id="upload-input"
                  onChange={(e) => {
                    setUploadFiles(e.target.files);
                    setUploadMsg(null);
                  }}
                />
                <label htmlFor="upload-input" className="cursor-pointer">
                  {uploadFiles && uploadFiles.length > 0 ? (
                    <div className="text-sm text-green-700">
                      ✅ 已选择 {uploadFiles.length} 个文件：
                      {Array.from(uploadFiles).map((f, i) => (
                        <span key={i} className="block text-xs mt-1">{f.name} ({(f.size / 1024).toFixed(0)} KB)</span>
                      ))}
                    </div>
                  ) : (
                    <div>
                      <p className="text-gray-500 text-sm">拖拽 Excel 文件到此处，或<span className="text-blue-600">点击选择</span></p>
                      <p className="text-gray-400 text-xs mt-1">支持 .xlsx / .xls 格式</p>
                    </div>
                  )}
                </label>
              </div>

              {/* 上传按钮 */}
              {uploadFiles && uploadFiles.length > 0 && (
                <div className="mt-3 flex items-center gap-3">
                  <button
                    onClick={handleUpload}
                    disabled={uploading}
                    className="px-5 py-2 bg-green-600 text-white rounded-lg text-sm font-medium hover:bg-green-700 disabled:opacity-50 transition-colors"
                  >
                    {uploading ? "⏳ 解析中..." : "✅ 确认上传并更新数据库"}
                  </button>
                  <button
                    onClick={() => { setUploadFiles(null); setUploadMsg(null); }}
                    className="px-3 py-2 text-sm text-gray-500 hover:text-gray-700"
                  >
                    取消
                  </button>
                </div>
              )}

              {uploadMsg && (
                <div className={`mt-3 p-3 rounded-lg text-sm ${
                  uploadMsg.type === "success" ? "bg-green-50 text-green-700 border border-green-200" : "bg-red-50 text-red-700 border border-red-200"
                }`}>
                  {uploadMsg.type === "success" ? "✅ " : "❌ "}
                  {uploadMsg.text}
                </div>
              )}
            </div>
          )}
        </div>

        {/* 最优推荐 */}
        {best && (
          <div className="bg-gradient-to-r from-yellow-50 to-orange-50 border border-yellow-200 rounded-xl p-5 mb-6">
            <div className="flex items-center gap-3">
              <span className="text-3xl">🏆</span>
              <div className="flex-1">
                <span className="text-sm text-yellow-700 font-medium">最优选择</span>
                <h3 className="text-lg font-bold text-gray-900">
                  {best.supplier} — {best.channel_name}
                </h3>
                <div className="flex flex-wrap gap-3 mt-1 text-sm text-gray-600">
                  <span>💰 <strong>{best.unit_price} {best.price_unit}</strong></span>
                  <span>⏱ {best.transit_time_desc || `${best.transit_time_min}-${best.transit_time_max}天`}</span>
                  <span>🚢 {best.vessel_config || best.transport_mode || "-"}</span>
                  <span>📦 {best.origin_region}</span>
                  {best.claim_rule && <span>📋 {best.claim_rule}</span>}
                </div>
              </div>
              <div className={`px-3 py-1 rounded-full text-xs font-medium ${supplierBadge(best.supplier).bg} ${supplierBadge(best.supplier).text} border ${supplierBadge(best.supplier).border}`}>
                最低价
              </div>
            </div>
          </div>
        )}

        {/* 统计行 */}
        {total > 0 && (
          <div className="text-sm text-gray-500 mb-3">
            共找到 <strong className="text-gray-700">{total}</strong> 条匹配记录
            {results.length < total && (
              <span>（当前显示前 {results.length} 条，可调整「显示条数」查看更多）</span>
            )}
          </div>
        )}

        {/* 结果表格 */}
        {results.length > 0 && (
          <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="bg-gray-50 border-b border-gray-200">
                    <th className="text-left px-3 py-3 font-medium text-gray-600 w-10">#</th>
                    <th className="text-left px-3 py-3 font-medium text-gray-600">供应商</th>
                    <th className="text-left px-3 py-3 font-medium text-gray-600 min-w-[160px]">渠道名</th>
                    <th className="text-left px-3 py-3 font-medium text-gray-600">运输方式</th>
                    <th className="text-left px-3 py-3 font-medium text-gray-600 min-w-[140px]">配置</th>
                    <th className="text-right px-3 py-3 font-medium text-gray-600">单价</th>
                    <th className="text-center px-3 py-3 font-medium text-gray-600">时效</th>
                    <th className="text-left px-3 py-3 font-medium text-gray-600">赔付规则</th>
                    <th className="text-left px-3 py-3 font-medium text-gray-600">发货仓</th>
                    <th className="text-center px-3 py-3 font-medium text-gray-600">生效日期</th>
                  </tr>
                </thead>
                <tbody>
                  {results.map((r, i) => {
                    const badge = supplierBadge(r.supplier);
                    const isBest = best && r.supplier === best.supplier && r.channel_name === best.channel_name;
                    return (
                      <tr key={i} className={`border-b border-gray-100 hover:bg-gray-50 transition-colors ${isBest ? "bg-yellow-50/50" : ""}`}>
                        <td className="px-3 py-3 text-gray-400">{i + 1}</td>
                        <td className="px-3 py-3">
                          <span className={`inline-block px-2 py-0.5 rounded text-xs font-medium ${badge.bg} ${badge.text} border ${badge.border}`}>
                            {r.supplier}
                          </span>
                        </td>
                        <td className="px-3 py-3 font-medium text-gray-800">{r.channel_name}</td>
                        <td className="px-3 py-3 text-gray-600 text-xs">{r.transport_mode || "海运"}</td>
                        <td className="px-3 py-3 text-gray-600 text-xs">{r.vessel_config || "-"}</td>
                        <td className="px-3 py-3 text-right">
                          <span className={`font-bold ${isBest ? "text-green-600" : "text-gray-900"}`}>
                            {r.unit_price}
                          </span>
                          <span className="text-gray-400 ml-1 text-xs">{r.price_unit}</span>
                        </td>
                        <td className="px-3 py-3 text-center text-gray-600 text-xs">
                          {r.transit_time_min ? `${r.transit_time_min}-${r.transit_time_max}天` : r.transit_time_desc || "-"}
                        </td>
                        <td className="px-3 py-3 text-gray-500 text-xs leading-relaxed max-w-[180px] truncate" title={r.claim_rule}>
                          {r.claim_rule || "-"}
                        </td>
                        <td className="px-3 py-3 text-gray-600 text-xs">{r.origin_region}</td>
                        <td className="px-3 py-3 text-center text-gray-400 text-xs">{r.effective_date || "-"}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* 无结果提示 */}
        {!loading && total === 0 && !error && (
          <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-12 text-center">
            <p className="text-gray-400 text-lg">🔍 选择线路和运输方式，输入条件后点击查询</p>
            <p className="text-gray-300 text-sm mt-2">
              {country === "美国" ? "例如：美线 → 海运 → ONT8 + 深圳 + 100KG + EXX" :
               country === "英国" ? "例如：英国线 → 卡航 → BHX4 + 深圳 + 100KG" :
               country === "欧洲" ? "例如：欧洲线 → 铁路 → DTM2 + 深圳 + 100KG" :
               country === "加拿大" ? "例如：加拿大线 → 海运 → YYZ1 + 深圳 + 100KG" :
               country === "墨西哥" ? "例如：墨西哥线 → 海运 → 墨西哥城 + 100KG" :
               country === "巴西" ? "例如：巴西线 → 海运 → GRU1 + 深圳 + 100KG" :
               "例如：澳洲线 → 海运 → SYD1 + 深圳 + 100KG"}
            </p>
          </div>
        )}

        {/* 加载状态 */}
        {loading && (
          <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-12 text-center">
            <p className="text-gray-400">⏳ 查询中...</p>
          </div>
        )}

        {/* JSON导出弹窗 */}
        {showExport && (
          <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center" onClick={() => setShowExport(false)}>
            <div className="bg-white rounded-xl shadow-2xl max-w-3xl w-full mx-4 max-h-[80vh] flex flex-col" onClick={(e) => e.stopPropagation()}>
              <div className="flex items-center justify-between px-6 py-4 border-b border-gray-200">
                <h3 className="font-semibold text-gray-800">JSON 导出 ({results.length} 条)</h3>
                <button onClick={() => setShowExport(false)} className="text-gray-400 hover:text-gray-600 text-xl leading-none">&times;</button>
              </div>
              <div className="p-6 overflow-auto flex-1">
                <textarea
                  readOnly
                  value={exportData}
                  className="w-full h-96 font-mono text-xs bg-gray-50 border border-gray-200 rounded-lg p-3 outline-none"
                  onFocus={(e) => e.target.select()}
                />
              </div>
              <div className="px-6 py-3 border-t border-gray-200 flex gap-3">
                <button
                  onClick={() => { navigator.clipboard.writeText(exportData); }}
                  className="px-4 py-2 text-sm bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
                >
                  📋 复制到剪贴板
                </button>
                <button onClick={() => setShowExport(false)} className="px-4 py-2 text-sm text-gray-600 bg-gray-100 rounded-lg hover:bg-gray-200 transition-colors">关闭</button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
