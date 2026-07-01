/**
 * 比价查询 API
 * GET /api/price-query?dest=ONT8&origin=深圳&weight=100&vessel=EXX&method=卡派
 */

import { NextRequest, NextResponse } from "next/server";
import { getData, type PriceEntry } from "@/lib/price-store";

interface PriceEntryWithCountry extends PriceEntry {
  country?: string;
  transport_mode?: string;
  tax_mode?: string;
}

interface QueryParams {
  dest?: string;
  origin?: string;
  weight?: number;
  vessel?: string;
  method?: string;
  supplier?: string;
  top?: number;
  best?: boolean;
  country?: string;
  transport_mode?: string;
}

function loadData(): PriceEntry[] {
  const store = getData();
  return store.data;
}

// ── 城市到供应商区域映射 ──
const CITY_TO_ORIGIN: Record<string, Record<string, string[]>> = {
  etton: {
    深圳: ["东莞", "中山", "广州"],
    东莞: ["东莞", "中山", "广州"],
    广州: ["东莞", "中山", "广州"],
    中山: ["东莞", "中山", "广州"],
    惠州: ["东莞", "中山", "广州"],
    义乌: ["嘉兴", "义乌"],
    嘉兴: ["嘉兴", "义乌"],
    杭州: ["嘉兴", "义乌"],
    宁波: ["嘉兴", "义乌"],
    上海: ["嘉兴", "义乌"],
    苏州: ["嘉兴", "义乌"],
    汕头: ["汕头", "厦门", "泉州"],
    厦门: ["汕头", "厦门", "泉州"],
    泉州: ["汕头", "厦门", "泉州"],
    福州: ["汕头", "厦门", "泉州"],
    武汉: ["武汉", "长沙"],
    长沙: ["武汉", "长沙"],
  },
  tiantu: {
    深圳: ["深圳", "广州", "中山", "东莞南城", "惠州"],
    东莞: ["深圳", "广州", "中山", "东莞南城", "惠州"],
    广州: ["深圳", "广州", "中山", "东莞南城", "惠州"],
    中山: ["深圳", "广州", "中山", "东莞南城", "惠州"],
    惠州: ["深圳", "广州", "中山", "东莞南城", "惠州"],
    义乌: ["义乌", "上海", "宁波", "苏州", "杭州", "绍兴"],
    上海: ["义乌", "上海", "宁波", "苏州", "杭州", "绍兴"],
    杭州: ["义乌", "上海", "宁波", "苏州", "杭州", "绍兴"],
    宁波: ["义乌", "上海", "宁波", "苏州", "杭州", "绍兴"],
    厦门: ["厦门", "泉州", "福州"],
    泉州: ["厦门", "泉州", "福州"],
    福州: ["厦门", "泉州", "福州"],
    汕头: ["汕头"],
    重庆: ["重庆"],
    武汉: ["武汉", "长沙", "成都"],
    长沙: ["武汉", "长沙", "成都"],
    青岛: ["青岛", "郑州", "温州", "台州", "连云港", "南京", "合肥"],
    济南: ["济南", "潍坊"],
    天津: ["天津", "南昌", "石家庄"],
    西安: ["西安", "沧州", "保定"],
  },
  yingmei: {
    深圳: ["东莞", "宝安", "中山", "广州", "南城", "汕头", "深圳"],
    东莞: ["东莞", "宝安", "中山", "广州", "南城", "汕头", "深圳"],
    宝安: ["东莞", "宝安", "中山", "广州", "南城", "汕头", "深圳"],
    广州: ["东莞", "宝安", "中山", "广州", "南城", "汕头", "深圳"],
    中山: ["东莞", "宝安", "中山", "广州", "南城", "汕头", "深圳"],
    汕头: ["东莞", "宝安", "中山", "广州", "南城", "汕头", "深圳"],
    义乌: ["义乌", "上海", "宁波", "杭州", "温州"],
    上海: ["义乌", "上海", "宁波", "杭州", "温州"],
    杭州: ["义乌", "上海", "宁波", "杭州", "温州"],
    宁波: ["义乌", "上海", "宁波", "杭州", "温州"],
    温州: ["义乌", "上海", "宁波", "杭州", "温州"],
    福州: ["福州", "厦门", "泉州", "合肥", "青岛", "温州", "汕头"],
    厦门: ["福州", "厦门", "泉州", "合肥", "青岛", "温州", "汕头"],
    泉州: ["福州", "厦门", "泉州", "合肥", "青岛", "温州", "汕头"],
    合肥: ["福州", "厦门", "泉州", "合肥", "青岛", "温州", "汕头"],
    青岛: ["福州", "厦门", "泉州", "合肥", "青岛", "温州", "汕头"],
  },
};

function getSupplierKey(supplier: string): string {
  const s = supplier.toLowerCase();
  if (s.includes("易通") || s.includes("etton")) return "etton";
  if (s.includes("天图") || s.includes("tiantu")) return "tiantu";
  if (s.includes("英美") || s.includes("yingmei")) return "yingmei";
  if (s.includes("皓辉") || s.includes("haohui")) return "haohui";
  if (s.includes("皓鹏") || s.includes("haopeng")) return "haopeng";
  if (s.includes("星链") || s.includes("xinglian")) return "xinglian";
  if (s.includes("心一") || s.includes("xinyi")) return "xinyi";
  if (s.includes("航乐") || s.includes("hangle") || s.includes("yue")) return "hangle";
  return "";
}

// ── 查询逻辑 ──
function query(params: QueryParams): { results: PriceEntry[]; total: number; best: PriceEntry | null } {
  const data = loadData();
  let results = [...data];

  // 0. 国家过滤 (归一化: "欧洲" 和 "欧线" 视为同一区域)
  if (params.country) {
    const country = params.country;
    results = results.filter((r) => {
      const rc = (r as PriceEntryWithCountry).country || "";
      if (country === "欧线" || country === "欧洲") return rc === "欧线" || rc === "欧洲";
      return rc === country;
    });
  }

  // 0.5 运输方式过滤
  if (params.transport_mode) {
    const mode = params.transport_mode;
    results = results.filter((r) => {
      const tm = (r as PriceEntryWithCountry).transport_mode || "海运";
      if (mode === "海运") return tm === "海运";
      if (mode === "空运") return tm === "空运";
      if (mode === "卡航" || mode === "专车" || mode === "卡航/专车") return tm === "卡航" || tm === "卡车" || tm.includes("专车");
      if (mode === "铁路" || mode === "铁运") return tm === "铁路" || tm === "铁运";
      return tm.includes(mode);
    });
  }

  // 1. 目的仓
  if (params.dest) {
    const dest = params.dest.toUpperCase();
    results = results.filter((r) => {
      if (r.destination_type === "none" || r.destination_code === "*") return false;
      return r.destination_code.toUpperCase() === dest;
    });
  }

  // 2. 发货城市
  if (params.origin) {
    const origin = params.origin;
    const ettonCities = (CITY_TO_ORIGIN.etton[origin] || []).map((c) => c.toLowerCase());
    const tiantuCities = (CITY_TO_ORIGIN.tiantu[origin] || []).map((c) => c.toLowerCase());
    const yingmeiCities = (CITY_TO_ORIGIN.yingmei[origin] || []).map((c) => c.toLowerCase());
    const searchLower = origin.toLowerCase();

    results = results.filter((r) => {
      if (!r.origin_cities || r.origin_cities.length === 0) return true;
      const cities = r.origin_cities.map((c) => c.toLowerCase());
      const region = r.origin_region.toLowerCase();
      const supplierKey = getSupplierKey(r.supplier);

      let targetCities: string[];
      if (supplierKey === "etton") targetCities = ettonCities;
      else if (supplierKey === "tiantu") targetCities = tiantuCities;
      else if (supplierKey === "yingmei") targetCities = yingmeiCities;
      else targetCities = [];

      if (cities.some((c) => c.includes(searchLower) || searchLower.includes(c))) return true;
      if (targetCities.length > 0 && targetCities.some((tc) => cities.some((c) => c.includes(tc) || tc.includes(c)))) return true;
      if (region.includes(searchLower)) return true;
      return false;
    });
  }

  // 3. 重量
  if (params.weight) {
    const w = params.weight;
    results = results.filter((r) => {
      if (r.price_unit === "元/CBM") return false;
      return r.min_quantity_value <= w;
    });
    const grouped: Record<string, PriceEntry> = {};
    for (const r of results) {
      const key = `${r.supplier}|${r.channel_name}|${r.destination_code}|${r.origin_region}`;
      if (!grouped[key] || r.min_quantity_value > grouped[key].min_quantity_value) {
        grouped[key] = r;
      }
    }
    results = Object.values(grouped);
  }

  // 4. 船司
  if (params.vessel) {
    const v = params.vessel.toLowerCase();
    results = results.filter((r) => {
      const vesselConfig = (r.vessel_config || "").toLowerCase();
      const channelName = (r.channel_name || "").toLowerCase();
      return vesselConfig.includes(v) || channelName.includes(v);
    });
  }

  // 5. 送仓方式
  if (params.method) {
    const m = params.method;
    results = results.filter((r) => {
      const dm = (r.delivery_method || "").toLowerCase();
      if (m.includes("卡派")) return dm.includes("卡派") || dm.includes("拆派");
      if (m.includes("海派")) return dm.includes("海派") || dm.includes("快递派");
      if (m.includes("整柜") || m.includes("直送")) return dm.includes("整柜") || dm.includes("直送");
      if (m.includes("自提")) return dm.includes("自提");
      return dm.includes(m.toLowerCase());
    });
  }

  // 6. 供应商 (支持逗号分隔多选)
  if (params.supplier) {
    const suppliers = params.supplier.split(/[,，]/).map(s => s.trim().toLowerCase()).filter(Boolean);
    const supplierMap: Record<string, string> = {
      "etton": "易通", "易通": "易通",
      "天图": "天图", "tiantu": "天图",
      "英美": "英美", "yingmei": "英美",
      "皓辉": "皓辉", "haohui": "皓辉",
      "皓鹏": "皓鹏", "haopeng": "皓鹏",
      "星链": "星链", "xinglian": "星链",
      "心一": "心一", "xinyi": "心一",
      "航乐": "航乐", "hangle": "航乐",
      "丰运": "丰运", "fengyun": "丰运",
      "华威尔": "华威尔", "huaweier": "华威尔",
      "凯鑫": "凯鑫", "kaixin": "凯鑫",
      "新胜": "新胜", "xinsheng": "新胜",
      "美琦": "美琦", "meiqi": "美琦",
    };
    const targetSuppliers = suppliers.map(s => supplierMap[s] || s).filter(Boolean);
    if (targetSuppliers.length > 0) {
      results = results.filter((r) => {
        const sup = r.supplier;
        return targetSuppliers.some(ts => sup.includes(ts));
      });
    }
  }

  // 7. 排序: 单价升序 → 时效升序
  results.sort((a, b) => {
    if (a.unit_price !== b.unit_price) return a.unit_price - b.unit_price;
    return (a.transit_time_min || 999) - (b.transit_time_min || 999);
  });

  const total = results.length;

  // 8. Top N / best
  const best = results.length > 0 ? results[0] : null;
  if (params.best && results.length > 0) {
    results = [results[0]];
  } else if (params.top && params.top > 0) {
    results = results.slice(0, params.top);
  }

  return { results, total, best };
}

// ── API Route Handler (GET /api/price-query) ──
export async function GET(request: NextRequest) {
  try {
    const searchParams = request.nextUrl.searchParams;
    const params: QueryParams = {
      dest: searchParams.get("dest") || undefined,
      country: searchParams.get("country") || "美国",
      origin: searchParams.get("origin") || undefined,
      weight: searchParams.get("weight") ? parseFloat(searchParams.get("weight")!) : undefined,
      vessel: searchParams.get("vessel") || undefined,
      method: searchParams.get("method") || undefined,
      supplier: searchParams.get("supplier") || undefined,
      transport_mode: searchParams.get("transport_mode") || undefined,
      top: searchParams.get("top") ? parseInt(searchParams.get("top")!) : undefined,
      best: searchParams.get("best") === "1" || searchParams.get("best") === "true",
    };

    const { results, total, best } = query(params);

    const store = getData();

    // 如果请求meta信息，返回供应商×国家×日期映射
    if (searchParams.get("meta") === "1") {
      const supplierMeta: Record<string, { countries: string[]; latestDate: string }> = {};
      for (const r of store.data) {
        const key = r.supplier;
        if (!supplierMeta[key]) supplierMeta[key] = { countries: [], latestDate: "" };
        const entry = r as PriceEntryWithCountry;
        const cn = entry.country || "美国";
        if (!supplierMeta[key].countries.includes(cn)) supplierMeta[key].countries.push(cn);
        if (entry.effective_date && entry.effective_date > supplierMeta[key].latestDate) {
          supplierMeta[key].latestDate = entry.effective_date;
        }
      }
      return NextResponse.json({
        success: true,
        meta: supplierMeta,
        stats: { total: store.total_records, generated_at: store.generated_at },
      });
    }

    return NextResponse.json({
      success: true,
      query: params,
      results,
      total,
      best,
      stats: { total: store.total_records, generated_at: store.generated_at },
    });
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : "未知错误";
    return NextResponse.json(
      { success: false, error: message },
      { status: 500 }
    );
  }
}
