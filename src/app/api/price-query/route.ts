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
  dg?: boolean;
  commercial?: boolean;
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
  if (s.includes("劲港") || s.includes("jingang")) return "jingang";
  if (s.includes("瑞秋") || s.includes("ruiqiu")) return "ruiqiu";
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

  // 0.3 DG 过滤: 普货查询时排除DG/纯电，DG专线时仅显示DG/纯电
  if (params.dg) {
    results = results.filter((r) => {
      const cn = (r.channel_name || "").toUpperCase();
      const dm = (r.delivery_method || "").toUpperCase();
      return cn.includes("DG") || cn.includes("纯电") || dm.includes("DG");
    });
  } else {
    results = results.filter((r) => {
      const cn = (r.channel_name || "").toUpperCase();
      const dm = (r.delivery_method || "").toUpperCase();
      if (cn.includes("DG") || cn.includes("纯电") || dm.includes("DG")) return false;
      return true;
    });
  }

  // 0.4 商业地址过滤
  if (params.commercial) {
    results = results.filter((r) => {
      const dm = (r.delivery_method || "");
      const cn = (r.channel_name || "");
      const dt = (r.destination_type || "");
      return dm.includes("商业") || dm.includes("商私") || cn.includes("商业") || cn.includes("商私") || dt === "zip_zone" || dt === "commercial";
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

  // 0.6 过滤 0 价格脏数据
  results = results.filter((r) => r.unit_price > 0);

  // 1. 目的仓（模糊匹配：代码 / 复合代码 / 区域 / 城市名 / 包含）
  if (params.dest) {
    const dest = params.dest.toUpperCase();
    // 加拿大机场代码→城市名映射（天图等供应商用城市名而非代码）
    const AIRPORT_TO_CITIES: Record<string, string[]> = {
      "YYZ": ["TORONTO","AJAX","BRAMPTON","MARKHAM","MISSISSAUGA","OAKVILLE","RICHMOND HILL","VAUGHAN","PICKERING","SCARBOROUGH","NORTH YORK","ETOBICOKE","WHITBY","OSHAWA","CALEDON","BURLINGTON","MILTON","AURORA","NEWMARKET","KING CITY"],
      "YVR": ["VANCOUVER","RICHMOND","BURNABY","SURREY","LANGLEY","ABBOTSFORD","COQUITLAM","DELTA","NEW WESTMINSTER","PORT COQUITLAM","PITT MEADOW","MISSION","CHILLIWACK","WHISTLER","KAMLOOPS","KELOWNA"],
      "YYC": ["CALGARY","AIRDRIE","ALDERSYDE","COCHRANE","HIGH RIVER","RED DEER","OKOTOKS","CHESTERMERE","BROOKS"],
      "YEG": ["EDMONTON","SHERWOOD PARK","ST ALBERT","LEDUC","FORT MCMURRAY","COLD LAKE"],
      "YOW": ["OTTAWA","GATINEAU","NEPEAN","KANATA","ORLEANS"],
      "YUL": ["MONTREAL","LAVAL","BROSSARD","DORVAL"],
      "YXE": ["SASKATOON"],
      "YQR": ["REGINA"],
      "YWG": ["WINNIPEG","HEADINGLEY","SELKIRK"],
      "YHZ": ["HALIFAX","DARTMOUTH"],
      "YQX": ["GANDER"],
      "YYT": ["ST JOHNS"],
    };
    // 搜"YYZ1"时也匹配YYZ的城市映射（加拿大仓库代码常带数字后缀）
    const destBase = dest.replace(/\d+$/, ""); // YYZ1 → YYZ, YVR2 → YVR
    const airportCities = (AIRPORT_TO_CITIES[dest] || AIRPORT_TO_CITIES[destBase] || []).map(c => c.toUpperCase());
    results = results.filter((r) => {
      if (r.destination_type === "none" || r.destination_code === "*") return false;
      const code = r.destination_code.toUpperCase();
      const region = (r.destination_region || "").toUpperCase();
      if (code === dest) return true;
      // 复合代码拆分: "YYZ/YHM/YOO" 匹配 "YYZ"
      if (code.split("/").map((s: string) => s.trim()).some((part: string) => part === dest)) return true;
      // 机场代码→城市名: 搜"YYZ"匹配 "BRAMPTON, ON"
      if (airportCities.length > 0 && airportCities.some(city => code.includes(city))) return true;
      // 区域模糊匹配: 搜"美西"匹配 destination_region="美西"
      if (region && region.includes(dest)) return true;
      // 代码包含搜索词: "ONT" 匹配 "ONT8"
      if (code.includes(dest)) return true;
      return false;
    });
  }

  // 2. 发货城市（支持模糊匹配：华南→深圳/东莞/广州/中山，中山→中山）
  if (params.origin) {
    const origin = params.origin;
    const searchLower = origin.toLowerCase();

    // 通用区域→城市映射（所有供应商通用）
    const REGION_TO_CITIES: Record<string, string[]> = {
      "华南": ["深圳", "东莞", "广州", "中山", "惠州", "汕头", "佛山", "珠海", "江门"],
      "华东": ["义乌", "上海", "宁波", "杭州", "苏州", "温州", "绍兴", "嘉兴", "南京", "合肥"],
      "华中": ["武汉", "长沙", "郑州", "成都", "重庆"],
      "华北": ["青岛", "天津", "济南", "潍坊", "南昌", "石家庄", "西安", "沧州", "保定", "连云港", "台州"],
      "福建": ["厦门", "泉州", "福州"],
    };
    const regionCities = (REGION_TO_CITIES[origin] || []).map(c => c.toLowerCase());

    const ettonCities = (CITY_TO_ORIGIN.etton[origin] || []).map((c) => c.toLowerCase());
    const tiantuCities = (CITY_TO_ORIGIN.tiantu[origin] || []).map((c) => c.toLowerCase());
    const yingmeiCities = (CITY_TO_ORIGIN.yingmei[origin] || []).map((c) => c.toLowerCase());

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

      // 直接城市匹配: 搜"中山"匹配 origin_cities 中的"中山"
      if (cities.some((c) => c.includes(searchLower) || searchLower.includes(c))) return true;
      // 供应商特定映射
      if (targetCities.length > 0 && targetCities.some((tc) => cities.some((c) => c.includes(tc) || tc.includes(c)))) return true;
      // 区域名称匹配: 搜"华南"匹配 origin_region="华南"
      if (region.includes(searchLower)) return true;
      // 通用区域→城市匹配: 搜"华南"时，所有含深圳/东莞/广州/中山的记录都返回
      if (regionCities.length > 0 && regionCities.some((rc) => cities.some((c) => c.includes(rc) || rc.includes(c)))) return true;
      // 反向: 搜"深圳"时也匹配 origin_region 包含的城市
      if (region && searchLower.length >= 2 && cities.some((c) => region.includes(c))) return true;
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

  // 4. 船司（别名映射 + 普船排除逻辑）
  if (params.vessel) {
    const v = params.vessel.toLowerCase();
    const VESSEL_ALIASES: Record<string, string[]> = {
      "美森": ["clx", "max", "matson"],
      "matson": ["美森", "clx", "max"],
      "clx": ["美森"],
      "max": ["美森"],
      "合德": ["hd"],
      "hd": ["合德"],
      "以星": ["zim", "zem"],
      "zim": ["以星", "zem"],
      "zem": ["以星", "zim"],
      "oa": ["cosco", "emc", "cma", "oocl"],
      "cosco": ["oa", "emc", "cma", "oocl"],
      "emc": ["oa", "cosco", "cma", "oocl"],
      "cma": ["oa", "cosco", "emc", "oocl"],
      "oocl": ["oa", "cosco", "emc", "cma"],
    };
    // 普船: 排除所有快船/名牌船司
    const PREMIUM_VESSELS = ["美森", "matson", "clx", "max", "exx", "合德", "hd", "以星", "zim", "zem", "oa", "cosco", "emc", "cma", "oocl"];
    const isPuChuan = v === "普船";
    const aliases = VESSEL_ALIASES[v] || [];
    results = results.filter((r) => {
      const vesselConfig = (r.vessel_config || "").toLowerCase();
      const channelName = (r.channel_name || "").toLowerCase();
      if (isPuChuan) {
        // 普船: 排除所有快船/名牌
        for (const pv of PREMIUM_VESSELS) {
          if (vesselConfig.includes(pv) || channelName.includes(pv)) return false;
        }
        return true;
      }
      if (vesselConfig.includes(v) || channelName.includes(v)) return true;
      for (const alias of aliases) {
        if (vesselConfig.includes(alias) || channelName.includes(alias)) return true;
      }
      return false;
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
      "etton": "易通", "易通": "易通", "etton易通": "易通",
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
      "劲港": "劲港", "jingang": "劲港",
      "瑞秋": "瑞秋", "ruiqiu": "瑞秋",
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
      dg: searchParams.get("dg") === "1" || searchParams.get("dg") === "true",
      commercial: searchParams.get("commercial") === "1" || searchParams.get("commercial") === "true",
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
