/**
 * 丰运跨境 — 欧洲线+英国线价格解析器
 *
 * 特点：城市分列（深圳/广州/义乌），多子表
 * 覆盖: 欧洲海运/空运/卡航/铁路 + 英国海运/卡航/铁路/空运
 */
const XLSX = require("xlsx");
const SUPPLIER = "丰运跨境";

const DEFAULT_CITIES = ["深圳", "东莞", "广州", "中山", "惠州"];

const EU_COUNTRIES = ["德国", "法国", "意大利", "西班牙", "波兰", "捷克", "荷兰", "奥地利",
  "比利时", "卢森堡", "丹麦", "瑞典", "芬兰", "匈牙利", "希腊", "葡萄牙", "爱尔兰"];

function parseCountries(cell) {
  const text = String(cell).replace(/\r?\n/g, " ").trim().replace(/\(FBA\)/gi, "").replace(/大货特价/gi, "").trim();
  for (const cn of EU_COUNTRIES) { if (text === cn) return [cn]; }
  const parts = text.split(/[,，、]/).map(s => s.trim()).filter(Boolean);
  const result = [];
  for (const p of parts) {
    for (const cn of EU_COUNTRIES) { if (p.includes(cn)) { result.push(cn); break; } }
  }
  return result.length > 0 ? result : [text];
}
function parseTransit(text) {
  const c = String(text || "").replace(/\r?\n/g, " ").trim();
  const m = c.match(/(\d+)\s*[-–~]\s*(\d+)\s*(?:个)?(?:自然日|天|工作日)/);
  if (m) return { min: parseInt(m[1]), max: parseInt(m[2]), desc: c };
  const s = c.match(/(\d+)\s*(?:个)?(?:自然日|天|工作日)/);
  if (s) return { min: parseInt(s[1]), max: parseInt(s[1]), desc: c };
  return { min: null, max: null, desc: c };
}
function mkr(o) {
  return {
    supplier: SUPPLIER, country: o.country || "欧线", channel_name: o.cn || "", transport_mode: o.tm || "海运",
    vessel_config: o.vc || "", vessel_tags: o.vt || [], delivery_method: o.dm || "快递派",
    destination_type: o.dt || "country", destination_code: o.dc || "", destination_region: o.dr || "",
    origin_region: o.or || "华南", origin_cities: o.oc || DEFAULT_CITIES.slice(0, 3),
    billing_type: o.bt || "包税", tax_mode: o.tx || o.bt || "包税",
    min_quantity: o.mq || "", min_quantity_value: o.mv || 0, unit_price: o.p || 0, price_unit: "元/KG",
    transit_time_min: o.tn || null, transit_time_max: o.tx2 || null, transit_time_desc: o.td || "",
    claim_rule: o.cr || "", effective_date: "", source_sheet: o.ss || "",
  };
}

// 通用解析：城市列结构 (深圳col,广州col,义乌col)
function parseCityColSheet(ws, sheetName, cfg) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 4) return [];
  const results = [];
  let currentCh = "";

  for (let ri = 1; ri < data.length; ri++) {
    const row = data[ri];
    const c0 = String(row[0] || "").replace(/\r?\n/g, " ").trim();
    const c1 = String(row[1] || "").replace(/\r?\n/g, " ").trim();

    // Skip non-data
    if (!c0 && !c1) continue;
    if (c0.includes("丰运跨境") || c0.includes("无时效") || c0.includes("拒收") || c0.includes("备注") || c0.includes("发车后") || c0.includes("货型")) continue;
    if (c1 === "国家" || c1 === "重量" || c1 === "渠道" || c1 === "服务") continue;

    // Detect sub-channel
    for (const ch of (cfg.subChannels || [])) {
      if (c0.includes(ch.kw) || c1.includes(ch.kw)) { currentCh = ch.name; break; }
    }

    if (!currentCh) continue;

    const countries = parseCountries(c0);
    if (countries.length === 0 || countries[0].length < 2) continue;

    const transit = parseTransit(String(row[cfg.transitCol] || row[cfg.transitCol + 1] || ""));
    const cities = [
      { name: "深圳", col: cfg.szCol || 2 },
      { name: "广州", col: cfg.gzCol || 5 },
      { name: "义乌", col: cfg.ywCol || 8 },
    ];

    for (const city of cities) {
      const cityList = city.name === "义乌" ? ["义乌", "上海", "宁波", "杭州"] :
        city.name === "广州" ? ["广州", "中山"] : ["深圳", "东莞"];
      for (const tier of cfg.tiers) {
        const price = parseFloat(row[city.col + tier.offset]);
        if (!isNaN(price) && price > 0) {
          for (const cn of countries) {
            results.push(mkr({
              country: cfg.country || "欧线", cn: `${sheetName}-${currentCh}`, tm: cfg.tm || "海运",
              vc: cfg.tm || "海运", vt: [cfg.tm || "海运"], dm: "快递派",
              dc: cn, dt: "country", dr: cn,
              or: city.name, oc: cityList,
              bt: currentCh.includes("递延") ? "递延" : "包税",
              tx: currentCh.includes("递延") ? "递延" : "包税",
              mq: tier.qty, mv: tier.val, p: price,
              tn: transit.min, tx2: transit.max, td: transit.desc,
              ss: sheetName,
            }));
          }
        }
      }
    }
  }
  return results;
}

// 英国海运/铁路 (递延+包税子表)
function parseUKSheet(ws, sheetName, transportMode, country) {
  return parseCityColSheet(ws, sheetName, {
    country, tm: transportMode,
    subChannels: [
      { kw: "递延", name: "递延" },
      { kw: "包税", name: "包税" },
    ],
    tiers: [{ offset: 0, qty: "21KG+", val: 21 }, { offset: 1, qty: "100KG+", val: 100 }],
    transitCol: 9, szCol: 3, gzCol: 6, ywCol: 9,
  });
}

function parseXinyun(filePath) {
  console.log("[丰运] 开始解析:", filePath);
  const wb = XLSX.readFile(filePath);
  const all = [];

  const configs = [
    { name: "欧洲海运一口价", fn: (ws) => parseCityColSheet(ws, "欧洲海运一口价", {
      country: "欧线", tm: "海运",
      subChannels: [
        { kw: "特惠-快递派", name: "特惠-快递派" },
        { kw: "特价", name: "大货特价" },
      ],
      tiers: [{ offset: 0, qty: "200KG+", val: 200 }],
      transitCol: 9, szCol: 2, gzCol: 5, ywCol: 8,
    })},
    { name: "欧洲空运一口价", fn: (ws) => parseCityColSheet(ws, "欧洲空运一口价", {
      country: "欧线", tm: "空运",
      subChannels: [
        { kw: "经济特惠", name: "经济特惠" },
        { kw: "特惠", name: "特惠" },
      ],
      tiers: [{ offset: 0, qty: "21KG+", val: 21 }, { offset: 1, qty: "45KG+", val: 45 },
        { offset: 2, qty: "71KG+", val: 71 }, { offset: 3, qty: "100KG+", val: 100 }],
      transitCol: 9, szCol: 1, gzCol: 5, ywCol: -1,
    })},
    { name: "中欧卡航一口价", fn: (ws) => parseCityColSheet(ws, "中欧卡航一口价", {
      country: "欧线", tm: "卡航",
      subChannels: [{ kw: "限时达", name: "限时达一口价" }],
      tiers: [{ offset: 0, qty: "21KG+", val: 21 }, { offset: 1, qty: "100KG+", val: 100 }],
      transitCol: 9, szCol: 2, gzCol: 4, ywCol: 6,
    })},
    { name: "欧洲铁路一口价", fn: (ws) => parseCityColSheet(ws, "欧洲铁路一口价", {
      country: "欧线", tm: "铁路",
      subChannels: [
        { kw: "快递派", name: "快递派" },
        { kw: "卡派", name: "卡派" },
      ],
      tiers: [{ offset: 0, qty: "21KG+", val: 21 }, { offset: 1, qty: "100KG+", val: 100 }],
      transitCol: 9, szCol: 2, gzCol: 5, ywCol: 8,
    })},
    { name: "英国海运一口价递延+包税", fn: (ws) => parseUKSheet(ws, "英国海运递延+包税", "海运", "英国") },
    { name: "英国铁路一口价包税+递延", fn: (ws) => parseUKSheet(ws, "英国铁路包税+递延", "铁路", "英国") },
    { name: "中英卡航一口价", fn: (ws) => parseUKSheet(ws, "中英卡航", "卡航", "英国") },
    { name: "英国空运普货一口价+带电 渠道", fn: (ws) => parseCityColSheet(ws, "英国空运", {
      country: "英国", tm: "空运",
      subChannels: [
        { kw: "普货", name: "普货" },
        { kw: "带电", name: "带电" },
      ],
      tiers: [{ offset: 0, qty: "21KG+", val: 21 }, { offset: 1, qty: "45KG+", val: 45 },
        { offset: 2, qty: "100KG+", val: 100 }],
      transitCol: 9, szCol: 2, gzCol: 5, ywCol: 8,
    })},
  ];

  for (const cfg of configs) {
    if (wb.SheetNames.includes(cfg.name)) {
      try {
        const r = cfg.fn(wb.Sheets[cfg.name]);
        console.log(`  [${cfg.name}] ${r.length} 条`);
        all.push(...r);
      } catch (err) { console.error(`  [${cfg.name}] 失败: ${err.message}`); }
    }
  }
  console.log(`[丰运] 总计 ${all.length} 条`);
  return all;
}
module.exports = { parseXinyun };
