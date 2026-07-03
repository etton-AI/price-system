/**
 * 瑞秋物流 — 美线(普货+DG) + 英欧线 + 加拿大线 + 日本线价格解析器
 */
const XLSX = require("xlsx");
const SUPPLIER = "瑞秋物流";

const SZ_CITIES = ["深圳", "东莞", "广州", "中山"];
const YW_CITIES = ["义乌", "上海", "宁波", "杭州", "温州"];

function mkr(o) {
  return {
    supplier: SUPPLIER, country: o.c || "美国", channel_name: o.cn || "", transport_mode: o.tm || "海运",
    vessel_config: o.vc || "", vessel_tags: o.vt || [], delivery_method: o.dm || "卡派",
    destination_type: o.dt || "warehouse", destination_code: o.dc || "", destination_region: o.dr || "",
    origin_region: o.or || "华南", origin_cities: o.oc || SZ_CITIES,
    billing_type: o.bt || "包税", tax_mode: o.tx || o.bt || "包税",
    min_quantity: o.mq || "", min_quantity_value: o.mv || 0, unit_price: o.p || 0, price_unit: o.pu || "元/KG",
    transit_time_min: o.tn || null, transit_time_max: o.tx2 || null, transit_time_desc: o.td || "",
    claim_rule: o.cr || "", effective_date: "", source_sheet: o.ss || "", dg_line: o.dg || false,
  };
}

// ═══════════════════════════════════════════
// 美线普货
// ═══════════════════════════════════════════

/** 美线海卡: warehouse rows with 货交东莞 + 货交义乌 price columns */
function parseUSSeaCard(ws, sheetName, channelName, vesselConfig) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 5) return [];
  const results = [];

  for (let ri = 4; ri < data.length; ri++) {
    const row = data[ri];
    const wh = String(row[4] || "").trim(); // col4=FBA仓代码
    if (!wh || wh.length < 3 || wh.includes("仓代码")) continue;
    if (wh.includes("注意事项") || wh.includes("赔偿")) continue;

    const dgPrice = parseFloat(String(row[7] || "").replace("+", "")); // 电池附加费

    // 货交东莞 price (col5)
    const pDg = parseFloat(row[5]);
    if (!isNaN(pDg) && pDg > 0) {
      const mode = channelName.includes("自税") ? "不包税" : "包税";
      results.push(mkr({
        c: "美国", cn: channelName, tm: "海运", vc: vesselConfig, vt: [vesselConfig],
        dm: "卡派", dc: wh, dt: "warehouse", dr: "美西",
        or: "华南", oc: SZ_CITIES, bt: mode, tx: mode,
        mq: String(row[3] || "50KG+").trim(), mv: parseFloat(String(row[3] || "50")),
        p: pDg, ss: sheetName,
      }));
      // DG variant
      if (!isNaN(dgPrice) && dgPrice > 0) {
        results.push(mkr({
          c: "美国", cn: channelName + "-DG", tm: "海运", vc: vesselConfig, vt: [vesselConfig],
          dm: "卡派", dc: wh, dt: "warehouse", dr: "美西",
          or: "华南", oc: SZ_CITIES, bt: mode, tx: mode,
          mq: String(row[3] || "50KG+").trim(), mv: parseFloat(String(row[3] || "50")),
          p: pDg + dgPrice, ss: sheetName, dg: true,
        }));
      }
    }
    // 货交义乌 price (col6)
    const pYw = parseFloat(row[6]);
    if (!isNaN(pYw) && pYw > 0) {
      const mode = channelName.includes("自税") ? "不包税" : "包税";
      results.push(mkr({
        c: "美国", cn: channelName, tm: "海运", vc: vesselConfig, vt: [vesselConfig],
        dm: "卡派", dc: wh, dt: "warehouse", dr: "美西",
        or: "华东", oc: YW_CITIES, bt: mode, tx: mode,
        mq: String(row[3] || "50KG+").trim(), mv: parseFloat(String(row[3] || "50")),
        p: pYw, ss: sheetName,
      }));
      if (!isNaN(dgPrice) && dgPrice > 0) {
        results.push(mkr({
          c: "美国", cn: channelName + "-DG", tm: "海运", vc: vesselConfig, vt: [vesselConfig],
          dm: "卡派", dc: wh, dt: "warehouse", dr: "美西",
          or: "华东", oc: YW_CITIES, bt: mode, tx: mode,
          mq: String(row[3] || "50KG+").trim(), mv: parseFloat(String(row[3] || "50")),
          p: pYw + dgPrice, ss: sheetName, dg: true,
        }));
      }
    }
  }
  return results;
}

/** 美线海派: region rows with channel columns */
function parseUSSeaParcel(ws, sheetName) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 4) return [];
  const results = [];
  // Row 1: channel names, Row 2: weight tiers
  const chRow = data[1] || [];
  const wtRow = data[2] || [];
  const channels = [];
  for (let ci = 1; ci < chRow.length; ci += 2) {
    const cn = String(chRow[ci] || "").trim();
    if (cn && cn.length > 2) {
      channels.push({ name: cn, col12: ci, col100: ci + 1,
        label12: String(wtRow[ci] || "12KG+").trim(), val12: 12,
        label100: String(wtRow[ci + 1] || "100KG+").trim(), val100: 100,
      });
    }
  }

  for (let ri = 3; ri < data.length; ri++) {
    const row = data[ri];
    const region = String(row[0] || "").trim();
    if (!region) continue;
    if (region.includes("船期") || region.includes("参考时效")) continue;

    let dr = "", or = "华南", oc = SZ_CITIES;
    if (region.includes("西部") || region.includes("8.9")) { dr = "美西"; }
    else if (region.includes("中部") || region.includes("5.6.7")) { dr = "美中"; }
    else if (region.includes("东部") || region.includes("0.1.2.3")) { dr = "美东"; }
    else continue;

    for (const ch of channels) {
      [ch.col12, ch.col100].forEach((col, idx) => {
        const p = parseFloat(row[col]);
        if (!isNaN(p) && p > 0) {
          results.push(mkr({
            c: "美国", cn: ch.name, tm: "海运", vc: "海运", vt: ["海运", "海派"],
            dm: "快递派", dc: dr, dt: "region", dr,
            or, oc, bt: "包税",
            mq: idx === 0 ? ch.label12 : ch.label100,
            mv: idx === 0 ? ch.val12 : ch.val100,
            p, ss: sheetName,
          }));
        }
      });
    }
  }
  return results;
}

// ═══════════════════════════════════════════
// 英欧线
// ═══════════════════════════════════════════

/** 欧洲FBA/海卡/海派: country×warehouse rows with 不包税/包税 price columns */
function parseEUSheet(ws, config) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 6) return [];
  const results = [];
  let currentChannel = config.channelName || "";

  for (let ri = 5; ri < data.length; ri++) {
    const row = data[ri];
    const col0 = String(row[0] || "").trim();
    const col1 = String(row[1] || "").trim();
    const col2 = String(row[2] || "").trim();
    const col3 = String(row[3] || "").trim();

    if (!col2 && !col3) continue;
    if (col0.includes("下单渠道") || col2.includes("国家") || col3.includes("仓库代码")) continue;

    if (col0 && col0.length > 2) currentChannel = col0;

    // col2 = country, col3 = warehouse code
    let countryName = col2;
    let warehouse = col3;

    // Handle country row: col3 empty but col2 has country name
    if (!warehouse && countryName && countryName.length >= 2 && countryName !== "亚马逊仓") {
      warehouse = countryName;
    } else if (!warehouse && countryName && countryName.length >= 2) {
      warehouse = countryName;
    }

    const country = config.country; // "欧线" or "英国"

    for (const tc of config.tierCols) {
      if (tc.col >= row.length) continue;
      const p = parseFloat(row[tc.col]);
      if (isNaN(p) || p <= 0) continue;

      results.push(mkr({
        c: country, cn: currentChannel, tm: config.transportMode || "海运",
        vc: config.transportMode || "海运", vt: [config.transportMode || "海运"],
        dm: config.deliveryMethod || "卡派",
        dc: warehouse, dt: warehouse.match(/[A-Z]{2,}\d/) ? "warehouse" : "country", dr: warehouse,
        or: "华南", oc: ["深圳", "东莞", "广州", "中山"],
        bt: tc.billingType, tx: tc.taxMode || tc.billingType,
        mq: tc.label, mv: tc.value, p, pu: tc.priceUnit || "元/KG",
        ss: config.sheetName,
      }));
    }
  }
  return results;
}

/** Build EU tier cols from row 4 weight headers */
function scanEUTiers(row, startCol, count, billingType) {
  const tiers = [];
  for (let i = 0; i < count; i++) {
    const cell = String(row[startCol + i] || "").trim();
    const kgm = cell.match(/(\d+)\s*KG/i);
    if (kgm) {
      tiers.push({ col: startCol + i, label: kgm[1] + "KG+", value: parseInt(kgm[1]),
        priceUnit: "元/KG", billingType, taxMode: billingType });
    }
  }
  return tiers;
}

// ═══════════════════════════════════════════
// 加拿大线
// ═══════════════════════════════════════════

function parseCASheet(ws, sheetName, channelName) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 4) return [];
  const results = [];

  for (let ri = 2; ri < data.length; ri++) {
    const row = data[ri];
    const col0 = String(row[0] || "").trim();
    const col1 = String(row[1] || "").trim();
    if (!col1 || col1.length < 2) continue;
    if (col1.includes("仓库") || col1.includes("地址") || col1.includes("渠道")) continue;

    let dest = col1;
    if (dest.match(/[A-Z]{2,}\d/)) dest = dest.match(/[A-Z]{2,}\d/)[0];

    // Scan for price columns (3+)
    for (let ci = 2; ci < Math.min(row.length, 10); ci++) {
      const p = parseFloat(row[ci]);
      if (isNaN(p) || p <= 0 || p > 99999) continue;
      let qty = "21KG+", val = 21;
      if (ci >= 5) { qty = "100KG+"; val = 100; }
      results.push(mkr({
        c: "加拿大", cn: channelName, tm: "海运", vc: "海运", vt: ["海运"],
        dm: sheetName.includes("海派") ? "快递派" : "卡派",
        dc: dest, dt: "warehouse", dr: dest,
        or: "华南", oc: SZ_CITIES, bt: "包税",
        mq: qty, mv: val, p, ss: sheetName,
      }));
    }
  }
  return results;
}

// ═══════════════════════════════════════════
// 日本线
// ═══════════════════════════════════════════

function parseJPSheet(ws, sheetName, channelName) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 5) return [];
  const results = [];

  for (let ri = 5; ri < data.length; ri++) {
    const row = data[ri];
    const col1 = String(row[1] || "").trim();
    if (!col1 || col1.startsWith("1）") || col1.startsWith("2）")) continue;

    const tiers = [
      { col: 7, label: "21KG+", value: 21 },
      { col: 8, label: "51KG+", value: 51 },
      { col: 9, label: "100KG+", value: 100 },
    ];

    for (const t of tiers) {
      const p = parseFloat(row[t.col]);
      if (!isNaN(p) && p > 0) {
        results.push(mkr({
          c: "日本", cn: channelName, tm: "海运", vc: "海运", vt: ["海运"],
          dm: "快递派", dc: "日本", dt: "country", dr: "日本",
          or: "华南", oc: SZ_CITIES, bt: "包税",
          mq: t.label, mv: t.value, p, ss: sheetName,
        }));
      }
    }
  }
  return results;
}

// ═══════════════════════════════════════════
// 主入口
// ═══════════════════════════════════════════

function parseRuiqiu(filePath) {
  console.log("[瑞秋] 开始解析:", filePath);
  const wb = XLSX.readFile(filePath);
  const all = [];
  const sn = wb.SheetNames;

  // ── 美线普货 ──
  if (sn.includes("LA美森海卡")) {
    all.push(...parseUSSeaCard(wb.Sheets["LA美森海卡"], "LA美森海卡", "LA美森海卡", "美森MATSON CLX"));
    console.log(`  [LA美森海卡] ${all.length} 条`);
  }
  if (sn.includes("LA以星海卡")) {
    const prev = all.length;
    all.push(...parseUSSeaCard(wb.Sheets["LA以星海卡"], "LA以星海卡", "LA以星海卡", "以星ZIM"));
    console.log(`  [LA以星海卡] ${all.length - prev} 条`);
  }
  if (sn.includes("LA普船海卡")) {
    const prev = all.length;
    all.push(...parseUSSeaCard(wb.Sheets["LA普船海卡"], "LA普船海卡", "LA普船海卡", "OA普船"));
    console.log(`  [LA普船海卡] ${all.length - prev} 条`);
  }
  if (sn.includes("LA海派专线")) {
    const prev = all.length;
    all.push(...parseUSSeaParcel(wb.Sheets["LA海派专线"], "LA海派专线"));
    console.log(`  [LA海派专线] ${all.length - prev} 条`);
  }
  // 其他美线海卡sheets
  ["OAK普船海卡", "NY普船海卡", "SAV普船海卡"].forEach(sh => {
    if (sn.includes(sh)) {
      const prev = all.length;
      const vc = sh.includes("OAK") ? "OA/OAK" : sh.includes("NY") ? "OA/NY" : "OA/SAV";
      all.push(...parseUSSeaCard(wb.Sheets[sh], sh, sh, vc));
      console.log(`  [${sh}] ${all.length - prev} 条`);
    }
  });

  // ── 美线DG ──
  ["LA以星DG海卡", "LA普船DG海卡", "NY普船DG海卡"].forEach(sh => {
    if (sn.includes(sh)) {
      const prev = all.length;
      const vc = sh.includes("以星") ? "以星ZIM-DG" : sh.includes("NY") ? "OA/NY-DG" : "OA/LA-DG";
      all.push(...parseUSSeaCard(wb.Sheets[sh], sh, sh + "-DG", vc));
      console.log(`  [${sh}] ${all.length - prev} 条 → DG`);
    }
  });

  // ── 英欧线 ──
  const euSheetConfigs = [
    { name: "欧洲FBA专线", channel: "欧洲FBA专线", country: "欧线", tm: "海运", dm: "卡派" },
    { name: "欧洲非FBA海卡专线", channel: "欧洲非FBA海卡专线", country: "欧线", tm: "海运", dm: "卡派" },
    { name: "欧洲海派专线", channel: "欧洲海派专线", country: "欧线", tm: "海运", dm: "快递派" },
    { name: "英国普船自税", channel: "英国普船自税", country: "英国", tm: "海运", dm: "快递派" },
  ];

  for (const cfg of euSheetConfigs) {
    if (!sn.includes(cfg.name)) continue;
    const ws = wb.Sheets[cfg.name];
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    if (data.length < 6) continue;
    const tierRow = data[4] || [];
    // Scan tier columns from the row
    const tierCols = [];
    // Detect 不包税 cols (4-8) and 包税 cols (8-12)
    for (let ci = 4; ci < Math.min(tierRow.length, 12); ci++) {
      const cell = String(tierRow[ci] || "").trim();
      const kgm = cell.match(/(\d+)\s*KG/i);
      if (kgm) {
        const bt = ci < 7 ? "不包税" : "包税";
        tierCols.push({ col: ci, label: kgm[1] + "KG+", value: parseInt(kgm[1]),
          priceUnit: "元/KG", billingType: bt, taxMode: bt });
      }
    }
    if (tierCols.length > 0) {
      const prev = all.length;
      all.push(...parseEUSheet(ws, { sheetName: cfg.name, country: cfg.country, transportMode: cfg.tm, deliveryMethod: cfg.dm, channelName: cfg.channel, tierCols }));
      console.log(`  [${cfg.name}] ${all.length - prev} 条 → ${cfg.country}`);
    }
  }

  // ── 加拿大线 ──
  ["加拿大海卡", "加拿大海派", "加拿大DG海卡", "加拿大DG海派"].forEach(sh => {
    if (sn.includes(sh)) {
      const prev = all.length;
      all.push(...parseCASheet(wb.Sheets[sh], sh, sh));
      console.log(`  [${sh}] ${all.length - prev} 条 → 加拿大`);
    }
  });

  // ── 日本线 ──
  if (sn.includes("五日达")) {
    const prev = all.length;
    all.push(...parseJPSheet(wb.Sheets["五日达"], "五日达", "日本海派-五日达"));
    console.log(`  [五日达] ${all.length - prev} 条 → 日本`);
  }
  if (sn.includes("七日达")) {
    const prev = all.length;
    all.push(...parseJPSheet(wb.Sheets["七日达"], "七日达", "日本海派-七日达"));
    console.log(`  [七日达] ${all.length - prev} 条 → 日本`);
  }

  console.log(`[瑞秋] 总计 ${all.length} 条`);
  return all;
}

module.exports = { parseRuiqiu };
