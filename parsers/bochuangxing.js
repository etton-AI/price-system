/**
 * 博创兴 — 美线+加拿大线价格解析器
 * 格式: 仓库行 × 深圳/义乌城市列 (KG+CBM)
 */
const XLSX = require("xlsx");
const SUPPLIER = "博创兴";

const SZ_CITIES = ["深圳", "东莞", "广州", "中山"];
const YW_CITIES = ["义乌", "宁波", "上海", "杭州"];

function mkr(o) {
  return {
    supplier: SUPPLIER, country: o.c || "美国", channel_name: o.cn || "", transport_mode: o.tm || "海运",
    vessel_config: o.vc || "", vessel_tags: o.vt || [], delivery_method: o.dm || "卡派",
    destination_type: o.dt || "warehouse", destination_code: o.dc || "", destination_region: o.dr || "",
    origin_region: o.or || "华南", origin_cities: o.oc || SZ_CITIES,
    billing_type: o.bt || "包税", tax_mode: o.tx || o.bt || "包税",
    min_quantity: o.mq || "", min_quantity_value: o.mv || 0, unit_price: o.p || 0, price_unit: o.pu || "元/KG",
    transit_time_min: o.tn || null, transit_time_max: o.tx2 || null, transit_time_desc: o.td || "",
    claim_rule: o.cr || "", effective_date: "", source_sheet: o.ss || "",
  };
}

/**
 * 通用仓库行解析
 * tierConfigs: [{szKgCol, szCbmCol, ywKgCol, ywCbmCol, kgLabel, kgVal, cbmLabel, cbmVal}]
 */
function parseWarehouseSheet(ws, config) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 5) return [];
  const results = [];
  let currentChannel = config.channelName;

  for (let ri = config.dataStartRow || 5; ri < data.length; ri++) {
    const row = data[ri];
    const col0 = String(row[0] || "").trim();
    if (!col0 || col0.length < 3) continue;
    if (col0.includes("仓库代码") || col0.includes("注意") || col0.includes("赔偿")) continue;
    // Channel name detection (rows with non-warehouse text)
    if (col0.includes("普船") || col0.includes("快船") || col0.includes("卡派") || col0.includes("专线")) {
      currentChannel = col0; continue;
    }
    // Must be warehouse code
    if (!col0.match(/^[A-Z]{2,}\d/) && !col0.match(/^[A-Z]{3,}$/)) continue;

    const wh = col0;

    for (const tc of config.tierConfigs) {
      // 深圳 KG
      if (tc.szKgCol < row.length) {
        const p = parseFloat(row[tc.szKgCol]);
        if (!isNaN(p) && p > 0 && p < 99999) {
          results.push(mkr({ c: config.country, cn: currentChannel, tm: config.tm || "海运",
            vc: config.vc || "海运", vt: [config.vc || "海运"], dm: config.dm || "卡派",
            dc: wh, dt: "warehouse", dr: config.dr || "美西", or: "华南", oc: SZ_CITIES,
            mq: tc.kgLabel, mv: tc.kgVal, p, pu: "元/KG", ss: config.sheetName }));
        }
      }
      // 深圳 CBM
      if (tc.szCbmCol && tc.szCbmCol < row.length) {
        const p = parseFloat(row[tc.szCbmCol]);
        if (!isNaN(p) && p > 0 && p < 99999) {
          results.push(mkr({ c: config.country, cn: currentChannel, tm: config.tm || "海运",
            vc: config.vc || "海运", vt: [config.vc || "海运"], dm: config.dm || "卡派",
            dc: wh, dt: "warehouse", dr: config.dr || "美西", or: "华南", oc: SZ_CITIES,
            bt: "不含税CBM", tx: "不含税CBM", mq: tc.cbmLabel, mv: tc.cbmVal, p, pu: "元/CBM", ss: config.sheetName }));
        }
      }
      // 义乌 KG
      if (tc.ywKgCol < row.length) {
        const p = parseFloat(row[tc.ywKgCol]);
        if (!isNaN(p) && p > 0 && p < 99999) {
          results.push(mkr({ c: config.country, cn: currentChannel, tm: config.tm || "海运",
            vc: config.vc || "海运", vt: [config.vc || "海运"], dm: config.dm || "卡派",
            dc: wh, dt: "warehouse", dr: config.dr || "美西", or: "华东", oc: YW_CITIES,
            mq: tc.kgLabel, mv: tc.kgVal, p, pu: "元/KG", ss: config.sheetName }));
        }
      }
      // 义乌 CBM
      if (tc.ywCbmCol && tc.ywCbmCol < row.length) {
        const p = parseFloat(row[tc.ywCbmCol]);
        if (!isNaN(p) && p > 0 && p < 99999) {
          results.push(mkr({ c: config.country, cn: currentChannel, tm: config.tm || "海运",
            vc: config.vc || "海运", vt: [config.vc || "海运"], dm: config.dm || "卡派",
            dc: wh, dt: "warehouse", dr: config.dr || "美西", or: "华东", oc: YW_CITIES,
            bt: "不含税CBM", tx: "不含税CBM", mq: tc.cbmLabel, mv: tc.cbmVal, p, pu: "元/CBM", ss: config.sheetName }));
        }
      }
    }
  }
  return results;
}

/** 海派 region rows with channel columns */
function parseSeaParcel(ws, sheetName) {
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (data.length < 6) return [];
  const results = [];
  // Row 3: channel names (美森海派, 以星合德海派, etc.)
  const chRow = data[3] || [];
  const wtRow = data[5] || []; // weight: 50KG+
  const channels = [];
  for (let ci = 1; ci < chRow.length; ci += 2) {
    const cn = String(chRow[ci] || "").trim();
    if (cn && cn.length > 2 && !cn.includes("深圳") && !cn.includes("义乌")) {
      channels.push({ name: cn, szCol: ci, ywCol: ci + 1 });
    }
  }

  for (let ri = 6; ri < data.length; ri++) {
    const row = data[ri];
    const region = String(row[0] || "").trim();
    if (!region) continue;
    let dr = "";
    if (region.includes("西岸") || region.includes("8.9")) dr = "美西";
    else if (region.includes("中部") || region.includes("5.6.7")) dr = "美中";
    else if (region.includes("东岸") || region.includes("0.1.2.3")) dr = "美东";
    else if (region.includes("西岸北") || region.includes("97.98.99")) dr = "美西北";
    else continue;

    for (const ch of channels) {
      // 深圳
      const pSz = parseFloat(row[ch.szCol]);
      if (!isNaN(pSz) && pSz > 0) {
        results.push(mkr({ c: "美国", cn: ch.name, tm: "海运", vc: "海运", vt: ["海运", "海派"],
          dm: "快递派", dc: dr, dt: "region", dr, or: "华南", oc: SZ_CITIES,
          mq: "50KG+", mv: 50, p: pSz, ss: sheetName }));
      }
      // 义乌
      const pYw = parseFloat(row[ch.ywCol]);
      if (!isNaN(pYw) && pYw > 0) {
        results.push(mkr({ c: "美国", cn: ch.name, tm: "海运", vc: "海运", vt: ["海运", "海派"],
          dm: "快递派", dc: dr, dt: "region", dr, or: "华东", oc: YW_CITIES,
          mq: "50KG+", mv: 50, p: pYw, ss: sheetName }));
      }
    }
  }
  return results;
}

// ── 主入口 ──
function parseBochuangxing(filePath) {
  console.log("[博创兴] 开始解析:", filePath);
  const wb = XLSX.readFile(filePath);
  const all = [];

  // 王牌渠道: 美森正班卡派 + 以星合德卡派 (col0=仓库, col1-4=深圳KG/深圳CBM/义乌KG/义乌CBM)
  if (wb.SheetNames.includes("王牌渠道")) {
    const r = parseWarehouseSheet(wb.Sheets["王牌渠道"], {
      sheetName: "王牌渠道", country: "美国", tm: "海运", dm: "卡派", dr: "美西",
      channelName: "王牌渠道-美森正班",
      tierConfigs: [
        { szKgCol: 1, szCbmCol: 2, ywKgCol: 3, ywCbmCol: 4, kgLabel: "21KG+", kgVal: 21, cbmLabel: "0.5CBM+", cbmVal: 0.5 },
        { szKgCol: 7, szCbmCol: 8, ywKgCol: 9, ywCbmCol: 10, kgLabel: "21KG+", kgVal: 21, cbmLabel: "0.5CBM+", cbmVal: 0.5 },
      ],
      dataStartRow: 5,
    });
    console.log(`  [王牌渠道] ${r.length} 条`);
    all.push(...r);
  }

  // 美西专线 (OA快船)
  if (wb.SheetNames.includes("美西专线")) {
    const r = parseWarehouseSheet(wb.Sheets["美西专线"], {
      sheetName: "美西专线", country: "美国", tm: "海运", dm: "卡派", dr: "美西",
      channelName: "美西OA快船", vc: "OA快船",
      tierConfigs: [{ szKgCol: 1, szCbmCol: 2, ywKgCol: 3, ywCbmCol: 4, kgLabel: "21KG+", kgVal: 21, cbmLabel: "0.5CBM+", cbmVal: 0.5 }],
      dataStartRow: 6,
    });
    console.log(`  [美西专线] ${r.length} 条`);
    all.push(...r);
  }

  // 经济线
  if (wb.SheetNames.includes("经济线")) {
    const r = parseWarehouseSheet(wb.Sheets["经济线"], {
      sheetName: "经济线", country: "美国", tm: "海运", dm: "卡派", dr: "美西",
      channelName: "经济线-LA普船", vc: "普船LA",
      tierConfigs: [{ szKgCol: 2, szCbmCol: 3, ywKgCol: 4, ywCbmCol: 5, kgLabel: "21KG+", kgVal: 21, cbmLabel: "1CBM+", cbmVal: 1 }],
      dataStartRow: 5,
    });
    console.log(`  [经济线] ${r.length} 条`);
    all.push(...r);
  }

  // 美中专线 (芝加哥)
  if (wb.SheetNames.includes("美中专线")) {
    const r = parseWarehouseSheet(wb.Sheets["美中专线"], {
      sheetName: "美中专线", country: "美国", tm: "海运", dm: "卡派", dr: "美中",
      channelName: "芝加哥普船", vc: "普船CHI",
      tierConfigs: [
        { szKgCol: 1, szCbmCol: 2, ywKgCol: 3, ywCbmCol: 4, kgLabel: "21KG+", kgVal: 21, cbmLabel: "0.5CBM+", cbmVal: 0.5 },
        { szKgCol: 7, szCbmCol: 8, ywKgCol: 9, ywCbmCol: 10, kgLabel: "21KG+", kgVal: 21, cbmLabel: "0.5CBM+", cbmVal: 0.5 },
      ],
      dataStartRow: 6,
    });
    console.log(`  [美中专线] ${r.length} 条`);
    all.push(...r);
  }

  // 美东专线 (纽约)
  if (wb.SheetNames.includes("美东专线")) {
    const r = parseWarehouseSheet(wb.Sheets["美东专线"], {
      sheetName: "美东专线", country: "美国", tm: "海运", dm: "卡派", dr: "美东",
      channelName: "纽约普船", vc: "普船NY",
      tierConfigs: [
        { szKgCol: 1, szCbmCol: 2, ywKgCol: 3, ywCbmCol: 4, kgLabel: "21KG+", kgVal: 21, cbmLabel: "0.5CBM+", cbmVal: 0.5 },
        { szKgCol: 7, szCbmCol: 8, ywKgCol: 9, ywCbmCol: 10, kgLabel: "21KG+", kgVal: 21, cbmLabel: "0.5CBM+", cbmVal: 0.5 },
      ],
      dataStartRow: 6,
    });
    console.log(`  [美东专线] ${r.length} 条`);
    all.push(...r);
  }

  // 海派专线: region × channel
  if (wb.SheetNames.includes("海派专线")) {
    const r = parseSeaParcel(wb.Sheets["海派专线"], "海派专线");
    console.log(`  [海派专线] ${r.length} 条`);
    all.push(...r);
  }

  // 萨瓦纳专线
  if (wb.SheetNames.includes("萨瓦纳专线")) {
    const r = parseWarehouseSheet(wb.Sheets["萨瓦纳专线"], {
      sheetName: "萨瓦纳专线", country: "美国", tm: "海运", dm: "卡派", dr: "美东",
      channelName: "萨瓦纳普船", vc: "普船SAV",
      tierConfigs: [{ szKgCol: 1, szCbmCol: 2, ywKgCol: 3, ywCbmCol: 4, kgLabel: "21KG+", kgVal: 21, cbmLabel: "0.5CBM+", cbmVal: 0.5 }],
      dataStartRow: 5,
    });
    console.log(`  [萨瓦纳专线] ${r.length} 条`);
    all.push(...r);
  }

  // 美东海卡
  if (wb.SheetNames.includes("美东海卡")) {
    const r = parseWarehouseSheet(wb.Sheets["美东海卡"], {
      sheetName: "美东海卡", country: "美国", tm: "海运", dm: "卡派", dr: "美东",
      channelName: "美东海卡", vc: "普船NY",
      tierConfigs: [{ szKgCol: 1, szCbmCol: 2, ywKgCol: 3, ywCbmCol: 4, kgLabel: "21KG+", kgVal: 21, cbmLabel: "0.5CBM+", cbmVal: 0.5 }],
      dataStartRow: 5,
    });
    console.log(`  [美东海卡] ${r.length} 条`);
    all.push(...r);
  }

  // 加拿大直航
  if (wb.SheetNames.includes("加拿大直航")) {
    const data = XLSX.utils.sheet_to_json(wb.Sheets["加拿大直航"], { header: 1, defval: "" });
    const caResults = [];
    for (let ri = 5; ri < data.length; ri++) {
      const row = data[ri];
      const wh = String(row[0] || "").trim();
      if (!wh.match(/^[A-Z]{2,}\d/)) continue;
      [1, 2].forEach(ci => {
        const p = parseFloat(row[ci]);
        if (!isNaN(p) && p > 0) {
          caResults.push(mkr({ c: "加拿大", cn: "加拿大直航卡派", tm: "海运", vc: "海运", vt: ["海运"],
            dm: "卡派", dc: wh, dt: "warehouse", dr: "多伦多",
            or: ci === 1 ? "华南" : "华东", oc: ci === 1 ? SZ_CITIES : YW_CITIES,
            mq: "101KG+", mv: 101, p, ss: "加拿大直航" }));
        }
      });
    }
    console.log(`  [加拿大直航] ${caResults.length} 条 → 加拿大`);
    all.push(...caResults);
  }

  console.log(`[博创兴] 总计 ${all.length} 条`);
  return all;
}

module.exports = { parseBochuangxing };
