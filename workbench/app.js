const finalHeaders = ["SKU", "款式标题", "长(in)", "宽(in)", "高(in)", "重(lb)", "定价"];
const reportHeaders = ["项目", "数值", "说明"];
const currencyRateRmbToUsd = 7.2;
const firstLegShippingUsd = 0.3;
const referralFeeRate = 0.18;
const returnRate = 0.1;
const disposalFeeUsd = 0.25;
const dimensionalWeightDivisor = 139;
const refundCommissionLossRate = referralFeeRate;
const breakEvenFactor = (1 - returnRate) * (1 - referralFeeRate) - returnRate * referralFeeRate * refundCommissionLossRate;

let finalExcelUrl = "";
let reportExcelUrl = "";
let reportTextUrl = "";
let currentPurchaseRows = [];
let currentProductMode = "standard";
let currentObjectUrl = "";
let currentImageFile = null;
let currentCompetitors = [];
let ocrWorker = null;
let ocrReady = false;
let tesseractLoadPromise = null;
let editingPricingHistoryIndex = null;
let isRestoringWorkbenchDraft = false;
let editingOutputPricingHistoryKey = null;
let outputPricingHistoryRecords = [];
let outputPricingHistoryOverrides = {};
let hiddenOutputPricingHistoryNames = new Set();

const pricingHistoryStorageKey = "priceWorkbench.pricingHistoryProducts";
const workbenchDraftStorageKey = "priceWorkbench.currentDraft";
const outputPricingHistoryOverridesStorageKey = "priceWorkbench.outputHistoryOverrides";
const hiddenOutputPricingHistoryStorageKey = "priceWorkbench.hiddenOutputHistoryNames";
const defaultPricingHistoryProducts = [
  { name: "宠物 AirTag 项圈", purchaseCostRmb: "待补", shippingFee: 1.77, finalPrice: 4.89 },
  { name: "宠物 AirTag 夜光项圈", purchaseCostRmb: "待补", shippingFee: 0.88, finalPrice: 5.49 },
  { name: "健身摇摇杯", purchaseCostRmb: "待补", shippingFee: 2.6, finalPrice: 3.59 },
  { name: "猫玩具大棉签", purchaseCostRmb: "待补", shippingFee: 0.88, finalPrice: 5.99 },
  { name: "黄麻猫玩具", purchaseCostRmb: "待补", shippingFee: 0.88, finalPrice: "$3.99-$4.99" },
  { name: "水晶鞋装饰", purchaseCostRmb: "待补", shippingFee: 3.72, finalPrice: 8.99 },
  { name: "运动毛巾", purchaseCostRmb: "待补", shippingFee: 0.88, finalPrice: 3.99 },
  { name: "防磨贴", purchaseCostRmb: 10.875, shippingFee: 0.88, finalPrice: 5.61 },
  { name: "花束卡片夹", purchaseCostRmb: "¥1.50-¥3.00", shippingFee: 0.5, finalPrice: 2.99 },
  { name: "园艺手套", purchaseCostRmb: 2.3667, shippingFee: 1.77, finalPrice: 4.99 },
  { name: "玻璃喷壶", purchaseCostRmb: 5.225, shippingFee: 3.22, finalPrice: 8.82 },
  { name: "玻璃壶", purchaseCostRmb: 5.225, shippingFee: 3.22, finalPrice: 8.82 },
  { name: "瑜伽砖", purchaseCostRmb: 3.4, shippingFee: 5.11, finalPrice: 12.53 }
];
let pricingHistoryProducts = [...defaultPricingHistoryProducts];

function moneyNumber(value) {
  const parsed = Number(String(value || "").replace(/[^\d.-]/g, ""));
  return Number.isFinite(parsed) ? parsed : 0;
}

function num(value, fallback = 0) {
  const parsed = Number(String(value ?? "").replace(/[^\d.-]/g, ""));
  return Number.isFinite(parsed) ? parsed : fallback;
}

function fmt(value, digits = 2) {
  if (digits <= 0) {
    return String(Math.round(Number(value || 0)));
  }
  return Number(value || 0).toFixed(digits).replace(/\.?0+$/g, "");
}

function money(value, digits = 2) {
  return `$${Number(value || 0).toFixed(digits)}`;
}

function signedMoney(value, digits = 2) {
  const amount = Number(value || 0);
  const sign = amount > 0 ? "+" : amount < 0 ? "-" : "";
  return `${sign}${money(Math.abs(amount), digits)}`;
}

function percent(value, digits = 1) {
  return `${fmt(Number(value || 0) * 100, digits)}%`;
}

function roundUpToCents(value) {
  return Math.ceil(Number(value || 0) * 100) / 100;
}

function cmToIn(value) {
  return fmt(num(value) / 2.54, 2);
}

function gToLb(value) {
  return fmt(num(value) / 453.59237, 3);
}

function gToOz(value) {
  return num(value) / 28.349523125;
}

function cmToRawIn(value) {
  return num(value) / 2.54;
}

function billableWeightFromDims(dims) {
  const actualOz = gToOz(dims.weightG);
  const dimensionalLb = cmToRawIn(dims.lengthCm) * cmToRawIn(dims.widthCm) * cmToRawIn(dims.heightCm) / dimensionalWeightDivisor;
  const dimensionalOz = dimensionalLb * 16;
  return {
    actualOz,
    dimensionalOz,
    billableOz: Math.max(actualOz, dimensionalOz)
  };
}

function $(id) {
  return document.getElementById(id);
}

function isBundleMode() {
  return currentProductMode === "bundle";
}

function escapeAttr(value) {
  return escapeXml(value).replace(/'/g, "&#39;");
}

function escapeXml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function cell(value, type = "String") {
  return `<Cell><Data ss:Type="${type}">${escapeXml(value)}</Data></Cell>`;
}

function buildExcelXml(sheetName, headers, rows) {
  const headerXml = `<Row>${headers.map((item) => cell(item)).join("")}</Row>`;
  const bodyXml = rows.map((row) => {
    return `<Row>${row.map((item, index) => cell(item, index > 1 ? numberType(item) : "String")).join("")}</Row>`;
  }).join("");
  return `<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:x="urn:schemas-microsoft-com:office:excel"
  xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
  <Worksheet ss:Name="${escapeXml(sheetName)}">
    <Table>${headerXml}${bodyXml}</Table>
  </Worksheet>
</Workbook>`;
}

function numberType(value) {
  return value !== "" && Number.isFinite(Number(value)) ? "Number" : "String";
}

function makeDownload(content, mimeType, oldUrl) {
  if (oldUrl) {
    URL.revokeObjectURL(oldUrl);
  }
  return URL.createObjectURL(new Blob([content], { type: mimeType }));
}

function finalRows() {
  return [...document.querySelectorAll("#finalPricingRows tr")].filter((row) => !row.querySelector(".empty")).map((row) => {
    const cells = row.querySelectorAll("td");
    if (cells.length < 7) {
      return [];
    }
    return [
      cells[0].textContent.trim(),
      cells[1].textContent.trim(),
      cells[2].textContent.trim(),
      cells[3].textContent.trim(),
      cells[4].textContent.trim(),
      cells[5].textContent.trim(),
      cells[6].querySelector("input")?.value.trim() || ""
    ];
  }).filter((row) => row.length);
}

function activePurchaseRows() {
  return currentPurchaseRows;
}

function productLabel() {
  const typed = $("productName").value.trim();
  const firstTitle = activePurchaseRows()[0]?.title || "";
  if (!typed || typed === "防磨贴") {
    return firstTitle || typed || "当前产品";
  }
  return typed;
}

function syncProductNameFromRows(rows) {
  const input = $("productName");
  const firstTitle = rows[0]?.title || "";
  if (input && firstTitle && !input.value.trim()) {
    input.value = firstTitle;
  }
}

function resetCompetitorData(message = "产品已更新，请重新上传当前产品的竞品资料。") {
  currentCompetitors = [];
  const fileInput = $("competitorFiles");
  if (fileInput) {
    fileInput.value = "";
  }
  $("competitorFileSummary").textContent = "暂无竞品文件，上传后自动解析当前产品竞品。";
  $("competitorParseStatus").textContent = message;
  $("competitorStatus").textContent = "待上传";
  $("competitorStatus").classList.remove("ready");
  updateCompetitorSummary(0);
}

function targetMarginNumber() {
  const raw = String($("targetMargin").value || "").trim();
  const parsed = raw ? num(raw, 15) : 15;
  return parsed > 1 ? parsed / 100 : parsed;
}

function suggestedPrice() {
  const prices = finalRows().map((row) => num(row[6])).filter((value) => value > 0);
  return prices.length ? prices[0] : 0;
}

function salePackQty() {
  const input = $("salePackQty");
  return Math.max(1, num(input?.value, 1));
}

function comparisonQty() {
  return Math.max(1, num($("comparisonUnitQty").value, 1));
}

function profileText(productName = "", rows = [], competitors = []) {
  return [
    productName,
    ...rows.flatMap((row) => [row.title, row.spec]),
    ...competitors.flatMap((item) => [item.title, item.label])
  ].filter(Boolean).join(" ").toLowerCase();
}

function inferProductProfile(productName = "", rows = [], competitors = []) {
  const text = profileText(productName, rows, competitors);
  const profile = {
    unitLabel: "件",
    saleUnitLabel: "件",
    strategyNoun: "差异化卖点",
    valueAnchor: "材质、规格、使用场景和评价积累",
    launchFocus: "先用可接受利润的入门价验证转化，再根据评价、广告成本和竞品变动调整价格",
    competitorBasis: "同规格折算价",
    reviewStep: 0.5
  };

  const matchers = [
    {
      pattern: /手套|glove|gloves|pair|pairs|双/,
      values: {
        unitLabel: "双",
        saleUnitLabel: "包",
        strategyNoun: "材质、防滑、尺码和多双装便利",
        valueAnchor: "防滑涂层、透气性、尺码覆盖和多双装消耗场景",
        launchFocus: "优先守住多双装折算价和基础利润，评价稳定后再测试更高毛利价",
        competitorBasis: "每双折算价",
        reviewStep: 0.7
      }
    },
    {
      pattern: /喷壶|玻璃壶|spray bottle|mister|sprayer|bottle|瓶/,
      values: {
        unitLabel: "瓶",
        saleUnitLabel: "件",
        strategyNoun: "容量、材质、喷雾效果和外观",
        valueAnchor: "容量、玻璃材质、喷头质感、颜色和使用场景",
        launchFocus: "先对齐主流单瓶价位和配送成本，确认转化后再围绕外观与容量做提价测试",
        competitorBasis: "每瓶折算价",
        reviewStep: 1
      }
    },
    {
      pattern: /瑜伽砖|yoga block|block|砖/,
      values: {
        unitLabel: "块",
        saleUnitLabel: "件",
        strategyNoun: "尺寸、材质密度、防滑和颜色",
        valueAnchor: "尺寸、EVA 密度、防滑触感、边角处理和颜色",
        launchFocus: "先卡住同尺寸单块价格带，评价积累后再测试颜色或套装溢价",
        competitorBasis: "每块折算价",
        reviewStep: 1
      }
    },
    {
      pattern: /卡片夹|花束夹|card holder|card pick|floral pick|pick/,
      values: {
        unitLabel: "支",
        saleUnitLabel: "包",
        strategyNoun: "数量、长度、材质和使用场景",
        valueAnchor: "每包数量、长度、金属/塑料材质和花艺场景",
        launchFocus: "先用清晰的数量折算价拿到花艺耗材用户，评价稳定后测试更高数量包或材质溢价",
        competitorBasis: "每支折算价",
        reviewStep: 0.6
      }
    },
    {
      pattern: /防磨贴|水胶体|贴|片|pad|pads|patch|patches|bandage|bandages/,
      values: {
        unitLabel: "片",
        saleUnitLabel: "套",
        strategyNoun: "组合规格、单片成本和试用门槛",
        valueAnchor: "片数、尺寸组合、亲肤材料和旅行/磨脚场景",
        launchFocus: "首发用低试用门槛跑转化，评价稳定后再测试组合套装溢价",
        competitorBasis: "每片折算价",
        reviewStep: 0.4
      }
    }
  ];

  for (const matcher of matchers) {
    if (matcher.pattern.test(text)) {
      return { ...profile, ...matcher.values };
    }
  }
  return profile;
}

function sellingUnitLabel(productName) {
  const rows = activePurchaseRows();
  if (rows.length === 1 && rows[0].bundle) {
    return rows[0].bundleUnitLabel || "片";
  }
  return inferProductProfile(productName, activePurchaseRows(), activeCompetitors()).unitLabel;
}

function saleUnitLabel(productName) {
  const rows = activePurchaseRows();
  if (rows.length === 1 && rows[0].bundle) {
    return "套";
  }
  return inferProductProfile(productName, activePurchaseRows(), activeCompetitors()).saleUnitLabel;
}

function activeCompetitors() {
  return currentCompetitors;
}

function parsePackCount(text, fallback = 1) {
  const source = String(text || "").replace(/(\d+)\s*ml/gi, "");
  const patterns = [
    /(\d+)\s*[-–—]?\s*(?:pcs|pc|pieces|piece|count|ct)\b/i,
    /(\d+)\s*[-–—]?\s*(?:pack|packs|pair|pairs)\b/i,
    /(?:pack|packs|set|sets|bundle|bundles|box|boxes)\s+of\s+(\d+)\b/i,
    /(\d+)\s*(?:片|件|个|支|套|双|包)/
  ];
  for (const pattern of patterns) {
    const match = source.match(pattern);
    if (match) {
      return Math.max(1, num(match[1], fallback));
    }
  }
  return fallback;
}

function shortCompetitorLabel(title, index) {
  const cleaned = String(title || "")
    .replace(/\s+/g, " ")
    .replace(/\bAmazon\.com\b/gi, "")
    .replace(/\s*[:|_].*$/g, "")
    .trim();
  return cleaned ? cleaned.slice(0, 28) : `竞品${index + 1}`;
}

function firstMatch(text, patterns) {
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) {
      return match[1] || match[0];
    }
  }
  return "";
}

function extractCompetitorPriceText(text) {
  const scopedPatterns = [
    /<span[^>]*class=["'][^"']*aok-offscreen[^"']*["'][^>]*>\s*\$(\d+(?:\.\d{1,2})?)\s*<\/span>[\s\S]{0,240}<span[^>]*class=["'][^"']*priceToPay[^"']*["'][^>]*>/i,
    /<div[^>]*id=["']corePrice[^"']*["'][^>]*>[\s\S]{0,1600}?<span[^>]*class=["'][^"']*(?:aok-offscreen|a-offscreen)[^"']*["'][^>]*>\s*\$(\d+(?:\.\d{1,2})?)\s*<\/span>/i,
    /<div[^>]*id=["']corePriceDisplay_desktop_feature_div["'][^>]*>[\s\S]{0,1600}?<span[^>]*class=["'][^"']*(?:aok-offscreen|a-offscreen)[^"']*["'][^>]*>\s*\$(\d+(?:\.\d{1,2})?)\s*<\/span>/i
  ];
  const scoped = firstMatch(text, scopedPatterns);
  if (scoped) {
    return scoped;
  }
  return firstMatch(text, [
    /aria-label=["'](?:current price\s*)?\$(\d+(?:\.\d{1,2})?)["']/i,
    /<span[^>]*class=["'][^"']*a-offscreen[^"']*["'][^>]*>\s*\$(\d+(?:\.\d{1,2})?)\s*<\/span>/i,
    /\$(\d+(?:\.\d{1,2})?)/
  ]);
}

function parseCompetitorText(text, fileName, index) {
  const plain = cleanTextFromHtml(text);
  const title = firstMatch(text, [
    /id=["']productTitle["'][^>]*>([\s\S]*?)<\/span>/i,
    /<title[^>]*>([\s\S]*?)<\/title>/i
  ]) || plain.split(/\r?\n/).find((line) => line.trim().length > 8) || fileName;
  const cleanTitle = cleanTextFromHtml(title).replace(/\s+/g, " ").trim();
  const priceText = extractCompetitorPriceText(text);
  const price = moneyNumber(priceText);
  if (!price) {
    return null;
  }
  const ratingText = firstMatch(text, [
    /(\d(?:\.\d)?)\s*out of\s*5\s*stars/i,
    /([1-5](?:\.\d)?)\s*分/
  ]);
  const salesText = firstMatch(plain, [
    /([\d,]+\+?)\s*bought\s+in\s+past\s+month/i,
    /过去\s*一个月\s*([\d,]+\+?)\s*(?:人)?(?:购买|买过)/i,
    /月销\s*([\d,]+\+?)/i
  ]);
  const packCount = parsePackCount(`${cleanTitle} ${fileName}`, 1);
  return {
    label: shortCompetitorLabel(cleanTitle || fileName, index),
    title: cleanTitle || fileName,
    price,
    packCount,
    rating: ratingText ? `${ratingText}分` : "评分待确认",
    sales: salesText ? `${salesText.replace(/,/g, "")}/月` : "销量待确认"
  };
}

async function parseCompetitorFile(file, index) {
  const suffix = file.name.split(".").pop().toLowerCase();
  if (!["html", "htm", "txt", "csv"].includes(suffix)) {
    return null;
  }
  const text = await file.text();
  return parseCompetitorText(text, file.name, index);
}

function sellingUnitCostUsd(rows, unitQty) {
  if (unitQty <= 0) {
    return 0;
  }
  const positiveCosts = rows.map((row) => num(row.cost)).filter((value) => value > 0);
  if (!positiveCosts.length) {
    return 0;
  }
  if (rows.length === 1 && rows[0].bundle) {
    return num(rows[0].bundleSaleUnitCostRmb || rows[0].cost) / currencyRateRmbToUsd;
  }
  if (rows.length > 1) {
    const highestCost = Math.max(...rows.map((row) => num(row.cost)).filter((value) => value > 0));
    return highestCost * unitQty / currencyRateRmbToUsd;
  }
  return positiveCosts[0] * unitQty / currencyRateRmbToUsd;
}

function rowProfitRows(rows, unitQty, price, shippingFee) {
  if (rows.length === 1 && rows[0].bundle) {
    const costUsd = num(rows[0].bundleSaleUnitCostRmb || rows[0].cost) / currencyRateRmbToUsd;
    const profit = price * breakEvenFactor - (costUsd + firstLegShippingUsd + shippingFee + disposalFeeUsd * returnRate);
    return [{
      title: rows[0].spec || rows[0].title || "组合套装",
      costUsd,
      profit
    }].filter((row) => row.costUsd > 0);
  }
  return rows.map((row) => {
    const costUsd = num(row.cost) * unitQty / currencyRateRmbToUsd;
    const profit = price * breakEvenFactor - (costUsd + firstLegShippingUsd + shippingFee + disposalFeeUsd * returnRate);
    return {
      title: row.spec || row.title || "当前款式",
      costUsd,
      profit
    };
  }).filter((row) => row.costUsd > 0);
}

function fbaFeeForOz(weightOz, price) {
  if (weightOz <= 0) {
    return { fee: 0, tier: "待填写尺寸重量" };
  }
  if (weightOz <= 4) {
    return price <= 3
      ? { fee: 0.5, tier: "4oz及以下且售价 <= $3" }
      : { fee: 0.88, tier: "4oz及以下且售价 > $3" };
  }
  if (weightOz <= 8) return { fee: 1.77, tier: "4+~8oz" };
  if (weightOz <= 12) return { fee: 2.6, tier: "8+~12oz" };
  if (weightOz <= 16) return { fee: 3.22, tier: "12+~16oz" };
  return { fee: 3.72, tier: "1lb以上" };
}

function breakEvenWithFee(saleUnitCostUsd, weightOz) {
  if (saleUnitCostUsd <= 0 || weightOz <= 0) {
    return { breakEven: 0, shippingFee: 0, shippingTier: weightOz <= 0 ? "待填写尺寸重量" : "待计算", fixedCostUsd: 0 };
  }
  let tier = fbaFeeForOz(weightOz, 2.99);
  let breakEven = 0;
  for (let index = 0; index < 4; index += 1) {
    const fixedCostUsd = firstLegShippingUsd + tier.fee + disposalFeeUsd * returnRate;
    breakEven = Math.max(0, (saleUnitCostUsd + fixedCostUsd) / breakEvenFactor);
    const nextTier = fbaFeeForOz(weightOz, breakEven);
    if (nextTier.fee === tier.fee) {
      return { breakEven, shippingFee: tier.fee, shippingTier: tier.tier, fixedCostUsd };
    }
    tier = nextTier;
  }
  const fixedCostUsd = firstLegShippingUsd + tier.fee + disposalFeeUsd * returnRate;
  return { breakEven, shippingFee: tier.fee, shippingTier: tier.tier, fixedCostUsd };
}

function saleQtyForRows(rows) {
  const bundleRow = rows.length === 1 && rows[0].bundle ? rows[0] : null;
  return bundleRow ? Math.max(0, num(bundleRow.bundleUnitCount, salePackQty())) : salePackQty();
}

function priceForTargetMargin(saleUnitCostUsd, weightOz, margin) {
  return targetPriceDetails(saleUnitCostUsd, weightOz, margin).suggestedPrice;
}

function targetPriceDetails(saleUnitCostUsd, weightOz, margin) {
  if (saleUnitCostUsd <= 0 || weightOz <= 0) {
    return {
      rawPrice: 0,
      suggestedPrice: 0,
      shippingFee: 0,
      shippingTier: "待计算",
      fixedCostUsd: 0
    };
  }
  let price = breakEvenWithFee(saleUnitCostUsd, weightOz).breakEven / Math.max(0.2, 1 - margin);
  for (let index = 0; index < 5; index += 1) {
    const fee = fbaFeeForOz(weightOz, price).fee;
    const fixedCostUsd = saleUnitCostUsd + firstLegShippingUsd + fee + disposalFeeUsd * returnRate;
    price = fixedCostUsd / Math.max(0.1, breakEvenFactor - margin);
  }
  const feeResult = fbaFeeForOz(weightOz, price);
  const fixedCostUsd = saleUnitCostUsd + firstLegShippingUsd + feeResult.fee + disposalFeeUsd * returnRate;
  const rawPrice = fixedCostUsd / Math.max(0.1, breakEvenFactor - margin);
  const suggestedPrice = roundUpToCents(rawPrice);
  const suggestedFeeResult = fbaFeeForOz(weightOz, suggestedPrice);
  const suggestedFixedCostUsd = saleUnitCostUsd + firstLegShippingUsd + suggestedFeeResult.fee + disposalFeeUsd * returnRate;
  return {
    rawPrice,
    suggestedPrice,
    shippingFee: suggestedFeeResult.fee,
    shippingTier: suggestedFeeResult.tier,
    fixedCostUsd: suggestedFixedCostUsd
  };
}

function suggestedDefaultPriceForRows(rows) {
  const saleQty = saleQtyForRows(rows);
  const dims = currentDimensions();
  if (saleQty <= 0 || dims.lengthCm <= 0 || dims.widthCm <= 0 || dims.heightCm <= 0 || dims.weightG <= 0) {
    return 0;
  }
  const saleUnitCostUsd = sellingUnitCostUsd(rows, saleQty);
  const weight = billableWeightFromDims(dims);
  const weightOz = weight.billableOz;
  return priceForTargetMargin(saleUnitCostUsd, weightOz, targetMarginNumber());
}

function suggestedDefaultPriceForRow(row, rows) {
  if (row?.bundle) {
    return suggestedDefaultPriceForRows([row]);
  }
  const saleQty = saleQtyForRows(rows);
  const dims = currentDimensions();
  if (!row || saleQty <= 0 || dims.lengthCm <= 0 || dims.widthCm <= 0 || dims.heightCm <= 0 || dims.weightG <= 0) {
    return 0;
  }
  const saleUnitCostUsd = num(row.cost) * saleQty / currencyRateRmbToUsd;
  const weight = billableWeightFromDims(dims);
  const weightOz = weight.billableOz;
  return priceForTargetMargin(saleUnitCostUsd, weightOz, targetMarginNumber());
}

function priceSourceForIndex(index) {
  const input = document.querySelectorAll("#finalPricingRows .price-input")[index];
  if (!input || !num(input.value)) {
    return "待填写";
  }
  return input.dataset.autoPrice === "true" ? "系统按目标利润率自动生成" : "价格确认表手动输入";
}

function currentPriceSource() {
  return priceSourceForIndex(0);
}

function refreshAutoPrices() {
  const rows = activePurchaseRows();
  if (!rows.length) {
    return;
  }
  document.querySelectorAll("#finalPricingRows .price-input").forEach((input, index) => {
    if (input.dataset.autoPrice === "true" || !num(input.value)) {
      const defaultPrice = suggestedDefaultPriceForRow(rows[index], rows);
      input.value = defaultPrice ? fmt(defaultPrice, 2) : "";
      input.dataset.autoPrice = "true";
    }
  });
  resetDownloadState();
  refreshReportDraft();
}

function rowPricingDetails(rows, saleQty, weightOz, margin) {
  const finalPriceRows = finalRows();
  return rows.map((row, index) => {
    const saleUnitCostUsd = row.bundle
      ? num(row.bundleSaleUnitCostRmb || row.cost) / currencyRateRmbToUsd
      : num(row.cost) * saleQty / currencyRateRmbToUsd;
    const breakEvenDetail = breakEvenWithFee(saleUnitCostUsd, weightOz);
    const targetDetail = targetPriceDetails(saleUnitCostUsd, weightOz, margin);
    const confirmedPrice = num(finalPriceRows[index]?.[6]) || targetDetail.suggestedPrice;
    const currentFee = fbaFeeForOz(weightOz, confirmedPrice);
    const returnDisposalReserve = disposalFeeUsd * returnRate;
    const currentFixedCostUsd = saleUnitCostUsd + firstLegShippingUsd + currentFee.fee + returnDisposalReserve;
    const currentProfit = confirmedPrice * breakEvenFactor - currentFixedCostUsd;
    const targetProfit = targetDetail.suggestedPrice * breakEvenFactor - targetDetail.fixedCostUsd;
    return {
      row,
      index,
      label: row.title || row.spec || `款式${index + 1}`,
      sku: row.sku || "",
      costRmb: row.bundle ? num(row.bundleSaleUnitCostRmb || row.cost) : num(row.cost),
      saleUnitCostUsd,
      breakEvenPrice: breakEvenDetail.breakEven,
      breakEvenShippingFee: breakEvenDetail.shippingFee,
      breakEvenShippingTier: breakEvenDetail.shippingTier,
      suggestedMinPrice: roundUpToCents(breakEvenDetail.breakEven),
      targetRawPrice: targetDetail.rawPrice,
      targetPrice: targetDetail.suggestedPrice,
      targetShippingFee: targetDetail.shippingFee,
      targetShippingTier: targetDetail.shippingTier,
      targetFixedCostUsd: targetDetail.fixedCostUsd,
      targetProfit,
      targetProfitMargin: targetDetail.suggestedPrice > 0 ? targetProfit / targetDetail.suggestedPrice : 0,
      confirmedPrice,
      priceSource: priceSourceForIndex(index),
      currentShippingFee: currentFee.fee,
      currentShippingTier: currentFee.tier,
      currentFixedCostUsd,
      currentProfit,
      currentProfitMargin: confirmedPrice > 0 ? currentProfit / confirmedPrice : 0
    };
  });
}

function groupedPricingDetails(rowDetails) {
  const groups = new Map();
  for (const item of rowDetails) {
    const key = [
      fmt(item.costRmb, 4),
      fmt(item.saleUnitCostUsd, 4),
      fmt(item.breakEvenShippingFee, 2),
      item.breakEvenShippingTier,
      fmt(item.targetPrice, 2),
      fmt(item.targetShippingFee, 2),
      item.targetShippingTier,
      fmt(item.confirmedPrice, 2),
      fmt(item.currentShippingFee, 2),
      item.currentShippingTier,
      item.priceSource
    ].join("|");
    if (!groups.has(key)) {
      groups.set(key, {
        ...item,
        items: []
      });
    }
    groups.get(key).items.push(item);
  }
  return [...groups.values()];
}

function pricingGroupLabel(group) {
  return group.items
    .map((item) => item.row?.title || item.label || `款式${item.index + 1}`)
    .join("；");
}

function reportData() {
  const rows = activePurchaseRows();
  const hasPurchaseRows = rows.length > 0;
  const bundleRow = rows.length === 1 && rows[0].bundle ? rows[0] : null;
  const dims = currentDimensions();
  const totalQty = rows.reduce((sum, row) => sum + num(row.quantity), 0);
  const totalCost = bundleRow
    ? num(bundleRow.quantity) * num(bundleRow.bundleSaleUnitCostRmb || bundleRow.cost)
    : rows.reduce((sum, row) => sum + num(row.quantity) * num(row.cost), 0);
  const avgCost = bundleRow
    ? num(bundleRow.bundleSaleUnitCostRmb || bundleRow.cost)
    : totalQty ? totalCost / totalQty : rows.reduce((sum, row) => sum + num(row.cost), 0) / Math.max(rows.length, 1);
  const price = suggestedPrice();
  const margin = targetMarginNumber();
  const productName = productLabel();
  const saleQty = bundleRow ? Math.max(1, num(bundleRow.bundleUnitCount, salePackQty())) : salePackQty();
  const unitQty = comparisonQty();
  const profile = inferProductProfile(productName, rows, activeCompetitors());
  const unitLabel = bundleRow?.bundleUnitLabel || profile.unitLabel;
  const saleLabel = bundleRow ? "套" : profile.saleUnitLabel;
  const pricingCostRow = bundleRow || rows.reduce((highest, row) => !highest || num(row.cost) > num(highest.cost) ? row : highest, null);
  const saleUnitCostUsd = sellingUnitCostUsd(rows, saleQty);
  const weight = billableWeightFromDims(dims);
  const weightOz = weight.billableOz;
  const returnDisposalReserve = disposalFeeUsd * returnRate;
  const breakEvenResult = breakEvenWithFee(saleUnitCostUsd, weightOz);
  const breakEvenShippingFee = breakEvenResult.shippingFee;
  const breakEvenShippingTier = breakEvenResult.shippingTier;
  const fixedCostUsd = breakEvenResult.fixedCostUsd;
  const breakEven = breakEvenResult.breakEven;
  const currentFeeResult = fbaFeeForOz(weightOz, price);
  const currentShippingFee = currentFeeResult.fee;
  const currentShippingTier = currentFeeResult.tier;
  const currentFixedCostUsd = saleUnitCostUsd + firstLegShippingUsd + currentShippingFee + returnDisposalReserve;
  const currentNetRevenue = price * breakEvenFactor;
  const currentProfit = currentNetRevenue - currentFixedCostUsd;
  const currentProfitMargin = price > 0 ? currentProfit / price : 0;
  const targetDetail = targetPriceDetails(saleUnitCostUsd, weightOz, margin);
  const targetPrice = targetDetail.suggestedPrice;
  const targetProfit = targetPrice * breakEvenFactor - targetDetail.fixedCostUsd;
  const targetProfitMargin = targetPrice > 0 ? targetProfit / targetPrice : 0;
  const suggestedMinPrice = roundUpToCents(breakEven);
  const competitors = activeCompetitors().map((competitor) => ({
    ...competitor,
    unitPrice: competitor.price / Math.max(1, competitor.packCount),
    comparisonPrice: competitor.price / Math.max(1, competitor.packCount) * unitQty
  }));
  const hasCompetitors = competitors.length > 0;
  const competitorCount = `${competitors.length} 个`;
  const lowestCompetitorPrice = hasCompetitors ? Math.min(...competitors.map((competitor) => competitor.price)) : 0;
  const currentUnitPrice = saleQty > 0 ? price / saleQty : 0;
  const currentComparisonPrice = currentUnitPrice * unitQty;
  const targetUnitPrice = saleQty > 0 ? targetPrice / saleQty : 0;
  const targetComparisonPrice = targetUnitPrice * unitQty;
  const lowestComparisonCompetitor = hasCompetitors
    ? competitors.reduce((lowest, competitor) => competitor.comparisonPrice < lowest.comparisonPrice ? competitor : lowest, competitors[0])
    : null;
  const lowestTotalCompetitor = hasCompetitors
    ? competitors.reduce((lowest, competitor) => competitor.price < lowest.price ? competitor : lowest, competitors[0])
    : null;
  const closestUnitCompetitor = hasCompetitors
    ? competitors.reduce((closest, competitor) => {
      return Math.abs(competitor.comparisonPrice - currentComparisonPrice) < Math.abs(closest.comparisonPrice - currentComparisonPrice) ? competitor : closest;
    }, competitors[0])
    : null;
  const competitorComparisonPrice = closestUnitCompetitor?.comparisonPrice || 0;
  const competitorUnitPrice = closestUnitCompetitor?.unitPrice || 0;
  const packagePriceGap = lowestTotalCompetitor ? price - lowestTotalCompetitor.price : 0;
  const unitPriceGap = closestUnitCompetitor ? currentComparisonPrice - competitorComparisonPrice : 0;
  const targetPackagePriceGap = lowestTotalCompetitor ? targetPrice - lowestTotalCompetitor.price : 0;
  const targetUnitPriceGap = lowestComparisonCompetitor ? targetComparisonPrice - lowestComparisonCompetitor.comparisonPrice : 0;
  const currentLowestUnitPriceGap = lowestComparisonCompetitor ? currentComparisonPrice - lowestComparisonCompetitor.comparisonPrice : 0;
  const unitPriceConclusion = !closestUnitCompetitor
    ? "请先上传并解析当前产品的竞品文件，再生成竞品价格结论。"
    : unitPriceGap <= 0
      ? `当前 ${unitQty}${unitLabel} 折算价不高于 ${closestUnitCompetitor.label} 竞品，价格位置有竞争力。`
      : `当前 ${unitQty}${unitLabel} 折算价高于 ${closestUnitCompetitor.label} 竞品，需要靠${profile.strategyNoun}或评价来支撑。`;
  const profitRows = rowProfitRows(rows, saleQty, price, currentShippingFee);
  const highCostProfit = profitRows.reduce((highest, row) => !highest || row.costUsd > highest.costUsd ? row : highest, null);
  const lowCostProfit = profitRows.reduce((lowest, row) => !lowest || row.costUsd < lowest.costUsd ? row : lowest, null);
  const rowDetails = rowPricingDetails(rows, saleQty, weightOz, margin);
  const pricingGroups = groupedPricingDetails(rowDetails);
  const postReviewPrice = price < 2.99 ? 2.99 : roundUpToCents(price + profile.reviewStep);

  return {
    profile,
    bundleRow,
    productName,
    rows,
    hasPurchaseRows,
    styleCount: rows.length,
    totalQty,
    totalCost,
    avgCost,
    dims,
    shippingFee: currentShippingFee,
    shippingTier: currentShippingTier,
    breakEvenShippingFee,
    breakEvenShippingTier,
    currentShippingFee,
    currentShippingTier,
    breakEven,
    targetPrice,
    targetRawPrice: targetDetail.rawPrice,
    targetShippingFee: targetDetail.shippingFee,
    targetShippingTier: targetDetail.shippingTier,
    targetFixedCostUsd: targetDetail.fixedCostUsd,
    targetProfit,
    targetProfitMargin,
    suggestedMinPrice,
    priceSource: currentPriceSource(),
    price,
    currentFixedCostUsd,
    currentNetRevenue,
    currentProfit,
    currentProfitMargin,
    margin,
    competitorCount,
    lowestCompetitorPrice,
    saleQty,
    unitQty,
    unitLabel,
    saleLabel,
    pricingCostRow,
    saleUnitCostUsd,
    returnDisposalReserve,
    fixedCostUsd,
    competitors,
    hasCompetitors,
    lowestTotalCompetitor,
    lowestComparisonCompetitor,
    closestUnitCompetitor,
    currentUnitPrice,
    currentComparisonPrice,
    targetUnitPrice,
    targetComparisonPrice,
    competitorUnitPrice,
    competitorComparisonPrice,
    packagePriceGap,
    unitPriceGap,
    targetPackagePriceGap,
    targetUnitPriceGap,
    currentLowestUnitPriceGap,
    unitPriceConclusion,
    rowDetails,
    pricingGroups,
    highCostProfit,
    lowCostProfit,
    postReviewPrice,
    weightOz,
    actualWeightOz: weight.actualOz,
    dimensionalWeightOz: weight.dimensionalOz
  };
}

function reportRows() {
  const data = reportData();
  const isBundle = Boolean(data.bundleRow);
  return [
    ["产品", data.productName, "按产品名称或识别表第一行款式标题生成"],
    [isBundle ? "销售SKU数" : "采购款数", data.styleCount, isBundle ? "套装组件已合并为父 SKU" : "来自采购单识别结果"],
    [isBundle ? "可售套数" : "采购总数", data.totalQty, isBundle ? "按组件采购数量和每套用量取最少可组成套数" : "来自采购单识别结果"],
    [isBundle ? "每套成本" : "平均成本", fmt(data.avgCost, 4), isBundle ? "按组件用量合计" : "按货品成本 / 采购总数估算"],
    ["销售单位成本", money(data.saleUnitCostUsd, 2), data.bundleRow ? `按套装每${data.saleLabel}成本折算` : `按最高成本款 ${data.pricingCostRow?.spec || data.pricingCostRow?.title || "当前款式"} × ${data.saleQty}${data.unitLabel} 保守折算`],
    ["当前售价配送费", money(data.currentShippingFee, 2), `当前确认价 ${money(data.price, 2)} 对应 ${data.currentShippingTier}`],
    ["保本测算配送费", money(data.breakEvenShippingFee, 2), `保本价 ${money(data.breakEven, 2)} 对应 ${data.breakEvenShippingTier}`],
    ["保本定价", money(data.breakEven, 2), "按截图里的保本公式折算"],
    ["建议不低于", money(data.suggestedMinPrice, 2), "向上取到两位小数"],
    ["目标利润建议价", money(data.targetPrice, 2), `目标利润率 ${fmt(data.margin * 100, 0)}%，原始价 ${money(data.targetRawPrice, 2)} 后向上取到两位小数`],
    ["目标价配送费", money(data.targetShippingFee, 2), `建议售价 ${money(data.targetPrice, 2)} 对应 ${data.targetShippingTier}`],
    ["目标价利润率", percent(data.targetProfitMargin), `利润 ${money(data.targetProfit, 2)} / 售价 ${money(data.targetPrice, 2)}`],
    ["当前确认价", money(data.price, 2), data.priceSource],
    ["当前确认价利润率", percent(data.currentProfitMargin), `利润 ${money(data.currentProfit, 2)} / 售价 ${money(data.price, 2)}`],
    ["本品售卖数量", `${data.saleQty}${data.unitLabel}/${data.saleLabel}`, "用于利润和保本计算"],
    ["竞品对比量", `${data.unitQty}${data.unitLabel}`, "用于竞品价格折算"],
    ["建议价折算价", money(data.targetComparisonPrice, 2), `建议售价 / ${data.saleQty}${data.unitLabel} × ${data.unitQty}${data.unitLabel}`],
    ["当前折算价", money(data.currentComparisonPrice, 2), `当前确认价 / ${data.saleQty}${data.unitLabel} × ${data.unitQty}${data.unitLabel}`],
    ["最低竞品锚点", data.lowestComparisonCompetitor?.label || "待上传", `${data.profile.competitorBasis}最低折算价`],
    ["建议价折算价差", signedMoney(data.targetUnitPriceGap, 3), data.hasCompetitors ? `相对 ${data.lowestComparisonCompetitor.label}` : "待上传竞品"],
    ["当前折算价差", signedMoney(data.currentLowestUnitPriceGap, 3), data.unitPriceConclusion],
    ...(data.pricingGroups.length > 1 ? [
      ["", "", ""],
      ["分组定价明细", "保本价 / 目标利润建议价 / 当前确认价", "不同成本、配送费档位或确认价条件才拆组"],
      ...data.pricingGroups.map((group) => [
        pricingGroupLabel(group),
        `成本 ${fmt(group.costRmb, 4)} RMB；销售单位成本 ${money(group.saleUnitCostUsd, 2)}；保本价 ${money(group.breakEvenPrice, 2)}；目标利润建议价 ${money(group.targetPrice, 2)}；当前确认价 ${money(group.confirmedPrice, 2)}`,
        `包含 ${group.items.length} 款；目标利润率 ${percent(group.targetProfitMargin)}；当前利润 ${money(group.currentProfit, 2)}；当前利润率 ${percent(group.currentProfitMargin)}`
      ])
    ] : [])
  ];
}

function reportText() {
  const data = reportData();
  const highCostLine = data.highCostProfit
    ? `按当前配送费模型，成本最高的 ${data.highCostProfit.title} 约有 ${money(data.highCostProfit.profit, 2)}/每${data.saleLabel} 利润；`
    : "";
  const lowCostLine = data.lowCostProfit
    ? `低成本款约有 ${money(data.lowCostProfit.profit, 2)}/每${data.saleLabel} 利润。`
    : "";
  return [
    `通用产品定价台 - ${data.productName} 定价分析报告`,
    "",
    "保本价计算过程",
    `尺寸：${fmt(data.dims.lengthCm, 2)} × ${fmt(data.dims.widthCm, 2)} × ${fmt(data.dims.heightCm, 2)} cm，实重 ${fmt(data.dims.weightG, 2)}g = ${fmt(data.actualWeightOz, 2)}oz；体积重 ${fmt(data.dimensionalWeightOz, 2)}oz；计费重 ${fmt(data.weightOz, 2)}oz。`,
    `保本价对应配送费：${data.breakEvenShippingTier}，配送费 ${money(data.breakEvenShippingFee, 2)}。`,
    `净入账系数 = (1 - ${fmt(returnRate * 100, 0)}%) × (1 - ${fmt(referralFeeRate * 100, 0)}%) - ${fmt(returnRate * 100, 0)}% × ${fmt(referralFeeRate * 100, 0)}% × ${fmt(refundCommissionLossRate * 100, 0)}% = ${fmt(breakEvenFactor, 5)}。`,
    data.bundleRow
      ? `保本售价使用套装成本：${money(data.saleUnitCostUsd, 2)}。`
      : `保本售价使用最高成本款 ${data.pricingCostRow?.spec || data.pricingCostRow?.title || "当前款式"}：${fmt(data.pricingCostRow?.cost, 4)} RMB × ${data.saleQty}${data.unitLabel} / ${currencyRateRmbToUsd} = ${money(data.saleUnitCostUsd, 2)}。`,
    `保本售价 = (销售单位成本 + ${fmt(firstLegShippingUsd, 2)} + ${fmt(data.breakEvenShippingFee, 2)} + ${fmt(data.returnDisposalReserve, 3)}) / ${fmt(breakEvenFactor, 5)}。`,
    `代入 ${money(data.saleUnitCostUsd, 2)}：(${money(data.saleUnitCostUsd, 2)} + ${money(data.fixedCostUsd, 3)}) / ${fmt(breakEvenFactor, 5)} = ${money(data.breakEven, 2)}。`,
    `建议不低于：${money(data.suggestedMinPrice, 2)}。`,
    "",
    "目标利润率下的建议售价",
    `目标利润率：${percent(data.margin)}。建议售价对应配送费：${data.targetShippingTier}，配送费 ${money(data.targetShippingFee, 2)}。`,
    `固定成本 = ${money(data.saleUnitCostUsd, 3)} + ${money(firstLegShippingUsd, 2)} + ${money(data.targetShippingFee, 2)} + ${money(data.returnDisposalReserve, 3)} = ${money(data.targetFixedCostUsd, 3)}。`,
    `目标利润价 = ${money(data.targetFixedCostUsd, 3)} / (${fmt(breakEvenFactor, 5)} - ${fmt(data.margin, 2)}) = ${money(data.targetRawPrice, 2)}；向上取到两位小数后，建议售价 ${money(data.targetPrice, 2)}。`,
    `建议售价利润 = ${money(data.targetPrice, 2)} × ${fmt(breakEvenFactor, 5)} - ${money(data.targetFixedCostUsd, 3)} = ${money(data.targetProfit, 2)}，利润率 ${percent(data.targetProfitMargin)}。`,
    "",
    "建议售价与竞品对比",
    data.hasCompetitors
      ? `竞品锚点我从你给的 ${data.competitors.length} 个页面快照里看到：`
      : "还没有解析到当前产品的竞品锚点。请上传当前产品的竞品 HTML 或文本文件。",
    ...(data.hasCompetitors ? data.competitors.map((item) => `${item.label}：${money(item.price, 2)} / ${item.packCount}${data.unitLabel}，单${data.unitLabel}价 ${money(item.unitPrice, 2)}/${data.unitLabel}，按 ${data.unitQty}${data.unitLabel} 折算为 ${money(item.comparisonPrice, 2)}，${item.rating} / ${item.sales}`) : []),
    data.hasCompetitors
      ? `系统建议售价 ${money(data.targetPrice, 2)}：按 ${data.unitQty}${data.unitLabel} 折算为 ${money(data.targetComparisonPrice, 2)}，比最低折算竞品 ${data.lowestComparisonCompetitor.label} ${signedMoney(data.targetUnitPriceGap, 2)}。整包价比最低整包竞品 ${data.lowestTotalCompetitor.label} ${signedMoney(data.targetPackagePriceGap, 2)}。`
      : `系统建议售价 ${money(data.targetPrice, 2)}：按 ${data.unitQty}${data.unitLabel} 折算为 ${money(data.targetComparisonPrice, 2)}。`,
    "",
    "当前/人工确认价利润率",
    `当前确认价：${money(data.price, 2)}（${data.priceSource}），对应配送费：${data.currentShippingTier}，配送费 ${money(data.currentShippingFee, 2)}。`,
    `当前固定成本 = ${money(data.saleUnitCostUsd, 3)} + ${money(firstLegShippingUsd, 2)} + ${money(data.currentShippingFee, 2)} + ${money(data.returnDisposalReserve, 3)} = ${money(data.currentFixedCostUsd, 3)}。`,
    `当前利润 = ${money(data.price, 2)} × ${fmt(breakEvenFactor, 5)} - ${money(data.currentFixedCostUsd, 3)} = ${money(data.currentProfit, 2)}。`,
    `当前利润率 = ${money(data.currentProfit, 2)} / ${money(data.price, 2)} = ${percent(data.currentProfitMargin)}。`,
    data.hasCompetitors
      ? `当前确认价按 ${data.unitQty}${data.unitLabel} 折算为 ${money(data.currentComparisonPrice, 2)}，比最低折算竞品 ${data.lowestComparisonCompetitor.label} ${signedMoney(data.currentLowestUnitPriceGap, 2)}。`
      : `当前确认价按 ${data.unitQty}${data.unitLabel} 折算为 ${money(data.currentComparisonPrice, 2)}。`,
    data.hasCompetitors ? `${highCostLine}${lowCostLine}` : "",
    ...(data.pricingGroups.length > 1 ? [
      "",
      "分组定价明细",
      ...data.pricingGroups.map((group) => [
        `包含款式：${pricingGroupLabel(group)}`,
        `成本 ${fmt(group.costRmb, 4)} RMB`,
        `销售单位成本 ${money(group.saleUnitCostUsd, 2)}`,
        `保本价 ${money(group.breakEvenPrice, 2)}（${group.breakEvenShippingTier} / 配送费 ${money(group.breakEvenShippingFee, 2)}）`,
        `目标利润建议价 ${money(group.targetPrice, 2)}（利润率 ${percent(group.targetProfitMargin)}）`,
        `当前确认价 ${money(group.confirmedPrice, 2)}（${group.priceSource}）`,
        `当前利润 ${money(group.currentProfit, 2)}，利润率 ${percent(group.currentProfitMargin)}`
      ].join("；"))
    ] : []),
    data.hasCompetitors
      ? `策略：首发 ${money(data.price, 2)} 先验证转化和评价，累计 20-30 个评价后测试 ${money(data.postReviewPrice, 2)}。不要只追 ${data.lowestComparisonCompetitor.label} 的 ${money(data.lowestComparisonCompetitor.comparisonPrice, 2)}/${data.unitQty}${data.unitLabel}，重点看${data.profile.valueAnchor}。${data.profile.launchFocus}。`
      : "策略：先补齐当前产品竞品锚点，再给最终价格策略。",
    "",
    "最终价格确认",
    ...finalRows().map((row) => row.join(" / "))
  ].join("\n");
}

function renderReportContent() {
  const data = reportData();
  if (!data.hasPurchaseRows) {
    $("reportContent").innerHTML = `
      <article class="process-block">
        <h3>配送费分析过程</h3>
        <p class="lead-text">请先上传或录入当前产品采购单，并确认尺寸重量。</p>
        <p>系统会在识别到采购成本、售卖数量、尺寸重量后，再按表内规则生成保本价、配送费和利润底线。</p>
      </article>

      <article class="process-block">
        <h3>竞品与利润分析过程</h3>
        <p class="lead-text">请上传当前产品的竞品 HTML / 文本文件。</p>
        <p>竞品价格会按“竞品售价 ÷ 竞品售卖数量 × 你填的对比量”折算，不再使用任何固定产品示例。</p>
      </article>
    `;
    renderCurrentProfitPanel(data);
    return;
  }
  const highCostProfit = data.highCostProfit ? money(data.highCostProfit.profit, 2) : "-";
  const lowCostProfit = data.lowCostProfit ? money(data.lowCostProfit.profit, 2) : "-";
  const highCostTitle = data.highCostProfit?.title || "成本最高款";
  const weightTierText = data.weightOz <= 4 ? "计费重落在 4oz 及以下" : `计费重对应 ${data.currentShippingTier} 档`;
  const bundleSummaryHtml = data.bundleRow
    ? `<li>套装结构：${escapeXml(bundleRowSummary(data.bundleRow))}。</li>`
    : "";
  const competitorRowsHtml = data.hasCompetitors
    ? data.competitors.map((competitor) => `
      <tr>
        <td>${escapeXml(competitor.label)}</td>
        <td><mark>${money(competitor.price, 2)}</mark> / ${competitor.packCount}${data.unitLabel}</td>
        <td><mark>${money(competitor.unitPrice, 2)}/${data.unitLabel}</mark></td>
        <td><mark>${money(competitor.comparisonPrice, 2)}</mark></td>
        <td>${escapeXml(competitor.rating)} / ${escapeXml(competitor.sales)}</td>
      </tr>
    `).join("")
    : '<tr><td colspan="5" class="empty">暂无当前产品竞品。请上传竞品 HTML / 文本后再生成报告。</td></tr>';
  const priceCompareRowsHtml = data.hasCompetitors
    ? `
      <tr>
        <td>系统建议售价</td>
        <td><mark>${money(data.targetPrice, 2)}</mark></td>
        <td><mark>${money(data.targetComparisonPrice, 2)}</mark></td>
        <td>${escapeXml(data.lowestComparisonCompetitor.label)}：<mark>${money(data.lowestComparisonCompetitor.comparisonPrice, 2)}</mark></td>
        <td><mark>${signedMoney(data.targetUnitPriceGap, 2)}</mark></td>
      </tr>
      <tr>
        <td>当前确认售价</td>
        <td><mark>${money(data.price, 2)}</mark></td>
        <td><mark>${money(data.currentComparisonPrice, 2)}</mark></td>
        <td>${escapeXml(data.lowestComparisonCompetitor.label)}：<mark>${money(data.lowestComparisonCompetitor.comparisonPrice, 2)}</mark></td>
        <td><mark>${signedMoney(data.currentLowestUnitPriceGap, 2)}</mark></td>
      </tr>
    `
    : `
      <tr>
        <td>系统建议售价</td>
        <td><mark>${money(data.targetPrice, 2)}</mark></td>
        <td><mark>${money(data.targetComparisonPrice, 2)}</mark></td>
        <td>待上传竞品</td>
        <td>-</td>
      </tr>
      <tr>
        <td>当前确认售价</td>
        <td><mark>${money(data.price, 2)}</mark></td>
        <td><mark>${money(data.currentComparisonPrice, 2)}</mark></td>
        <td>待上传竞品</td>
        <td>-</td>
      </tr>
    `;
  const pricingGroupRowsHtml = data.pricingGroups.length > 1
    ? data.pricingGroups.map((group) => `
      <tr>
        <td>${escapeXml(pricingGroupLabel(group))}</td>
        <td><mark>${fmt(group.costRmb, 4)} RMB</mark></td>
        <td><mark>${money(group.saleUnitCostUsd, 2)}</mark></td>
        <td>${money(group.breakEvenPrice, 2)}<br><span class="muted">${escapeXml(group.breakEvenShippingTier)} / ${money(group.breakEvenShippingFee, 2)}</span></td>
        <td>${money(group.targetPrice, 2)}<br><span class="muted">${escapeXml(group.targetShippingTier)} / 利润率 ${percent(group.targetProfitMargin)}</span></td>
        <td>${money(group.confirmedPrice, 2)}<br><span class="muted">${escapeXml(group.priceSource)}</span></td>
        <td><mark>${percent(group.currentProfitMargin)}</mark><br><span class="muted">利润 ${money(group.currentProfit, 2)}</span></td>
      </tr>
    `).join("")
    : "";
  const pricingGroupsSectionHtml = data.pricingGroups.length > 1 ? `
    <article class="process-block">
      <h3>分组定价明细</h3>
      <p class="lead-text">不同成本、配送费档位或确认价条件才拆组；包含款式使用识别结果里的款式标题。</p>
      <div class="comparison-table-wrap">
        <table class="comparison-table">
          <thead>
            <tr>
              <th>包含款式</th>
              <th>单件成本</th>
              <th>销售单位成本</th>
              <th>保本价</th>
              <th>目标利润建议价</th>
              <th>当前确认价</th>
              <th>当前利润率</th>
            </tr>
          </thead>
          <tbody>${pricingGroupRowsHtml}</tbody>
        </table>
      </div>
    </article>
  ` : "";
  $("reportContent").innerHTML = `
    <article class="process-block">
      <h3>保本价计算</h3>
      <p class="lead-text">${escapeXml(data.productName)} 的保本价是 <mark>${money(data.breakEven, 2)}</mark>，向上取到两位小数后建议不低于 <mark>${money(data.suggestedMinPrice, 2)}</mark>。</p>
      <ul>
        <li>尺寸：${fmt(data.dims.lengthCm, 2)} × ${fmt(data.dims.widthCm, 2)} × ${fmt(data.dims.heightCm, 2)} cm，实重 ${fmt(data.dims.weightG, 2)}g = ${fmt(data.actualWeightOz, 2)}oz；体积重 ${fmt(data.dimensionalWeightOz, 2)}oz；计费重 ${fmt(data.weightOz, 2)}oz，${escapeXml(weightTierText)}。</li>
        <li>保本价对应配送费：<mark>${escapeXml(data.breakEvenShippingTier)}</mark>，配送费 <mark>${money(data.breakEvenShippingFee, 2)}</mark>。</li>
        <li>销售单位成本：<mark>${money(data.saleUnitCostUsd, 2)}</mark>。${data.bundleRow ? "按套装成本折算。" : `按最高成本款 ${escapeXml(data.pricingCostRow?.spec || data.pricingCostRow?.title || "当前款式")}：${fmt(data.pricingCostRow?.cost, 4)} RMB × ${data.saleQty}${data.unitLabel} / ${currencyRateRmbToUsd} 保守折算。`}</li>
        ${bundleSummaryHtml}
        <li>头程 <mark>${money(firstLegShippingUsd, 2)}</mark>。</li>
        <li>退货预留：退货率 <mark>${fmt(returnRate * 100, 0)}%</mark> × 弃置费 <mark>${money(disposalFeeUsd, 2)}</mark> = <mark>${money(data.returnDisposalReserve, 3)}</mark>。</li>
        <li>净入账系数：<mark>(1 - ${fmt(returnRate * 100, 0)}%) × (1 - ${fmt(referralFeeRate * 100, 0)}%) - ${fmt(returnRate * 100, 0)}% × ${fmt(referralFeeRate * 100, 0)}% × ${fmt(refundCommissionLossRate * 100, 0)}% = ${fmt(breakEvenFactor, 5)}</mark>。</li>
      </ul>
      <div class="formula-box">
        <p>保本售价 = (销售单位成本 + 头程 + 保本价配送费 + 退货预留) / 净入账系数</p>
        <p><mark>(${money(data.saleUnitCostUsd, 2)} + ${money(firstLegShippingUsd, 2)} + ${money(data.breakEvenShippingFee, 2)} + ${money(data.returnDisposalReserve, 3)}) / ${fmt(breakEvenFactor, 5)} = ${money(data.breakEven, 2)}</mark></p>
      </div>
    </article>

    <article class="process-block">
      <h3>目标利润建议售价</h3>
      <p class="lead-text">目标利润率 <mark>${percent(data.margin)}</mark> 下，系统建议售价是 <mark>${money(data.targetPrice, 2)}</mark>。</p>
      <ul>
        <li>建议售价对应配送费：<mark>${escapeXml(data.targetShippingTier)}</mark>，配送费 <mark>${money(data.targetShippingFee, 2)}</mark>。</li>
        <li>建议售价固定成本：<mark>${money(data.saleUnitCostUsd, 3)} + ${money(firstLegShippingUsd, 2)} + ${money(data.targetShippingFee, 2)} + ${money(data.returnDisposalReserve, 3)} = ${money(data.targetFixedCostUsd, 3)}</mark>。</li>
      </ul>
      <div class="formula-box">
        <p>目标利润价 = 固定成本 / (净入账系数 - 目标利润率)</p>
        <p><mark>${money(data.targetFixedCostUsd, 3)} / (${fmt(breakEvenFactor, 5)} - ${fmt(data.margin, 2)}) = ${money(data.targetRawPrice, 2)}</mark></p>
        <p>向上取到两位小数后：<mark>${money(data.targetPrice, 2)}</mark>。</p>
        <p>建议售价利润 = 售价 × 净入账系数 - 固定成本 = <mark>${money(data.targetPrice, 2)} × ${fmt(breakEvenFactor, 5)} - ${money(data.targetFixedCostUsd, 3)} = ${money(data.targetProfit, 2)}</mark>，利润率 <mark>${percent(data.targetProfitMargin)}</mark>。</p>
      </div>
    </article>

    <article class="process-block">
      <h3>建议售价与竞品对比</h3>
      <p class="lead-text">${data.hasCompetitors ? `竞品锚点我从你给的 <mark>${data.competitors.length}</mark> 个页面快照里看到：` : "还没有解析到当前产品的竞品锚点。"}</p>
      <div class="comparison-table-wrap">
        <table class="comparison-table">
          <thead>
            <tr>
              <th>竞品</th>
              <th>售价/数量</th>
              <th>单${data.unitLabel}价</th>
              <th>${data.unitQty}${data.unitLabel}折算价</th>
              <th>评分/销量</th>
            </tr>
          </thead>
          <tbody>${competitorRowsHtml}</tbody>
        </table>
      </div>

      <p class="lead-text">按当前竞品对比量 <mark>${data.unitQty}${data.unitLabel}</mark>，建议价和确认价的位置如下：</p>
      <div class="comparison-table-wrap">
        <table class="comparison-table">
          <thead>
            <tr>
              <th>价格口径</th>
              <th>整包售价</th>
              <th>${data.unitQty}${data.unitLabel}折算价</th>
              <th>最低竞品锚点</th>
              <th>折算价差</th>
            </tr>
          </thead>
          <tbody>${priceCompareRowsHtml}</tbody>
        </table>
      </div>
      ${data.hasCompetitors ? `<p>整包价门槛：系统建议售价比最低整包竞品 ${escapeXml(data.lowestTotalCompetitor.label)} 的 ${money(data.lowestTotalCompetitor.price, 2)} ${signedMoney(data.targetPackagePriceGap, 2)}；当前确认价比最低整包竞品 ${signedMoney(data.packagePriceGap, 2)}。</p>` : ""}
    </article>

    ${pricingGroupsSectionHtml}
  `;
  renderCurrentProfitPanel(data);
}

function renderCurrentProfitPanel(data = reportData()) {
  const panel = $("currentProfitPanel");
  if (!panel) return;
  if (!data.hasPurchaseRows) {
    panel.innerHTML = `
      <article class="process-block">
        <h3>当前确认价利润率</h3>
        <p class="lead-text">等待最终价格行生成后，这里会显示当前确认价对应的利润率。</p>
      </article>
    `;
    return;
  }
  const highCostProfit = data.highCostProfit ? money(data.highCostProfit.profit, 2) : "-";
  const lowCostProfit = data.lowCostProfit ? money(data.lowCostProfit.profit, 2) : "-";
  const highCostTitle = data.highCostProfit?.title || "成本最高款";
  panel.innerHTML = `
    <article class="process-block">
      <h3>当前确认价利润率</h3>
      <p class="lead-text">当前确认价是 <mark>${money(data.price, 2)}</mark>（${escapeXml(data.priceSource)}）。如果你后期人工修改最终定价，这一块会按修改后的价格重新计算利润率。</p>
      <ul>
        <li>当前确认价对应配送费：<mark>${escapeXml(data.currentShippingTier)}</mark>，配送费 <mark>${money(data.currentShippingFee, 2)}</mark>。</li>
        <li>当前固定成本：<mark>${money(data.saleUnitCostUsd, 3)} + ${money(firstLegShippingUsd, 2)} + ${money(data.currentShippingFee, 2)} + ${money(data.returnDisposalReserve, 3)} = ${money(data.currentFixedCostUsd, 3)}</mark>。</li>
      </ul>
      <div class="formula-box">
        <p>当前利润 = 当前售价 × 净入账系数 - 当前固定成本</p>
        <p><mark>${money(data.price, 2)} × ${fmt(breakEvenFactor, 5)} - ${money(data.currentFixedCostUsd, 3)} = ${money(data.currentProfit, 2)}</mark></p>
        <p>当前利润率 = 当前利润 / 当前售价 = <mark>${money(data.currentProfit, 2)} / ${money(data.price, 2)} = ${percent(data.currentProfitMargin)}</mark></p>
      </div>
      <p>按当前售价配送费 <mark>${money(data.currentShippingFee, 2)}</mark> 计算，成本最高的 ${escapeXml(highCostTitle)} 约 <mark>${highCostProfit}/每${data.saleLabel}</mark> 利润；低成本款约 <mark>${lowCostProfit}/每${data.saleLabel}</mark>。</p>
      <p class="strategy-text">${data.hasCompetitors ? `我的策略：首发 <mark>${money(data.price, 2)}</mark> 先验证转化和评价，累计 20-30 个评价后测试 <mark>${money(data.postReviewPrice, 2)}</mark>。不要只追 ${escapeXml(data.lowestComparisonCompetitor.label)} 的 <mark>${money(data.lowestComparisonCompetitor.comparisonPrice, 2)}/${data.unitQty}${data.unitLabel}</mark>，重点看${escapeXml(data.profile.valueAnchor)}。${escapeXml(data.profile.launchFocus)}。` : "我的策略：先补齐当前产品竞品锚点，再输出最终价格策略。"}</p>
    </article>
  `;
}

function markReportStale() {
  const button = $("generateReportBtn");
  if (!button) return;
  button.textContent = "生成报告";
  button.classList.remove("is-confirmed");
}

function refreshReportDraft() {
  renderReportContent();
  markReportStale();
  renderPricingHistoryNav();
}

function setLink(link, href) {
  link.href = href;
  link.classList.remove("is-hidden");
}

function safeDownloadName(name) {
  return String(name || "产品").replace(/[\\/:*?"<>|\s]+/g, "_").replace(/^_+|_+$/g, "") || "产品";
}

function setSiteStatus(message, tone = "ok") {
  const status = $("siteStatus");
  status.textContent = message;
  status.className = `site-status ${tone}`.trim();
}

function smoothScrollTo(id) {
  const element = $(id);
  if (element) {
    element.scrollIntoView({ behavior: "smooth", block: "start" });
  }
}

function moneyRange(values) {
  const usable = values.filter((value) => Number.isFinite(value) && value > 0);
  if (!usable.length) {
    return "待计算";
  }
  const min = Math.min(...usable);
  const max = Math.max(...usable);
  return Math.abs(max - min) < 0.005 ? money(min, 2) : `${money(min, 2)}-${money(max, 2)}`;
}

function rmbRange(values, digits = 2) {
  const usable = values.filter((value) => Number.isFinite(value) && value > 0);
  if (!usable.length) {
    return "待计算";
  }
  const min = Math.min(...usable);
  const max = Math.max(...usable);
  const formatRmb = (value) => `¥${Number(value || 0).toFixed(digits).replace(/\.?0+$/g, "")}`;
  return Math.abs(max - min) < 0.0005 ? formatRmb(min) : `${formatRmb(min)}-${formatRmb(max)}`;
}

function formatHistoryValue(value, formatter) {
  if (typeof value === "number") {
    return formatter(value);
  }
  return String(value || "待补");
}

function parseEditableHistoryValue(value) {
  const text = String(value ?? "").trim();
  if (!text || text === "待补") {
    return "待补";
  }
  const parsed = Number(text.replace(/[^\d.-]/g, ""));
  return Number.isFinite(parsed) && !/[~-]/.test(text) ? parsed : text;
}

function loadPricingHistoryProducts() {
  try {
    const saved = JSON.parse(localStorage.getItem(pricingHistoryStorageKey) || "null");
    if (Array.isArray(saved)) {
      pricingHistoryProducts = saved.filter((item) => item && item.name);
    }
  } catch (error) {
    pricingHistoryProducts = [...defaultPricingHistoryProducts];
  }
  try {
    const overrides = JSON.parse(localStorage.getItem(outputPricingHistoryOverridesStorageKey) || "{}");
    outputPricingHistoryOverrides = overrides && typeof overrides === "object" && !Array.isArray(overrides) ? overrides : {};
  } catch (error) {
    outputPricingHistoryOverrides = {};
  }
  try {
    const hidden = JSON.parse(localStorage.getItem(hiddenOutputPricingHistoryStorageKey) || "[]");
    hiddenOutputPricingHistoryNames = new Set(Array.isArray(hidden) ? hidden : []);
  } catch (error) {
    hiddenOutputPricingHistoryNames = new Set();
  }
}

async function loadOutputPricingHistory() {
  try {
    const response = await fetch("/api/history");
    if (!response.ok) {
      throw new Error("历史记录接口暂时不可用");
    }
    const payload = await response.json();
    outputPricingHistoryRecords = Array.isArray(payload.records)
      ? payload.records.filter((item) => item && item.name && !item.error)
      : [];
    renderPricingHistoryNav();
  } catch (error) {
    outputPricingHistoryRecords = [];
    console.warn(error);
  }
}

function savePricingHistoryProducts() {
  localStorage.setItem(pricingHistoryStorageKey, JSON.stringify(pricingHistoryProducts));
}

function saveOutputPricingHistoryPrefs() {
  localStorage.setItem(outputPricingHistoryOverridesStorageKey, JSON.stringify(outputPricingHistoryOverrides));
  localStorage.setItem(hiddenOutputPricingHistoryStorageKey, JSON.stringify([...hiddenOutputPricingHistoryNames]));
}

function editPricingHistoryProduct(index) {
  editingPricingHistoryIndex = pricingHistoryProducts[index] ? index : null;
  editingOutputPricingHistoryKey = null;
  renderPricingHistoryNav();
}

function editOutputPricingHistoryProduct(key) {
  editingOutputPricingHistoryKey = key || null;
  editingPricingHistoryIndex = null;
  renderPricingHistoryNav();
}

function cancelPricingHistoryEdit() {
  editingPricingHistoryIndex = null;
  editingOutputPricingHistoryKey = null;
  renderPricingHistoryNav();
}

function savePricingHistoryEdit(index) {
  const item = pricingHistoryProducts[index];
  const card = document.querySelector(`[data-pricing-edit-form="${index}"]`);
  if (!item || !card) {
    return;
  }
  const name = card.querySelector('[data-edit-field="name"]')?.value.trim() || item.name;
  const purchaseCostRmb = card.querySelector('[data-edit-field="purchase"]')?.value || "待补";
  const shippingFee = card.querySelector('[data-edit-field="shipping"]')?.value || "待补";
  const finalPrice = card.querySelector('[data-edit-field="price"]')?.value || "待补";
  pricingHistoryProducts[index] = {
    name,
    purchaseCostRmb: parseEditableHistoryValue(purchaseCostRmb),
    shippingFee: parseEditableHistoryValue(shippingFee),
    finalPrice: parseEditableHistoryValue(finalPrice)
  };
  editingPricingHistoryIndex = null;
  savePricingHistoryProducts();
  renderPricingHistoryNav();
}

function saveOutputPricingHistoryEdit(key) {
  const card = [...document.querySelectorAll("[data-pricing-output-edit-form]")]
    .find((element) => element.dataset.pricingOutputEditForm === key);
  const source = outputPricingHistoryRecords.find((item) => item.name === key);
  if (!card || !source) {
    return;
  }
  const name = card.querySelector('[data-edit-field="name"]')?.value.trim() || source.name;
  const purchaseCostRmb = card.querySelector('[data-edit-field="purchase"]')?.value || "待补";
  const shippingFee = card.querySelector('[data-edit-field="shipping"]')?.value || "待补";
  const finalPrice = card.querySelector('[data-edit-field="price"]')?.value || "待补";
  outputPricingHistoryOverrides[key] = {
    name,
    purchaseCostRmb: parseEditableHistoryValue(purchaseCostRmb),
    shippingFee: parseEditableHistoryValue(shippingFee),
    finalPrice: parseEditableHistoryValue(finalPrice)
  };
  editingOutputPricingHistoryKey = null;
  saveOutputPricingHistoryPrefs();
  renderPricingHistoryNav();
}

function deletePricingHistoryProduct(index) {
  const item = pricingHistoryProducts[index];
  if (!item) {
    return;
  }
  if (!window.confirm(`删除“${item.name}”？`)) {
    return;
  }
  pricingHistoryProducts.splice(index, 1);
  if (editingPricingHistoryIndex === index) {
    editingPricingHistoryIndex = null;
  }
  savePricingHistoryProducts();
  renderPricingHistoryNav();
}

function deleteOutputPricingHistoryProduct(key) {
  const source = outputPricingHistoryRecords.find((item) => item.name === key);
  const displayName = outputPricingHistoryOverrides[key]?.name || source?.name || key;
  if (!source) {
    return;
  }
  if (!window.confirm(`从导航里删除“${displayName}”？结果文件会保留。`)) {
    return;
  }
  hiddenOutputPricingHistoryNames.add(key);
  if (editingOutputPricingHistoryKey === key) {
    editingOutputPricingHistoryKey = null;
  }
  saveOutputPricingHistoryPrefs();
  renderPricingHistoryNav();
}

function currentPricingNavItem() {
  if (!currentPurchaseRows.length) {
    return null;
  }
  try {
    const data = reportData();
    const prices = finalRows().map((row) => num(row[6])).filter((value) => value > 0);
    const fees = data.rowDetails?.map((item) => item.currentShippingFee) || [];
    const costs = data.rowDetails?.map((item) => item.costRmb) || currentPurchaseRows.map((row) => num(row.cost));
    return {
      name: data.productName || productLabel(),
      purchaseCostText: rmbRange(costs, 4),
      shippingFeeText: moneyRange(fees.length ? fees : [data.currentShippingFee]),
      finalPriceText: moneyRange(prices.length ? prices : [data.price || data.targetPrice]),
      current: true
    };
  } catch (error) {
    return null;
  }
}

function normalizedHistoryRecord(record) {
  const historyKey = record.name;
  const override = outputPricingHistoryOverrides[historyKey] || {};
  return {
    ...record,
    historyKey,
    name: override.name || record.name,
    purchaseCostText: formatHistoryValue(override.purchaseCostRmb ?? record.purchaseCostRmb, (value) => rmbRange([value], 4)),
    shippingFeeText: formatHistoryValue(override.shippingFee ?? record.shippingFee, (value) => money(value, 2)),
    finalPriceText: formatHistoryValue(override.finalPrice ?? record.finalPrice, (value) => money(value, 2)),
    current: false,
    firstCalculatedAt: Number(record.firstCalculatedAt || record.updatedAt || 0),
    orderSequence: Number(record.orderSequence || 0),
    updatedAt: Number(record.updatedAt || 0),
    source: record.source || "manual"
  };
}

function loadPricingHistoryRecord(record) {
  if (!record?.rows?.length) {
    setSiteStatus(`“${record?.name || "这条记录"}”只有摘要，没有找到完整结果文件。`, "warn");
    smoothScrollTo("reportPanel");
    return;
  }
  $("productName").value = record.name || "";
  currentProductMode = "standard";
  const dims = record.dimensions || {};
  [["lengthCm", dims.lengthCm], ["widthCm", dims.widthCm], ["heightCm", dims.heightCm], ["weightG", dims.weightG]].forEach(([id, value]) => {
    if ($(id) && value !== undefined && value !== null && value !== "") {
      $(id).value = value;
    }
  });
  if ($("salePackQty")) {
    $("salePackQty").value = record.salePackQty || 1;
  }
  if ($("comparisonUnitQty")) {
    $("comparisonUnitQty").value = record.comparisonQty || record.salePackQty || 1;
  }
  if ($("targetMargin") && record.targetMargin !== undefined && record.targetMargin !== null && record.targetMargin !== "") {
    const margin = num(record.targetMargin);
    $("targetMargin").value = margin > 0 && margin <= 1 ? `${fmt(margin * 100, 0)}%` : record.targetMargin;
  }
  renderRecognizedRows(record.rows);
  updateSummaryFromRows(record.rows);
  const priceInputs = document.querySelectorAll("#finalPricingRows .price-input");
  (record.finalRows || []).forEach((row, index) => {
    const input = priceInputs[index];
    if (input && row.price !== undefined && row.price !== null && row.price !== "") {
      input.value = fmt(row.price, 2);
      input.dataset.autoPrice = "false";
    }
  });
  currentCompetitors = (record.competitors || []).filter((item) => num(item.price) > 0).map((item, index) => ({
    label: item.label || shortCompetitorLabel(item.title, index),
    title: item.title || item.label || `竞品${index + 1}`,
    price: num(item.price),
    packCount: Math.max(1, num(item.packCount, 1)),
    rating: item.rating || "评分待确认",
    sales: item.sales || "销量待确认"
  }));
  $("competitorFileSummary").textContent = currentCompetitors.length
    ? `已从历史结果载入 ${currentCompetitors.length} 个竞品。`
    : "这条历史记录没有竞品数据。";
  $("competitorParseStatus").textContent = "历史定价记录已载入。";
  $("competitorStatus").textContent = currentCompetitors.length ? "历史记录" : "待上传";
  $("competitorStatus").classList.toggle("ready", currentCompetitors.length > 0);
  updateCompetitorSummary(currentCompetitors.length);
  renderReportContent();
  $("baseStatus").textContent = "历史记录";
  $("baseStatus").classList.add("ready");
  $("generateReportBtn").textContent = "历史报告已载入";
  $("generateReportBtn").classList.add("is-confirmed");
  resetDownloadState();
  renderPricingHistoryNav();
  setSiteStatus(`已载入“${record.name}”的历史定价结果。`);
  smoothScrollTo("reportPanel");
}

function renderPricingHistoryNav() {
  const list = $("pricingHistoryList");
  if (!list) {
    return;
  }
  const currentItem = currentPricingNavItem();
  const currentName = currentItem?.name || "";
  const outputItems = outputPricingHistoryRecords
    .filter((item) => !hiddenOutputPricingHistoryNames.has(item.name))
    .map((item) => {
      const normalized = normalizedHistoryRecord(item);
      return {
        ...normalized,
        current: normalized.name === currentName || item.name === currentName
      };
    })
    .sort((a, b) => (
      Number(b.firstCalculatedAt || 0) - Number(a.firstCalculatedAt || 0)
      || Number(b.orderSequence || 0) - Number(a.orderSequence || 0)
    ));
  const hasCurrentOutputItem = outputItems.some((item) => item.current);
  const usedNames = new Set([...outputItems.map((item) => item.name)].filter(Boolean));
  const historyItems = pricingHistoryProducts
    .map((item, sourceIndex) => ({ item, sourceIndex }))
    .filter(({ item }) => !usedNames.has(item.name))
    .map(({ item, sourceIndex }) => ({
      name: item.name,
      purchaseCostText: formatHistoryValue(item.purchaseCostRmb, (value) => rmbRange([value], 4)),
      shippingFeeText: formatHistoryValue(item.shippingFee, (value) => money(value, 2)),
      finalPriceText: formatHistoryValue(item.finalPrice, (value) => money(value, 2)),
      current: false,
      source: "manual",
      sourceIndex,
      editing: sourceIndex === editingPricingHistoryIndex
    }));
  const items = currentItem && !hasCurrentOutputItem ? [currentItem, ...outputItems, ...historyItems] : [...outputItems, ...historyItems];
  if (!items.length) {
    list.innerHTML = '<p class="empty-nav">暂无已完成的定价分析。</p>';
    return;
  }
  list.innerHTML = items.map((item, index) => `
    <div class="pricing-history-item ${item.current ? "is-current" : ""}">
      <button class="pricing-history-main" type="button" data-pricing-nav="${index}">
        <span class="pricing-history-title">${escapeXml(item.current ? `${item.name}（当前）` : item.name)}</span>
        <span class="pricing-history-meta">
          <span>采购价<strong>${escapeXml(item.purchaseCostText)}</strong></span>
          <span>配送费<strong>${escapeXml(item.shippingFeeText)}</strong></span>
          <span>最终售价<strong>${escapeXml(item.finalPriceText)}</strong></span>
        </span>
      </button>
      ${item.current && item.source !== "output" ? "" : `
        ${(item.source === "output" ? item.historyKey === editingOutputPricingHistoryKey : item.editing) ? `
          <span class="pricing-history-edit-form" ${item.source === "output" ? `data-pricing-output-edit-form="${escapeAttr(item.historyKey)}"` : `data-pricing-edit-form="${item.sourceIndex}"`}>
            <label>产品<input data-edit-field="name" type="text" value="${escapeAttr(item.name)}"></label>
            <label>采购价<input data-edit-field="purchase" type="text" value="${escapeAttr(item.purchaseCostText)}"></label>
            <label>配送费<input data-edit-field="shipping" type="text" value="${escapeAttr(item.shippingFeeText)}"></label>
            <label>售价<input data-edit-field="price" type="text" value="${escapeAttr(item.finalPriceText)}"></label>
            <span class="pricing-history-actions">
              <button type="button" class="mini-button primary-mini" ${item.source === "output" ? `data-pricing-output-save="${escapeAttr(item.historyKey)}"` : `data-pricing-save="${item.sourceIndex}"`}>保存</button>
              <button type="button" class="mini-button" data-pricing-cancel>取消</button>
            </span>
          </span>
        ` : `
          <span class="pricing-history-actions">
            <button type="button" class="mini-button" ${item.source === "output" ? `data-pricing-output-edit="${escapeAttr(item.historyKey)}"` : `data-pricing-edit="${item.sourceIndex}"`}>修改</button>
            <button type="button" class="mini-button danger" ${item.source === "output" ? `data-pricing-output-delete="${escapeAttr(item.historyKey)}"` : `data-pricing-delete="${item.sourceIndex}"`}>删除</button>
          </span>
        `}
      `}
    </div>
  `).join("");
  document.querySelectorAll("[data-pricing-nav]").forEach((button, index) => {
    button.addEventListener("click", () => {
      const item = items[index];
      if (item.current) {
        smoothScrollTo("finalStage");
        return;
      }
      if (item.source === "output") {
        loadPricingHistoryRecord(item);
        return;
      }
      setSiteStatus(`“${item.name}”只有导航摘要，没有找到完整结果文件。`, "warn");
      smoothScrollTo("reportPanel");
    });
  });
  document.querySelectorAll("[data-pricing-edit]").forEach((button) => {
    button.addEventListener("click", () => editPricingHistoryProduct(Number(button.dataset.pricingEdit)));
  });
  document.querySelectorAll("[data-pricing-output-edit]").forEach((button) => {
    button.addEventListener("click", () => editOutputPricingHistoryProduct(button.dataset.pricingOutputEdit));
  });
  document.querySelectorAll("[data-pricing-save]").forEach((button) => {
    button.addEventListener("click", () => savePricingHistoryEdit(Number(button.dataset.pricingSave)));
  });
  document.querySelectorAll("[data-pricing-output-save]").forEach((button) => {
    button.addEventListener("click", () => saveOutputPricingHistoryEdit(button.dataset.pricingOutputSave));
  });
  document.querySelectorAll("[data-pricing-cancel]").forEach((button) => {
    button.addEventListener("click", cancelPricingHistoryEdit);
  });
  document.querySelectorAll("[data-pricing-delete]").forEach((button) => {
    button.addEventListener("click", () => deletePricingHistoryProduct(Number(button.dataset.pricingDelete)));
  });
  document.querySelectorAll("[data-pricing-output-delete]").forEach((button) => {
    button.addEventListener("click", () => deleteOutputPricingHistoryProduct(button.dataset.pricingOutputDelete));
  });
}

function setOcrStatus(message, isBusy = false) {
  const status = $("ocrStatus");
  const button = $("runOcrBtn");
  if (status) {
    status.textContent = message;
    status.classList.toggle("is-busy", isBusy);
  }
  if (button) {
    button.disabled = isBusy || !currentImageFile;
  }
}

function setPurchasePreview(html = "") {
  const preview = $("purchasePreview");
  if (!preview) return;
  preview.innerHTML = html;
  preview.classList.toggle("is-hidden", !html);
}

function showFilePreview(file, suffix) {
  if (currentObjectUrl) {
    URL.revokeObjectURL(currentObjectUrl);
    currentObjectUrl = "";
  }
  currentImageFile = null;
  if (["png", "jpg", "jpeg", "webp"].includes(suffix)) {
    currentImageFile = file;
    currentObjectUrl = URL.createObjectURL(file);
    setPurchasePreview(`<img src="${currentObjectUrl}" alt="采购单图片预览">`);
    setOcrStatus("图片已选择，准备识别");
    return;
  }
  if (suffix === "pdf") {
    setPurchasePreview(`<p>已选择 PDF：${escapeXml(file.name)}。当前前端会保留文件名；PDF 明细请改用手动添加或批量添加采购行继续。</p>`);
    setOcrStatus("PDF 暂不支持 OCR");
    return;
  }
  setOcrStatus("当前文件不需要 OCR");
  setPurchasePreview("");
}

function otsuThreshold(grays) {
  const histogram = new Array(256).fill(0);
  for (const gray of grays) {
    histogram[gray] += 1;
  }
  const total = grays.length;
  let sum = 0;
  for (let level = 0; level < 256; level += 1) {
    sum += level * histogram[level];
  }
  let backgroundWeight = 0;
  let backgroundSum = 0;
  let maxVariance = 0;
  let threshold = 180;
  for (let level = 0; level < 256; level += 1) {
    backgroundWeight += histogram[level];
    if (!backgroundWeight) continue;
    const foregroundWeight = total - backgroundWeight;
    if (!foregroundWeight) break;
    backgroundSum += level * histogram[level];
    const backgroundMean = backgroundSum / backgroundWeight;
    const foregroundMean = (sum - backgroundSum) / foregroundWeight;
    const variance = backgroundWeight * foregroundWeight * (backgroundMean - foregroundMean) ** 2;
    if (variance > maxVariance) {
      maxVariance = variance;
      threshold = level;
    }
  }
  return Math.max(145, Math.min(215, threshold + 18));
}

function eraseTableRules(data, width, height) {
  const darkRows = new Array(height).fill(0);
  const darkColumns = new Array(width).fill(0);
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const offset = (y * width + x) * 4;
      if (data[offset] < 128) {
        darkRows[y] += 1;
        darkColumns[x] += 1;
      }
    }
  }
  const rowsToErase = darkRows
    .map((count, index) => count > width * 0.45 ? index : -1)
    .filter((index) => index >= 0);
  const columnsToErase = darkColumns
    .map((count, index) => count > height * 0.35 ? index : -1)
    .filter((index) => index >= 0);

  const whitenPixel = (x, y) => {
    if (x < 0 || x >= width || y < 0 || y >= height) return;
    const offset = (y * width + x) * 4;
    data[offset] = 255;
    data[offset + 1] = 255;
    data[offset + 2] = 255;
  };

  for (const y of rowsToErase) {
    for (let dy = -1; dy <= 1; dy += 1) {
      for (let x = 0; x < width; x += 1) {
        whitenPixel(x, y + dy);
      }
    }
  }
  for (const x of columnsToErase) {
    for (let dx = -1; dx <= 1; dx += 1) {
      for (let y = 0; y < height; y += 1) {
        whitenPixel(x + dx, y);
      }
    }
  }
}

function imageFileToCanvas(file, maxWidth = 2800) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    const objectUrl = URL.createObjectURL(file);
    image.onload = () => {
      URL.revokeObjectURL(objectUrl);
      const targetWidth = 1800;
      const scale = Math.min(maxWidth / image.naturalWidth, Math.max(1, targetWidth / image.naturalWidth));
      const canvas = document.createElement("canvas");
      canvas.width = Math.max(1, Math.round(image.naturalWidth * scale));
      canvas.height = Math.max(1, Math.round(image.naturalHeight * scale));
      const ctx = canvas.getContext("2d", { willReadFrequently: true });
      ctx.imageSmoothingEnabled = false;
      ctx.drawImage(image, 0, 0, canvas.width, canvas.height);

      const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
      const data = imageData.data;
      const grays = [];
      for (let index = 0; index < data.length; index += 4) {
        const gray = data[index] * 0.299 + data[index + 1] * 0.587 + data[index + 2] * 0.114;
        grays.push(Math.round(gray));
      }
      const threshold = otsuThreshold(grays);
      for (let index = 0; index < data.length; index += 4) {
        const gray = grays[index / 4];
        const boosted = gray > threshold ? 255 : 0;
        data[index] = boosted;
        data[index + 1] = boosted;
        data[index + 2] = boosted;
      }
      eraseTableRules(data, canvas.width, canvas.height);
      ctx.putImageData(imageData, 0, 0);
      resolve(canvas);
    };
    image.onerror = () => {
      URL.revokeObjectURL(objectUrl);
      reject(new Error("图片读取失败"));
    };
    image.src = objectUrl;
  });
}

async function getOcrWorker() {
  await loadTesseractScript();
  if (!window.Tesseract) {
    throw new Error("OCR 库没有加载成功，请确认网络后刷新页面");
  }
  if (ocrWorker && ocrReady) {
    return ocrWorker;
  }
  setOcrStatus("正在加载 OCR 识别库 0%", true);
  ocrWorker = await Tesseract.createWorker("chi_sim+eng", 1, {
    workerPath: "https://cdn.jsdelivr.net/npm/tesseract.js@7/dist/worker.min.js",
    corePath: "https://cdn.jsdelivr.net/npm/tesseract.js-core@7/tesseract-core-simd.wasm.js",
    langPath: "https://tessdata.projectnaptha.com/4.0.0",
    logger: (message) => {
      if (message.status) {
        const progress = Number.isFinite(message.progress) ? ` ${Math.round(message.progress * 100)}%` : "";
        setOcrStatus(`${message.status}${progress}`, true);
      }
    }
  });
  await ocrWorker.setParameters({
    tessedit_pageseg_mode: "6",
    preserve_interword_spaces: "1"
  });
  ocrReady = true;
  return ocrWorker;
}

function loadTesseractScript() {
  if (window.Tesseract) {
    return Promise.resolve();
  }
  if (tesseractLoadPromise) {
    return tesseractLoadPromise;
  }
  tesseractLoadPromise = new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "https://cdn.jsdelivr.net/npm/tesseract.js@7/dist/tesseract.min.js";
    script.onload = () => resolve();
    script.onerror = () => reject(new Error("浏览器 OCR 脚本加载失败"));
    document.head.appendChild(script);
  });
  return tesseractLoadPromise;
}

async function runServerOcr(file) {
  if (!/^https?:/.test(window.location.protocol)) {
    throw new Error("当前不是本地服务页面");
  }
  const formData = new FormData();
  formData.append("file", file);
  const response = await fetch("/api/ocr", {
    method: "POST",
    body: formData
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok || !payload.ok) {
    throw new Error(payload.error || "服务端 OCR 失败");
  }
  return payload.text || "";
}

function normalizeOcrText(text) {
  return normalizeOrderText(text)
    .replace(/[|｜]/g, " ")
    .replace(/[。]/g, ".")
    .replace(/规\s*格\s*型\s*号/g, "规格型号")
    .replace(/款\s*式/g, "款式")
    .replace(/([A-Z])\]\s*-\s*(\d+)/gi, "$1J-$2")
    .replace(/\bH\]\s*(\d+)/gi, "HJ-$1")
    .replace(/(\d)\s+00ml/gi, "$100ml")
    .replace(/货\s*品\s*名\s*称/g, "货品名称")
    .replace(/实\s*付\s*款/g, "实付款")
    .replace(/元\s*\/\s*PCS/gi, "")
    .replace(/[ \t]{2,}/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

async function runImageOcr({ autoParse = true } = {}) {
  if (!currentImageFile) {
    setSiteStatus("请先上传图片采购单。", "warn");
    setOcrStatus("未选择图片");
    return;
  }
  try {
    renderRecognizedRows([]);
    updateSummaryFromRows([]);
    $("baseStatus").textContent = "OCR 中";
    $("baseStatus").classList.remove("ready");
    setSiteStatus("正在识别图片文字，第一次加载中文 OCR 会稍慢。", "warn");
    setOcrStatus("正在预处理图片", true);

    let rawText = "";
    try {
      setOcrStatus("正在调用本地 OCR", true);
      rawText = await runServerOcr(currentImageFile);
    } catch (serverError) {
      setOcrStatus("本地 OCR 不可用，改用浏览器 OCR", true);
      const canvas = await imageFileToCanvas(currentImageFile);
      const worker = await getOcrWorker();
      setOcrStatus("正在识别图片文字", true);
      const result = await worker.recognize(canvas);
      rawText = result.data?.text || "";
    }
    const text = normalizeOcrText(rawText);

    if (!text) {
      setSiteStatus("图片 OCR 没有识别到文字，请换更清晰截图或手动输入。", "warn");
      setOcrStatus("未识别到文字");
      $("baseStatus").textContent = "待补录";
      return;
    }

    if (!autoParse) {
      setSiteStatus("图片文字已识别。");
      setOcrStatus("OCR 完成");
      $("baseStatus").textContent = "待解析";
      return;
    }

    const rows = parsePurchaseOrderText(text, "ocr.txt");
    if (!rows.length) {
      setSiteStatus("图片文字已识别，但没有自动解析出采购行；请检查 OCR 文本或手动添加采购行。", "warn");
      setOcrStatus("OCR 完成，待人工复核");
      $("baseStatus").textContent = "待补录";
      return;
    }

    renderRecognizedRows(rows);
    updateSummaryFromRows(rows);
    $("baseStatus").textContent = "待确认";
    $("baseStatus").classList.remove("ready");
    setSiteStatus(`OCR 已识别并解析出 ${rows.length} 款产品，请检查表格后确认。`);
    setOcrStatus(`OCR 完成：${rows.length} 款`);
  } catch (error) {
    setSiteStatus(`图片 OCR 失败：${error.message}`, "warn");
    setOcrStatus("OCR 失败");
    $("baseStatus").textContent = "待补录";
  }
}

function generateFinalExcel() {
  const xml = buildExcelXml("最终价格确认", finalHeaders, finalRows());
  finalExcelUrl = makeDownload(xml, "application/vnd.ms-excel;charset=utf-8", finalExcelUrl);
  $("downloadExcelLink").download = `${safeDownloadName(productLabel())}_最终价格确认表.xls`;
  setLink($("downloadExcelLink"), finalExcelUrl);
  $("exportStatus").textContent = `已生成 ${finalRows().length} 条 SKU 的 Excel 表格。`;
  renderPricingHistoryNav();
  setSiteStatus("最终价格确认表已生成，请点击底部的“下载 Excel”。");
  $("downloadExcelLink").scrollIntoView({ behavior: "smooth", block: "center" });
}

function generateReportExcel() {
  renderReportContent();
  const xml = buildExcelXml("定价分析报告", reportHeaders, reportRows());
  reportExcelUrl = makeDownload(xml, "application/vnd.ms-excel;charset=utf-8", reportExcelUrl);
  $("downloadReportExcelLink").download = `${safeDownloadName(productLabel())}_定价分析报告.xls`;
  setLink($("downloadReportExcelLink"), reportExcelUrl);
  setSiteStatus("定价分析 Excel 已生成，请点击“下载分析 Excel”。");
}

function generateReportText() {
  renderReportContent();
  reportTextUrl = makeDownload(reportText(), "text/plain;charset=utf-8", reportTextUrl);
  $("downloadReportTextLink").download = `${safeDownloadName(productLabel())}_定价分析报告.txt`;
  setLink($("downloadReportTextLink"), reportTextUrl);
  setSiteStatus("定价分析报告文本已生成，请点击“下载报告文本”。");
}

function syncSkuToFinal(index, value) {
  const target = document.querySelector(`[data-final-sku="${index}"]`);
  if (target) {
    target.textContent = value.trim() || "-";
  }
}

function syncTitleToFinal(index, value) {
  const target = document.querySelector(`[data-final-title="${index}"]`);
  if (target) {
    target.textContent = value.trim() || "-";
  }
}

function updatePurchaseRow(index, field, value) {
  const row = currentPurchaseRows[Number(index)];
  if (!row) {
    return;
  }
  row[field] = value;
  if (field === "quantity" || field === "cost") {
    updateSummaryFromRows(currentPurchaseRows);
    renderFinalRows(currentPurchaseRows);
    initPriceInputs();
  }
  if (field === "sku") {
    syncSkuToFinal(index, value);
  }
  if (field === "title" || field === "spec") {
    syncTitleToFinal(index, finalTitle(row));
  }
  resetDownloadState();
  refreshReportDraft();
  saveWorkbenchDraft();
}

function currentDimensions() {
  return {
    lengthCm: num($("lengthCm").value, 0),
    widthCm: num($("widthCm").value, 0),
    heightCm: num($("heightCm").value, 0),
    weightG: num($("weightG").value, 0)
  };
}

function bundleComponentInputs() {
  return [...document.querySelectorAll(".bundle-component-row")].map((row, index) => ({
    name: row.querySelector(".bundle-component-name")?.value.trim() || `款式${index + 1}`,
    qtyPerBundle: num(row.querySelector(".bundle-component-qty")?.value, 0),
    unitCostRmb: num(row.querySelector(".bundle-component-cost")?.value, 0),
    purchaseQty: num(row.querySelector(".bundle-component-purchase")?.value, 0)
  })).filter((item) => item.qtyPerBundle > 0 || item.unitCostRmb > 0 || item.purchaseQty > 0 || item.name);
}

function bundleCostRmb(components) {
  return components.reduce((sum, item) => sum + item.qtyPerBundle * item.unitCostRmb, 0);
}

function bundleUnitCount(components) {
  return components.reduce((sum, item) => sum + item.qtyPerBundle, 0);
}

function bundleAvailableQty(components) {
  const counts = components
    .filter((item) => item.qtyPerBundle > 0 && item.purchaseQty > 0)
    .map((item) => Math.floor(item.purchaseQty / item.qtyPerBundle));
  return counts.length ? Math.min(...counts) : 0;
}

function bundleSpecText(components, unitLabel) {
  const unitCount = bundleUnitCount(components);
  const uniqueQty = [...new Set(components.map((item) => item.qtyPerBundle).filter((value) => value > 0))];
  if (components.length && uniqueQty.length === 1) {
    return `${components.length}款各${fmt(uniqueQty[0], 0)}${unitLabel}，共${fmt(unitCount, 0)}${unitLabel}/套`;
  }
  return `${components.length}款组合，共${fmt(unitCount, 0)}${unitLabel}/套`;
}

function bundleRowFromInputs() {
  const components = bundleComponentInputs().filter((item) => item.qtyPerBundle > 0);
  const unitLabel = $("bundleUnitLabel").value.trim() || "片";
  const unitCount = bundleUnitCount(components);
  const costRmb = bundleCostRmb(components);
  if (!components.length) {
    throw new Error("请至少填写 1 个套装组件和每套用量");
  }
  if (unitCount <= 0) {
    throw new Error("套装每套总数量必须大于 0");
  }
  if (costRmb <= 0) {
    throw new Error("请填写组件单件成本，系统需要计算每套成本");
  }
  const sku = $("bundleSku").value.trim() || `Bundle-${components.length}Style-${fmt(unitCount, 0)}pcs`;
  const title = $("bundleTitle").value.trim() || $("productName").value.trim() || "组合套装";
  const spec = $("bundleSpec").value.trim() || bundleSpecText(components, unitLabel);
  const quantity = bundleAvailableQty(components);
  return {
    sku,
    itemCode: "",
    title,
    spec,
    quantity: quantity || "",
    cost: costRmb,
    bundle: true,
    bundleUnitLabel: unitLabel,
    bundleUnitCount: unitCount,
    bundleSaleUnitCostRmb: costRmb,
    averageUnitCostRmb: costRmb / unitCount,
    components
  };
}

function bundleRowSummary(row) {
  if (!row?.bundle) {
    return "";
  }
  const componentText = row.components
    .map((item) => `${item.name}${fmt(item.qtyPerBundle, 0)}${row.bundleUnitLabel}`)
    .join(" + ");
  const availableText = row.quantity ? `；可组成 ${row.quantity} 套` : "";
  return `${componentText}；每套成本 ${fmt(row.bundleSaleUnitCostRmb, 4)} RMB；平均 ${fmt(row.averageUnitCostRmb, 4)} RMB/${row.bundleUnitLabel}${availableText}`;
}

function updateBundlePreview() {
  const preview = $("bundlePreview");
  if (!preview) return;
  try {
    const components = bundleComponentInputs().filter((item) => item.qtyPerBundle > 0);
    if (!components.length) {
      preview.textContent = "填写组件后生成一条父 SKU 套装行。";
      return;
    }
    const unitLabel = $("bundleUnitLabel").value.trim() || "片";
    const unitCount = bundleUnitCount(components);
    const costRmb = bundleCostRmb(components);
    const available = bundleAvailableQty(components);
    const average = unitCount ? costRmb / unitCount : 0;
    preview.textContent = `当前合计 ${fmt(unitCount, 0)}${unitLabel}/套，每套成本 ${fmt(costRmb, 4)} RMB，平均 ${fmt(average, 4)} RMB/${unitLabel}${available ? `，可组成 ${available} 套` : ""}。`;
  } catch (error) {
    preview.textContent = error.message;
  }
}

function resetTransientWorkbenchState() {
  currentPurchaseRows = [];
  currentCompetitors = [];
  currentProductMode = "standard";
  ["productName", "manualBulkRows", "manualSku", "manualItemCode", "manualTitle", "manualSpec", "manualQty", "manualCost"].forEach((id) => {
    const input = $(id);
    if (input) input.value = "";
  });
  ["lengthCm", "widthCm", "heightCm", "weightG"].forEach((id) => {
    const input = $(id);
    if (input) input.value = "0";
  });
  ["purchaseOrderFile", "competitorFiles"].forEach((id) => {
    const input = $(id);
    if (input) input.value = "";
  });
  document.querySelectorAll("input, textarea").forEach((input) => {
    input.setAttribute("autocomplete", "off");
  });
}

function currentFinalPriceDraft() {
  return [...document.querySelectorAll("#finalPricingRows .price-input")].map((input) => ({
    value: input.value,
    autoPrice: input.dataset.autoPrice !== "false"
  }));
}

function saveWorkbenchDraft() {
  if (isRestoringWorkbenchDraft) {
    return;
  }
  const draft = {
    savedAt: Date.now(),
    productMode: currentProductMode,
    productName: $("productName")?.value || "",
    dimensions: currentDimensions(),
    salePackQty: $("salePackQty")?.value || "1",
    comparisonUnitQty: $("comparisonUnitQty")?.value || "1",
    targetMargin: $("targetMargin")?.value || "15%",
    purchaseRows: currentPurchaseRows,
    competitors: currentCompetitors,
    finalPrices: currentFinalPriceDraft()
  };
  localStorage.setItem(workbenchDraftStorageKey, JSON.stringify(draft));
}

function applyFinalPriceDraft(finalPrices = []) {
  if (!Array.isArray(finalPrices) || !finalPrices.length) {
    return;
  }
  document.querySelectorAll("#finalPricingRows .price-input").forEach((input, index) => {
    const saved = finalPrices[index];
    if (!saved || saved.value === undefined || saved.value === "") {
      return;
    }
    input.value = saved.value;
    input.dataset.autoPrice = saved.autoPrice === false ? "false" : "true";
  });
}

function restoreWorkbenchDraft() {
  let draft = null;
  try {
    draft = JSON.parse(localStorage.getItem(workbenchDraftStorageKey) || "null");
  } catch (error) {
    draft = null;
  }
  if (!draft || !Array.isArray(draft.purchaseRows) || !draft.purchaseRows.length) {
    renderRecognizedRows([]);
    updateSummaryFromRows([]);
    return;
  }

  isRestoringWorkbenchDraft = true;
  $("productName").value = draft.productName || "";
  $("lengthCm").value = draft.dimensions?.lengthCm ?? 0;
  $("widthCm").value = draft.dimensions?.widthCm ?? 0;
  $("heightCm").value = draft.dimensions?.heightCm ?? 0;
  $("weightG").value = draft.dimensions?.weightG ?? 0;
  $("salePackQty").value = draft.salePackQty || "1";
  $("comparisonUnitQty").value = draft.comparisonUnitQty || "1";
  $("targetMargin").value = draft.targetMargin || "15%";
  currentCompetitors = Array.isArray(draft.competitors) ? draft.competitors : [];
  setProductMode(draft.productMode || "standard");
  renderRecognizedRows(draft.purchaseRows);
  applyFinalPriceDraft(draft.finalPrices);
  updateSummaryFromRows(draft.purchaseRows);
  updateCompetitorSummary(currentCompetitors.length);
  renderReportContent();
  renderPricingHistoryNav();
  $("baseStatus").textContent = "待确认";
  $("baseStatus").classList.remove("ready");
  setSiteStatus("已恢复上次未完成的识别结果，可以继续编辑。");
  isRestoringWorkbenchDraft = false;
}

function skuPrefixFromSpec(spec) {
  const normalized = String(spec || "")
    .toUpperCase()
    .replace(/MM|CM|IN/g, "")
    .replace(/[×*]/g, "X")
    .replace(/[^\u4e00-\u9fa5A-Z0-9X-]+/g, "")
    .replace(/款式/g, "STYLE")
    .replace(/彩色/g, "COLOR")
    .replace(/毫升|ML/g, "ML")
    .replace(/\*/g, "X")
    .replace(/^-+|-+$/g, "")
    .slice(0, 18);
  return normalized || "ITEM";
}

function makeSku(spec, index, itemCode = "") {
  const prefix = String(itemCode || "").trim().toUpperCase().replace(/[^\u4e00-\u9fa5A-Z0-9-]+/g, "");
  return `${prefix || `PO-${skuPrefixFromSpec(spec)}`}-${String(index + 1).padStart(3, "0")}`;
}

function rowIdentity(row) {
  const titleKey = row.itemCode || row.spec ? "" : row.title || "";
  return [
    row.bundle ? "BUNDLE" : "STANDARD",
    row.itemCode || "",
    titleKey,
    row.spec || "",
    row.quantity || "",
    row.cost || "",
    row.bundle ? JSON.stringify(row.components || []) : ""
  ].join("|").toUpperCase();
}

function finalTitle(row) {
  const title = String(row.title || "采购款式").trim();
  const spec = String(row.spec || "").trim();
  return spec && !title.includes(spec) ? `${title} - ${spec}` : title;
}

function looksLikeDynamicShell(text) {
  const hasOrderData = /规格|数量|实付|实付款|金额|单价|商品|凉拖|透明|合计|运费/.test(text);
  const hasDynamicShell = /<app-root|buyerOrderPrint|vm-seller-print|app\.nocache\.js|vite-legacy-entry/.test(text);
  return hasDynamicShell && !hasOrderData;
}

function cleanTextFromHtml(text) {
  const doc = new DOMParser().parseFromString(text, "text/html");
  return doc.body?.innerText || text;
}

function normalizeOrderText(text) {
  return String(text || "")
    .replace(/\u00a0/g, " ")
    .replace(/[，]/g, ",")
    .replace(/[；]/g, ";")
    .replace(/[×]/g, "*")
    .replace(/规格\s*型号/g, "规格型号")
    .replace(/数\s*量/g, "数量")
    .replace(/单\s*价/g, "单价")
    .replace(/金\s*额/g, "金额")
    .replace(/货\s*号/g, "货号")
    .replace(/货品\s*名称/g, "货品名称");
}

function fieldValue(text, labels) {
  for (const label of labels) {
    const pattern = new RegExp(`${label}\\s*[：:]?\\s*([^\\n\\r]+)`);
    const match = text.match(pattern);
    if (match) {
      return match[1]
        .replace(/规格型号|数量|单价|优惠|金额|货品合计|货品总量|运费|实付款.*/g, "")
        .trim();
    }
  }
  return "";
}

function parseLabelledRows(text) {
  const normalized = normalizeOrderText(text);
  const defaultItemCode = fieldValue(normalized, ["货号", "商品货号", "款号"]);
  const defaultTitle = fieldValue(normalized, ["货品名称", "商品名称", "产品名称", "品名"]);
  const rowPattern = /(?:货号\s*[：:]?\s*(?<itemCode>[A-Za-z0-9-]+)\s*)?(?:货品名称\s*[：:]?\s*(?<titleBefore>.*?))?规格型号\s*[：:]?\s*(?<spec>.*?)(?:\s+数量\s*[：:]?\s*(?<qty>\d+(?:\.\d+)?)|\s+单价\s*[：:]?\s*(?<priceOnly>\d+(?:\.\d+)?))(?:.*?数量\s*[：:]?\s*(?<qtyLater>\d+(?:\.\d+)?))?(?:.*?单价\s*[：:]?\s*(?<price>\d+(?:\.\d+)?))?(?:.*?优惠\s*[：:]?\s*(?<discount>-?\d+(?:\.\d+)?))?(?:.*?金额\s*[：:]?\s*(?<amount>\d+(?:\.\d+)?))?/g;
  const rows = [];
  const used = new Set();
  let match;

  while ((match = rowPattern.exec(normalized)) !== null) {
    const groups = match.groups || {};
    const rawSpec = cleanLooseSpec(groups.spec);
    const quantity = num(groups.qty || groups.qtyLater, 0);
    const price = moneyNumber(groups.price || groups.priceOnly);
    const amount = moneyNumber(groups.amount);
    const title = cleanLooseTitle(groups.titleBefore) || defaultTitle || "采购款式";
    const itemCode = (groups.itemCode || defaultItemCode || "").trim();
    const key = `${itemCode}-${title}-${rawSpec}-${quantity}-${amount}-${price}`;

    if (!rawSpec || used.has(key)) {
      continue;
    }
    used.add(key);
    rows.push({
      itemCode,
      title,
      spec: rawSpec,
      quantity: quantity || "",
      cost: quantity && amount ? amount / quantity : price || "",
      amount: amount || "",
      discount: groups.discount || ""
    });
  }

  return rows;
}

function cleanLooseSpec(value) {
  return String(value || "")
    .replace(/数量\s*[：:]?.*/g, "")
    .replace(/单价\s*[：:]?.*/g, "")
    .replace(/优惠\s*[：:]?.*/g, "")
    .replace(/金额\s*[：:]?.*/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/^[-:：]+|[-:：]+$/g, "");
}

function cleanLooseTitle(value) {
  return String(value || "")
    .replace(/规格型号.*/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function parseTableLikeRows(lines) {
  const rows = [];
  const used = new Set();
  const pattern = /^(?:(?<itemCode>[A-Za-z]{1,8}-?\d{1,8})\s+)?(?<title>[\u4e00-\u9fa5A-Za-z0-9（）() -]{2,80}?)\s+(?<spec>(?:规格型号\s*[：:]?\s*)?[^数量单价金额]{2,40}?(?:款式|ml|ML|cm|mm|彩色|黑色|白色|透明|[xX*])[^数量单价金额]{0,40})\s+数量?\s*(?<qty>\d+(?:\.\d+)?)\s+单价?\s*(?<price>\d+(?:\.\d+)?)?(?:\s+优惠?\s*(?<discount>-?\d+(?:\.\d+)?))?\s+金额?\s*(?<amount>\d+(?:\.\d+)?)$/;

  for (const line of lines) {
    const match = line.match(pattern);
    if (!match?.groups) {
      continue;
    }
    const quantity = num(match.groups.qty, 0);
    const amount = moneyNumber(match.groups.amount);
    const price = moneyNumber(match.groups.price);
    const spec = cleanLooseSpec(match.groups.spec.replace(/^规格型号\s*[：:]?\s*/, ""));
    const key = `${match.groups.itemCode || ""}-${match.groups.title}-${spec}-${quantity}-${amount}`;
    if (!quantity || !spec || used.has(key)) {
      continue;
    }
    used.add(key);
    rows.push({
      itemCode: match.groups.itemCode || "",
      title: cleanLooseTitle(match.groups.title),
      spec,
      quantity,
      cost: amount ? amount / quantity : price || "",
      amount: amount || "",
      discount: match.groups.discount || ""
    });
  }

  return rows;
}

function compactOcrLine(line) {
  let normalized = normalizeOrderText(line)
    .replace(/[|｜]/g, " ")
    .replace(/[。]/g, "")
    .replace(/规\s*格\s*型\s*号/g, "规格型号")
    .replace(/货\s*品\s*名\s*称/g, "货品名称")
    .replace(/货\s*品\s*总\s*量/g, "货品总量")
    .replace(/实\s*付\s*款/g, "实付款")
    .replace(/款\s*式/g, "款式")
    .replace(/元\s*\/\s*PCS/gi, "")
    .replace(/([A-Z])\]\s*-\s*(\d+)/gi, "$1J-$2")
    .replace(/\bH\]\s*(\d+)/gi, "HJ-$1")
    .replace(/(\d)\s+00ml/gi, "$100ml")
    .replace(/(\d+)\s*ml/gi, "$1ml")
    .replace(/(\d{2,4}ml)\s+款式/gi, "$1款式")
    .replace(/00ml\s+款式/gi, "00ml款式")
    .replace(/款式\s*(\d+)/g, "款式$1")
    .replace(/[ \t]{2,}/g, " ")
    .trim();

  for (let index = 0; index < 4; index += 1) {
    normalized = normalized.replace(/([\u4e00-\u9fa5])\s+([\u4e00-\u9fa5])/g, "$1$2");
  }
  return normalized;
}

function normalizeItemCode(value) {
  const fixed = String(value || "")
    .replace(/]/g, "J")
    .replace(/\s+/g, "")
    .toUpperCase();
  const match = fixed.match(/[A-Z]{1,4}-\d{1,8}/);
  return match ? match[0] : "";
}

function guessOcrTitle(text) {
  const compact = compactOcrLine(text);
  if (/复古透明浮雕玻璃/.test(compact)) {
    return "复古透明浮雕玻璃喷壶";
  }
  const cleaned = compact
    .replace(/[A-Z]{1,4}-\d{1,8}/gi, " ")
    .replace(/规格型号\s*[：:]?\s*\d?/g, " ")
    .replace(/\d{2,4}ml款式\d+;?/gi, " ")
    .replace(/\d+\s+\d+(?:\.\d+)?\s+-?\d+(?:\.\d+)?\s+\d+(?:\.\d+)?/g, " ")
    .replace(/彩色|黑色|白色|透明|红色|绿色|蓝色|粉色|黄色|紫色/g, " ")
    .replace(/[^\u4e00-\u9fa5A-Za-z0-9]+/g, " ")
    .trim();
  return cleaned.slice(0, 32) || "图片识别款式";
}

function normalizeOcrColorName(value) {
  const compact = String(value || "").replace(/\s+/g, "");
  const shortColors = {
    金: "金色",
    黑: "黑色",
    白: "白色",
    红: "红色",
    绿: "绿色",
    蓝: "蓝色",
    粉: "粉色",
    黄: "黄色",
    紫: "紫色",
    灰: "灰色",
    银: "银色",
    棕: "棕色",
    米: "米色",
    橙: "橙色",
    透: "透明"
  };
  return shortColors[compact] || compact;
}

function extractOcrColor(block) {
  const tight = compactOcrLine(block).replace(/\s+/g, "");
  const colorCapture = "(透明|彩色|金色|黑色|白色|红色|绿色|蓝色|粉色|黄色|紫色|灰色|银色|棕色|咖啡色|米色|橙色|金|黑|白|红|绿|蓝|粉|黄|紫|灰|银|棕|米|橙|透)";
  const patterns = [
    new RegExp(`颜色[:：][^\\n。；;]{0,120}[（(]${colorCapture}`),
    new RegExp(`(?:款式|规格|彩花|花|颜色)[^\\n。；;（）()]{0,40}[（(]${colorCapture}`),
    new RegExp(`[（(]${colorCapture}(?:色|明)?[）)]?`),
    new RegExp(`颜色[:：][^\\n。；;]{0,40}${colorCapture}`)
  ];

  for (const pattern of patterns) {
    const match = tight.match(pattern);
    if (match) {
      return normalizeOcrColorName(match[1]);
    }
  }
  return "";
}

function extractOcrSize(block) {
  const tight = compactOcrLine(block).replace(/\s+/g, "");
  const match = tight.match(/(?:尺码|尺寸|码数)[:：]?([A-Za-z0-9.-]{1,8})/);
  return match ? match[1] : "";
}

function parseOcrAmountInfo(block) {
  const compact = compactOcrLine(block);
  const patterns = [
    /(?<qty>\d{1,5})\s+(?<price>\d+(?:\.\d+)?)\s*元\s*(?:\/\s*[\u4e00-\u9fa5A-Za-z]+)?\s*(?<discount>-?\d+(?:\.\d+)?)\s+(?<amount>\d+(?:\.\d+)?)/g,
    /(?<qty>\d{1,5})\s+(?<price>\d+(?:\.\d+)?)\s+(?<discount>-?\d+(?:\.\d+)?)\s+(?<amount>\d+(?:\.\d+)?)(?=\s|$)/g
  ];
  let best = null;

  for (const pattern of patterns) {
    let match;
    while ((match = pattern.exec(compact)) !== null) {
      best = match.groups;
    }
    if (best) {
      break;
    }
  }

  if (!best) {
    return null;
  }
  const quantity = num(best.qty);
  const price = moneyNumber(best.price);
  const amount = moneyNumber(best.amount);
  if (!quantity || (!price && !amount)) {
    return null;
  }
  return {
    quantity,
    price,
    amount,
    discount: best.discount || "",
    cost: amount ? amount / quantity : price
  };
}

function normalizeLooseOcrItemCode(block) {
  const alphaCode = normalizeItemCode(block.match(/[A-Z][A-Z\]]\s*-\s*\d{1,8}/i)?.[0] || "");
  if (alphaCode) {
    return alphaCode;
  }
  const spaced = String(block || "").replace(/[，,。]/g, " ");
  const looseCode = spaced.match(/\d{2,8}\s*[*xX×]\s*\d{1,6}(?:\s*[\u4e00-\u9fa5A-Za-z0-9]){0,4}/)?.[0] || "";
  return looseCode
    .replace(/[×xX]/g, "*")
    .replace(/[^\u4e00-\u9fa5A-Za-z0-9*.-]+/g, "")
    .toUpperCase();
}

function cleanColorSizeOcrTitle(block, itemCode) {
  let cleaned = compactOcrLine(block)
    .replace(/序号.*?金额\s*\(?元?\)?/g, " ")
    .replace(/货品合计.*$/g, " ")
    .replace(/实付款.*$/g, " ")
    .replace(/货品总量.*$/g, " ")
    .replace(/[（(]\s*(?:金|黑|白|红|绿|蓝|粉|黄|紫|灰|银|棕|米|橙|透)(?:色|明)?/g, " ")
    .replace(/(?:金|黑|白|红|绿|蓝|粉|黄|紫|灰|银|棕|米|橙|透)(?:色|明)?\s*[）)]/g, " ")
    .replace(/明\s*[）)]/g, " ")
    .replace(/[，,\s]\d+\s*[工丁]\s*[，,\s]/g, " ")
    .replace(/鞋\s*花\s*(?:色|明)/g, "鞋花")
    .replace(/颜色\s*[:：]\s*[^\n ]+/g, " ")
    .replace(/\d{1,5}\s+\d+(?:\.\d+)?\s*元\s*(?:\/\s*[\u4e00-\u9fa5A-Za-z]+)?\s*-?\d+(?:\.\d+)?\s+\d+(?:\.\d+)?/g, " ")
    .replace(/\d{1,5}\s+\d+(?:\.\d+)?\s+-?\d+(?:\.\d+)?\s+\d+(?:\.\d+)?/g, " ")
    .replace(/尺码\s*[:：]?\s*[A-Za-z0-9.-]+[^。；;\n)]*[）)]?/g, " ")
    .replace(/码数偏小\d+号/g, " ")
    .replace(/\d{2,8}\s*[*xX×]\s*\d{1,6}(?:\s*[\u4e00-\u9fa5A-Za-z0-9]){0,4}/g, " ")
    .replace(/[A-Za-z]+/g, " ")
    .replace(/[^\u4e00-\u9fa5]+/g, " ")
    .replace(/颜色|规格|数量|单价|优惠|金额|货号|货品名称|彩\s*花|花\s*色/g, " ")
    .replace(/金色|黑色|白色|红色|绿色|蓝色|粉色|黄色|紫色|灰色|银色|棕色|咖啡色|米色|橙色|透明/g, " ")
    .replace(/\s+/g, "");

  if (itemCode) {
    cleaned = cleaned.replace(itemCode.replace(/[^\u4e00-\u9fa5A-Za-z0-9]/g, ""), "");
  }
  return cleaned.slice(0, 42) || "图片识别款式";
}

function splitColorSizeOcrBlocks(text) {
  const normalized = normalizeOrderText(text);
  const byColorLabel = normalized
    .split(/(?=^\s*颜色\s*[:：])/m)
    .map((part) => part.trim())
    .filter(Boolean);
  if (byColorLabel.length > 1) {
    return byColorLabel;
  }

  const lines = normalized.split(/\r?\n/);
  const blocks = [];
  let current = [];
  for (const line of lines) {
    if (/颜色\s*[:：]/.test(line) && current.length) {
      blocks.push(current.join("\n"));
      current = [line];
      continue;
    }
    current.push(line);
  }
  if (current.length) {
    blocks.push(current.join("\n"));
  }
  return blocks.map((part) => part.trim()).filter(Boolean);
}

function parseColorSizeOcrRows(text) {
  const rows = [];
  const used = new Set();

  for (const block of splitColorSizeOcrBlocks(text)) {
    const amountInfo = parseOcrAmountInfo(block);
    const color = extractOcrColor(block);
    const size = extractOcrSize(block);
    const hasVariantCue = /颜色|尺码|尺寸|码数/.test(block);
    if (!amountInfo || (!hasVariantCue && !color && !size)) {
      continue;
    }

    const itemCode = normalizeLooseOcrItemCode(block);
    const specParts = [];
    if (color) {
      specParts.push(`颜色：${color}`);
    }
    if (size) {
      specParts.push(`尺码：${size}`);
    }
    const spec = specParts.join("；") || `OCR款式${rows.length + 1}`;
    const row = {
      itemCode,
      title: cleanColorSizeOcrTitle(block, itemCode) || guessOcrTitle(block),
      spec,
      quantity: amountInfo.quantity,
      cost: amountInfo.cost,
      amount: amountInfo.amount || "",
      discount: amountInfo.discount
    };
    const key = rowIdentity(row);
    if (row.quantity > 0 && row.cost && !used.has(key)) {
      used.add(key);
      rows.push(row);
    }
  }

  return rows;
}

function parseOcrTableRows(text) {
  const sourceLines = text
    .split(/\r?\n/)
    .map(compactOcrLine)
    .filter(Boolean);
  const lines = sourceLines.flatMap((line, index) => {
    const previous = sourceLines[index - 1] || "";
    const next = sourceLines[index + 1] || "";
    const merged = `${previous} ${line}`.trim();
    return [
      { line, colorSource: `${line} ${next}`.trim(), titleSource: `${previous} ${line} ${next}`.trim() },
      { line: merged, colorSource: `${line} ${next}`.trim(), titleSource: `${previous} ${line} ${next}`.trim() }
    ];
  });
  const rows = [];
  const used = new Set();
  const colorPattern = /(彩色|金色|黑色|白色|透明|红色|绿色|蓝色|粉色|黄色|紫色|灰色|银色|棕色|咖啡色|米色|橙色)/;
  const priceLinePattern = /(?<spec>(?:\d\s*)?00ml款式\d+;?|\d{2,4}ml款式\d+;?).*?(?<qty>\d{1,5})\s+(?<price>\d+(?:\.\d+)?)\s+(?<discount>-\d+(?:\.\d+)?)\s+(?<amount>\d+(?:\.\d+)?)/gi;

  lines.forEach((entry) => {
    let match;
    while ((match = priceLinePattern.exec(entry.line)) !== null) {
      const windowText = entry.titleSource;
      let spec = match.groups.spec.replace(/\s+/g, "");
      if (/^00ml/i.test(spec)) {
        const prefix = windowText.match(/规格型号\s*[：:]?\s*(\d)\s/i)?.[1] || "2";
        spec = `${prefix}${spec}`;
      }
      const color = entry.colorSource.match(colorPattern)?.[1] || "";
      if (color && !spec.includes(color)) {
        spec = `${spec} ${color}`;
      }
      const itemCode = normalizeItemCode(windowText.match(/[A-Z][A-Z\]]\s*-\s*\d{1,8}/i)?.[0] || "");
      const quantity = num(match.groups.qty);
      const amount = moneyNumber(match.groups.amount);
      const price = moneyNumber(match.groups.price);
      const row = {
        itemCode,
        title: guessOcrTitle(windowText),
        spec,
        quantity,
        cost: amount && quantity ? amount / quantity : price,
        amount,
        discount: match.groups.discount || ""
      };
      const key = rowIdentity(row);
      if (quantity > 0 && row.cost && !used.has(key)) {
        used.add(key);
        rows.push(row);
      }
    }
  });

  return rows;
}

function parsePurchaseOrderText(text, fileName = "") {
  if (looksLikeDynamicShell(text)) {
    return [];
  }
  const plain = normalizeOrderText(fileName.toLowerCase().endsWith(".html") || fileName.toLowerCase().endsWith(".htm")
    ? cleanTextFromHtml(text)
    : text);
  const lines = plain
    .split(/\r?\n/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .filter(Boolean);

  const rows = [];
  const used = new Set();

  const parsedRows = [
    ...parseColorSizeOcrRows(plain),
    ...parseLabelledRows(plain),
    ...parseTableLikeRows(lines),
    ...parseOcrTableRows(plain)
  ];

  for (const row of parsedRows) {
    const key = rowIdentity(row);
    if (!used.has(key)) {
      used.add(key);
      rows.push(row);
    }
  }

  if (!rows.length) {
    const unit = "(?:mm|MM|cm|CM|in|IN)?";
    const linePattern = new RegExp(`(?<title>[\\u4e00-\\u9fa5A-Za-z0-9 -]{0,32}?)\\s*(?<spec>\\d+(?:\\.\\d+)?\\s*${unit}\\s*[xX*×]\\s*\\d+(?:\\.\\d+)?\\s*${unit})\\D{0,50}(?<qty>\\d{1,5})\\D{0,35}(?<price>\\d+(?:\\.\\d{1,4})?)\\D{0,35}(?<amount>\\d+(?:\\.\\d{1,2})?)`, "g");

    for (const line of lines) {
      let match;
      while ((match = linePattern.exec(line)) !== null) {
        const spec = match.groups.spec.replace(/\s+/g, "").replace("×", "*");
        const qty = num(match.groups.qty);
        const price = moneyNumber(match.groups.price);
        const amount = moneyNumber(match.groups.amount);
        const row = {
            title: normalizeTitle(match.groups.title, spec, rows.length),
            spec,
            quantity: qty,
            cost: amount > 0 ? amount / qty : price
          };
        const key = rowIdentity(row);
        if (qty > 0 && (price > 0 || amount > 0) && !used.has(key)) {
          used.add(key);
          rows.push(row);
        }
      }
    }
  }

  if (!rows.length) {
    const specMatches = [...plain.matchAll(/\d+(?:\.\d+)?\s*(?:mm|MM|cm|CM|in|IN)?\s*[xX*×]\s*\d+(?:\.\d+)?\s*(?:mm|MM|cm|CM|in|IN)?/g)];
    specMatches.slice(0, 20).forEach((match, index) => {
      rows.push({
        title: `采购款式-${index + 1}`,
        spec: match[0].replace(/\s+/g, "").replace("×", "*"),
        quantity: "",
        cost: ""
      });
    });
  }

  return rows.slice(0, 30);
}

function rowFromManualInputs() {
  const sku = $("manualSku").value.trim();
  const itemCode = $("manualItemCode").value.trim();
  const title = $("manualTitle").value.trim();
  const spec = $("manualSpec").value.trim();
  const quantity = $("manualQty").value.trim();
  const cost = $("manualCost").value.trim();
  if (!title && !spec) {
    throw new Error("请至少填写款式标题或规格");
  }
  return {
    sku,
    itemCode,
    title: title || `采购款式-${currentPurchaseRows.length + 1}`,
    spec: spec || "手动规格",
    quantity,
    cost
  };
}

function splitBulkColumns(line) {
  const trimmed = String(line || "").trim();
  if (!trimmed) {
    return [];
  }
  if (!/[\t,，|｜]/.test(trimmed)) {
    return [trimmed];
  }
  const delimiter = trimmed.includes("\t") ? /\t/ : /[|｜]/.test(trimmed) ? /[|｜]/ : /[,，]/;
  return trimmed.split(delimiter).map((item) => item.trim());
}

function looksLikeBulkHeader(columns) {
  const text = columns.join("").toLowerCase();
  return /sku|货号|款式|标题|规格|数量|成本|单价|price|qty/.test(text)
    && !columns.some((item) => /\d+(?:\.\d+)?/.test(item));
}

function rowFromBulkColumns(columns, index) {
  if (columns.length === 1) {
    return rowFromLooseBulkLine(columns[0], index);
  }
  const values = columns.map(cleanBulkValue).filter(Boolean);
  if (values.length >= 6 && isBulkSku(values[0])) {
    const [sku, itemCode, title, spec, quantity, cost] = values;
    return {
      sku: cleanBulkValue(sku),
      itemCode: cleanBulkValue(itemCode),
      title: cleanBulkValue(title) || `采购款式-${currentPurchaseRows.length + index + 1}`,
      spec: cleanBulkValue(spec) || "手动规格",
      quantity: cleanBulkNumber(quantity),
      cost: cleanBulkNumber(cost)
    };
  }
  return rowFromDelimitedBulkColumns(values, index);
}

function rowFromDelimitedBulkColumns(values, index) {
  const costIndex = findLastNumericIndex(values);
  const quantityIndex = findLastNumericIndex(values, costIndex - 1);
  if (quantityIndex < 0 || costIndex < 0) {
    throw new Error("批量粘贴至少需要 款式/规格、数量、成本价；SKU 可在表格里人工填写");
  }

  const quantity = cleanBulkNumber(values[quantityIndex]);
  const cost = cleanBulkNumber(values[costIndex]);
  const descriptive = values.filter((_, valueIndex) => valueIndex !== quantityIndex && valueIndex !== costIndex);
  let sku = "";
  if (descriptive.length && isBulkSku(descriptive[0])) {
    sku = descriptive.shift();
  }

  const fallback = cleanBulkValue($("productName")?.value) || `采购款式-${currentPurchaseRows.length + index + 1}`;
  const title = descriptive[0] || fallback;
  const spec = descriptive.length > 1 ? descriptive.slice(1).join(" ") : title || "手动规格";
  return {
    sku,
    itemCode: "",
    title,
    spec,
    quantity,
    cost
  };
}

function numericLabelValue(text, labels) {
  const pattern = labels.join("|");
  const match = String(text || "").match(new RegExp(`(?:${pattern})\\s*[：:]?\\s*(-?\\d+(?:\\.\\d+)?)`, "i"));
  return match ? cleanBulkValue(match[1]) : "";
}

function removeLabeledNumber(text, labels) {
  const pattern = labels.join("|");
  return String(text || "").replace(new RegExp(`(?:${pattern})\\s*[：:]?\\s*-?\\d+(?:\\.\\d+)?`, "gi"), " ");
}

function splitTitleAndSpec(text, index) {
  const words = cleanBulkValue(text)
    .replace(/^(款式|标题|品名|规格)\s*[：:]?/g, "")
    .split(/\s+/)
    .filter(Boolean);
  const fallback = `采购款式-${currentPurchaseRows.length + index + 1}`;
  if (!words.length) {
    return { title: fallback, spec: "手动规格" };
  }
  if (words.length === 1) {
    return { title: words[0], spec: words[0] };
  }
  const dimensionStart = words.findIndex((word) => /\d+(?:\.\d+)?\s*(cm|厘米|mm|毫米|in|inch|英寸|g|kg|lb|oz|ml|l|升)\b/i.test(word));
  if (dimensionStart > 0) {
    return {
      title: words.slice(0, dimensionStart).join(" "),
      spec: words.slice(dimensionStart).join(" ")
    };
  }
  return {
    title: words.slice(0, -1).join(" "),
    spec: words[words.length - 1]
  };
}

function rowFromLooseBulkLine(line, index) {
  const normalized = String(line || "")
    .replace(/[，,；;]/g, " ")
    .replace(/：/g, ":")
    .replace(/\s+/g, " ")
    .trim();

  const skuMatch = normalized.match(/(?:^|\s)(?:sku|货号)\s*:\s*(\S+)/i);
  const explicitSku = skuMatch ? cleanBulkValue(skuMatch[1]) : "";
  let quantity = numericLabelValue(normalized, ["数量", "qty"]);
  let cost = numericLabelValue(normalized, ["成本价", "成本", "单价", "price"]);

  let body = normalized
    .replace(/(?:^|\s)(?:sku|货号)\s*:\s*\S+/gi, " ")
    .replace(/\b(?:sku|货号)\b/gi, " ");
  body = removeLabeledNumber(body, ["数量", "qty"]);
  body = removeLabeledNumber(body, ["成本价", "成本", "单价", "price"]);
  body = body
    .replace(/\b(?:款式|标题|品名|规格)\s*:/gi, " ")
    .replace(/\s+/g, " ")
    .trim();

  const tailMatch = body.match(/^(.*?)\s+(\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)$/);
  if (tailMatch) {
    body = cleanBulkValue(tailMatch[1]);
    quantity = quantity || cleanBulkValue(tailMatch[2]);
    cost = cost || cleanBulkValue(tailMatch[3]);
  }

  const parts = body.split(/\s+/).filter(Boolean);
  let sku = explicitSku;
  if (!sku && parts[0] && /^[A-Z0-9][A-Z0-9_-]{2,}$/i.test(parts[0]) && parts.length >= 2) {
    sku = cleanBulkValue(parts.shift());
  }
  body = parts.join(" ");

  if (body && quantity && cost) {
    const { title, spec } = splitTitleAndSpec(body, index);
    return {
      sku,
      itemCode: "",
      title,
      spec,
      quantity,
      cost
    };
  }

  throw new Error("批量粘贴至少需要 款式/规格、数量、成本价；SKU 可写可不写，也可以在表格里人工填写");
}

function cleanBulkValue(value) {
  return String(value ?? "").replace(/^["']|["']$/g, "").trim();
}

function cleanBulkNumber(value) {
  const match = cleanBulkValue(value).match(/-?\d+(?:\.\d+)?/);
  return match ? match[0] : "";
}

function findLastNumericIndex(values, startIndex = values.length - 1) {
  for (let index = Math.min(startIndex, values.length - 1); index >= 0; index -= 1) {
    if (cleanBulkNumber(values[index])) {
      return index;
    }
  }
  return -1;
}

function isBulkSku(value) {
  const cleaned = cleanBulkValue(value);
  return /^[A-Z0-9][A-Z0-9_-]{2,}$/i.test(cleaned) && /[A-Za-z]/.test(cleaned);
}

function rowsFromBulkInput(text) {
  const rows = [];
  const lines = String(text || "").split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  lines.forEach((line) => {
    const columns = splitBulkColumns(line);
    if (!columns.length || looksLikeBulkHeader(columns)) {
      return;
    }
    const row = rowFromBulkColumns(columns, rows.length);
    if (!row.sku && !row.title && !row.spec) {
      return;
    }
    rows.push(row);
  });
  return rows;
}

function normalizeTitle(rawTitle, spec, index) {
  const cleaned = String(rawTitle || "")
    .replace(spec, "")
    .replace(/[：:|｜,，]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return cleaned || `采购款式-${index + 1}`;
}

function renderRecognizedRows(rows) {
  const dims = currentDimensions();
  const previousSignature = currentPurchaseRows.map(rowIdentity).join("||");
  const nextSignature = rows.map(rowIdentity).join("||");
  if (previousSignature && nextSignature && previousSignature !== nextSignature) {
    resetCompetitorData();
  }
  currentPurchaseRows = rows;
  if (!rows.length) {
    $("recognizedRows").innerHTML = '<tr><td colspan="9" class="empty">暂无采购行，请上传采购单或手动添加采购行。</td></tr>';
    $("recognizedSummaryText").textContent = "暂无采购行，请上传采购单或手动添加采购行。";
    renderFinalRows([]);
    saveWorkbenchDraft();
    return;
  }
  syncProductNameFromRows(rows);
  $("recognizedRows").innerHTML = rows.map((row, index) => {
    const sku = row.sku || "";
    const quantityValue = row.quantity === "" ? "" : fmt(row.quantity, 0);
    const costValue = row.cost === "" ? "" : fmt(row.bundle ? row.bundleSaleUnitCostRmb || row.cost : row.cost, 4);
    return `
      <tr>
        <td><input class="table-input sku-input" data-row-edit="${index}" data-field="sku" value="${escapeXml(sku)}" aria-label="产品 SKU"></td>
        <td><input class="table-input title-input" data-row-edit="${index}" data-field="title" value="${escapeXml(row.title)}" aria-label="款式标题"></td>
        <td><input class="table-input spec-input" data-row-edit="${index}" data-field="spec" value="${escapeXml(row.spec)}" aria-label="款式规格"></td>
        <td class="numeric"><input class="table-input numeric-input" type="number" min="0" step="1" data-row-edit="${index}" data-field="quantity" value="${escapeXml(quantityValue)}" aria-label="${escapeXml(sku)} 采购数量"></td>
        <td class="numeric"><input class="table-input numeric-input strong-input" type="number" min="0" step="0.0001" data-row-edit="${index}" data-field="cost" value="${escapeXml(costValue)}" aria-label="${escapeXml(sku)} 成本价"></td>
        <td class="numeric">${fmt(dims.lengthCm, 2)}</td>
        <td class="numeric">${fmt(dims.widthCm, 2)}</td>
        <td class="numeric">${fmt(dims.heightCm, 2)}</td>
        <td class="numeric">${fmt(dims.weightG, 2)}</td>
      </tr>
    `;
  }).join("");
  const bundleRow = rows.length === 1 && rows[0].bundle ? rows[0] : null;
  $("recognizedSummaryText").textContent = bundleRow
    ? `已生成 1 个组合套装父 SKU：${bundleRow.spec}，${bundleRowSummary(bundleRow)}。`
    : `从采购单识别出 ${rows.length} 款产品，尺寸重量已按手动输入补齐。`;
  renderFinalRows(rows);
  initEditableFields();
  refreshReportDraft();
  saveWorkbenchDraft();
}

function renderFinalRows(rows) {
  const dims = currentDimensions();
  if (!rows.length) {
    $("finalPricingRows").innerHTML = '<tr><td colspan="7" class="empty">暂无最终价格行。</td></tr>';
    resetDownloadState();
    renderPricingHistoryNav();
    return;
  }
  const previousPrices = finalRows().map((row) => row[6]);
  const previousAutoFlags = [...document.querySelectorAll("#finalPricingRows .price-input")]
    .map((input) => input.dataset.autoPrice === "true");
  $("finalPricingRows").innerHTML = rows.map((row, index) => {
    const sku = row.sku || "";
    const defaultPrice = suggestedDefaultPriceForRow(row, rows);
    const previousPrice = previousPrices[index] || "";
    const wasManual = previousPrice && previousAutoFlags[index] === false;
    const priceValue = wasManual ? previousPrice : defaultPrice ? fmt(defaultPrice, 2) : "";
    const autoPrice = wasManual ? "false" : "true";
    return `
      <tr>
        <td data-final-sku="${index}">${escapeXml(sku || "-")}</td>
        <td data-final-title="${index}">${escapeXml(finalTitle(row))}</td>
        <td class="numeric">${cmToIn(dims.lengthCm)}</td>
        <td class="numeric">${cmToIn(dims.widthCm)}</td>
        <td class="numeric">${cmToIn(dims.heightCm)}</td>
        <td class="numeric">${gToLb(dims.weightG)}</td>
        <td><input class="price-input" type="number" value="${escapeXml(priceValue)}" min="0" step="0.01" data-auto-price="${autoPrice}" aria-label="${escapeXml(sku || `第${index + 1}行`)} 定价"></td>
      </tr>
    `;
  }).join("");
  resetDownloadState();
  renderPricingHistoryNav();
}

function updateSummaryFromRows(rows) {
  const cards = document.querySelectorAll(".summary-grid strong");
  const bundleRow = rows.length === 1 && rows[0].bundle ? rows[0] : null;
  const totalQty = rows.reduce((sum, row) => sum + num(row.quantity), 0);
  const totalAmount = bundleRow
    ? num(bundleRow.quantity) * num(bundleRow.bundleSaleUnitCostRmb || bundleRow.cost)
    : rows.reduce((sum, row) => sum + (num(row.quantity) * num(row.cost)), 0);
  if (cards[0]) cards[0].textContent = bundleRow ? "1 个SKU" : `${rows.length} 款`;
  if (cards[1]) cards[1].textContent = totalQty ? `${totalQty} ${bundleRow ? "套" : "件"}` : "-";
  if (cards[2]) cards[2].textContent = totalAmount ? `${fmt(totalAmount, 2)} RMB` : "-";
  renderPricingHistoryNav();
}

async function handlePurchaseOrderFile(file) {
  const suffix = file.name.split(".").pop().toLowerCase();
  $("purchaseOrderFileName").textContent = `已选择：${file.name}`;
  showFilePreview(file, suffix);

  if (["pdf", "png", "jpg", "jpeg", "webp"].includes(suffix)) {
    renderRecognizedRows([]);
    updateSummaryFromRows([]);
    $("baseStatus").textContent = "待补录";
    $("baseStatus").classList.remove("ready");
    if (["png", "jpg", "jpeg", "webp"].includes(suffix)) {
      await runImageOcr({ autoParse: true });
    } else {
      setSiteStatus("已选择 PDF 采购单。PDF OCR 下一步再接；现在请手动添加或批量添加采购行。", "warn");
    }
    return;
  }

  const text = await file.text();
  const rows = parsePurchaseOrderText(text, file.name);
  if (!rows.length) {
    const message = looksLikeDynamicShell(text)
      ? "这个 HTML 是 1688 动态页面壳，里面没有订单明细数据。已清空示例行，请手动添加或批量添加采购行继续测试。"
      : "没有从采购单中识别到规格行。请手动添加或批量添加采购行继续测试。";
    renderRecognizedRows([]);
    updateSummaryFromRows([]);
    $("baseStatus").textContent = "待补录";
    $("baseStatus").classList.remove("ready");
    setSiteStatus(message, "warn");
    return;
  }

  renderRecognizedRows(rows);
  updateSummaryFromRows(rows);
  $("baseStatus").textContent = "待确认";
  $("baseStatus").classList.remove("ready");
  setSiteStatus(`已从 ${file.name} 识别 ${rows.length} 款产品，请检查表格后确认。`);
}

function addManualRow() {
  try {
    const row = rowFromManualInputs();
    const rows = [...currentPurchaseRows, row];
    renderRecognizedRows(rows);
    updateSummaryFromRows(rows);
    $("manualSku").value = "";
    $("manualItemCode").value = "";
    $("manualTitle").value = "";
    $("manualSpec").value = "";
    $("manualQty").value = "";
    $("manualCost").value = "";
    $("baseStatus").textContent = "待确认";
    $("baseStatus").classList.remove("ready");
    setSiteStatus("已添加 1 条采购行，请检查识别结果表。");
  } catch (error) {
    setSiteStatus(error.message, "warn");
  }
}

function addBulkRows() {
  try {
    const text = $("manualBulkRows").value.trim();
    if (!text) {
      throw new Error("请先粘贴多行采购数据");
    }
    const parsedRows = rowsFromBulkInput(text);
    if (!parsedRows.length) {
      throw new Error("没有解析到可添加的采购行");
    }
    const rows = [...currentPurchaseRows, ...parsedRows];
    renderRecognizedRows(rows);
    updateSummaryFromRows(rows);
    $("manualBulkRows").value = "";
    $("baseStatus").textContent = "待确认";
    $("baseStatus").classList.remove("ready");
    setSiteStatus(`已批量添加 ${parsedRows.length} 条采购行，请检查 SKU、款式和规格。`);
  } catch (error) {
    setSiteStatus(error.message, "warn");
  }
}

function setProductMode(mode) {
  currentProductMode = mode === "bundle" ? "bundle" : "standard";
  const standardButton = $("standardModeBtn");
  const bundleButton = $("bundleModeBtn");
  const bundleCard = $("bundleCard");
  standardButton.classList.toggle("active", !isBundleMode());
  bundleButton.classList.toggle("active", isBundleMode());
  standardButton.setAttribute("aria-pressed", String(!isBundleMode()));
  bundleButton.setAttribute("aria-pressed", String(isBundleMode()));
  bundleCard.classList.toggle("is-hidden", !isBundleMode());
  setSiteStatus(isBundleMode()
    ? "已切换到组合套装模式：填写父 SKU 和组件后生成一条套装行。"
    : "已切换到普通单品模式：采购行按普通款式录入。");
  updateBundlePreview();
  saveWorkbenchDraft();
}

function buildBundleRow() {
  try {
    const row = bundleRowFromInputs();
    renderRecognizedRows([row]);
    updateSummaryFromRows([row]);
    $("salePackQty").value = fmt(row.bundleUnitCount, 0);
    $("comparisonUnitQty").value = fmt(row.bundleUnitCount, 0);
    updateCompetitorSummary();
    $("baseStatus").textContent = "待确认";
    $("baseStatus").classList.remove("ready");
    setSiteStatus(`已生成组合套装父 SKU：${row.sku}，请检查尺寸重量和竞品口径。`);
    refreshReportDraft();
  } catch (error) {
    setSiteStatus(error.message, "warn");
  }
}

function clearBundleInputs() {
  ["bundleSku", "bundleTitle", "bundleSpec"].forEach((id) => {
    $(id).value = "";
  });
  document.querySelectorAll(".bundle-component-row").forEach((row, index) => {
    row.querySelector(".bundle-component-name").value = "";
    row.querySelector(".bundle-component-qty").value = "8";
    row.querySelector(".bundle-component-cost").value = "";
    row.querySelector(".bundle-component-purchase").value = "";
  });
  renderRecognizedRows([]);
  updateSummaryFromRows([]);
  updateBundlePreview();
  $("baseStatus").textContent = "待补录";
  $("baseStatus").classList.remove("ready");
  setSiteStatus("套装输入已清空，可以重新填写组件。");
}

function refreshDimensions() {
  const rows = currentPurchaseRows;
  renderRecognizedRows(rows);
  updateSummaryFromRows(rows);
  setSiteStatus("尺寸重量已更新，最终表已同步转换为 in/lb。");
  refreshReportDraft();
  saveWorkbenchDraft();
}

function resetDownloadState() {
  const link = $("downloadExcelLink");
  link.classList.add("is-hidden");
  if (finalExcelUrl) {
    URL.revokeObjectURL(finalExcelUrl);
    finalExcelUrl = "";
  }
  $("exportStatus").textContent = "价格已修改，等待重新确认。";
  setSiteStatus("价格或 SKU 已修改，请重新确认并生成 Excel。", "warn");
}

function showGeneratedReportState() {
  if ($("competitorStatus").textContent === "待上传") {
    $("competitorParseStatus").textContent = "还没有当前产品竞品资料，报告只生成配送费和利润底线，竞品策略待补充。";
    updateCompetitorSummary();
  }
  renderReportContent();
  $("generateReportBtn").textContent = "报告已生成";
  $("generateReportBtn").classList.add("is-confirmed");
  $("generateReportBtn").disabled = false;
  setSiteStatus("定价分析报告已生成，可以导出 Excel 或报告文本。");
  smoothScrollTo("finalStage");
}

function updateCompetitorSummary(count) {
  const summaryCards = document.querySelectorAll(".competitor-summary strong");
  const unitQty = comparisonQty();
  const unitLabel = sellingUnitLabel(productLabel());
  const competitors = activeCompetitors();
  const lowestComparisonPrice = competitors.reduce((lowest, competitor) => {
    if (unitQty <= 0) {
      return lowest;
    }
    const comparisonPrice = competitor.price / competitor.packCount * unitQty;
    return comparisonPrice < lowest ? comparisonPrice : lowest;
  }, Infinity);
  const visibleCount = count === undefined ? competitors.length : count;
  if (summaryCards[0]) summaryCards[0].textContent = `${visibleCount} 个`;
  if (summaryCards[1]) summaryCards[1].textContent = Number.isFinite(lowestComparisonPrice) ? money(lowestComparisonPrice, 2) : "待上传";
  if (summaryCards[2]) summaryCards[2].textContent = `${unitQty} ${unitLabel}`;
}

async function handleCompetitorFiles(event) {
  const files = [...event.target.files];
  const count = files.length;
  if (!count) {
    currentCompetitors = [];
    $("competitorFileSummary").textContent = "暂无竞品文件。";
    $("competitorParseStatus").textContent = "等待补充竞品资料。";
    $("competitorStatus").textContent = "待上传";
    $("competitorStatus").classList.remove("ready");
    updateCompetitorSummary();
    refreshReportDraft();
    saveWorkbenchDraft();
    return;
  }

  $("competitorFileSummary").textContent = `已选择：${count} 个竞品文件。`;
  $("competitorParseStatus").textContent = "正在解析竞品售价和售卖数量。";
  const parsed = (await Promise.all(files.map((file, index) => parseCompetitorFile(file, index)))).filter(Boolean);
  currentCompetitors = parsed;
  if (!parsed.length) {
    $("competitorParseStatus").textContent = "没有从竞品文件中解析到售价。请上传包含价格的 HTML / 文本，或后续补手动竞品输入。";
    $("competitorStatus").textContent = "待补录";
    $("competitorStatus").classList.remove("ready");
    updateCompetitorSummary(0);
    refreshReportDraft();
    saveWorkbenchDraft();
    setSiteStatus("竞品文件已接收，但没有解析到可用售价。", "warn");
    smoothScrollTo("reportPanel");
    return;
  }
  $("competitorFileSummary").textContent = `已解析：${parsed.length} 个竞品锚点。`;
  $("competitorParseStatus").textContent = `已按当前产品解析竞品售价和售卖数量：${parsed.map((item) => `${item.label} ${money(item.price, 2)} / ${item.packCount}`).join("；")}`;
  $("competitorStatus").textContent = "已整理";
  $("competitorStatus").classList.add("ready");
  updateCompetitorSummary(parsed.length);
  refreshReportDraft();
  saveWorkbenchDraft();
  setSiteStatus(`已补充并解析 ${parsed.length} 个竞品文件，下一步点击“生成报告”。`);
  smoothScrollTo("reportPanel");
}

function initFileInputs() {
  $("purchaseOrderFile").addEventListener("change", async (event) => {
    const file = event.target.files[0];
    if (!file) {
      $("purchaseOrderFileName").textContent = "等待上传采购单";
      return;
    }
    try {
      await handlePurchaseOrderFile(file);
    } catch (error) {
      setSiteStatus(`采购单读取失败：${error.message}`, "warn");
    }
  });

  $("competitorFiles").addEventListener("change", handleCompetitorFiles);
}

function initButtons() {
  $("confirmBaseBtn").addEventListener("click", () => {
    $("baseStatus").textContent = "已确认";
    $("baseStatus").classList.add("ready");
    setSiteStatus("基础表格已确认，可以继续补充竞品并生成报告。");
    smoothScrollTo("reportPanel");
  });

  $("generateReportBtn").addEventListener("click", showGeneratedReportState);
  $("reportExcelBtn").addEventListener("click", generateReportExcel);
  $("reportTextBtn").addEventListener("click", generateReportText);
  $("generateExcelBtn").addEventListener("click", generateFinalExcel);
  $("runOcrBtn").addEventListener("click", () => runImageOcr({ autoParse: true }));
  $("addManualRowBtn").addEventListener("click", addManualRow);
  $("addBulkRowsBtn").addEventListener("click", addBulkRows);
  $("standardModeBtn").addEventListener("click", () => setProductMode("standard"));
  $("bundleModeBtn").addEventListener("click", () => setProductMode("bundle"));
  $("buildBundleBtn").addEventListener("click", buildBundleRow);
  $("clearBundleBtn").addEventListener("click", clearBundleInputs);
  ["bundleSku", "bundleTitle", "bundleSpec", "bundleUnitLabel"].forEach((id) => {
    $(id).addEventListener("input", updateBundlePreview);
  });
  document.querySelectorAll(".bundle-component-row input").forEach((input) => {
    input.addEventListener("input", updateBundlePreview);
  });
  $("productName").addEventListener("input", () => {
    refreshReportDraft();
    saveWorkbenchDraft();
  });
  $("salePackQty").addEventListener("input", () => {
    refreshAutoPrices();
    saveWorkbenchDraft();
  });
  $("salePackQty").addEventListener("change", () => {
    refreshAutoPrices();
    saveWorkbenchDraft();
  });
  $("comparisonUnitQty").addEventListener("input", () => {
    updateCompetitorSummary();
    refreshReportDraft();
    saveWorkbenchDraft();
  });
  $("comparisonUnitQty").addEventListener("change", () => {
    updateCompetitorSummary();
    refreshReportDraft();
    saveWorkbenchDraft();
  });
  $("targetMargin").addEventListener("input", () => {
    refreshAutoPrices();
    saveWorkbenchDraft();
  });
  ["lengthCm", "widthCm", "heightCm", "weightG"].forEach((id) => {
    $(id).addEventListener("input", refreshDimensions);
    $(id).addEventListener("change", refreshDimensions);
  });
}

function initEditableFields() {
  document.querySelectorAll("[data-row-edit]").forEach((input) => {
    input.addEventListener("input", () => {
      updatePurchaseRow(input.dataset.rowEdit, input.dataset.field, input.value);
    });
  });

  initPriceInputs();
}

function initPriceInputs() {
  document.querySelectorAll(".price-input").forEach((input) => {
    input.addEventListener("input", () => {
      input.dataset.autoPrice = "false";
      resetDownloadState();
      refreshReportDraft();
      saveWorkbenchDraft();
    });
  });
}

document.addEventListener("DOMContentLoaded", () => {
  loadPricingHistoryProducts();
  resetTransientWorkbenchState();
  initFileInputs();
  initButtons();
  initEditableFields();
  updateBundlePreview();
  restoreWorkbenchDraft();
  updateCompetitorSummary();
  renderPricingHistoryNav();
  loadOutputPricingHistory();
});

window.priceWorkbench = {
  parsePurchaseOrderText,
  parseCompetitorText,
  renderRecognizedRows,
  updateSummaryFromRows
};
