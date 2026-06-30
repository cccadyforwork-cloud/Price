const finalHeaders = ["SKU", "款式标题", "长(in)", "宽(in)", "高(in)", "重(lb)", "定价"];
const reportHeaders = ["项目", "数值", "说明"];
const currencyRateRmbToUsd = 7.2;
const firstLegShippingUsd = 0.3;
const referralFeeRate = 0.18;
const returnRate = 0.1;
const disposalFeeUsd = 0.25;
const breakEvenFactor = 0.73476;
const sampleOrderText = `货号 HJ-11
货品名称 复古透明浮雕玻璃喷壶 喷水壶室内园艺按压式玻璃浇水壶批发浇水壶
规格型号：200ml款式3; 彩色 数量 2 单价 4.00 优惠 -0.25 金额 7.75
规格型号：200ml款式1; 彩色 数量 2 单价 4.00 优惠 -0.25 金额 7.75
规格型号：200ml款式4; 彩色 数量 2 单价 4.00 优惠 -0.25 金额 7.75
规格型号：200ml款式2; 彩色 数量 2 单价 4.00 优惠 -0.25 金额 7.75
货品合计 32.00 元 货品总量 8 运费 10.80 元 优惠 -1.00 元 实付款 41.80 元`;

let finalExcelUrl = "";
let reportExcelUrl = "";
let reportTextUrl = "";
let currentPurchaseRows = [];
let currentObjectUrl = "";
let currentImageFile = null;
let currentCompetitors = [];
let ocrWorker = null;
let ocrReady = false;
let tesseractLoadPromise = null;

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

function roundUpToEnding99(value) {
  const dollars = Math.floor(Number(value || 0));
  const candidate = dollars + 0.99;
  return candidate >= value ? candidate : dollars + 1.99;
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

function $(id) {
  return document.getElementById(id);
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
  return currentPurchaseRows.length ? currentPurchaseRows : defaultPurchaseRows();
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
  if (input && firstTitle && (!input.value.trim() || input.value.trim() === "防磨贴")) {
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
  const parsed = num($("targetMargin").value, 25);
  return parsed > 1 ? parsed / 100 : parsed;
}

function suggestedPrice() {
  const prices = finalRows().map((row) => num(row[6])).filter((value) => value > 0);
  return prices.length ? prices[0] : 5.61;
}

function salePackQty() {
  const input = $("salePackQty");
  return Math.max(1, num(input?.value, 30));
}

function comparisonQty() {
  return Math.max(1, num($("comparisonUnitQty").value, 1));
}

function sellingUnitLabel(productName) {
  return /贴|片|pad|patch/i.test(productName) ? "片" : "件";
}

function activeCompetitors() {
  return currentCompetitors;
}

function parsePackCount(text, fallback = 1) {
  const source = String(text || "").replace(/(\d+)\s*ml/gi, "");
  const patterns = [
    /(\d+)\s*(?:pcs|pc|pieces|piece|count|ct)\b/i,
    /(\d+)\s*(?:pack|packs|pair|pairs)\b/i,
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

function parseCompetitorText(text, fileName, index) {
  const plain = cleanTextFromHtml(text);
  const title = firstMatch(text, [
    /id=["']productTitle["'][^>]*>([\s\S]*?)<\/span>/i,
    /<title[^>]*>([\s\S]*?)<\/title>/i
  ]) || plain.split(/\r?\n/).find((line) => line.trim().length > 8) || fileName;
  const cleanTitle = cleanTextFromHtml(title).replace(/\s+/g, " ").trim();
  const priceText = firstMatch(text, [
    /aria-label=["'](?:current price\s*)?\$(\d+(?:\.\d{1,2})?)["']/i,
    /<span[^>]*class=["'][^"']*a-offscreen[^"']*["'][^>]*>\$(\d+(?:\.\d{1,2})?)<\/span>/i,
    /\$(\d+(?:\.\d{1,2})?)/
  ]);
  const price = moneyNumber(priceText);
  if (!price) {
    return null;
  }
  const ratingText = firstMatch(text, [
    /(\d(?:\.\d)?)\s*out of\s*5\s*stars/i,
    /([1-5](?:\.\d)?)\s*分/
  ]);
  const reviewText = firstMatch(text, [
    /id=["']acrCustomerReviewText["'][^>]*>([\d,]+)\s*(?:ratings?|reviews?)<\/span>/i,
    /([\d,]+)\s*(?:ratings?|reviews?)/i,
    /([\d,]+)\s*评/
  ]);
  const packCount = parsePackCount(`${cleanTitle} ${fileName}`, 1);
  return {
    label: shortCompetitorLabel(cleanTitle || fileName, index),
    title: cleanTitle || fileName,
    price,
    packCount,
    rating: ratingText || reviewText ? `${ratingText ? `${ratingText}分` : "-"} / ${reviewText ? `${reviewText.replace(/,/g, "")}评` : "-"} ` : "待确认"
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
  const positiveCosts = rows.map((row) => num(row.cost)).filter((value) => value > 0);
  if (!positiveCosts.length) {
    return 0;
  }
  if (rows.length > 1) {
    const pcsPerStyle = unitQty / rows.length;
    return rows.reduce((sum, row) => sum + num(row.cost) * pcsPerStyle, 0) / currencyRateRmbToUsd;
  }
  return positiveCosts[0] * unitQty / currencyRateRmbToUsd;
}

function rowProfitRows(rows, unitQty, price, shippingFee) {
  const pcsPerStyle = rows.length > 1 ? unitQty / rows.length : unitQty;
  return rows.map((row) => {
    const costUsd = num(row.cost) * pcsPerStyle / currencyRateRmbToUsd;
    const profit = price * breakEvenFactor - (costUsd + firstLegShippingUsd + shippingFee + disposalFeeUsd * returnRate);
    return {
      title: row.spec || row.title || "当前款式",
      costUsd,
      profit
    };
  }).filter((row) => row.costUsd > 0);
}

function fbaFeeForOz(weightOz, price) {
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

function reportData() {
  const rows = activePurchaseRows();
  const hasPurchaseRows = rows.length > 0;
  const dims = currentDimensions();
  const totalQty = rows.reduce((sum, row) => sum + num(row.quantity), 0);
  const totalCost = rows.reduce((sum, row) => sum + num(row.quantity) * num(row.cost), 0);
  const avgCost = totalQty ? totalCost / totalQty : rows.reduce((sum, row) => sum + num(row.cost), 0) / Math.max(rows.length, 1);
  const price = suggestedPrice();
  const margin = targetMarginNumber();
  const productName = productLabel();
  const saleQty = salePackQty();
  const unitQty = comparisonQty();
  const unitLabel = sellingUnitLabel(productName);
  const saleUnitCostUsd = sellingUnitCostUsd(rows, saleQty);
  const weightOz = gToOz(dims.weightG);
  const returnDisposalReserve = disposalFeeUsd * returnRate;
  const breakEvenResult = breakEvenWithFee(saleUnitCostUsd, weightOz);
  const shippingFee = breakEvenResult.shippingFee;
  const shippingTier = breakEvenResult.shippingTier;
  const fixedCostUsd = breakEvenResult.fixedCostUsd;
  const breakEven = breakEvenResult.breakEven;
  const targetPrice = Math.max(roundUpToEnding99(breakEven), breakEven / Math.max(0.2, 1 - margin));
  const suggestedMinPrice = roundUpToEnding99(breakEven);
  const competitors = activeCompetitors().map((competitor) => ({
    ...competitor,
    unitPrice: competitor.price / Math.max(1, competitor.packCount),
    comparisonPrice: competitor.price / Math.max(1, competitor.packCount) * unitQty
  }));
  const hasCompetitors = competitors.length > 0;
  const competitorCount = `${competitors.length} 个`;
  const lowestCompetitorPrice = hasCompetitors ? Math.min(...competitors.map((competitor) => competitor.price)) : 0;
  const currentUnitPrice = price / saleQty;
  const currentComparisonPrice = currentUnitPrice * unitQty;
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
  const unitPriceConclusion = !closestUnitCompetitor
    ? "请先上传并解析当前产品的竞品文件，再生成竞品价格结论。"
    : unitPriceGap <= 0
      ? `当前 ${unitQty}${unitLabel} 折算价不高于 ${closestUnitCompetitor.label} 竞品，价格位置偏锋利。`
      : `当前 ${unitQty}${unitLabel} 折算价高于 ${closestUnitCompetitor.label} 竞品，需要靠低总价、差异化或评价来支撑。`;
  const profitRows = rowProfitRows(rows, saleQty, price, shippingFee);
  const highCostProfit = profitRows.reduce((highest, row) => !highest || row.costUsd > highest.costUsd ? row : highest, null);
  const lowCostProfit = profitRows.reduce((lowest, row) => !lowest || row.costUsd < lowest.costUsd ? row : lowest, null);
  const postReviewPrice = price < 2.99 ? 2.99 : roundUpToEnding99(price + 0.4);

  return {
    productName,
    rows,
    hasPurchaseRows,
    styleCount: rows.length,
    totalQty,
    totalCost,
    avgCost,
    dims,
    shippingFee,
    shippingTier,
    breakEven,
    targetPrice,
    suggestedMinPrice,
    price,
    margin,
    competitorCount,
    lowestCompetitorPrice,
    saleQty,
    unitQty,
    unitLabel,
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
    competitorUnitPrice,
    competitorComparisonPrice,
    packagePriceGap,
    unitPriceGap,
    unitPriceConclusion,
    highCostProfit,
    lowCostProfit,
    postReviewPrice,
    weightOz
  };
}

function reportRows() {
  const data = reportData();
  return [
    ["产品", data.productName, "按产品名称或识别表第一行款式标题生成"],
    ["采购款数", data.styleCount, "来自采购单识别结果"],
    ["采购总数", data.totalQty, "来自采购单识别结果"],
    ["平均成本", fmt(data.avgCost, 4), "按货品成本 / 采购总数估算"],
    ["销售单位成本", money(data.saleUnitCostUsd, 2), `按本品售卖数量 ${data.saleQty}${data.unitLabel} 折算`],
    ["配送费", money(data.shippingFee, 2), "按重量和售价档位估算"],
    ["保本定价", money(data.breakEven, 2), "按截图里的保本公式折算"],
    ["建议不低于", money(data.suggestedMinPrice, 2), "向上取 .99 上架价"],
    ["当前确认价", money(data.price, 2), "来自价格确认表默认或手动定价"],
    ["本品售卖数量", `${data.saleQty}${data.unitLabel}`, "用于利润和保本计算"],
    ["竞品对比量", `${data.unitQty}${data.unitLabel}`, "用于竞品价格折算"],
    ["当前折算价", money(data.currentComparisonPrice, 2), `当前确认价 / ${data.saleQty}${data.unitLabel} × ${data.unitQty}${data.unitLabel}`],
    ["竞品锚点", data.closestUnitCompetitor?.label || "待上传", `${data.unitQty}${data.unitLabel} 折算价最接近当前价格`],
    ["折算价差", signedMoney(data.unitPriceGap, 3), data.unitPriceConclusion]
  ];
}

function reportText() {
  const data = reportData();
  const highCostLine = data.highCostProfit
    ? `按你表里的低价配送费模型，成本最高的 ${data.highCostProfit.title} 约有 ${money(data.highCostProfit.profit, 2)}/套 利润；`
    : "";
  const lowCostLine = data.lowCostProfit
    ? `低成本款约有 ${money(data.lowCostProfit.profit, 2)}/套 利润。`
    : "";
  return [
    `通用产品定价台 - ${data.productName} 定价分析报告`,
    "",
    "配送费分析过程",
    `按你表里的参数算，${data.productName} 保本定价约为 ${money(data.breakEven, 2)}，建议上架不要低于 ${money(data.suggestedMinPrice, 2)}。`,
    `尺寸：${fmt(data.dims.lengthCm, 2)} × ${fmt(data.dims.widthCm, 2)} × ${fmt(data.dims.heightCm, 2)} cm，重量 ${fmt(data.dims.weightG, 2)}g = 约 ${fmt(data.weightOz, 2)}oz。`,
    `按表里非服饰类 ${data.shippingTier} 的配送费 ${money(data.shippingFee, 2)}。`,
    `保本售价 = (单件成本 + ${fmt(firstLegShippingUsd, 2)} + ${fmt(data.shippingFee, 2)} + ${fmt(data.returnDisposalReserve, 3)}) / ${fmt(breakEvenFactor, 5)}。`,
    `代入 ${money(data.saleUnitCostUsd, 2)}：(${money(data.saleUnitCostUsd, 2)} + ${money(data.fixedCostUsd, 3)}) / ${fmt(breakEvenFactor, 5)} = ${money(data.breakEven, 2)}。`,
    "",
    "竞品与利润分析过程",
    data.hasCompetitors
      ? `竞品锚点我从你给的 ${data.competitors.length} 个页面快照里看到：`
      : "还没有解析到当前产品的竞品锚点。请上传当前产品的竞品 HTML 或文本文件。",
    ...(data.hasCompetitors ? data.competitors.map((item) => `${item.label}：${money(item.price, 2)} / ${item.packCount}${data.unitLabel}，单${data.unitLabel}价 ${money(item.unitPrice, 2)}/${data.unitLabel}，按 ${data.unitQty}${data.unitLabel} 折算为 ${money(item.comparisonPrice, 2)}，${item.rating}`) : []),
    data.hasCompetitors
      ? `你的 ${data.saleQty}${data.unitLabel}装卖 ${money(data.price, 2)}：单${data.unitLabel}价 ${money(data.currentUnitPrice, 2)}/${data.unitLabel}，按 ${data.unitQty}${data.unitLabel} 折算为 ${money(data.currentComparisonPrice, 2)}，基本对标 ${data.closestUnitCompetitor.label} 竞品。`
      : `你的 ${data.saleQty}${data.unitLabel}装卖 ${money(data.price, 2)}：单${data.unitLabel}价 ${money(data.currentUnitPrice, 2)}/${data.unitLabel}，按 ${data.unitQty}${data.unitLabel} 折算为 ${money(data.currentComparisonPrice, 2)}。`,
    data.hasCompetitors ? `总价门槛：当前售价比最低竞品 ${money(data.lowestTotalCompetitor.price, 2)} 对比。` : "",
    data.hasCompetitors ? `${highCostLine}${lowCostLine}` : "",
    data.hasCompetitors
      ? `策略：首发 ${money(data.price, 2)} 跑评论和转化，累计 20-30 个评价后测试 ${money(data.postReviewPrice, 2)}。不要为了追 ${data.lowestComparisonCompetitor.label} 的 ${money(data.lowestComparisonCompetitor.comparisonPrice, 2)}/${data.unitQty}${data.unitLabel} 去打，应该主打低总价、单尺寸刚需、试用门槛低。`
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
    return;
  }
  const highCostProfit = data.highCostProfit ? money(data.highCostProfit.profit, 2) : "-";
  const lowCostProfit = data.lowCostProfit ? money(data.lowCostProfit.profit, 2) : "-";
  const highCostTitle = data.highCostProfit?.title || "成本最高款";
  const weightTierText = data.weightOz <= 4 ? "落在 4oz 及以下" : `对应 ${data.shippingTier} 档`;
  const totalPriceCompareText = data.hasCompetitors && data.price < data.lowestTotalCompetitor.price
    ? `比 ${money(data.lowestTotalCompetitor.price, 2)} 和 ${money(data.competitors[data.competitors.length - 1].price, 2)} 的竞品下单门槛低很多`
    : data.hasCompetitors
      ? `没有明显低于 ${money(data.lowestTotalCompetitor.price, 2)} 的最低总价锚点，需要靠利润或差异化支撑`
      : "暂无当前产品竞品锚点";
  const competitorRowsHtml = data.hasCompetitors
    ? data.competitors.map((competitor) => `
      <tr>
        <td>${escapeXml(competitor.label)}</td>
        <td><mark>${money(competitor.price, 2)}</mark> / ${competitor.packCount}${data.unitLabel}</td>
        <td><mark>${money(competitor.unitPrice, 2)}/${data.unitLabel}</mark></td>
        <td><mark>${money(competitor.comparisonPrice, 2)}</mark></td>
        <td>${escapeXml(competitor.rating)}</td>
      </tr>
    `).join("")
    : '<tr><td colspan="5" class="empty">暂无当前产品竞品。请上传竞品 HTML / 文本后再生成报告。</td></tr>';
  $("reportContent").innerHTML = `
    <article class="process-block">
      <h3>配送费分析过程</h3>
      <p class="lead-text">按你表里的参数算，${escapeXml(data.productName)} 保本定价约为 <mark>${money(data.breakEven, 2)}</mark>，建议上架不要低于 <mark>${money(data.suggestedMinPrice, 2)}</mark>。</p>
      <p class="lead-text">我用的表内假设是：</p>
      <ul>
        <li>尺寸：${fmt(data.dims.lengthCm, 2)} × ${fmt(data.dims.widthCm, 2)} × ${fmt(data.dims.heightCm, 2)} cm，重量 ${fmt(data.dims.weightG, 2)}g = 约 ${fmt(data.weightOz, 2)}oz，${escapeXml(weightTierText)}。</li>
        <li>按表里非服饰类 <mark>${escapeXml(data.shippingTier)}</mark> 的配送费 <mark>${money(data.shippingFee, 2)}</mark>。</li>
        <li>单件成本沿用表里 <mark>${money(data.saleUnitCostUsd, 2)}</mark>。</li>
        <li>头程 <mark>${money(firstLegShippingUsd, 2)}</mark>。</li>
        <li>退货率 <mark>${fmt(returnRate * 100, 0)}%</mark>。</li>
        <li>弃置费 <mark>${money(disposalFeeUsd, 2)}</mark>。</li>
        <li>亚马逊抽佣 <mark>${fmt(referralFeeRate * 100, 0)}%</mark>。</li>
      </ul>
      <p class="lead-text">公式折算后是：</p>
      <p><mark>保本售价 = (单件成本 + ${fmt(firstLegShippingUsd, 2)} + ${fmt(data.shippingFee, 2)} + ${fmt(data.returnDisposalReserve, 3)}) / ${fmt(breakEvenFactor, 5)}</mark></p>
      <p>代入 <mark>${money(data.saleUnitCostUsd, 2)}</mark>：</p>
      <p><mark>(${money(data.saleUnitCostUsd, 2)} + ${money(data.fixedCostUsd, 3)}) / ${fmt(breakEvenFactor, 5)} = ${money(data.breakEven, 2)}</mark></p>
      <p>如果 ${escapeXml(data.productName)} 实际采购成本不是 <mark>${money(data.saleUnitCostUsd, 2)}</mark>，把真实单件成本替进去就行。比如成本每增加 <mark>$0.10</mark>，保本价大约增加 <mark>${money(0.1 / breakEvenFactor, 2)}</mark>。</p>
    </article>

    <article class="process-block">
      <h3>竞品与利润分析过程</h3>
      <p class="lead-text">${data.hasCompetitors ? `竞品锚点我从你给的 <mark>${data.competitors.length}</mark> 个页面快照里看到：` : "还没有解析到当前产品的竞品锚点。"}</p>
      <div class="comparison-table-wrap">
        <table class="comparison-table">
          <thead>
            <tr>
              <th>竞品</th>
              <th>售价/数量</th>
              <th>单${data.unitLabel}价</th>
              <th>${data.unitQty}${data.unitLabel}折算价</th>
              <th>评分/评论</th>
            </tr>
          </thead>
          <tbody>${competitorRowsHtml}</tbody>
        </table>
      </div>

      <p class="lead-text">你的 ${data.saleQty}${data.unitLabel}装卖 <mark>${money(data.price, 2)}</mark>，当前竞品对比量是 <mark>${data.unitQty}${data.unitLabel}</mark>：</p>
      <ul>
        <li>本品单${data.unitLabel}价：<mark>${money(data.currentUnitPrice, 2)}/${data.unitLabel}</mark>；按 ${data.unitQty}${data.unitLabel} 折算为 <mark>${money(data.currentComparisonPrice, 2)}</mark>${data.hasCompetitors ? `，基本对标 ${escapeXml(data.closestUnitCompetitor.label)} 竞品` : "，等待竞品锚点对比"}。</li>
        ${data.hasCompetitors ? `<li>总价明显低：${escapeXml(totalPriceCompareText)}。</li>` : ""}
        <li>按你表里的低价配送费模型，成本最高的 ${escapeXml(highCostTitle)} 仍有约 <mark>${highCostProfit}/套</mark> 利润；低成本款利润约 <mark>${lowCostProfit}/套</mark>。</li>
      </ul>
      ${data.hasCompetitors ? `<p>不建议首发 <mark>${money(data.postReviewPrice, 2)}</mark>，虽然利润更好，但单${data.unitLabel}价会到 <mark>${money(data.postReviewPrice / data.saleQty, 2)}/${data.unitLabel}</mark>，按 ${data.unitQty}${data.unitLabel} 折算是 <mark>${money(data.postReviewPrice / data.saleQty * data.unitQty, 2)}</mark>，面对 ${escapeXml(data.closestUnitCompetitor.label)} 的 <mark>${money(data.closestUnitCompetitor.comparisonPrice, 2)}</mark> 不够锋利。</p>` : ""}
      <p class="strategy-text">${data.hasCompetitors ? `我的策略：首发 <mark>${money(data.price, 2)}</mark> 跑评论和转化，累计 20-30 个评价后涨到 <mark>${money(data.postReviewPrice, 2)}</mark>。不要为了追 ${escapeXml(data.lowestComparisonCompetitor.label)} 的 <mark>${money(data.lowestComparisonCompetitor.comparisonPrice, 2)}/${data.unitQty}${data.unitLabel}</mark> 去打，应该主打“低总价、单尺寸刚需、试用门槛低”。` : "我的策略：先补齐当前产品竞品锚点，再输出最终价格策略。"}</p>
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
    setPurchasePreview(`<p>已选择 PDF：${escapeXml(file.name)}。当前前端会保留文件名，请把 PDF 中的订单文字粘贴到下方，或用手动添加采购行继续测试。</p>`);
    setOcrStatus("PDF 暂不支持 OCR");
    return;
  }
  setOcrStatus("当前文件不需要 OCR");
  setPurchasePreview("");
}

function imageFileToCanvas(file, maxWidth = 2200) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    const objectUrl = URL.createObjectURL(file);
    image.onload = () => {
      URL.revokeObjectURL(objectUrl);
      const scale = Math.min(1, maxWidth / image.naturalWidth);
      const canvas = document.createElement("canvas");
      canvas.width = Math.max(1, Math.round(image.naturalWidth * scale));
      canvas.height = Math.max(1, Math.round(image.naturalHeight * scale));
      const ctx = canvas.getContext("2d", { willReadFrequently: true });
      ctx.drawImage(image, 0, 0, canvas.width, canvas.height);

      const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
      const data = imageData.data;
      for (let index = 0; index < data.length; index += 4) {
        const gray = data[index] * 0.299 + data[index + 1] * 0.587 + data[index + 2] * 0.114;
        const boosted = gray > 180 ? 255 : gray < 105 ? 0 : gray;
        data[index] = boosted;
        data[index + 1] = boosted;
        data[index + 2] = boosted;
      }
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
    $("orderTextInput").value = text;

    if (!text) {
      setSiteStatus("图片 OCR 没有识别到文字，请换更清晰截图或手动输入。", "warn");
      setOcrStatus("未识别到文字");
      $("baseStatus").textContent = "待补录";
      return;
    }

    if (!autoParse) {
      setSiteStatus("图片文字已识别，请检查 OCR 文本后点击解析。");
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

function currentDimensions() {
  return {
    lengthCm: num($("lengthCm").value, 8),
    widthCm: num($("widthCm").value, 12),
    heightCm: num($("heightCm").value, 1.5),
    weightG: num($("weightG").value, 55)
  };
}

function defaultPurchaseRows() {
  return [...document.querySelectorAll("#recognizedRows tr")].map((row, index) => {
    const cells = row.querySelectorAll("td");
    if (cells.length < 9 || row.querySelector(".empty")) {
      return null;
    }
    return {
      title: cells[1].textContent.trim(),
      spec: cells[2].textContent.trim(),
      quantity: cells[3].textContent.trim(),
      cost: cells[4].textContent.trim(),
      sku: cells[0].querySelector("input")?.value || makeSku(cells[2].textContent.trim(), index),
      itemCode: skuPrefixFromSpec(cells[2].textContent.trim())
    };
  }).filter(Boolean);
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
    row.itemCode || "",
    titleKey,
    row.spec || "",
    row.quantity || "",
    row.cost || ""
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
  const colorPattern = /(彩色|黑色|白色|透明|红色|绿色|蓝色|粉色|黄色|紫色)/;
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

  for (const row of [...parseLabelledRows(plain), ...parseTableLikeRows(lines), ...parseOcrTableRows(plain)]) {
    const key = rowIdentity(row);
    if (!used.has(key)) {
      used.add(key);
      rows.push(row);
    }
  }

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
  const itemCode = $("manualItemCode").value.trim();
  const title = $("manualTitle").value.trim();
  const spec = $("manualSpec").value.trim();
  const quantity = $("manualQty").value.trim();
  const cost = $("manualCost").value.trim();
  if (!title && !spec) {
    throw new Error("请至少填写款式标题或规格");
  }
  return {
    itemCode,
    title: title || `采购款式-${currentPurchaseRows.length + 1}`,
    spec: spec || "手动规格",
    quantity,
    cost
  };
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
    $("recognizedRows").innerHTML = '<tr><td colspan="9" class="empty">暂无采购行，请粘贴订单文本或手动添加采购行。</td></tr>';
    $("recognizedSummaryText").textContent = "暂无采购行，请粘贴订单文本或手动添加采购行。";
    renderFinalRows([]);
    return;
  }
  syncProductNameFromRows(rows);
  $("recognizedRows").innerHTML = rows.map((row, index) => {
    const sku = row.sku || makeSku(row.spec, index, row.itemCode);
    return `
      <tr>
        <td><input data-final-sku-source="${index}" value="${escapeXml(sku)}"></td>
        <td>${escapeXml(row.title)}</td>
        <td>${escapeXml(row.spec)}</td>
        <td class="numeric">${escapeXml(row.quantity)}</td>
        <td class="numeric strong">${row.cost === "" ? "" : fmt(row.cost, 4)}</td>
        <td class="numeric">${fmt(dims.lengthCm, 2)}</td>
        <td class="numeric">${fmt(dims.widthCm, 2)}</td>
        <td class="numeric">${fmt(dims.heightCm, 2)}</td>
        <td class="numeric">${fmt(dims.weightG, 2)}</td>
      </tr>
    `;
  }).join("");
  $("recognizedSummaryText").textContent = `从采购单识别出 ${rows.length} 款产品，尺寸重量已按手动输入补齐。`;
  renderFinalRows(rows);
  initEditableFields();
  refreshReportDraft();
}

function renderFinalRows(rows) {
  const dims = currentDimensions();
  if (!rows.length) {
    $("finalPricingRows").innerHTML = '<tr><td colspan="7" class="empty">暂无最终价格行。</td></tr>';
    resetDownloadState();
    return;
  }
  $("finalPricingRows").innerHTML = rows.map((row, index) => {
    const sku = row.sku || makeSku(row.spec, index, row.itemCode);
    return `
      <tr>
        <td data-final-sku="${index}">${escapeXml(sku)}</td>
        <td>${escapeXml(finalTitle(row))}</td>
        <td class="numeric">${cmToIn(dims.lengthCm)}</td>
        <td class="numeric">${cmToIn(dims.widthCm)}</td>
        <td class="numeric">${cmToIn(dims.heightCm)}</td>
        <td class="numeric">${gToLb(dims.weightG)}</td>
        <td><input class="price-input" type="number" value="5.61" min="0" step="0.01" aria-label="${escapeXml(sku)} 定价"></td>
      </tr>
    `;
  }).join("");
  resetDownloadState();
}

function updateSummaryFromRows(rows) {
  const cards = document.querySelectorAll(".summary-grid strong");
  const totalQty = rows.reduce((sum, row) => sum + num(row.quantity), 0);
  const totalAmount = rows.reduce((sum, row) => sum + (num(row.quantity) * num(row.cost)), 0);
  if (cards[0]) cards[0].textContent = `${rows.length} 款`;
  if (cards[1]) cards[1].textContent = totalQty ? `${totalQty} 件` : "-";
  if (cards[2]) cards[2].textContent = totalAmount ? `${fmt(totalAmount, 2)} RMB` : "-";
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
      setSiteStatus("已选择 PDF 采购单。PDF OCR 下一步再接；现在请粘贴订单文字或手动添加采购行。", "warn");
    }
    return;
  }

  const text = await file.text();
  const rows = parsePurchaseOrderText(text, file.name);
  if (!rows.length) {
    const message = looksLikeDynamicShell(text)
      ? "这个 HTML 是 1688 动态页面壳，里面没有订单明细数据。已清空示例行，请粘贴订单文字，或手动添加采购行继续测试。"
      : "没有从采购单中识别到规格行。请粘贴订单文字，或手动添加采购行继续测试。";
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

function parsePastedOrderText() {
  const text = $("orderTextInput").value.trim();
  if (!text) {
    setSiteStatus("请先粘贴订单文字。", "warn");
    return;
  }
  const rows = parsePurchaseOrderText(text, "pasted.txt");
  if (!rows.length) {
    setSiteStatus("粘贴文本里没有识别到采购行。可以改用手动添加采购行。", "warn");
    return;
  }
  renderRecognizedRows(rows);
  updateSummaryFromRows(rows);
  $("baseStatus").textContent = "待确认";
  $("baseStatus").classList.remove("ready");
  setSiteStatus(`已从粘贴文本识别 ${rows.length} 款产品，请检查表格后确认。`);
}

function loadSampleOrderText() {
  $("orderTextInput").value = sampleOrderText;
  parsePastedOrderText();
}

function addManualRow() {
  try {
    const row = rowFromManualInputs();
    const rows = [...currentPurchaseRows, row];
    renderRecognizedRows(rows);
    updateSummaryFromRows(rows);
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

function refreshDimensions() {
  const rows = currentPurchaseRows.length ? currentPurchaseRows : defaultPurchaseRows();
  renderRecognizedRows(rows);
  updateSummaryFromRows(rows);
  setSiteStatus("尺寸重量已更新，最终表已同步转换为 in/lb。");
  refreshReportDraft();
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
    const comparisonPrice = competitor.price / competitor.packCount * unitQty;
    return comparisonPrice < lowest ? comparisonPrice : lowest;
  }, Infinity);
  const visibleCount = count === undefined ? competitors.length : count;
  if (summaryCards[0]) summaryCards[0].textContent = `${visibleCount} 个`;
  if (summaryCards[1]) summaryCards[1].textContent = Number.isFinite(lowestComparisonPrice) ? money(lowestComparisonPrice, 2) : "待上传";
  if (summaryCards[2]) summaryCards[2].textContent = `${unitQty || 1} ${unitLabel}`;
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
  $("saveDraftBtn").addEventListener("click", () => {
    $("baseStatus").textContent = "草稿已保存";
    $("baseStatus").classList.add("ready");
    setSiteStatus("草稿已保存。");
  });

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
  $("parseOrderTextBtn").addEventListener("click", parsePastedOrderText);
  $("runOcrBtn").addEventListener("click", () => runImageOcr({ autoParse: true }));
  $("loadSampleOrderBtn").addEventListener("click", loadSampleOrderText);
  $("addManualRowBtn").addEventListener("click", addManualRow);
  $("productName").addEventListener("input", refreshReportDraft);
  $("salePackQty").addEventListener("input", refreshReportDraft);
  $("salePackQty").addEventListener("change", refreshReportDraft);
  $("comparisonUnitQty").addEventListener("input", () => {
    updateCompetitorSummary();
    refreshReportDraft();
  });
  $("comparisonUnitQty").addEventListener("change", () => {
    updateCompetitorSummary();
    refreshReportDraft();
  });
  $("targetMargin").addEventListener("input", refreshReportDraft);
  ["lengthCm", "widthCm", "heightCm", "weightG"].forEach((id) => {
    $(id).addEventListener("change", refreshDimensions);
  });
}

function initEditableFields() {
  document.querySelectorAll("[data-final-sku-source]").forEach((input) => {
    input.addEventListener("input", () => {
      syncSkuToFinal(input.dataset.finalSkuSource, input.value);
      resetDownloadState();
    });
  });

  document.querySelectorAll(".price-input").forEach((input) => {
    input.addEventListener("input", () => {
      resetDownloadState();
      refreshReportDraft();
    });
  });
}

document.addEventListener("DOMContentLoaded", () => {
  currentPurchaseRows = defaultPurchaseRows();
  initFileInputs();
  initButtons();
  initEditableFields();
  updateCompetitorSummary();
  renderReportContent();
});

window.priceWorkbench = {
  parsePurchaseOrderText,
  parseCompetitorText,
  renderRecognizedRows,
  updateSummaryFromRows
};
