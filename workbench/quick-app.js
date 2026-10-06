const SHIPPING_TIERS = {
  under4: { id: "under4", label: "≤4oz", feeLabel: "$0.50 / $0.88", dimensions: [10, 10, 2], weightG: 100 },
  "4to8": { id: "4to8", label: "4–8oz", fee: 1.77, dimensions: [15, 10, 4], weightG: 200 },
  "8to12": { id: "8to12", label: "8–12oz", fee: 2.6, dimensions: [18, 12, 5], weightG: 300 },
  "12to16": { id: "12to16", label: "12–16oz", fee: 3.22, dimensions: [22, 15, 6], weightG: 430 },
  "1to1_25": { id: "1to1_25", label: "1–1.25lb", fee: 3.72, dimensions: [24, 15, 7], weightG: 510 }
};

const appState = {
  pricing: {
    currency_rate_rmb_to_usd: 7.2,
    referral_fee_rate: 0.15,
    return_rate: 0.05,
    return_rate_price_threshold_usd: 3,
    first_leg_shipping_usd: 0.4,
    disposal_fee_usd: 0
  },
  quickRecords: [],
  outputRecords: [],
  normalizedRecords: [],
  selectedRecordId: "",
  selectedTierId: "under4",
  historyDays: "all",
  historyPage: 1,
  pageSize: 15,
  bulkEditOpen: false,
  bulkSelectedIds: new Set(),
  editingRecordId: "",
  latestSavedId: ""
};

const $ = (id) => document.getElementById(id);
const money = (value) => Number.isFinite(Number(value)) && Number(value) > 0 ? `$${Number(value).toFixed(2)}` : "—";
const rmb = (value) => Number.isFinite(Number(value)) && Number(value) > 0 ? `¥${Number(value).toFixed(2).replace(/\.00$/, "")}` : "—";
const escapeHtml = (value) => String(value ?? "")
  .replaceAll("&", "&amp;")
  .replaceAll("<", "&lt;")
  .replaceAll(">", "&gt;")
  .replaceAll('"', "&quot;")
  .replaceAll("'", "&#039;");

function setStatus(message, tone = "") {
  const element = $("appStatus");
  element.textContent = message;
  element.className = `app-status ${tone ? `is-${tone}` : ""}`.trim();
}

function toTimestampMs(value) {
  const number = Number(value || 0);
  if (!number) return 0;
  return number < 100000000000 ? number * 1000 : number;
}

function formatDate(record) {
  if (record.dateLabel) return record.dateLabel;
  const timestamp = toTimestampMs(record.createdAt || record.updatedAt);
  if (!timestamp) return "日期待补";
  return new Intl.DateTimeFormat("zh-CN", { month: "2-digit", day: "2-digit" }).format(new Date(timestamp));
}

function rangeText(rows, key) {
  const values = rows.map((row) => Number(row[key])).filter((value) => Number.isFinite(value) && value > 0);
  if (!values.length) return "—";
  const minimum = Math.min(...values);
  const maximum = Math.max(...values);
  return Math.abs(maximum - minimum) < 0.005 ? money(minimum) : `${money(minimum)}–${money(maximum)}`;
}

function rmbRange(rows, key) {
  const values = rows.map((row) => Number(row[key])).filter((value) => Number.isFinite(value) && value > 0);
  if (!values.length) return "—";
  const minimum = Math.min(...values);
  const maximum = Math.max(...values);
  return Math.abs(maximum - minimum) < 0.005 ? rmb(minimum) : `${rmb(minimum)}–${rmb(maximum)}`;
}

function costRange(rows) {
  const values = rows.map((row) => Number(row.costRmb)).filter((value) => Number.isFinite(value) && value > 0);
  if (!values.length) return "—";
  const minimum = Math.min(...values);
  const maximum = Math.max(...values);
  return Math.abs(maximum - minimum) < 0.005 ? rmb(minimum) : `${rmb(minimum)}–${rmb(maximum)}`;
}

function feeForTier(tier, price) {
  if (!tier) return 0;
  if (tier.id === "under4") return price <= 3 ? 0.5 : 0.88;
  return Number(tier.fee || 0);
}

function pricingFactors(returnRate = null) {
  const pricing = appState.pricing;
  const referral = Number(pricing.referral_fee_rate ?? 0.15);
  const returns = returnRate === null ? Number(pricing.return_rate ?? 0.05) : Number(returnRate);
  const factor = (1 - returns) * (1 - referral) - returns * referral * referral;
  return {
    currencyRate: Number(pricing.currency_rate_rmb_to_usd || 7.2),
    referral,
    returns,
    firstLeg: Number(pricing.first_leg_shipping_usd || 0.3),
    disposal: Number(pricing.disposal_fee_usd ?? 0),
    factor
  };
}

function returnRateThreshold() {
  return Number(appState.pricing.return_rate_price_threshold_usd ?? 3);
}

function roundUpCents(value) {
  return Math.ceil((Number(value || 0) - 1e-9) * 100) / 100;
}

function calculatePriceInRegime(costRmbValue, tier, margin, returnRate, referencePrice) {
  const factors = pricingFactors(returnRate);
  const costUsd = Number(costRmbValue || 0) / factors.currencyRate;
  const shippingFee = feeForTier(tier, referencePrice);
  const fixedCost = costUsd + factors.firstLeg + shippingFee + factors.disposal * factors.returns;
  const rawPrice = fixedCost / Math.max(0.1, factors.factor - margin);
  const suggestedPrice = roundUpCents(rawPrice);
  const profitUsd = suggestedPrice * factors.factor - fixedCost;
  return { price: suggestedPrice, rawPrice, shippingFee, costUsd, fixedCost, profitUsd, profitRmb: profitUsd * factors.currencyRate, returnRate: factors.returns };
}

function calculatePrice(costRmbValue, tier, margin) {
  const threshold = returnRateThreshold();
  const configuredReturnRate = Number(appState.pricing.return_rate ?? 0.05);
  const withoutReturns = calculatePriceInRegime(costRmbValue, tier, margin, 0, threshold);
  if (withoutReturns.price <= threshold) return withoutReturns;

  const withReturns = calculatePriceInRegime(costRmbValue, tier, margin, configuredReturnRate, threshold + 0.01);
  const suggestedPrice = Math.max(roundUpCents(threshold + 0.01), withReturns.price);
  const factors = pricingFactors(configuredReturnRate);
  const shippingFee = feeForTier(tier, suggestedPrice);
  const fixedCost = withReturns.costUsd + factors.firstLeg + shippingFee + factors.disposal * factors.returns;
  const profitUsd = suggestedPrice * factors.factor - fixedCost;
  return {
    ...withReturns,
    price: suggestedPrice,
    shippingFee,
    fixedCost,
    profitUsd,
    profitRmb: profitUsd * factors.currencyRate,
    returnRate: configuredReturnRate
  };
}

function calculateSku(row, tier) {
  const margin40 = calculatePrice(row.costRmb, tier, 0.4);
  const margin30 = calculatePrice(row.costRmb, tier, 0.3);
  const margin20 = calculatePrice(row.costRmb, tier, 0.2);
  const margin10 = calculatePrice(row.costRmb, tier, 0.1);
  const breakEven = calculatePrice(row.costRmb, tier, 0);
  return {
    ...row,
    margin40: margin40.price,
    margin30: margin30.price,
    margin20: margin20.price,
    margin10: margin10.price,
    profit40Rmb: margin40.profitRmb,
    profit30Rmb: margin30.profitRmb,
    profit20Rmb: margin20.profitRmb,
    profit10Rmb: margin10.profitRmb,
    breakEven: breakEven.price,
    shippingFee: margin30.shippingFee,
    fees: {
      margin40: margin40.shippingFee,
      margin30: margin30.shippingFee,
      margin20: margin20.shippingFee,
      margin10: margin10.shippingFee,
      breakEven: breakEven.shippingFee
    },
    returnRates: {
      margin40: margin40.returnRate,
      margin30: margin30.returnRate,
      margin20: margin20.returnRate,
      margin10: margin10.returnRate,
      breakEven: breakEven.returnRate
    }
  };
}

function tierFromFee(value) {
  const fee = Number(value);
  if ([0.5, 0.88, 0.99].some((item) => Math.abs(fee - item) < 0.01)) return SHIPPING_TIERS.under4;
  if ([1.77, 2.05].some((item) => Math.abs(fee - item) < 0.01)) return SHIPPING_TIERS["4to8"];
  if ([2.6, 2.84].some((item) => Math.abs(fee - item) < 0.01)) return SHIPPING_TIERS["8to12"];
  if ([3.22, 3.48].some((item) => Math.abs(fee - item) < 0.01)) return SHIPPING_TIERS["12to16"];
  if ([3.72, 4.14].some((item) => Math.abs(fee - item) < 0.01)) return SHIPPING_TIERS["1to1_25"];
  return null;
}

function tierById(id) {
  return SHIPPING_TIERS[id] || null;
}

function tierForRecord(record) {
  return tierById(record.tierId) || tierFromFee(record.shippingFee) || Object.values(SHIPPING_TIERS).find((tier) => tier.label === record.tierLabel) || SHIPPING_TIERS.under4;
}

function tierOptionsHtml(selectedId = "under4") {
  return Object.values(SHIPPING_TIERS).map((tier) => {
    const fee = tier.id === "under4" ? tier.feeLabel : money(tier.fee);
    return `<option value="${tier.id}" ${tier.id === selectedId ? "selected" : ""}>${escapeHtml(tier.label)} · ${escapeHtml(fee)}</option>`;
  }).join("");
}

function recordShippingFee(tier, rows) {
  if (tier.id !== "under4") return tier.fee;
  const fees = rows.map((row) => Number(row.shippingFee)).filter((value) => Number.isFinite(value) && value > 0);
  if (!fees.length) return null;
  return Math.max(...fees);
}

function recalculatedRows(rows, tier) {
  return rows.map((row) => calculateSku({
    ...row,
    sku: String(row.sku || "").trim(),
    title: String(row.title || "").trim(),
    costRmb: Number(row.costRmb)
  }, tier));
}

function editableRecord(record) {
  return {
    ...record,
    rows: (record.rows || []).map((row) => ({
      sku: row.sku || "",
      title: row.title || "",
      costRmb: row.costRmb || "",
      purchaseLink: row.purchaseLink || "",
      linkedQty: row.linkedQty ?? null
    }))
  };
}

function normalizeQuickRecord(record) {
  const tier = tierForRecord(record);
  const rows = Array.isArray(record.rows) ? record.rows.map((row, index) => {
    const costRmb = Number(row.costRmb ?? row.cost) || null;
    const calculated40 = costRmb && tier ? calculatePrice(costRmb, tier, 0.4) : null;
    return {
      ...row,
      sku: row.sku || `SKU-${index + 1}`,
      title: row.title || "默认款",
      costRmb,
      margin40: Number(row.margin40) || calculated40?.price || null,
      margin30: Number(row.margin30 ?? row.margin15) || null,
      margin20: Number(row.margin20) || null,
      margin10: Number(row.margin10) || null,
      profit40Rmb: Number(row.profit40Rmb) || calculated40?.profitRmb || null,
      profit30Rmb: Number(row.profit30Rmb) || null,
      profit20Rmb: Number(row.profit20Rmb) || null,
      profit10Rmb: Number(row.profit10Rmb) || null,
      breakEven: Number(row.breakEven) || null,
      shippingFee: Number(row.shippingFee ?? record.shippingFee) || null,
      fees: { ...(row.fees || {}), ...(calculated40 ? { margin40: calculated40.shippingFee } : {}) },
      purchaseLink: row.purchaseLink || "",
      linkedQty: row.linkedQty ?? null
    };
  }) : [];
  return {
    ...record,
    id: record.id || `quick-${record.name}`,
    source: record.source || "quick",
    createdAt: record.createdAt || record.updatedAt || Date.now() / 1000,
    tierLabel: record.tierLabel || SHIPPING_TIERS[record.tierId]?.label || "配送档位待确认",
    rows
  };
}

function normalizeOutputRecord(record) {
  const sourceRows = Array.isArray(record.rows) ? record.rows : [];
  const finalRows = Array.isArray(record.finalRows) ? record.finalRows : [];
  const fee = Number(finalRows.find((row) => Number(row.fbaFee))?.fbaFee || record.shippingFee) || 0;
  const tier = tierFromFee(fee);
  const rows = sourceRows.map((row, index) => {
    const base = { sku: row.sku || `SKU-${index + 1}`, title: row.spec || row.title || "默认款", costRmb: Number(row.cost) || null };
    if (!base.costRmb || !tier) {
      return { ...base, shippingFee: fee || null, margin40: null, margin30: null, margin20: null, margin10: null, profit40Rmb: null, profit30Rmb: null, profit20Rmb: null, profit10Rmb: null, breakEven: null };
    }
    return calculateSku(base, tier);
  });
  return {
    id: `output-${record.name}`,
    name: record.name,
    source: "output",
    createdAt: record.firstCalculatedAt || record.updatedAt,
    updatedAt: record.updatedAt,
    tierLabel: tier?.label || (fee ? `历史配送费 $${fee.toFixed(2)}` : "配送档位待确认"),
    shippingFee: fee || null,
    rows,
    notes: rows.some((row) => !row.margin30) ? "历史文件缺少可用于五档重算的成本或配送档位，请人工补充。" : "根据历史结果中的成本和配送档位，按当前系统参数重算 40% / 30% / 20% / 10% 利润价格。"
  };
}

function rebuildNormalizedRecords() {
  const quick = appState.quickRecords.map(normalizeQuickRecord);
  const usedNames = new Set(quick.map((record) => record.name));
  const output = appState.outputRecords.filter((record) => !usedNames.has(record.name)).map(normalizeOutputRecord);
  appState.normalizedRecords = [...quick, ...output].sort((a, b) => toTimestampMs(b.createdAt) - toTimestampMs(a.createdAt));
}

function recordMatchesDate(record) {
  if (appState.historyDays === "all") return true;
  const timestamp = toTimestampMs(record.createdAt || record.updatedAt);
  if (!timestamp) return true;
  const cutoff = Date.now() - Number(appState.historyDays) * 86400000;
  return timestamp >= cutoff;
}

function visibleRecords() {
  const query = $("historySearch").value.trim().toLowerCase();
  return appState.normalizedRecords.filter((record) => {
    if (!recordMatchesDate(record)) return false;
    if (!query) return true;
    const searchable = [record.name, ...record.rows.flatMap((row) => [row.sku, row.title])].join(" ").toLowerCase();
    return searchable.includes(query);
  });
}

function pagedHistoryRecords(records) {
  const pageCount = Math.max(1, Math.ceil(records.length / appState.pageSize));
  appState.historyPage = Math.min(Math.max(1, appState.historyPage), pageCount);
  const start = (appState.historyPage - 1) * appState.pageSize;
  return records.slice(start, start + appState.pageSize);
}

function renderPagination(totalRecords) {
  const pageCount = Math.max(1, Math.ceil(totalRecords / appState.pageSize));
  $("pageLabel").textContent = totalRecords ? `第 ${appState.historyPage} / ${pageCount} 页` : "第 0 / 0 页";
  $("prevPageBtn").disabled = appState.historyPage <= 1;
  $("nextPageBtn").disabled = appState.historyPage >= pageCount || totalRecords === 0;
  $("pageSizeSelect").value = String(appState.pageSize);
}

function renderHistory() {
  const filteredRecords = visibleRecords();
  const records = pagedHistoryRecords(filteredRecords);
  $("historyCount").textContent = `${filteredRecords.length} 个产品`;
  $("historyEmpty").classList.toggle("is-hidden", filteredRecords.length > 0);
  $("historyRows").innerHTML = records.map((record) => {
    const selected = appState.selectedRecordId === record.id;
    const review = record.status === "needs_review";
    const sourceLabel = review ? "手写整理 · 待确认" : record.source === "output" ? "历史结果" : "快速定价";
    const feeText = Number(record.shippingFee) > 0 ? ` · ${money(record.shippingFee)}` : "";
    return `
      <tr data-record-id="${escapeHtml(record.id)}" class="${selected ? "is-selected" : ""}">
        <td class="product-cell"><strong>${escapeHtml(record.name)}</strong><span class="${review ? "review-dot" : ""}">${escapeHtml(sourceLabel)}</span></td>
        <td>${escapeHtml(formatDate(record))}</td>
        <td>${record.rows.length}款</td>
        <td><span class="tier-pill">${escapeHtml(record.tierLabel)}${escapeHtml(feeText)}</span></td>
        <td class="tier-40">${rangeText(record.rows, "margin40")}<br><span class="muted">预计毛利 ${rmbRange(record.rows, "profit40Rmb")}</span></td>
        <td class="tier-30">${rangeText(record.rows, "margin30")}<br><span class="muted">预计毛利 ${rmbRange(record.rows, "profit30Rmb")}</span></td>
        <td class="tier-20">${rangeText(record.rows, "margin20")}<br><span class="muted">预计毛利 ${rmbRange(record.rows, "profit20Rmb")}</span></td>
        <td class="tier-10">${rangeText(record.rows, "margin10")}<br><span class="muted">预计毛利 ${rmbRange(record.rows, "profit10Rmb")}</span></td>
        <td class="floor-price">${rangeText(record.rows, "breakEven")}</td>
      </tr>`;
  }).join("");
  $("historyRows").querySelectorAll("tr[data-record-id]").forEach((row) => {
    row.addEventListener("click", () => selectRecord(row.dataset.recordId));
  });
  if (appState.selectedRecordId && !records.some((record) => record.id === appState.selectedRecordId)) {
    if (!filteredRecords.some((record) => record.id === appState.selectedRecordId)) $("historyReport").classList.add("is-hidden");
  }
  renderPagination(filteredRecords.length);
  renderBulkEditRows();
}

function parameterRows(record) {
  const factors = pricingFactors();
  const shippingFees = record.rows.map((row) => Number(row.shippingFee || record.shippingFee)).filter((value) => value > 0);
  const feeText = shippingFees.length ? [...new Set(shippingFees.map((value) => money(value)))].join(" / ") : "待确认";
  return [
    ["成本范围", costRange(record.rows)],
    ["配送档位", record.tierLabel || "待确认"],
    ["配送费", feeText],
    ["人民币兑美元", `1 USD = ¥${factors.currencyRate.toFixed(2)}`],
    ["中国头程", money(factors.firstLeg)],
    ["亚马逊抽佣", `${(factors.referral * 100).toFixed(0)}%`],
    ["退货率", `0%（售价≤${money(returnRateThreshold())}） / ${(factors.returns * 100).toFixed(0)}%（售价>${money(returnRateThreshold())}）`],
    ["弃置费", factors.disposal > 0 ? money(factors.disposal) : "$0.00（不计）"]
  ];
}

function renderSkuRows(targetId, record) {
  $(targetId).innerHTML = record.rows.map((row) => {
    const fee = Number(row.shippingFee || row.fees?.margin30 || row.fees?.margin15 || record.shippingFee);
    return `<tr>
      <td>${escapeHtml(row.sku || "—")}</td><td>${escapeHtml(row.title || "—")}</td><td>${rmb(row.costRmb)}</td>
      <td>${money(fee)}</td><td class="tier-40">${money(row.margin40)}<br><span class="muted">毛利 ${rmb(row.profit40Rmb)}</span></td><td class="tier-30">${money(row.margin30)}<br><span class="muted">毛利 ${rmb(row.profit30Rmb)}</span></td><td class="tier-20">${money(row.margin20)}<br><span class="muted">毛利 ${rmb(row.profit20Rmb)}</span></td><td class="tier-10">${money(row.margin10)}<br><span class="muted">毛利 ${rmb(row.profit10Rmb)}</span></td><td class="floor-price">${money(row.breakEven)}</td>
    </tr>`;
  }).join("");
}

function renderBulkEditRows() {
  if (!$("bulkEditRows")) return;
  $("bulkEditPanel").classList.toggle("is-hidden", !appState.bulkEditOpen);
  if (!appState.bulkEditOpen) return;
  const records = pagedHistoryRecords(visibleRecords());
  $("bulkTierSelect").innerHTML = tierOptionsHtml($("bulkTierSelect").value || "under4");
  $("bulkEditRows").innerHTML = records.map((record) => {
    const selected = appState.bulkSelectedIds.has(record.id);
    return `<tr>
      <td><input data-bulk-id="${escapeHtml(record.id)}" type="checkbox" ${selected ? "checked" : ""}></td>
      <td><strong>${escapeHtml(record.name)}</strong></td>
      <td>${record.rows.length}款</td>
      <td>${escapeHtml(record.tierLabel || "待确认")} ${Number(record.shippingFee) > 0 ? escapeHtml(money(record.shippingFee)) : ""}</td>
      <td>${costRange(record.rows)}</td>
      <td class="floor-price">${rangeText(record.rows, "breakEven")}</td>
    </tr>`;
  }).join("");
  $("bulkEditRows").querySelectorAll("input[data-bulk-id]").forEach((input) => {
    input.addEventListener("change", () => {
      if (input.checked) appState.bulkSelectedIds.add(input.dataset.bulkId);
      else appState.bulkSelectedIds.delete(input.dataset.bulkId);
      $("bulkEditStatus").textContent = `已选择 ${appState.bulkSelectedIds.size} 个产品。`;
    });
  });
  $("bulkEditStatus").textContent = `已选择 ${appState.bulkSelectedIds.size} 个产品。`;
}

function renderEditSkuRows(rows) {
  $("editSkuRows").innerHTML = rows.map((row, index) => `<tr data-index="${index}">
    <td><input data-edit-field="sku" value="${escapeHtml(row.sku || "")}" placeholder="SKU-${String(index + 1).padStart(3, "0")}"></td>
    <td><input data-edit-field="title" value="${escapeHtml(row.title || "")}" placeholder="款式标题"></td>
    <td><input data-edit-field="costRmb" type="number" min="0" step="0.01" value="${escapeHtml(row.costRmb ?? "")}" placeholder="0.00"></td>
    <td><button class="remove-row" type="button" aria-label="删除第${index + 1}个SKU">×</button></td>
  </tr>`).join("");
  $("editSkuRows").querySelectorAll(".remove-row").forEach((button) => {
    button.addEventListener("click", () => {
      const rowsFromDom = editRowsFromDom();
      rowsFromDom.splice(Number(button.closest("tr").dataset.index), 1);
      renderEditSkuRows(rowsFromDom.length ? rowsFromDom : [blankRow()]);
    });
  });
}

function editRowsFromDom() {
  return [...$("editSkuRows").querySelectorAll("tr")].map((row) => ({
    sku: row.querySelector('[data-edit-field="sku"]').value.trim(),
    title: row.querySelector('[data-edit-field="title"]').value.trim(),
    costRmb: row.querySelector('[data-edit-field="costRmb"]').value.trim()
  }));
}

function editableRows() {
  return editRowsFromDom().filter((row) => row.sku || row.title || Number(row.costRmb));
}

function recordById(recordId) {
  return appState.normalizedRecords.find((record) => record.id === recordId);
}

function startRecordEdit() {
  const record = recordById(appState.selectedRecordId);
  if (!record) return;
  const editable = editableRecord(record);
  const tier = tierForRecord(record);
  appState.editingRecordId = record.id;
  $("editProductName").value = editable.name || "";
  $("editDateLabel").value = editable.dateLabel || "";
  $("editTierSelect").innerHTML = tierOptionsHtml(tier.id);
  $("editNotes").value = editable.notes || "";
  $("recordEditStatus").textContent = "";
  renderEditSkuRows(editable.rows.length ? editable.rows : [blankRow()]);
  $("recordEditPanel").classList.remove("is-hidden");
  $("recordEditPanel").scrollIntoView({ behavior: "smooth", block: "start" });
}

function stopRecordEdit() {
  appState.editingRecordId = "";
  $("recordEditPanel").classList.add("is-hidden");
  $("recordEditStatus").textContent = "";
}

function buildEditedRecord(baseRecord, rows, tier) {
  const calculatedRows = recalculatedRows(rows, tier);
  return {
    ...baseRecord,
    name: $("editProductName").value.trim(),
    dateLabel: $("editDateLabel").value.trim(),
    tierId: tier.id,
    tierLabel: tier.label,
    shippingFee: recordShippingFee(tier, calculatedRows),
    rows: calculatedRows,
    notes: $("editNotes").value.trim()
  };
}

async function saveRecordEdit() {
  const baseRecord = recordById(appState.editingRecordId);
  if (!baseRecord) return;
  const rows = editableRows();
  const error = validateRows(rows);
  const tier = tierById($("editTierSelect").value);
  if (!$("editProductName").value.trim()) {
    $("recordEditStatus").textContent = "请填写产品名称。";
    return;
  }
  if (error) {
    $("recordEditStatus").textContent = error;
    return;
  }
  $("saveRecordEditBtn").disabled = true;
  $("recordEditStatus").textContent = "正在保存并重算…";
  try {
    const saved = await saveQuickRecord(buildEditedRecord(baseRecord, rows, tier));
    appState.selectedRecordId = saved.id;
    stopRecordEdit();
    renderHistory();
    selectRecord(saved.id, { scroll: false });
    setStatus(`“${saved.name}”已更新。`, "ok");
  } catch (errorObject) {
    $("recordEditStatus").textContent = errorObject.message || "保存失败";
  } finally {
    $("saveRecordEditBtn").disabled = false;
  }
}

function buildBulkEditedRecord(record, tier) {
  const rows = recalculatedRows(record.rows, tier);
  return {
    ...record,
    tierId: tier.id,
    tierLabel: tier.label,
    shippingFee: recordShippingFee(tier, rows),
    rows,
    notes: record.notes || ""
  };
}

async function saveBulkEdit() {
  const ids = [...appState.bulkSelectedIds];
  const tier = tierById($("bulkTierSelect").value);
  if (!ids.length) {
    $("bulkEditStatus").textContent = "请先选择要修改的产品。";
    return;
  }
  $("saveBulkEditBtn").disabled = true;
  $("bulkEditStatus").textContent = `正在保存 ${ids.length} 个产品…`;
  try {
    for (const id of ids) {
      const record = recordById(id);
      if (record) await saveQuickRecord(buildBulkEditedRecord(record, tier));
    }
    appState.bulkSelectedIds.clear();
    renderHistory();
    if (appState.selectedRecordId) selectRecord(appState.selectedRecordId, { scroll: false });
    $("bulkEditStatus").textContent = "批量保存完成。";
    setStatus(`已批量更新 ${ids.length} 个产品。`, "ok");
  } catch (errorObject) {
    $("bulkEditStatus").textContent = errorObject.message || "批量保存失败";
  } finally {
    $("saveBulkEditBtn").disabled = false;
  }
}

function renderRecordReport(record) {
  $("historyReport").classList.remove("is-hidden");
  $("reportProductName").textContent = `${record.name} · 定价分析报告`;
  $("reportMeta").textContent = `${formatDate(record)} · ${record.rows.length}个SKU · ${record.tierLabel}`;
  $("reportReviewTag").classList.toggle("is-hidden", record.status !== "needs_review");
  $("reportMargin40").textContent = rangeText(record.rows, "margin40");
  $("reportMargin30").textContent = rangeText(record.rows, "margin30");
  $("reportMargin20").textContent = rangeText(record.rows, "margin20");
  $("reportMargin10").textContent = rangeText(record.rows, "margin10");
  $("reportProfit40").textContent = `预计毛利 ${rmbRange(record.rows, "profit40Rmb")}`;
  $("reportProfit30").textContent = `预计毛利 ${rmbRange(record.rows, "profit30Rmb")}`;
  $("reportProfit20").textContent = `预计毛利 ${rmbRange(record.rows, "profit20Rmb")}`;
  $("reportProfit10").textContent = `预计毛利 ${rmbRange(record.rows, "profit10Rmb")}`;
  $("reportBreakEven").textContent = rangeText(record.rows, "breakEven");
  $("reportParameters").innerHTML = parameterRows(record).map(([label, value]) => `<dt>${escapeHtml(label)}</dt><dd>${escapeHtml(value)}</dd>`).join("");
  if (record.status === "needs_review") {
    const factors = pricingFactors();
    $("reportFormula").innerHTML = `<p><strong>当前展示：</strong>已用当前 40% / 30% / 20% / 10% 规则，根据录入的成本和配送档位重算。</p><p><strong>退货规则：</strong>售价不高于 ${money(returnRateThreshold())} 时退货率按 0% 计算；售价高于 ${money(returnRateThreshold())} 时按 ${(factors.returns * 100).toFixed(0)}% 计算。</p><p><strong>待确认项：</strong>“待确认”仅表示原始手写资料的产品名、成本或配送档位仍建议人工核对。</p><p><strong>预计毛利：</strong>40%、30%、20% 和 10% 价格下的利润已按当前汇率换算为人民币。</p>`;
  } else {
    const factors = pricingFactors();
    $("reportFormula").innerHTML = `<p><strong>单件美元成本</strong> = 人民币成本 ÷ ${factors.currencyRate.toFixed(2)}</p><p><strong>固定成本</strong> = 商品成本 + 头程 ${money(factors.firstLeg)} + 配送费；所有产品弃置费按 $0.00 计算。</p><p><strong>退货规则</strong>：售价不高于 ${money(returnRateThreshold())} 时退货率按 0% 计算；售价高于 ${money(returnRateThreshold())} 时按 ${(factors.returns * 100).toFixed(0)}% 计算。</p><p><strong>保本价</strong> = 固定成本 ÷ 净入账系数；<strong>利润价</strong> = 固定成本 ÷（净入账系数 − 目标利润率）</p><p>各档结果统一向上取整到美分；≤4oz根据计算售价自动切换 $0.50 / $0.88。</p>`;
  }
  $("reportNotesBlock").classList.toggle("is-hidden", !record.notes);
  $("reportNotes").textContent = record.notes || "";
  renderSkuRows("reportSkuRows", record);
  if (appState.editingRecordId !== record.id) stopRecordEdit();
}

function selectRecord(recordId, options = {}) {
  const record = appState.normalizedRecords.find((item) => item.id === recordId);
  if (!record) return;
  appState.selectedRecordId = record.id;
  renderHistory();
  renderRecordReport(record);
  if (options.scroll !== false) $("historyReport").scrollIntoView({ behavior: "smooth", block: "start" });
}

function switchView(view) {
  const history = view === "history";
  $("historyView").classList.toggle("is-hidden", !history);
  $("quickView").classList.toggle("is-hidden", history);
  $("pageTitle").textContent = history ? "产品定价记录" : "新品快速定价";
  $("pageSubtitle").textContent = history ? "先看产品汇总，再查看SKU和核心计算过程" : "输入SKU、款式标题和成本价即可批量计算";
  document.querySelectorAll(".nav-button[data-view]").forEach((button) => {
    const active = button.dataset.view === view;
    button.classList.toggle("is-active", active);
    button.setAttribute("aria-pressed", String(active));
  });
}

function blankRow() {
  return { sku: "", title: "", costRmb: "" };
}

function quickRowsFromDom() {
  return [...$("quickRows").querySelectorAll("tr")].map((row) => ({
    sku: row.querySelector('[data-field="sku"]').value.trim(),
    title: row.querySelector('[data-field="title"]').value.trim(),
    costRmb: row.querySelector('[data-field="costRmb"]').value.trim()
  }));
}

function renderQuickRows(rows = []) {
  const usable = rows.length ? rows : [blankRow(), blankRow(), blankRow()];
  $("quickRows").innerHTML = usable.map((row, index) => `<tr data-index="${index}">
    <td><input data-field="sku" value="${escapeHtml(row.sku)}" placeholder="SKU-${String(index + 1).padStart(3, "0")}"></td>
    <td><input data-field="title" value="${escapeHtml(row.title)}" placeholder="款式标题"></td>
    <td><input data-field="costRmb" type="number" min="0" step="0.01" value="${escapeHtml(row.costRmb)}" placeholder="0.00"></td>
    <td><button class="remove-row" type="button" aria-label="删除第${index + 1}行">×</button></td>
  </tr>`).join("");
  $("quickRows").querySelectorAll("input").forEach((input) => input.addEventListener("input", updateQuickSummary));
  $("quickRows").querySelectorAll(".remove-row").forEach((button) => button.addEventListener("click", () => {
    const row = button.closest("tr");
    const current = quickRowsFromDom();
    current.splice(Number(row.dataset.index), 1);
    renderQuickRows(current.length ? current : [blankRow()]);
  }));
  updateQuickSummary();
}

function validQuickRows() {
  return quickRowsFromDom().filter((row) => row.sku || row.title || Number(row.costRmb));
}

function validateRows(rows) {
  if (!rows.length) return "请至少添加一个SKU。";
  if (rows.some((row) => !row.sku)) return "每一行都需要填写产品SKU。";
  if (rows.some((row) => !row.title)) return "每一行都需要填写款式标题。";
  if (rows.some((row) => !(Number(row.costRmb) > 0))) return "每一行成本价都必须大于0。";
  const skus = rows.map((row) => row.sku.toLowerCase());
  if (new Set(skus).size !== skus.length) return "发现重复SKU，请修改后再计算。";
  return "";
}

function updateQuickSummary() {
  const rows = validQuickRows();
  const error = rows.length ? validateRows(rows) : "";
  $("quickRowSummary").textContent = `${rows.length} 个待计算SKU`;
  $("quickValidation").textContent = error;
  $("quickValidation").classList.toggle("is-error", Boolean(error));
}

function parseBulkText(text) {
  const lines = String(text || "").split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  const parsed = [];
  lines.forEach((line, index) => {
    let columns = line.includes("\t") ? line.split("\t") : line.split(/\s{2,}|[,，]/);
    columns = columns.map((value) => value.trim()).filter(Boolean);
    if (index === 0 && columns.join(" ").match(/sku.*(标题|款式).*(成本|价格)/i)) return;
    if (columns.length < 3) {
      const loose = line.split(/\s+/).filter(Boolean);
      if (loose.length >= 3) columns = [loose[0], loose.slice(1, -1).join(" "), loose.at(-1)];
    }
    if (columns.length >= 3) parsed.push({ sku: columns[0], title: columns.slice(1, -1).join(" "), costRmb: columns.at(-1).replace(/[¥￥元]/g, "") });
  });
  return parsed;
}

async function saveQuickRecord(record) {
  const response = await fetch("/api/quick-history", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ record })
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok || !payload.ok) throw new Error(payload.error || "保存定价记录失败");
  appState.quickRecords = Array.isArray(payload.records) ? payload.records : [payload.record, ...appState.quickRecords];
  rebuildNormalizedRecords();
  return normalizeQuickRecord(payload.record);
}

function notifyEmbeddedPricingRecord(record) {
  const params = new URLSearchParams(location.search);
  if (params.get("embed") !== "1" || !params.get("project_id") || !record?.id) return;
  let parentOrigin;
  try { parentOrigin = new URL(document.referrer).origin; } catch { return; }
  if (!["http://127.0.0.1:8770", "http://localhost:8770"].includes(parentOrigin)) return;
  window.parent.postMessage({
    type: "store-ops-pricing-record",
    projectId: params.get("project_id"),
    recordId: record.id,
    productName: record.name,
    skus: (record.rows || []).map((row) => row.sku).filter(Boolean)
  }, parentOrigin);
}

function notifyLatestEmbeddedPricingRecord() {
  const name = new URLSearchParams(location.search).get("product_name");
  if (!name) return;
  const records = appState.quickRecords.filter((record) => record?.name === name && record.source === "quick");
  records.sort((a, b) => Number(b.updatedAt || b.createdAt || 0) - Number(a.updatedAt || a.createdAt || 0));
  if (records[0]) notifyEmbeddedPricingRecord(records[0]);
}

async function calculateAndSave() {
  const name = $("quickProductName").value.trim();
  const rows = validQuickRows();
  const error = validateRows(rows);
  if (!name) {
    $("quickValidation").textContent = "请填写产品名称。";
    $("quickValidation").classList.add("is-error");
    $("quickProductName").focus();
    return;
  }
  if (error) {
    $("quickValidation").textContent = error;
    $("quickValidation").classList.add("is-error");
    return;
  }
  const tier = SHIPPING_TIERS[appState.selectedTierId];
  const calculatedRows = rows.map((row) => calculateSku({ sku: row.sku, title: row.title, costRmb: Number(row.costRmb) }, tier));
  const dateValue = $("quickPricingDate").value;
  const date = dateValue ? new Date(`${dateValue}T12:00:00`) : new Date();
  const record = {
    name,
    source: "quick",
    createdAt: date.getTime() / 1000,
    tierId: tier.id,
    tierLabel: tier.label,
    shippingFee: tier.id === "under4" ? null : tier.fee,
    assumptions: { ...appState.pricing, dimensions: tier.dimensions, weightG: tier.weightG },
    rows: calculatedRows,
    notes: "由新品快速定价生成。默认尺寸重量仅在后台用于匹配配送档位，页面不展示。"
  };
  $("calculateQuickBtn").disabled = true;
  setStatus("正在计算并保存…");
  try {
    const saved = await saveQuickRecord(record);
    appState.latestSavedId = saved.id;
    $("quickResult").classList.remove("is-hidden");
    $("quickResultName").textContent = `${saved.name} · 计算结果`;
    $("quickMargin40").textContent = rangeText(saved.rows, "margin40");
    $("quickMargin30").textContent = rangeText(saved.rows, "margin30");
    $("quickMargin20").textContent = rangeText(saved.rows, "margin20");
    $("quickMargin10").textContent = rangeText(saved.rows, "margin10");
    $("quickProfit40").textContent = `预计毛利 ${rmbRange(saved.rows, "profit40Rmb")}`;
    $("quickProfit30").textContent = `预计毛利 ${rmbRange(saved.rows, "profit30Rmb")}`;
    $("quickProfit20").textContent = `预计毛利 ${rmbRange(saved.rows, "profit20Rmb")}`;
    $("quickProfit10").textContent = `预计毛利 ${rmbRange(saved.rows, "profit10Rmb")}`;
    $("quickBreakEven").textContent = rangeText(saved.rows, "breakEven");
    renderSkuRows("quickResultRows", saved);
    renderHistory();
    setStatus(`“${saved.name}”已保存，共${saved.rows.length}个SKU。`, "ok");
    notifyEmbeddedPricingRecord(saved);
    $("quickResult").scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (errorObject) {
    setStatus(errorObject.message || "保存失败", "error");
  } finally {
    $("calculateQuickBtn").disabled = false;
  }
}

async function copySelectedReport() {
  const record = appState.normalizedRecords.find((item) => item.id === appState.selectedRecordId);
  if (!record) return;
  const text = `${record.name}\n40%利润价：${rangeText(record.rows, "margin40")}（预计毛利 ${rmbRange(record.rows, "profit40Rmb")}）\n30%利润价：${rangeText(record.rows, "margin30")}（预计毛利 ${rmbRange(record.rows, "profit30Rmb")}）\n20%利润价：${rangeText(record.rows, "margin20")}（预计毛利 ${rmbRange(record.rows, "profit20Rmb")}）\n10%利润价：${rangeText(record.rows, "margin10")}（预计毛利 ${rmbRange(record.rows, "profit10Rmb")}）\n保底价：${rangeText(record.rows, "breakEven")}`;
  try {
    await navigator.clipboard.writeText(text);
    setStatus("五档价格已复制。", "ok");
  } catch (error) {
    setStatus("浏览器未允许复制，请手动选择价格。", "warn");
  }
}

async function loadData() {
  setStatus("正在读取定价记录…");
  try {
    const [stateResponse, quickResponse, historyResponse] = await Promise.all([
      fetch("/api/state"),
      fetch("/api/quick-history"),
      fetch("/api/history")
    ]);
    if (stateResponse.ok) {
      const statePayload = await stateResponse.json();
      if (statePayload.pricing) appState.pricing = { ...appState.pricing, ...statePayload.pricing };
    }
    appState.quickRecords = quickResponse.ok ? (await quickResponse.json()).records || [] : [];
    appState.outputRecords = historyResponse.ok ? (await historyResponse.json()).records?.filter((record) => !record.error) || [] : [];
    rebuildNormalizedRecords();
    renderHistory();
    notifyLatestEmbeddedPricingRecord();
    const first = visibleRecords()[0];
    if (first) selectRecord(first.id, { scroll: false });
    setStatus(`已载入${appState.normalizedRecords.length}个产品记录。`, "ok");
  } catch (error) {
    setStatus("部分历史记录读取失败，快速定价仍可继续使用。", "warn");
    console.warn(error);
  }
}

function initEvents() {
  document.querySelectorAll(".nav-button[data-view]").forEach((button) => button.addEventListener("click", () => switchView(button.dataset.view)));
  document.querySelectorAll(".filter-button").forEach((button) => button.addEventListener("click", () => {
    if (!button.dataset.days) return;
    appState.historyDays = button.dataset.days === "all" ? "all" : Number(button.dataset.days);
    appState.historyPage = 1;
    document.querySelectorAll(".filter-button").forEach((item) => {
      if (!item.dataset.days) return;
      const active = item === button;
      item.classList.toggle("is-active", active);
      item.setAttribute("aria-pressed", String(active));
    });
    renderHistory();
  }));
  $("historySearch").addEventListener("input", () => {
    appState.historyPage = 1;
    renderHistory();
  });
  $("pageSizeSelect").addEventListener("change", () => {
    appState.pageSize = Number($("pageSizeSelect").value) || 15;
    appState.historyPage = 1;
    renderHistory();
  });
  $("prevPageBtn").addEventListener("click", () => {
    appState.historyPage -= 1;
    renderHistory();
  });
  $("nextPageBtn").addEventListener("click", () => {
    appState.historyPage += 1;
    renderHistory();
  });
  $("toggleBulkEditBtn").addEventListener("click", () => {
    appState.bulkEditOpen = !appState.bulkEditOpen;
    renderHistory();
  });
  $("closeBulkEditBtn").addEventListener("click", () => {
    appState.bulkEditOpen = false;
    renderHistory();
  });
  $("selectPageRecordsBtn").addEventListener("click", () => {
    pagedHistoryRecords(visibleRecords()).forEach((record) => appState.bulkSelectedIds.add(record.id));
    renderBulkEditRows();
  });
  $("clearBulkSelectionBtn").addEventListener("click", () => {
    appState.bulkSelectedIds.clear();
    renderBulkEditRows();
  });
  $("saveBulkEditBtn").addEventListener("click", saveBulkEdit);
  $("editRecordBtn").addEventListener("click", startRecordEdit);
  $("cancelRecordEditBtn").addEventListener("click", stopRecordEdit);
  $("addEditSkuBtn").addEventListener("click", () => renderEditSkuRows([...editRowsFromDom(), blankRow()]));
  $("saveRecordEditBtn").addEventListener("click", saveRecordEdit);
  $("copyReportBtn").addEventListener("click", copySelectedReport);
  $("shippingTierButtons").querySelectorAll(".tier-button").forEach((button) => button.addEventListener("click", () => {
    appState.selectedTierId = button.dataset.tier;
    $("shippingTierButtons").querySelectorAll(".tier-button").forEach((item) => {
      const active = item === button;
      item.classList.toggle("is-active", active);
      item.setAttribute("aria-pressed", String(active));
    });
  }));
  document.querySelectorAll(".entry-tab").forEach((button) => button.addEventListener("click", () => {
    const paste = button.dataset.entryMode === "paste";
    document.querySelectorAll(".entry-tab").forEach((item) => {
      const active = item === button;
      item.classList.toggle("is-active", active);
      item.setAttribute("aria-selected", String(active));
    });
    $("pastePanel").classList.toggle("is-hidden", !paste);
  }));
  $("parseBulkBtn").addEventListener("click", () => {
    const rows = parseBulkText($("bulkInput").value);
    if (!rows.length) {
      $("quickValidation").textContent = "没有识别到三列数据，请从Excel复制SKU、标题和成本价。";
      $("quickValidation").classList.add("is-error");
      return;
    }
    renderQuickRows(rows);
    $("quickValidation").textContent = `已识别${rows.length}行，可以直接修改表格。`;
    $("quickValidation").classList.remove("is-error");
  });
  $("addQuickRowBtn").addEventListener("click", () => renderQuickRows([...quickRowsFromDom(), blankRow()]));
  $("clearRowsBtn").addEventListener("click", () => {
    $("bulkInput").value = "";
    renderQuickRows([blankRow(), blankRow(), blankRow()]);
    $("quickResult").classList.add("is-hidden");
  });
  $("calculateQuickBtn").addEventListener("click", calculateAndSave);
  $("viewSavedRecordBtn").addEventListener("click", () => {
    if (!appState.latestSavedId) return;
    switchView("history");
    selectRecord(appState.latestSavedId);
  });
}

function applyLaunchContext() {
  const params = new URLSearchParams(window.location.search);
  const embedded = params.get("embed") === "1";
  if (embedded) document.documentElement.dataset.embed = "1";
  document.documentElement.dataset.theme = params.get("theme") === "dark" ? "dark" : "light";
  if (params.get("view") !== "quick") return;
  switchView("quick");
  const productName = (params.get("product_name") || "").trim().slice(0, 160);
  const skus = [...new Set(params.getAll("sku").map((value) => value.trim()).filter(Boolean))].slice(0, 50);
  if (productName) $("quickProductName").value = productName;
  if (skus.length) renderQuickRows(skus.map((sku) => ({ sku, title: productName, costRmb: "" })));
  if (embedded) $("quickProductName").focus({ preventScroll: true });
}

window.addEventListener("message", (event) => {
  if (document.documentElement.dataset.embed !== "1") return;
  if (!["http://127.0.0.1:8770", "http://localhost:8770"].includes(event.origin)) return;
  if (event.data?.type !== "store-ops-theme") return;
  document.documentElement.dataset.theme = event.data.theme === "dark" ? "dark" : "light";
});

document.addEventListener("DOMContentLoaded", () => {
  const today = new Date();
  $("quickPricingDate").value = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}-${String(today.getDate()).padStart(2, "0")}`;
  $("editTierSelect").innerHTML = tierOptionsHtml();
  $("bulkTierSelect").innerHTML = tierOptionsHtml();
  renderQuickRows();
  initEvents();
  applyLaunchContext();
  loadData();
});
