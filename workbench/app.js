const stateUrl = "/api/state";
let state = null;
let pricingMode = "bundle";

const icons = {
  save: '<svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2Z"/><path d="M17 21v-8H7v8"/><path d="M7 3v5h8"/></svg>',
  file: '<svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8Z"/><path d="M14 2v6h6"/><path d="M8 13h8"/><path d="M8 17h8"/><path d="M8 9h2"/></svg>',
  calculator: '<svg viewBox="0 0 24 24" fill="none" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="16" height="20" x="4" y="2" rx="2"/><line x1="8" x2="16" y1="6" y2="6"/><line x1="16" x2="16" y1="14" y2="18"/><path d="M8 10h.01"/><path d="M12 10h.01"/><path d="M16 10h.01"/><path d="M8 14h.01"/><path d="M12 14h.01"/><path d="M8 18h.01"/><path d="M12 18h.01"/></svg>'
};

function $(id) {
  return document.getElementById(id);
}

function num(value, fallback = 0) {
  if (value === "" || value === null || value === undefined) return fallback;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function money(value) {
  return Number(value || 0).toLocaleString("en-US", {
    style: "currency",
    currency: "USD",
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function pct(value) {
  return `${(Number(value || 0) * 100).toFixed(1)}%`;
}

function filename(path) {
  return String(path || "").split("/").pop();
}

function slugifySku(value) {
  const text = String(value || "")
    .trim()
    .replace(/[^\p{L}\p{N}]+/gu, "-")
    .replace(/^-+|-+$/g, "");
  return text || "Auto-SKU";
}

function setStatus(message, tone = "") {
  const band = $("statusBand");
  band.textContent = message;
  band.className = `status-band ${tone}`.trim();
}

function setBusy(isBusy) {
  ["saveBtn", "draftBtn", "pricingBtn"].forEach((id) => {
    $(id).disabled = isBusy;
  });
  document.querySelectorAll(".file-button button").forEach((button) => {
    button.disabled = isBusy;
  });
}

function renderIcons() {
  document.querySelectorAll("[data-icon]").forEach((node) => {
    node.innerHTML = icons[node.dataset.icon] || "";
  });
}

function manualDimensions() {
  state.workbench = state.workbench || {};
  if (!Array.isArray(state.workbench.manual_dimensions)) {
    state.workbench.manual_dimensions = [];
  }
  if (!state.workbench.manual_dimensions.length) {
    state.workbench.manual_dimensions.push({
      sku: "",
      title: "",
      length_cm: "",
      width_cm: "",
      height_cm: "",
      weight_g: ""
    });
  }
  return state.workbench.manual_dimensions;
}

function productLinks() {
  state.workbench = state.workbench || {};
  if (!Array.isArray(state.workbench.product_links)) {
    state.workbench.product_links = [""];
  }
  if (!state.workbench.product_links.length) {
    state.workbench.product_links.push("");
  }
  return state.workbench.product_links;
}

function bundlePackCount(components) {
  return components.reduce((sum, component) => sum + num(component.pcs_per_bundle), 0);
}

function renderFileList(id, files) {
  const list = $(id);
  const items = files || [];
  list.innerHTML = items.length
    ? items.slice(0, 4).map((file) => `<li title="${file.path}">${file.name}</li>`).join("")
    : "<li>暂无文件</li>";
}

function renderSourceFiles() {
  const files = state.inputFiles || {};
  renderFileList("purchaseOrderFiles", files.purchase_order);
  renderFileList("productLinkFiles", files.product_links);
  renderFileList("competitorFiles", files.competitors);
  $("builtinRules").innerHTML = `
    汇率 ${state.pricing.currency_rate_rmb_to_usd} / 佣金 ${pct(state.pricing.referral_fee_rate)} /
    退货 ${pct(state.pricing.return_rate)} / 目标利润 ${pct(state.pricing.target_margin_min)}
  `;
}

function renderInputs() {
  const product = state.product;
  $("sku").value = product.sku || "";
  $("variant").value = product.variant || "";
  $("packCountDisplay").value = bundlePackCount(product.components || []);
  $("lengthCm").value = product.package?.length_cm ?? "";
  $("widthCm").value = product.package?.width_cm ?? "";
  $("heightCm").value = product.package?.height_cm ?? "";
  $("weightG").value = product.package?.weight_g ?? "";
  renderManualDimensions();
  renderOutputs();
}

function renderManualDimensions() {
  const tbody = $("manualDimensionRows");
  tbody.innerHTML = "";
  manualDimensions().forEach((row, index) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><input data-manual-field="title" data-manual-index="${index}" value="${row.title || ""}"></td>
      <td><input data-manual-field="sku" data-manual-index="${index}" value="${row.sku || ""}"></td>
      <td><input data-manual-field="length_cm" data-manual-index="${index}" type="number" step="0.1" min="0" value="${row.length_cm ?? ""}"></td>
      <td><input data-manual-field="width_cm" data-manual-index="${index}" type="number" step="0.1" min="0" value="${row.width_cm ?? ""}"></td>
      <td><input data-manual-field="height_cm" data-manual-index="${index}" type="number" step="0.1" min="0" value="${row.height_cm ?? ""}"></td>
      <td><input data-manual-field="weight_g" data-manual-index="${index}" type="number" step="1" min="0" value="${row.weight_g ?? ""}"></td>
    `;
    tbody.appendChild(tr);
  });
  tbody.querySelectorAll("input").forEach((input) => {
    input.addEventListener("input", updateLocalFromForm);
  });
}

function renderOutputs() {
  const outputs = state.outputs || {};
  $("fileLinks").innerHTML = [
    outputs.draft?.exists ? `<a href="/files/output/${encodeURIComponent(filename(outputs.draft.path))}">数据表</a>` : "",
    outputs.result?.exists ? `<a href="/files/output/${encodeURIComponent(filename(outputs.result.path))}">分析表</a>` : "",
    outputs.report?.exists ? `<a href="/files/output/${encodeURIComponent(filename(outputs.report.path))}">报告</a>` : "",
    outputs.upload?.exists ? `<a href="/files/output/${encodeURIComponent(filename(outputs.upload.path))}">导入表</a>` : ""
  ].join("");

  const resultRows = outputs.result?.rows || [];
  renderAnalysisSummaries(resultRows[0]);
  $("scenarioRows").innerHTML = resultRows.length
    ? resultRows.map((item) => `
        <tr>
          <td>${item.SKU || "-"}</td>
          <td class="numeric">${item["包装长cm"] ?? "-"}</td>
          <td class="numeric">${item["包装宽cm"] ?? "-"}</td>
          <td class="numeric">${item["包装高cm"] ?? "-"}</td>
          <td class="numeric">${item["包装重量g"] ?? "-"}</td>
          <td class="numeric">${money(item["建议售价USD"])}</td>
          <td class="numeric">${money(item["FBA费"])}</td>
          <td class="numeric">${money(item["期望利润"])}</td>
          <td class="numeric">${pct(item["期望利润率"])}</td>
        </tr>
      `).join("")
    : '<tr><td colspan="9" class="empty">暂无定价分析</td></tr>';
}

function renderAnalysisSummaries(row) {
  if (!row) {
    $("deliverySummary").textContent = "产品配送费价格：等待生成定价分析表";
    $("competitorSummary").textContent = "竞品价格分析：等待生成定价分析表";
    $("pricingNarrative").textContent = "产品条件、利润、定价分析：等待生成定价分析表";
    return;
  }
  $("deliverySummary").textContent = `产品配送费价格：计费重量约 ${Number(row["计费重量lb"] || 0).toFixed(2)} lb，匹配 ${row["FBA档位"] || "FBA档位"}，FBA费 ${money(row["FBA费"])}。`;
  $("competitorSummary").textContent = `竞品价格分析：最低对比价约 ${money(row["竞品最低对比价USD"])}，当前建议价 ${money(row["建议售价USD"])}，差价 ${money(row["相对竞品差价"])}，结论是${row["结论"] || "待判断"}。`;
  $("pricingNarrative").textContent = `产品条件、利润、定价分析：按当前尺寸重量、采购成本和内置费率测算，建议价 ${money(row["建议售价USD"])} 时，期望利润约 ${money(row["期望利润"])}，利润率 ${pct(row["期望利润率"])}。`;
}

function updateLocalFromForm() {
  const product = state.product;
  product.variant = $("variant").value.trim();
  product.sku = $("sku").value.trim() || slugifySku(product.variant);
  product.package.length_cm = num($("lengthCm").value);
  product.package.width_cm = num($("widthCm").value);
  product.package.height_cm = num($("heightCm").value);
  product.package.weight_g = num($("weightG").value);
  product.comparison_unit_quantity = bundlePackCount(product.components || []);

  document.querySelectorAll("#manualDimensionRows input").forEach((input) => {
    const row = manualDimensions()[Number(input.dataset.manualIndex)];
    const field = input.dataset.manualField;
    row[field] = field === "sku" || field === "title" ? input.value.trim() : input.value;
  });
  manualDimensions().forEach((row) => {
    if (!row.sku && row.title) {
      row.sku = slugifySku(row.title);
    }
  });
}

async function postJson(url, payload = {}) {
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });
  const data = await response.json();
  if (!response.ok || data.ok === false) {
    throw new Error(data.error || data.command?.stderr || data.pricing?.stderr || "操作失败");
  }
  return data;
}

async function loadState(message = "工作台已加载。") {
  const response = await fetch(stateUrl);
  state = await response.json();
  renderSourceFiles();
  renderInputs();
  setStatus(message, "ok");
}

async function saveConfig() {
  updateLocalFromForm();
  setBusy(true);
  setStatus("正在保存输入...", "warn");
  try {
    await postJson("/api/config", { product: state.product });
    const data = await postJson("/api/workbench", { workbench: state.workbench });
    state = data.state;
    renderSourceFiles();
    renderInputs();
    setStatus("输入已保存。", "ok");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    setBusy(false);
  }
}

async function uploadFiles(form) {
  const type = form.dataset.uploadType;
  const input = form.querySelector('input[type="file"]');
  if (!input.files.length) {
    setStatus("请选择要上传的文件。", "warn");
    return;
  }
  setBusy(true);
  setStatus("正在上传资料...", "warn");
  try {
    const payload = new FormData();
    payload.append("type", type);
    [...input.files].forEach((file) => payload.append("files", file));
    const response = await fetch("/api/upload", {
      method: "POST",
      body: payload
    });
    const data = await response.json();
    if (!response.ok || data.ok === false) {
      throw new Error(data.error || "上传失败");
    }
    input.value = "";
    state = data.state;
    renderSourceFiles();
    renderInputs();
    setStatus(`已上传 ${data.saved.length} 个文件。`, "ok");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    setBusy(false);
  }
}

async function generateDraft() {
  await saveConfig();
  setBusy(true);
  setStatus("正在生成数据表...", "warn");
  try {
    const data = await postJson("/api/generate-draft", { mode: pricingMode });
    state = data.state;
    renderInputs();
    setStatus("数据表已生成。", "ok");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    setBusy(false);
  }
}

async function runPricing() {
  await saveConfig();
  setBusy(true);
  setStatus("正在生成定价分析表...", "warn");
  try {
    const data = await postJson("/api/run-pricing", { mode: pricingMode });
    state = data.state;
    renderInputs();
    setStatus("定价分析表已生成。", "ok");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    setBusy(false);
  }
}

document.addEventListener("DOMContentLoaded", async () => {
  renderIcons();
  $("saveBtn").addEventListener("click", saveConfig);
  $("draftBtn").addEventListener("click", generateDraft);
  $("pricingBtn").addEventListener("click", runPricing);
  ["sku", "variant", "lengthCm", "widthCm", "heightCm", "weightG"]
    .forEach((id) => $(id).addEventListener("input", updateLocalFromForm));
  document.querySelectorAll(".file-button").forEach((form) => {
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      uploadFiles(form);
    });
  });
  try {
    await loadState();
  } catch (error) {
    setStatus(error.message, "error");
  }
});
