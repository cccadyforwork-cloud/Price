const finalHeaders = ["SKU", "款式标题", "长(in)", "宽(in)", "高(in)", "重(lb)", "定价"];
const reportHeaders = ["项目", "数值", "说明"];

let finalExcelUrl = "";
let reportExcelUrl = "";
let reportTextUrl = "";

function $(id) {
  return document.getElementById(id);
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
  return [...document.querySelectorAll("#finalPricingRows tr")].map((row) => {
    const cells = row.querySelectorAll("td");
    return [
      cells[0].textContent.trim(),
      cells[1].textContent.trim(),
      cells[2].textContent.trim(),
      cells[3].textContent.trim(),
      cells[4].textContent.trim(),
      cells[5].textContent.trim(),
      cells[6].querySelector("input").value.trim()
    ];
  });
}

function reportRows() {
  return [
    ["配送费", "0.88", "非服饰类，售价 > $3，重量落在 4oz 及以下"],
    ["保本定价", "3.94", "按当前采购成本、配送费和平台费用测算"],
    ["建议底价", "3.99", "上架建议不要低于该价格"],
    ["目标定价", "5.61", "按目标利润率测试"],
    ["期望利润率", "25.0%", "按当前采购成本和配送费模型"],
    ["竞品最低价", "5.69", "30pcs 竞品锚点"],
    ["结论", "价格优势不足", "与最低竞品价差较小"]
  ];
}

function reportText() {
  return [
    "通用产品定价台 - 防磨贴定价分析报告",
    "",
    "配送费分析过程",
    "按表内参数算，防磨贴配送费预计为 $0.88，保本定价约为 $3.94，建议上架不要低于 $3.99。",
    "尺寸：8 x 12 x 1.5 cm，重量 55g = 约 1.94oz，落在 4oz 及以下。",
    "预计售价会高于 $3，所以按表里非服饰类、售价 > $3 的配送费 $0.88 测算。",
    "",
    "竞品与利润分析过程",
    "15PCS：$5.99，$0.40/片，3.6分 / 6评",
    "30pcs：$5.69，$0.19/片，4.2分 / 41评",
    "32PCS：$9.99，$0.31/片，4.1分 / 135评",
    "",
    "如果 30 片组合按 $5.61 上架，单片价约 $0.19/片，基本贴近 30pcs 竞品。",
    "按当前采购成本和配送费模型，期望利润率约 25.0%。",
    "但和最低竞品 $5.69 的差距很小，价格优势不明显。",
    "",
    "最终价格确认",
    ...finalRows().map((row) => row.join(" / "))
  ].join("\n");
}

function setLink(link, href) {
  link.href = href;
  link.classList.remove("is-hidden");
}

function setSiteStatus(message, tone = "ok") {
  const status = $("siteStatus");
  status.textContent = message;
  status.className = `site-status ${tone}`.trim();
}

function generateFinalExcel() {
  const xml = buildExcelXml("最终价格确认", finalHeaders, finalRows());
  finalExcelUrl = makeDownload(xml, "application/vnd.ms-excel;charset=utf-8", finalExcelUrl);
  setLink($("downloadExcelLink"), finalExcelUrl);
  $("exportStatus").textContent = `已生成 ${finalRows().length} 条 SKU 的 Excel 表格。`;
  setSiteStatus("最终价格确认表已生成，请点击底部的“下载 Excel”。");
  $("downloadExcelLink").scrollIntoView({ behavior: "smooth", block: "center" });
}

function generateReportExcel() {
  const xml = buildExcelXml("定价分析报告", reportHeaders, reportRows());
  reportExcelUrl = makeDownload(xml, "application/vnd.ms-excel;charset=utf-8", reportExcelUrl);
  setLink($("downloadReportExcelLink"), reportExcelUrl);
  setSiteStatus("定价分析 Excel 已生成，请点击“下载分析 Excel”。");
}

function generateReportText() {
  reportTextUrl = makeDownload(reportText(), "text/plain;charset=utf-8", reportTextUrl);
  setLink($("downloadReportTextLink"), reportTextUrl);
  setSiteStatus("定价分析报告文本已生成，请点击“下载报告文本”。");
}

function syncSkuToFinal(index, value) {
  const target = document.querySelector(`[data-final-sku="${index}"]`);
  if (target) {
    target.textContent = value.trim() || "-";
  }
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
  $("generateReportBtn").textContent = "报告已生成";
  $("generateReportBtn").classList.add("is-confirmed");
  setSiteStatus("定价分析报告已生成，可以导出 Excel 或报告文本。");
}

function initFileInputs() {
  $("purchaseOrderFile").addEventListener("change", (event) => {
    const file = event.target.files[0];
    $("purchaseOrderFileName").textContent = file ? `已选择：${file.name}` : "已放入：防磨贴采购单.pdf";
    if (file) {
      setSiteStatus(`已选择采购单：${file.name}`);
    }
  });

  $("competitorFiles").addEventListener("change", (event) => {
    const count = event.target.files.length;
    $("competitorFileSummary").textContent = count
      ? `已选择：${count} 个竞品文件，等待生成分析。`
      : "示例已读取：4 个竞品文件，最低对比价 $5.69。";
    if (count) {
      setSiteStatus(`已选择 ${count} 个竞品文件。`);
    }
  });
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
  });

  $("generateReportBtn").addEventListener("click", showGeneratedReportState);
  $("reportExcelBtn").addEventListener("click", generateReportExcel);
  $("reportTextBtn").addEventListener("click", generateReportText);
  $("generateExcelBtn").addEventListener("click", generateFinalExcel);
}

function initEditableFields() {
  document.querySelectorAll("[data-final-sku-source]").forEach((input) => {
    input.addEventListener("input", () => {
      syncSkuToFinal(input.dataset.finalSkuSource, input.value);
      resetDownloadState();
    });
  });

  document.querySelectorAll(".price-input").forEach((input) => {
    input.addEventListener("input", resetDownloadState);
  });
}

document.addEventListener("DOMContentLoaded", () => {
  initFileInputs();
  initButtons();
  initEditableFields();
});
