import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = "/Users/cc/Documents/GitHub/Price";
const inputPath = path.join(__dirname, "cleaned_products.json");
const outputDir = path.join(root, "outputs", "product_history_cleanup_20260809");
const outputPath = path.join(outputDir, "历史产品价格_待录入整理.xlsx");

const data = JSON.parse(await fs.readFile(inputPath, "utf8"));

function colName(index) {
  let name = "";
  let n = index + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function rangeFor(rowCount, colCount) {
  return `A1:${colName(colCount - 1)}${rowCount}`;
}

function asRows(headers, records, fields) {
  return [
    headers,
    ...records.map((record) => fields.map((field) => record[field] ?? null)),
  ];
}

function setupSheet(sheet, rowCount, colCount, headerFill = "#243B53") {
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(1);
  const used = sheet.getRange(rangeFor(rowCount, colCount));
  used.format.font = { name: "Arial", size: 10, color: "#243B53" };
  used.format.wrapText = true;
  used.format.borders = { preset: "inside", style: "thin", color: "#E6E8EB" };
  const header = sheet.getRange(`A1:${colName(colCount - 1)}1`);
  header.format.fill = { color: headerFill };
  header.format.font = { bold: true, color: "#111827", name: "Arial", size: 10 };
  header.format.rowHeight = 28;
  header.format.horizontalAlignment = "center";
  header.format.verticalAlignment = "center";
  used.format.autofitColumns();
  used.format.autofitRows();
}

function setWidths(sheet, widths) {
  widths.forEach((width, index) => {
    sheet.getRange(`${colName(index)}:${colName(index)}`).format.columnWidth = width;
  });
}

const workbook = Workbook.create();

const summarySheet = workbook.worksheets.add("产品汇总");
const summaryHeaders = [
  "建议历史ID",
  "来源",
  "日期标签",
  "产品编码/型号",
  "产品名称",
  "SKU数",
  "成本范围(CNY)",
  "采购数量合计",
  "采购链接数",
  "供应商",
  "采购员",
  "配送档位",
  "配送费(USD)",
  "录入状态",
  "可能匹配已有产品",
  "备注",
  "采购链接",
];
const summaryFields = [
  "product_record_id",
  "source_name",
  "date_label",
  "product_code",
  "product_name",
  "sku_count",
  "cost_range_rmb",
  "linked_qty_total",
  "purchase_link_count",
  "supplier",
  "buyer",
  "shipping_tier",
  "shipping_fee_usd",
  "status",
  "existing_match",
  "notes",
  "purchase_links",
];
const summaryRows = asRows(summaryHeaders, data.products, summaryFields);
summarySheet.getRange(rangeFor(summaryRows.length, summaryHeaders.length)).values = summaryRows;
setupSheet(summarySheet, summaryRows.length, summaryHeaders.length, "#1F4E5F");
setWidths(summarySheet, [16, 18, 14, 22, 18, 8, 15, 12, 10, 28, 12, 12, 12, 12, 18, 34, 44]);
summarySheet.getRange(`F2:I${summaryRows.length}`).format.horizontalAlignment = "right";
summarySheet.getRange(`M2:M${summaryRows.length}`).setNumberFormat("$0.00");

const detailSheet = workbook.worksheets.add("SKU明细");
const detailHeaders = [
  "建议历史ID",
  "来源",
  "产品编码/型号",
  "产品名称",
  "SKU",
  "品名",
  "款式/规格",
  "关联数量",
  "采购员",
  "采购成本(CNY)",
  "含税单价",
  "供应商",
  "采购链接",
  "配送档位",
  "配送费(USD)",
  "15%利润价",
  "10%利润价",
  "5%利润价",
  "保本价",
  "录入状态",
  "源表行号",
];
const detailFields = [
  "product_record_id",
  "source_name",
  "product_code",
  "product_name",
  "sku",
  "title",
  "variant",
  "linked_qty",
  "buyer",
  "cost_rmb",
  "tax_unit_price",
  "supplier",
  "purchase_link",
  "shipping_tier",
  "shipping_fee_usd",
  "margin15",
  "margin10",
  "margin5",
  "break_even",
  "status",
  "source_row",
];
const detailRows = asRows(detailHeaders, data.sku_rows, detailFields);
detailSheet.getRange(rangeFor(detailRows.length, detailHeaders.length)).values = detailRows;
setupSheet(detailSheet, detailRows.length, detailHeaders.length, "#2F5597");
setWidths(detailSheet, [16, 18, 22, 18, 26, 30, 20, 10, 10, 13, 11, 28, 42, 12, 12, 12, 12, 12, 10, 12, 10]);
detailSheet.getRange(`H2:K${detailRows.length}`).format.horizontalAlignment = "right";
detailSheet.getRange(`J2:K${detailRows.length}`).setNumberFormat("0.0000");
detailSheet.getRange(`O2:S${detailRows.length}`).setNumberFormat("$0.00");

const duplicateSheet = workbook.worksheets.add("重复合并提示");
const duplicateHeaders = ["产品名称", "1688 Offer ID/链接", "出现来源", "SKU数", "涉及SKU", "成本(CNY)", "处理建议"];
const duplicateFields = ["product_name", "offer_id", "source_names", "sku_count", "skus", "costs_rmb", "suggestion"];
const duplicateRecords = data.duplicates.map((record) => ({
  ...record,
  offer_id: /^\d+$/.test(String(record.offer_id || "")) ? `offer/${record.offer_id}` : record.offer_id,
}));
const duplicateRows = asRows(duplicateHeaders, duplicateRecords, duplicateFields);
duplicateSheet.getRange(rangeFor(Math.max(duplicateRows.length, 2), duplicateHeaders.length)).values =
  duplicateRows.length > 1 ? duplicateRows : [duplicateHeaders, ["没有发现跨表重复", "", "", "", "", "", ""]];
setupSheet(duplicateSheet, Math.max(duplicateRows.length, 2), duplicateHeaders.length, "#7A3E1D");
setWidths(duplicateSheet, [20, 20, 24, 10, 42, 18, 58]);

const notesSheet = workbook.worksheets.add("字段说明");
const notesRows = [
  ["项目", "说明"],
  ["整理范围", `共 ${data.summary.sku_rows} 条 SKU，整理成 ${data.summary.product_records} 个产品记录。`],
  ["录入原则", "本文件只整理基础产品资料，不计算配送费，不生成四档利润价，不写入历史价格。"],
  ["产品聚合", "新版采购上传表按“型号”聚合；旧领星同步表按“品名基础名 + 1688链接”聚合。"],
  ["待补字段", "配送档位、配送费(USD)、15%利润价、10%利润价、5%利润价、保本价。"],
  ["重复处理", "重复合并提示页列出两个源表中同产品/同采购链接的记录，后续录入前需要确认是否保留两套SKU。"],
  ["可能匹配已有产品", "用当前历史产品名做简单包含匹配，仅作为提醒，最终是否并入已有产品需要人工确认。"],
  ["源文件1", data.summary.source_files[0]],
  ["源文件2", data.summary.source_files[1]],
];
notesSheet.getRange(rangeFor(notesRows.length, 2)).values = notesRows;
setupSheet(notesSheet, notesRows.length, 2, "#4B5563");
setWidths(notesSheet, [20, 90]);

await fs.mkdir(outputDir, { recursive: true });

const inspectSummary = await workbook.inspect({
  kind: "table",
  range: "产品汇总!A1:Q12",
  include: "values",
  tableMaxRows: 12,
  tableMaxCols: 17,
  maxChars: 6000,
});
console.log(inspectSummary.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

for (const sheetName of ["产品汇总", "SKU明细", "重复合并提示", "字段说明"]) {
  const preview = await workbook.render({
    sheetName,
    range: sheetName === "SKU明细" ? "A1:U28" : undefined,
    autoCrop: "all",
    scale: 1,
    format: "png",
  });
  const bytes = new Uint8Array(await preview.arrayBuffer());
  await fs.writeFile(path.join(outputDir, `${sheetName}.png`), bytes);
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(JSON.stringify({ outputPath, summary: data.summary }, null, 2));
