import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = "/Users/cc/Documents/GitHub/Price";
const sourcePath = path.join(root, "config/quick_pricing_history.json");
const outputDir = path.join(root, "outputs/20260921_w1_sku_export");
const outputPath = path.join(outputDir, "9月W1新品单_SKU定价明细.xlsx");

const source = JSON.parse(await fs.readFile(sourcePath, "utf8"));
const record = source.records.find((item) => item.name === "9月W1新品单（≤4oz）");
if (!record) throw new Error("未找到 9月W1新品单（≤4oz）记录");
if (!Array.isArray(record.rows) || record.rows.length !== 23) {
  throw new Error(`SKU数量异常：${record.rows?.length ?? 0}`);
}

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("SKU明细");
sheet.showGridLines = false;
sheet.tabColor = "#1F4E78";

const font = "Arial";
const lastRow = 5 + record.rows.length;
const title = "9月W1新品单 SKU定价明细";
const context = `定价日期：2026-09-12    配送档位：${record.tierLabel}    SKU数量：${record.rows.length}`;

sheet.getRange("A2").values = [[title]];
sheet.getRange("A2:H2").format.font = { name: font, size: 15, bold: true, color: "#1F2937" };
sheet.getRange("A3").values = [[context]];
sheet.getRange("A3:H3").format.font = { name: font, size: 10, italic: true, color: "#5B6573" };
sheet.getRange("A4:H4").format.borders = {
  bottom: { style: "thin", color: "#9CA3AF" },
};

const headers = [[
  "SKU",
  "品名",
  "成本价（RMB）",
  "配送费（USD）",
  "15%利润价（USD）",
  "10%利润价（USD）",
  "5%利润价（USD）",
  "保本价（USD）",
]];
sheet.getRange("A5:H5").values = headers;

const rows = record.rows.map((row) => [
  row.sku,
  row.title,
  Number(row.costRmb),
  Number(row.shippingFee),
  Number(row.margin15),
  Number(row.margin10),
  Number(row.margin5),
  Number(row.breakEven),
]);
sheet.getRange(`A6:H${lastRow}`).values = rows;

const used = sheet.getRange(`A2:H${lastRow + 3}`);
used.format.font = { name: font, size: 10, color: "#1F2937" };
used.format.verticalAlignment = "center";

sheet.getRange("A2").format.font = { name: font, size: 15, bold: true, color: "#1F2937" };
sheet.getRange("A3").format.font = { name: font, size: 10, italic: true, color: "#5B6573" };
sheet.getRange("A5:H5").format = {
  fill: "#1F4E78",
  font: { name: font, size: 10, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: {
    insideVertical: { style: "thin", color: "#FFFFFF" },
    bottom: { style: "medium", color: "#153650" },
  },
};
sheet.getRange(`A6:B${lastRow}`).format.horizontalAlignment = "left";
sheet.getRange(`C6:H${lastRow}`).format.horizontalAlignment = "right";
sheet.getRange(`C6:C${lastRow}`).format.numberFormat = "0.000";
sheet.getRange(`D6:H${lastRow}`).format.numberFormat = '"$"0.00';
sheet.getRange(`A6:H${lastRow}`).format.borders = {
  insideHorizontal: { style: "thin", color: "#E5E7EB" },
  bottom: { style: "thin", color: "#CBD5E1" },
};

for (let row = 6; row <= lastRow; row += 2) {
  sheet.getRange(`A${row}:H${row}`).format.fill = "#F6F8FB";
}

sheet.getRange(`A${lastRow + 2}`).values = [["数据来源：自动定价工作台本地记录；导出日期：2026-09-21"]];
sheet.getRange(`A${lastRow + 2}:H${lastRow + 2}`).format.font = {
  name: font,
  size: 9,
  italic: true,
  color: "#6B7280",
};

sheet.getRange("A:A").format.columnWidth = 28;
sheet.getRange("B:B").format.columnWidth = 30;
sheet.getRange("C:C").format.columnWidth = 16;
sheet.getRange("D:D").format.columnWidth = 16;
sheet.getRange("E:H").format.columnWidth = 18;
sheet.getRange("2:2").format.rowHeight = 25;
sheet.getRange("3:3").format.rowHeight = 19;
sheet.getRange("5:5").format.rowHeight = 24;
sheet.getRange(`6:${lastRow}`).format.rowHeight = 21;
sheet.freezePanes.freezeRows(5);
sheet.freezePanes.freezeColumns(2);

const table = sheet.tables.add(`A5:H${lastRow}`, true, "W1SkuPricingTable");
table.style = "TableStyleMedium2";
table.showFilterButton = true;
table.showBandedColumns = false;

workbook.recalculate();

const inspection = await workbook.inspect({
  kind: "table",
  range: `SKU明细!A2:H${lastRow}`,
  include: "values,formulas",
  tableMaxRows: 30,
  tableMaxCols: 8,
  maxChars: 16000,
});
console.log(inspection.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

await fs.mkdir(outputDir, { recursive: true });
const preview = await workbook.render({
  sheetName: "SKU明细",
  range: `A1:H${lastRow + 2}`,
  scale: 1.25,
  format: "png",
});
await fs.writeFile(path.join(outputDir, "preview.png"), new Uint8Array(await preview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(JSON.stringify({ outputPath, previewPath: path.join(outputDir, "preview.png"), rows: record.rows.length }));
