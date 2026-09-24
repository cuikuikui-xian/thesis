import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = String.raw`C:\Users\32933\xwechat_files\wxid_bbyf1p55y50a22_da86\msg\file\2026-09\期刊目录修订表（2026年修订）-传递所.xlsx`;
const outputDir = String.raw`C:\Users\32933\Documents\ChatGPT\论文\outputs\journal_revision_device_20260914`;
const outputPath = `${outputDir}/期刊目录修订表（2026年修订）-已增加Device.xlsx`;

await fs.mkdir(outputDir, { recursive: true });
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
const sheet = workbook.worksheets.getItem("最具+知名目录");
// Extend the data block by inheriting the preceding row's formatting,
// borders, and validation rules.
sheet.getRange("A8:AC8").copyFrom(sheet.getRange("A7:AC7"), "all");

const revisionReason =
  "Device是Cell Press于2023年创办的综合性器件研究期刊，面向物理、化学、材料科学、生物、计算机与工程等领域的多学科应用技术研究，强调由基础发现走向真实应用的器件创新。期刊由Cell Press（Elsevier）出版，已被Web of Science核心合集ESCI收录，2025年影响因子为8.7、JCR Q1。其刊文方向覆盖能源转换与存储、电子器件热管理、传感器、柔性器件及先进制造等，与化工学院动力工程及工程热物理、能源化工和交叉器件研究高度契合。学院团队已在该刊发表超薄微型环路热管研究，体现了该刊对高水平工程器件成果的认可，建议新增为“最具”期刊。";

const row = [[
  5,
  "理学与工学门类",
  "Device",
  "最具",
  "增加",
  "化工",
  "动力工程及工程热物理",
  "Cell Press（Elsevier）",
  "2666-9986",
  "Web of Science核心合集（ESCI）",
  "Q1",
  "未填写",
  "未填写",
  8.0,
  8.7,
  "否",
  "未填写",
  "未填写",
  "Cell Press（Elsevier）",
  "未填写",
  "美国",
  "否（混合OA）",
  "否（OA可选）",
  "未填写",
  "否",
  "否",
  "未填写",
  revisionReason,
  "2023年创刊；国外英文期刊；Cell Press综合性器件研究期刊"
]];

sheet.getRange("A8:AC8").values = row;
sheet.getRange("A8:AC8").format.wrapText = true;
sheet.getRange("A8:AC8").format.verticalAlignment = "center";
sheet.getRange("A8:AC8").format.rowHeight = 245;

workbook.recalculate();

console.log((await workbook.inspect({
  kind: "table",
  range: "最具+知名目录!A1:AC8",
  include: "values,formulas",
  tableMaxRows: 10,
  tableMaxCols: 29,
  maxChars: 24000
})).ndjson);

console.log((await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  summary: "final formula error scan"
})).ndjson);

for (let i = 0; i < workbook.worksheets.items.length; i++) {
  const current = workbook.worksheets.getItemAt(i);
  const preview = await workbook.render({
    sheetName: current.name,
    autoCrop: "all",
    scale: 1,
    format: "png"
  });
  const safeName = current.name.replaceAll(/[\\/:*?"<>|]/g, "_");
  await fs.writeFile(`${outputDir}/after_${i}_${safeName}.png`, new Uint8Array(await preview.arrayBuffer()));
}

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(`OUTPUT=${outputPath}`);
