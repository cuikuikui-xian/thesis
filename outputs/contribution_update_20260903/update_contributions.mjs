import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = String.raw`C:\Users\32933\Documents\ChatGPT\论文\outputs\loop_heat_pipe_papers_20260902\环路热管论文信息整理_原版摘要_20260902.xlsx`;
const workDir = String.raw`C:\Users\32933\Documents\ChatGPT\论文\outputs\contribution_update_20260903`;

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);
const firstSheet = workbook.worksheets.getItemAt(0);

const sheets = await workbook.inspect({
  kind: "sheet",
  include: "id,name",
  maxChars: 3000,
});
console.log(sheets.ndjson);

const overview = await workbook.inspect({
  kind: "workbook,sheet,table",
  maxChars: 9000,
  tableMaxRows: 12,
  tableMaxCols: 18,
  tableMaxCellChars: 160,
});
console.log(overview.ndjson);

const contributionRegion = await workbook.inspect({
  kind: "region",
  sheetId: "论文信息",
  range: "S1:X8",
  maxChars: 14000,
});
console.log(contributionRegion.ndjson);

const contributionStyles = await workbook.inspect({
  kind: "computedStyle",
  sheetId: "论文信息",
  range: "S1:X8",
  maxChars: 6000,
});
console.log(contributionStyles.ndjson);
console.log("CONTRIBUTIONS=" + JSON.stringify(firstSheet.getRange("S3:S8").values));
console.log("ABSTRACTS=" + JSON.stringify(firstSheet.getRange("R3:R8").values));
const preview = await workbook.render({
  sheetName: firstSheet.name,
  autoCrop: "all",
  scale: 1,
  format: "png",
});
await fs.writeFile(`${workDir}/before.png`, new Uint8Array(await preview.arrayBuffer()));

const contributions = [
  "本人负责研究构思、实验设计、数据采集及论文撰写。首创主—辅毛细芯协同结构，研制0.7 mm、3.95 g样机，使导热系数提升378.6%、热阻下降79.2%，性能超过商业石墨烯均热材料15倍。",
  "本人负责实验设计、样机制备、数据采集及论文撰写。研制0.7 mm超薄柔性环路热管，揭示重力与弯折作用机制；优化倾角使导热系数提升109.26%，实现5 W/cm²高热流密度散热。",
  "本人参与实验设计、装置搭建、数据分析及论文撰写。通过整体烧结和微针肋底板创新，将最大稳定热负荷由300 W提高至500 W（增66.7%），相较槽式底板散热性能提升42%。",
  "本人参与毛细芯制备、性能测试、模型建立及论文撰写。制备并评价42组双孔镍芯，构建改进Kozeny–Carman模型，使92.86%的预测结果误差控制在20%以内，解决含造孔剂毛细芯性能预测难题。",
  "本人负责文献调研、综述框架设计、资料整理及论文撰写。系统整合228篇文献，构建结构设计—传热极限—数值模型—运行稳定性—工程应用的全链条框架，明确新材料与AI辅助优化方向。",
];

firstSheet.getRange("S4:S8").values = contributions.map((value) => [value]);

const updated = await workbook.inspect({
  kind: "table",
  range: "论文信息!S3:S8",
  include: "values,formulas",
  tableMaxRows: 8,
  tableMaxCols: 2,
  tableMaxCellChars: 220,
  maxChars: 5000,
});
console.log("UPDATED=" + updated.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
  maxChars: 3000,
});
console.log("ERRORS=" + errors.ndjson);

for (const sheetName of ["论文信息", "作者信息", "填报说明"]) {
  const rendered = await workbook.render({
    sheetName,
    autoCrop: "all",
    scale: 1,
    format: "png",
  });
  const safeName = sheetName.replaceAll(/[\\/:*?"<>|]/g, "_");
  await fs.writeFile(`${workDir}/after_${safeName}.png`, new Uint8Array(await rendered.arrayBuffer()));
}

const outputPath = `${workDir}/环路热管论文信息整理_本人贡献量化版_20260903.xlsx`;
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log("OUTPUT=" + outputPath);
