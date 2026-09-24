import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = String.raw`C:\Users\32933\Documents\ChatGPT\论文\outputs\contribution_update_20260903\环路热管论文信息整理_本人贡献量化版_20260903.xlsx`;
const workDir = String.raw`C:\Users\32933\Documents\ChatGPT\论文\outputs\hp1_addition_20260903`;
const outputPath = `${workDir}/环路热管论文信息整理_含化工学报论文_20260903.xlsx`;
const sourceUrl = "https://hgxb.cip.com.cn/CN/10.11949/0438-1157.20250399";

await fs.mkdir(workDir, { recursive: true });
const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const papers = workbook.worksheets.getItem("论文信息");
const authors = workbook.worksheets.getItem("作者信息");
const notes = workbook.worksheets.getItem("填报说明");

papers.getRange("A1").values = [["六篇环路热管论文信息（按申报字段整理）"]];
papers.getRange("A2").values = [["引用次数均为指定数据库在核验日期的快照；第5篇采用用户提供的WoS截图数值23，第6篇新发表论文的他引次数暂按0录入并标记待复核。黄色单元格表示提交前仍需复核。"]];

const abstract = "With the rapid advancement of 5G technology, electronic devices are evolving toward higher integration and miniaturization, leading to a significant increase in heat flux. However, traditional loop heat pipes (LHPs) face limitations in heat flux density due to inherent heat leakage issues, failing to meet emerging thermal demands. In previous studies, an innovative injector-integrated loop heat pipe (LHPI) was developed and demonstrated superior thermal management performance. However, early experiments utilized water as the working fluid, which suffers from low freezing points and ice formation in low-temperature environments, restricting its broader application. To address this, this study introduces a novel low-freezing-point refrigerant, HP-1, and systematically investigates the effects of heat loads (50—300 W) and heat sink temperatures (5—15℃) on LHPI performance. Experimental results reveal that the injector's operational modes—classified as low-efficiency, normal injection, restricted expansion, and superheated modes—significantly influence LHPI's heat transfer characteristics. At elevated heat sink temperatures, the transition to restricted expansion and superheated modes occurs earlier. Compared with water-based working fluids, HP-1 increases the heat flux of LHPI to 41.7 W/cm² at a base plate temperature of 85℃, and its low temperature adaptability is significantly enhanced, providing a new idea for passive heat dissipation of high-power electronic devices.";
const contribution = "本人负责实验设计、装置搭建、数据采集及论文撰写。首次采用环保工质HP-1并揭示四类引射模式，实现50.0 W/cm²最高热流密度；在85℃底板温度下仍达41.7 W/cm²，显著提升低温适应性。";

const paperTable = papers.tables.items[0];
paperTable.rows.add(null, [[
  6,
  "环保型制冷剂HP-1为工质的新型环路热管传热特性研究",
  "化工学报（CIESC Journal）",
  new Date("2025-11-25T00:00:00+08:00"),
  0,
  "待在WoS/CNKI核验",
  new Date("2026-09-03T00:00:00+08:00"),
  "是",
  "中文（含英文摘要）",
  "76",
  "11",
  "5645",
  "5654",
  "连续页码",
  "10.11949/0438-1157.20250399",
  sourceUrl,
  "loop heat pipe, ejector, HP-1, heat transfer",
  abstract,
  contribution,
  "第1作者",
  "官网PDF待下载",
  "题录、作者、英文原版摘要、关键词及出版日期均由《化工学报》官网核验；他引次数暂按0录入，正式提交前请在WoS或CNKI复核。",
  sourceUrl,
  null,
]]);

papers.getRange("X4").formulas = [["=COUNTIF('作者信息'!$A$4:$A$45,A4)"]];
papers.getRange("X4:X9").fillDown();
papers.getRange("D9").format.numberFormat = "yyyy-mm-dd";
papers.getRange("G9").format.numberFormat = "yyyy-mm-dd";
papers.getRange("E9").format.numberFormat = "0";
papers.getRange("A9:X9").format.rowHeight = 220;
papers.getRange("A9:X9").format.wrapText = true;
papers.getRange("A9:X9").format.verticalAlignment = "top";
papers.getRange("A9").format.horizontalAlignment = "center";
papers.getRange("E9:G9").format.fill = "#FFF2CC";
papers.getRange("U9:V9").format.fill = "#FFF2CC";

authors.getRange("A2").values = [["“本人”按Qingjie Cui（崔庆杰）标记；第6篇作者与单位已由《化工学报》官网核验，黄色内容为仍需最终复核的信息。"]];
const authorTable = authors.tables.items[0];
const unit12 = "西安交通大学化学工程与技术学院；西安交通大学氟氮化工新材料全国重点实验室";
const unit3 = "西安交通大学热流科学与工程教育部重点实验室";
authorTable.rows.add(null, [
  [6, "环保型制冷剂HP-1为工质的新型环路热管传热特性研究", 1, "Qingjie Cui（崔庆杰）", "第一作者", "是", unit12, "cuiqingjie@stu.xjtu.edu.cn", "作者简介及署名顺序由期刊官网核验", sourceUrl],
  [6, "环保型制冷剂HP-1为工质的新型环路热管传热特性研究", 2, "Li Shen（沈立）", "共同作者", "否", unit12, null, "署名顺序由期刊官网核验", sourceUrl],
  [6, "环保型制冷剂HP-1为工质的新型环路热管传热特性研究", 3, "Yicheng Ni（倪一程）", "共同作者", "否", unit12, null, "署名顺序由期刊官网核验", sourceUrl],
  [6, "环保型制冷剂HP-1为工质的新型环路热管传热特性研究", 4, "Yao Zhou（周尧）", "共同作者", "否", unit3, null, "署名顺序由期刊官网核验", sourceUrl],
  [6, "环保型制冷剂HP-1为工质的新型环路热管传热特性研究", 5, "Xiaoping Yang（杨小平）", "通讯作者", "否", unit12, "yxping@xjtu.edu.cn", "通讯作者由期刊官网核验", sourceUrl],
  [6, "环保型制冷剂HP-1为工质的新型环路热管传热特性研究", 6, "Yonghai Zhang（张永海）", "共同作者", "否", unit12, null, "署名顺序由期刊官网核验", sourceUrl],
  [6, "环保型制冷剂HP-1为工质的新型环路热管传热特性研究", 7, "Jinjia Wei（魏进家）", "共同作者", "否", unit12, null, "署名顺序由期刊官网核验", sourceUrl],
]);
authors.getRange("A39:J45").format.rowHeight = 42;
authors.getRange("A39:J45").format.wrapText = true;
authors.getRange("A39:J45").format.verticalAlignment = "top";
authors.getRange("A39:A45").format.horizontalAlignment = "center";
authors.getRange("C39:C45").format.horizontalAlignment = "center";
authors.getRange("F39:F45").format.horizontalAlignment = "center";
authors.getRange("H39:H45").format.font = { color: "#0563C1", size: 10 };
authors.getRange("J39:J45").format.font = { color: "#0563C1", underline: true, size: 8 };
authors.getRange("F39:F45").conditionalFormats.deleteAll();
authors.getRange("F39:F45").conditionalFormats.addCustom('=F39="是"', { fill: "#E2F0D9", font: { bold: true, color: "#006100" } });

notes.getRange("B4").values = [["第5篇为用户提供的Web of Science截图：23次，截图日期按2026-08-27记录。第6篇为2025年11月新发表中文论文，他引次数暂按0录入，正式提交前请在WoS或CNKI核验；其余4篇来源和日期见“论文信息”表。"]];
notes.getRange("B6").values = [["前5篇使用文章号；第6篇为连续页码5645—5654。申报系统中第6篇可直接分别填写起始页码和截止页码。"]];
notes.getRange("B7").values = [["前5篇摘要逐字提取自正式论文PDF；第6篇英文摘要逐字取自《化工学报》官网，仅统一空格与符号。关键词均控制在5项以内。"]];
notes.getRange("B8").values = [["6篇均按“本人职责＋关键机理/创新＋量化结果”的结构填写，并控制在100字符以内；涉及具体分工的表述请本人按实际情况确认后提交。"]];
notes.getRange("B9").values = [["第6篇作者单位、署名顺序、第一作者和通讯作者由《化工学报》官网核验；其他论文仍以正式PDF首页为最终依据。"]];
notes.getRange("B10").values = [["前5份正式论文PDF已与工作簿放在同一目标文件夹；第6篇《化工学报》官网PDF因服务器断流尚未完整下载，请后续从论文官网补入。上传申报系统前请检查单个文件不超过15 MB。"]];

for (const sheetName of ["论文信息", "作者信息", "填报说明"]) {
  const rendered = await workbook.render({ sheetName, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${workDir}/after_${sheetName}.png`, new Uint8Array(await rendered.arrayBuffer()));
}

const paperCheck = await workbook.inspect({
  kind: "table",
  range: "论文信息!A3:X9",
  include: "values,formulas",
  tableMaxRows: 9,
  tableMaxCols: 24,
  tableMaxCellChars: 180,
  maxChars: 12000,
});
console.log("PAPER_CHECK=" + paperCheck.ndjson);
const authorCheck = await workbook.inspect({
  kind: "table",
  range: "作者信息!A38:J45",
  include: "values,formulas",
  tableMaxRows: 10,
  tableMaxCols: 10,
  tableMaxCellChars: 140,
  maxChars: 8000,
});
console.log("AUTHOR_CHECK=" + authorCheck.ndjson);
const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
  maxChars: 3000,
});
console.log("ERRORS=" + errors.ndjson);

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log("OUTPUT=" + outputPath);
