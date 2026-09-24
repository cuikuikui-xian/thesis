import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = path.resolve("outputs/loop_heat_pipe_papers_20260902");
const outputPath = path.join(outputDir, "环路热管论文信息整理_原版摘要_20260902.xlsx");
await fs.mkdir(outputDir, { recursive: true });

const workbook = Workbook.create();
const papersSheet = workbook.worksheets.add("论文信息");
const authorsSheet = workbook.worksheets.add("作者信息");
const notesSheet = workbook.worksheets.add("填报说明");

const navy = "#17365D";
const blue = "#1F4E78";
const paleBlue = "#D9EAF7";
const paleGold = "#FFF2CC";
const paleGreen = "#E2F0D9";
const paleRed = "#FCE4D6";
const border = "#B7C9D6";
const text = "#1F2937";
const muted = "#5B6573";

const paperHeaders = [
  "序号", "论文名称", "发表刊物名称", "发表日期", "他引次数", "引用来源", "引用核验日期",
  "是否中文期刊", "论文语言", "卷号", "期号", "起始页码/文章号", "截止页码", "页码说明",
  "DOI", "全文/出版商链接", "论文关键词（≤5）", "论文摘要（原版英文）", "本人贡献（≤100字）",
  "本人署名位置", "附件状态", "数据核验备注", "主要来源", "作者人数"
];

const papers = [
  [
    1,
    "A thin and lightweight miniature loop heat pipe for cooling mobile electronic devices",
    "Device",
    new Date(2025, 7, 15),
    6,
    "ScienceDirect（公开页面 Cited by）",
    new Date(2026, 8, 2),
    "否",
    "外文（英语）",
    "3",
    "8",
    "100783",
    "",
    "文章号，无连续页码",
    "10.1016/j.device.2025.100783",
    "https://www.sciencedirect.com/science/article/pii/S2666998625000961",
    "miniature loop heat pipe, mobile electronics, nickel fiber wick, thermal management, passive phase-change cooling",
    "The escalating thermal challenges in compact electronics necessitate micrometer-scale thermal management innovations. We designed a thin (0.7 mm) and lightweight (3.95 g) miniature loop heat pipe (mLHP). We developed an integrated wick structure with a stepped-serrated main wick and a braided copper wire as the supplementary wick. The copper-based evaporator and polymer-based vapor-liquid lines/condenser are integrated via a solder-free bonding technique to evaporate heat transfer with a maximum equivalent thermal conductivity of 24,449 W/(m·K), exceeding commercial graphene heat spreaders more than 15-fold. Accelerated aging tests at 90°C demonstrate stable performance over 30 days, confirming structural robustness and process stability for scalable mass production.",
    "本人负责研究构思、实验设计、数据采集及论文撰写。开发主—辅毛细芯协同回液结构，研制0.7 mm、3.95 g微型环路热管，实现24,449 W/(m·K)等效导热系数并通过90 ℃、30 d老化验证。",
    "第1作者",
    "已附PDF（同文件夹）",
    "题录已由出版社页面核验；摘要逐字取自正式PDF的SUMMARY，仅规范换行和单位字符；他引次数正式填报前建议统一在WoS复核。",
    "https://doi.org/10.1016/j.device.2025.100783",
    null
  ],
  [
    2,
    "Bending the heat: Innovative ultra-thin flexible loop heat pipes for enhanced mobile device cooling",
    "Energy Conversion and Management",
    new Date(2025, 1, 1),
    11,
    "ResearchGate（公开页面）",
    new Date(2026, 8, 2),
    "否",
    "外文（英语）",
    "325",
    "",
    "119332",
    "",
    "文章号，无连续页码",
    "10.1016/j.enconman.2024.119332",
    "https://www.sciencedirect.com/science/article/pii/S0196890424012731",
    "ultra-thin flexible loop heat pipe, mobile device cooling, metal powder wick, tilt angle, bending angle",
    "The rapid advancement of 5G technology has significantly accelerated the progression of mobile devices, promoting the evolution of electronic products such as smartphones, tablet computers, and virtual reality (VR) and augmented reality (AR) eyewear towards an increasingly foldable design. However, due to the inherent constraints of spatial and structural characteristics, conventional thermal management solutions are no longer adequate to meet the performance requirements of these foldable devices. Therefore, it is imperative to develop efficient thermal management solutions that are compatible with the cross-hinge structures within chip design. This study proposes and fabricates a novel ultra-thin flexible loop heat pipe (UFLHP) with a thickness of merely 0.7 mm to address the heat transfer challenges posed by cross-hinge designs. By utilizing powder sintering and wire cutting techniques, an innovative approach has been developed for fabricating a metallic powder wick with a thickness of 0.4 mm. During the experiments, ethanol was employed as the working fluid to systematically investigate the effects of thermal loading, tilt angle, and bending angle on the steady-state heat transfer performance of the UFLHP. The experimental results indicate that the maximum heat flux density of the UFLHP reaches 5 W/cm². Under this heat flux density, the evaporator temperature of the UFLHP attains 72.15 °C, while the thermal resistance is measured at 2.48 K/W, resulting in an effective thermal conductivity of 10,273.27 W/(m·K). The tilt angle has a beneficial effect on the UFLHP’s performance under gravitational influence, while the bending angle adversely affects its performance due to increased flow resistance. This research provides a feasible solution for the heat dissipation challenges in foldable electronic devices.",
    "本人负责实验设计、样机制备、数据采集及论文撰写。研制0.7 mm超薄柔性环路热管，揭示倾角促回液、弯折增流阻机制，实现5 W/cm²热流密度和10,273.27 W/(m·K)等效导热系数。",
    "第1作者",
    "已附PDF（同文件夹）",
    "出版题录已核验；摘要逐字取自正式PDF，仅规范换行和单位字符；公开引用数与WoS可能不同，正式填报前需在WoS统一复核。",
    "https://doi.org/10.1016/j.enconman.2024.119332",
    null
  ],
  [
    3,
    "Effect of tilt angle and base plate design on the performance of a loop heat pipe driven by vapor–liquid injector",
    "Applied Thermal Engineering",
    new Date(2024, 9, 15),
    2,
    "ScienceDirect（公开页面 Cited by）",
    new Date(2026, 8, 2),
    "否",
    "外文（英语）",
    "255",
    "",
    "123921",
    "",
    "文章号，无连续页码",
    "10.1016/j.applthermaleng.2024.123921",
    "https://www.sciencedirect.com/science/article/abs/pii/S1359431124015898",
    "Loop heat pipe, Vapor-driven jet injector, Tilt angle, Integrated sintering, Base plate design",
    "With the rapid development of microelectronics technology and the increasing heat flux of electronic devices, loop heat pipes have garnered significant attention for their highly efficient passive heat transfer performance. However, the heat transfer performance of conventional loop heat pipes is impaired by the heat leakage, and the maximum heat flux is unable to meet the heat dissipation requirement of high-heat-flux devices. To solve this problem, a loop heat pipe with a vapor-driven injector (LHPI) was previously proposed. However, only the heat transfer performance of the horizontally-placed LHPI had been investigated in the preliminary stage, and the effect of gravity on the heat transfer performance was not further discussed. Now in this paper, a 360° rotatable experimental bench is designed to investigate the influence of the tilt angle on the performance of LHPI. In addition, a capillary wick is integrally sintered with the trough baseplate and the microcolumn baseplate of the LHPI, respectively, to reduce the thermal resistance between them. The results show that the heat transfer performance of LHPI is significantly affected by its tilt angle. In the horizontal condition, LHPI exhibits the lowest baseplate temperature and thermal resistance, and the best heat transfer performance, which verifies its applications in engineering. Integral sintering increases the maximum thermal load from 300 W to 350 W under stable operation conditions of LHPI. The maximum thermal load of the integrated micro pin-finned baseplate reaches 500 W, which improves the heat dissipation effect by 42 % in heat dissipation compared to the trough baseplate under the same conditions., providing an improvement angle for the traditional LHP.",
    "本人参与实验设计、装置搭建、数据分析及论文撰写。揭示倾角与底板结构对引射式环路热管性能的影响，通过整体烧结和微针肋底板将稳定热负荷由300 W提升至500 W，散热能力提高42%。",
    "第2作者",
    "已附PDF（同文件夹）",
    "题录、关键词及公开引用数已由出版社页面核验；摘要逐字取自正式PDF，仅规范换行和单位字符。",
    "https://doi.org/10.1016/j.applthermaleng.2024.123921",
    null
  ],
  [
    4,
    "Experimental and theoretical study on performance of bi-porous wick for passive phase-change devices",
    "Physics of Fluids",
    new Date(2024, 11, 3),
    9,
    "OUCI（公开引文聚合）",
    new Date(2026, 8, 2),
    "否",
    "外文（英语）",
    "36",
    "12",
    "123308",
    "",
    "文章号，无连续页码",
    "10.1063/5.0240095",
    "https://pubs.aip.org/pof/article/36/12/123308/3323852/Experimental-and-theoretical-study-on-performance",
    "bi-porous wick, capillary performance, permeability, pore-forming agent, passive phase-change device",
    "The driving force of the fluid in passive phase-change devices is provided by the porous wick, and thus, the porous wick is one of the key factors determining the performance of these devices. In this work, the properties of sintered nickel powder porous wicks were systematically studied to give a comprehensive evaluation. Three particle sizes (average diameter 2–20 μm), pore-forming agent concentration (0%–30%wt), and two preparation methods [loose sintering (LS) and cold press sintering] were utilized to prepare 42 samples. The results showed that with the increase in particle size and pore-forming agent concentration and the decrease in preparation pressure, the capillary performance was enhanced due to the improvement in the connectivity of internal pores. However, the strength of the wick decreased. Additionally, an improved model for the capillary wick performance and permeability was proposed by combining the traditional Kozeny–Carman permeability formula with an enhanced permeability equation based on the research outcomes of Byon and Kin. The improved model was specifically optimized for scenarios involving the pore-forming agent. The predicted results agreed well with the experimental data, thoroughly validating the effectiveness of the developed model.",
    "本人参与毛细芯制备、性能测试、模型建立及论文撰写。阐明粒径、造孔剂和成形压力对孔隙连通性、毛细性能及强度的耦合影响，构建适用于双孔毛细芯的改进Kozeny–Carman模型。",
    "第2作者",
    "已附PDF（同文件夹）",
    "作者、单位、通讯邮箱和出版日期已由AIP页面核验；摘要逐字取自正式PDF，仅规范换行和单位字符；引用数为OUCI快照。",
    "https://doi.org/10.1063/5.0240095",
    null
  ],
  [
    5,
    "A comprehensive review of loop heat pipe: From fundamental researches to applications",
    "International Journal of Heat and Fluid Flow",
    new Date(2025, 8, 1),
    23,
    "Web of Science（用户提供截图）",
    new Date(2026, 7, 27),
    "否",
    "外文（英语）",
    "115",
    "",
    "109880",
    "",
    "文章号；期刊仅标示September 2025，日期按当月1日录入",
    "10.1016/j.ijheatfluidflow.2025.109880",
    "https://www.sciencedirect.com/science/article/abs/pii/S0142727X25001389",
    "Loop heat pipe, theoretical models, Heat transfer limit, Wick, Temperature oscillation",
    "With the rapid development of clean energy and artificial intelligence technologies, the thermal management of electronic devices in energy systems, data centers, and new energy vehicles has become a critical challenge. Efficient, compact, and reliable heat dissipation solutions are urgently needed to address the increasing power density and operational demands of these systems. Loop heat pipes (LHP), as a passive phase-change thermal management device, have shown great potential in thermal management of electronic devices both on the ground and in space due to their high efficiency, compactness, and reliability. LHP has attracted widespread global attention, thus, this paper reviews the research progress of loop heat pipes in terms of structural design, heat transfer limit, numerical simulation, and operational stability. The design of loop heat pipes needs to consider specific application environments and work mass selection in order to optimize their thermal performance. It has been shown that the heat transfer efficiency of loop heat pipes can be significantly improved by optimizing the design of the evaporator and condenser, selecting the appropriate component materials, and improving the wick structure. Despite the significant progress, there are still many challenges in the application of loop heat pipes, such as the improvement of heat transfer limits and the suppression of instability. Future research should focus on the application of novel materials, structural optimization, and stability analysis in complex environments, while combining artificial intelligence and big data technologies to achieve more accurate performance prediction and optimal design. This paper provides an important theoretical foundation and development direction for the research and application of loop heat pipes.",
    "本人负责文献调研、综述框架设计、资料整理及论文撰写。系统梳理环路热管结构、传热极限、模型、稳定性与应用进展，揭示温度振荡和启动失效机制，提出新材料、复杂环境可靠性及AI辅助优化方向。",
    "第1作者",
    "已附PDF（同文件夹）",
    "他引次数23采用用户提供的WoS截图（2026-08-27）；摘要逐字取自正式PDF，仅规范换行和单位字符；公开数据库数值会有差异。",
    "https://doi.org/10.1016/j.ijheatfluidflow.2025.109880",
    null
  ]
];

const authorHeaders = ["论文序号", "论文名称", "作者序号", "姓名", "作者属性", "本人", "工作单位", "电子邮箱", "核验备注", "来源URL"];
const XJTU = "西安交通大学化学工程与技术学院";
const XJTU_LAB = "西安交通大学化学工程与技术学院；含氟含氮化学品国家重点实验室";
const QILU = "中国科学院空天信息创新研究院齐鲁研究院";
const TBD = "待依据论文PDF核实";

const authorRows = [];
const addAuthors = (paperNo, title, names) => {
  names.forEach((a, idx) => authorRows.push([
    paperNo, title, idx + 1, a.name, a.role ?? "共同作者", a.me ? "是" : "否",
    a.affiliation ?? TBD, a.email ?? "", a.note ?? "", a.source ?? papers[paperNo - 1][22]
  ]));
};

addAuthors(1, papers[0][1], [
  { name: "Qingjie Cui（崔庆杰）", role: "第一作者", me: true, affiliation: XJTU, note: "署名顺序已核验" },
  { name: "Ziyi You（尤紫溢）", affiliation: XJTU },
  { name: "Xiang Ma", affiliation: XJTU },
  { name: "Xiaoping Yang（杨小平）", role: "通讯作者 / Lead contact", affiliation: XJTU, email: "yxping@xjtu.edu.cn", note: "出版社页面列为Lead contact" },
  { name: "Yonghai Zhang（张永海）", affiliation: XJTU },
  { name: "Jinjia Wei（魏进家）", affiliation: XJTU, email: "jjwei@xjtu.edu.cn" },
  { name: "Qie Sun", affiliation: TBD },
  { name: "Mu Du", affiliation: TBD },
  { name: "Feng Zhang", affiliation: TBD },
  { name: "Fatuan Li", affiliation: TBD }
]);

addAuthors(2, papers[1][1], [
  { name: "Qingjie Cui（崔庆杰）", role: "第一作者", me: true, affiliation: XJTU },
  { name: "Xiang Ma", affiliation: XJTU },
  { name: "Ziyi You（尤紫溢）", affiliation: XJTU },
  { name: "Xiaoping Yang（杨小平）", affiliation: XJTU, email: "yxping@xjtu.edu.cn", note: "通讯属性待PDF复核" },
  { name: "Yonghai Zhang（张永海）", affiliation: XJTU },
  { name: "Jinjia Wei（魏进家）", affiliation: XJTU, email: "jjwei@xjtu.edu.cn", note: "通讯属性待PDF复核" }
]);

addAuthors(3, papers[2][1], [
  { name: "Xiaoping Yang（杨小平）", role: "第一作者", affiliation: XJTU, email: "yxping@xjtu.edu.cn" },
  { name: "Qingjie Cui（崔庆杰）", role: "第二作者", me: true, affiliation: XJTU },
  { name: "Yao Zhou", affiliation: TBD },
  { name: "Gaoxiang Wang", affiliation: XJTU },
  { name: "Lei Liu", affiliation: TBD },
  { name: "Jinjia Wei（魏进家）", role: "通讯作者", affiliation: XJTU, email: "jjwei@xjtu.edu.cn", note: "通讯属性据公开作者资料核验" }
]);

addAuthors(4, papers[3][1], [
  { name: "Xiaoping Yang（杨小平）", role: "第一作者", affiliation: XJTU, email: "yxping@xjtu.edu.cn" },
  { name: "Qingjie Cui（崔庆杰）", role: "第二作者", me: true, affiliation: XJTU },
  { name: "Ziyi You（尤紫溢）", affiliation: XJTU },
  { name: "Jie Liu（刘杰）", affiliation: QILU },
  { name: "Yonghai Zhang（张永海）", affiliation: XJTU },
  { name: "Jinjia Wei（魏进家）", role: "通讯作者", affiliation: XJTU, email: "jjwei@xjtu.edu.cn", note: "AIP页面核验" }
]);

addAuthors(5, papers[4][1], [
  { name: "Qingjie Cui（崔庆杰）", role: "第一作者", me: true, affiliation: XJTU_LAB },
  { name: "Ziyi You（尤紫溢）", affiliation: XJTU_LAB },
  { name: "Yicheng Ni", affiliation: XJTU_LAB },
  { name: "Xiaoping Yang（杨小平）", affiliation: XJTU_LAB, email: "yxping@xjtu.edu.cn", note: "通讯属性待PDF复核" },
  { name: "Yao Zhou", affiliation: TBD },
  { name: "Jinjia Wei（魏进家）", affiliation: XJTU_LAB, email: "jjwei@xjtu.edu.cn", note: "通讯属性待PDF复核" },
  { name: "Jiping Liu", affiliation: TBD, note: "通讯属性待PDF复核" }
]);

function styleTitle(sheet, range, title) {
  range.merge();
  range.values = [[title]];
  range.format.fill = navy;
  range.format.font = { bold: true, color: "#FFFFFF", size: 16 };
  range.format.verticalAlignment = "center";
  range.format.horizontalAlignment = "left";
  range.format.rowHeight = 30;
}

function styleHeader(range) {
  range.format.fill = blue;
  range.format.font = { bold: true, color: "#FFFFFF", size: 10 };
  range.format.wrapText = true;
  range.format.verticalAlignment = "center";
  range.format.horizontalAlignment = "center";
  range.format.borders = { preset: "all", style: "thin", color: border };
  range.format.rowHeight = 38;
}

// 论文信息
papersSheet.showGridLines = false;
styleTitle(papersSheet, papersSheet.getRange("A1:X1"), "五篇环路热管论文信息（按申报字段整理）");
papersSheet.getRange("A2:X2").merge();
papersSheet.getRange("A2:X2").values = [["引用次数均为指定数据库在核验日期的快照；第5篇采用用户提供的WoS截图数值23。黄色单元格表示提交前仍需复核。"]];
papersSheet.getRange("A2:X2").format = { fill: paleGold, font: { color: "#7F6000", italic: true }, wrapText: true, verticalAlignment: "center" };
papersSheet.getRange("A2:X2").format.rowHeight = 30;
papersSheet.getRange("A3:X3").values = [paperHeaders];
styleHeader(papersSheet.getRange("A3:X3"));
papersSheet.getRange("A4:X8").values = papers;
papersSheet.getRange("X4").formulas = [["=COUNTIF('作者信息'!$A$4:$A$38,A4)"]];
papersSheet.getRange("X4:X8").fillDown();
papersSheet.getRange("A4:X8").format = {
  font: { color: text, size: 10 }, wrapText: true, verticalAlignment: "top",
  borders: { preset: "all", style: "thin", color: border }
};
papersSheet.getRange("A4:X8").format.rowHeight = 220;
papersSheet.getRange("A4:A8").format.horizontalAlignment = "center";
papersSheet.getRange("D4:D8").setNumberFormat("yyyy-mm-dd");
papersSheet.getRange("G4:G8").setNumberFormat("yyyy-mm-dd");
papersSheet.getRange("E4:E8").setNumberFormat("0");
papersSheet.getRange("X4:X8").setNumberFormat("0");
papersSheet.getRange("P4:P8").format.font = { color: "#0563C1", underline: true, size: 9 };
papersSheet.getRange("W4:W8").format.font = { color: "#0563C1", underline: true, size: 9 };
papersSheet.getRange("V4:V8").format.fill = paleGold;
papersSheet.getRange("H4:H8").dataValidation = { rule: { type: "list", values: ["是", "否"] } };
papersSheet.getRange("I4:I8").dataValidation = { rule: { type: "list", values: ["中文", "外文（英语）", "外文（其他）"] } };
papersSheet.freezePanes.freezeRows(3);
papersSheet.freezePanes.freezeColumns(2);
papersSheet.tables.add("A3:X8", true, "PapersTable");

const paperWidths = [7, 42, 28, 13, 10, 24, 14, 12, 14, 8, 8, 16, 12, 24, 28, 36, 36, 85, 42, 14, 22, 42, 34, 10];
paperWidths.forEach((w, i) => papersSheet.getRangeByIndexes(0, i, 8, 1).format.columnWidth = w);

// 作者信息
authorsSheet.showGridLines = false;
styleTitle(authorsSheet, authorsSheet.getRange("A1:J1"), "作者顺序、作者属性及工作单位");
authorsSheet.getRange("A2:J2").merge();
authorsSheet.getRange("A2:J2").values = [["“本人”按Qingjie Cui（崔庆杰）标记；黄色内容为尚未从公开页面完整确认、需用论文PDF最终复核的信息。"]];
authorsSheet.getRange("A2:J2").format = { fill: paleGold, font: { color: "#7F6000", italic: true }, wrapText: true };
authorsSheet.getRange("A2:J2").format.rowHeight = 28;
authorsSheet.getRange("A3:J3").values = [authorHeaders];
styleHeader(authorsSheet.getRange("A3:J3"));
authorsSheet.getRange(`A4:J${authorRows.length + 3}`).values = authorRows;
authorsSheet.getRange(`A4:J${authorRows.length + 3}`).format = {
  font: { color: text, size: 10 }, wrapText: true, verticalAlignment: "top",
  borders: { preset: "all", style: "thin", color: border }
};
authorsSheet.getRange(`A4:J${authorRows.length + 3}`).format.rowHeight = 42;
authorsSheet.getRange(`A4:A${authorRows.length + 3}`).format.horizontalAlignment = "center";
authorsSheet.getRange(`C4:C${authorRows.length + 3}`).format.horizontalAlignment = "center";
authorsSheet.getRange(`F4:F${authorRows.length + 3}`).format.horizontalAlignment = "center";
authorsSheet.getRange(`F4:F${authorRows.length + 3}`).conditionalFormats.add("cellIs", {
  operator: "equal", formula: "\"是\"", format: { fill: paleGreen, font: { bold: true, color: "#375623" } }
});
authorsSheet.getRange(`G4:G${authorRows.length + 3}`).conditionalFormats.add("containsText", {
  text: "待依据", format: { fill: paleGold, font: { color: "#7F6000" } }
});
authorsSheet.getRange(`H4:H${authorRows.length + 3}`).format.font = { color: "#0563C1", size: 10 };
authorsSheet.getRange(`J4:J${authorRows.length + 3}`).format.font = { color: "#0563C1", underline: true, size: 8 };
authorsSheet.freezePanes.freezeRows(3);
authorsSheet.freezePanes.freezeColumns(4);
authorsSheet.tables.add(`A3:J${authorRows.length + 3}`, true, "AuthorsTable");
const authorWidths = [10, 45, 10, 25, 24, 9, 43, 26, 34, 28];
authorWidths.forEach((w, i) => authorsSheet.getRangeByIndexes(0, i, authorRows.length + 3, 1).format.columnWidth = w);

// 填报说明
notesSheet.showGridLines = false;
styleTitle(notesSheet, notesSheet.getRange("A1:F1"), "填报说明与复核清单");
notesSheet.getRange("A3:B3").values = [["项目", "说明"]];
styleHeader(notesSheet.getRange("A3:B3"));
const notes = [
  ["引用次数口径", "第5篇为用户提供的Web of Science截图：23次，截图日期按2026-08-27记录。其余4篇为公开页面快照，来源和日期见“论文信息”表；正式提交前建议登录WoS逐篇统一复核。"],
  ["发表日期", "第5篇期刊仅标示September 2025，为便于Excel日期排序暂按2025-09-01录入；如申报系统只接受具体日期，请按正式论文或数据库记录调整。"],
  ["页码填写", "5篇论文均使用文章号。根据申报页面提示，可将文章号填入“起始页码”；“截止页码”留空，并在“页码说明”中标注。"],
  ["摘要与关键词", "摘要逐字提取自用户G盘中的5份正式论文PDF，仅对PDF断行、断词和单位字符进行规范；Device论文原文栏目名为SUMMARY。关键词优先采用出版社关键词，均控制在5项以内。"],
  ["本人贡献", "5篇均按“本人职责＋关键机理/创新＋量化结果”的结构填写，并控制在100字符以内；涉及具体分工的表述请本人按实际情况确认后提交。"],
  ["作者单位", "公开页面未完整展示的单位标为“待依据论文PDF核实”。请以上传PDF首页的作者上标、单位和通讯邮箱为最终依据。"],
  ["附件", "5份正式论文PDF已与工作簿放在同一目标文件夹；上传申报系统前请再次检查单个文件不超过15 MB。"],
  ["使用方式", "“论文信息”表一行对应一篇论文；“作者信息”表一行对应一名作者，可直接筛选论文序号后按作者序号录入申报系统。"]
];
notesSheet.getRange(`A4:B${notes.length + 3}`).values = notes;
notesSheet.getRange(`A4:B${notes.length + 3}`).format = {
  font: { color: text, size: 11 }, wrapText: true, verticalAlignment: "top",
  borders: { preset: "all", style: "thin", color: border }
};
notesSheet.getRange(`A4:A${notes.length + 3}`).format = { fill: paleBlue, font: { bold: true, color: navy }, wrapText: true, verticalAlignment: "top", borders: { preset: "all", style: "thin", color: border } };
notesSheet.getRange(`A4:B${notes.length + 3}`).format.rowHeight = 56;
notesSheet.getRange("A13:F13").merge();
notesSheet.getRange("A13:F13").values = [["颜色提示：绿色 = 已标记本人；黄色 = 需复核；蓝色链接 = 可用于查验题录。"]];
notesSheet.getRange("A13:F13").format = { fill: paleGreen, font: { color: "#375623", bold: true }, wrapText: true };
notesSheet.getRange("A13:F13").format.rowHeight = 28;
notesSheet.getRange("A:A").format.columnWidth = 18;
notesSheet.getRange("B:B").format.columnWidth = 100;
notesSheet.getRange("C:F").format.columnWidth = 12;
notesSheet.freezePanes.freezeRows(3);

// Workbook verification before export
const paperInspect = await workbook.inspect({ kind: "region", sheetId: "论文信息", range: "A1:X8", maxChars: 12000 });
const authorInspect = await workbook.inspect({ kind: "region", sheetId: "作者信息", range: "A1:J12", maxChars: 8000 });
const formulaInspect = await workbook.inspect({ kind: "formula", sheetId: "论文信息", range: "X4:X8", maxChars: 3000 });
console.log("PAPER_INSPECT", paperInspect.ndjson ?? paperInspect);
console.log("AUTHOR_INSPECT", authorInspect.ndjson ?? authorInspect);
console.log("FORMULA_INSPECT", formulaInspect.ndjson ?? formulaInspect);

const errorScan = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  maxChars: 4000
});
console.log("ERROR_SCAN", errorScan.ndjson ?? errorScan);

for (const sheetName of ["论文信息", "作者信息", "填报说明"]) {
  const preview = await workbook.render({ sheetName, autoCrop: "all", scale: 0.8, format: "png" });
  const bytes = new Uint8Array(await preview.arrayBuffer());
  await fs.writeFile(path.join(outputDir, `${sheetName}.png`), bytes);
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(`OUTPUT=${outputPath}`);
