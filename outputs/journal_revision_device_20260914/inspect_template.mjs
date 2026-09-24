import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const inputPath = String.raw`C:\Users\32933\xwechat_files\wxid_bbyf1p55y50a22_da86\msg\file\2026-09\期刊目录修订表（2026年修订）-传递所.xlsx`;
const outputDir = String.raw`C:\Users\32933\Documents\ChatGPT\论文\outputs\journal_revision_device_20260914`;
await fs.mkdir(outputDir, { recursive: true });
const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(inputPath));
console.log((await workbook.inspect({ kind: "workbook,sheet,table", maxChars: 12000, tableMaxRows: 20, tableMaxCols: 20, tableMaxCellChars: 220 })).ndjson);
for (let i = 0; i < workbook.worksheets.items.length; i++) {
  const sheet = workbook.worksheets.getItemAt(i);
  console.log(`SHEET_${i}=${sheet.name}`);
  const used = sheet.getUsedRange();
  if (used) {
    console.log((await workbook.inspect({ kind: "region", sheetId: sheet.name, range: used.address, maxChars: 16000 })).ndjson);
  }
  const preview = await workbook.render({ sheetName: sheet.name, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(`${outputDir}/before_${i}_${sheet.name.replaceAll(/[\\/:*?"<>|]/g, "_")}.png`, new Uint8Array(await preview.arrayBuffer()));
}
