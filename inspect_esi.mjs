import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const path = "C:/Users/32933/Documents/ChatGPT/论文/2026-07-ESI-highly-cited-thresholds.xlsx";
const input = await FileBlob.load(path);
const workbook = await SpreadsheetFile.importXlsx(input);

const sheets = await workbook.inspect({
  kind: "sheet",
  include: "id,name",
  maxChars: 8000,
});
console.log("SHEETS");
console.log(sheets.ndjson);

for (const term of ["ENGINEERING", "Engineering", "2025", "109880", "Cui Qingjie"]) {
  const result = await workbook.inspect({
    kind: "match",
    searchTerm: term,
    options: { useRegex: false, maxResults: 100 },
    maxChars: 12000,
    summary: `matches for ${term}`,
  });
  console.log(`MATCH ${term}`);
  console.log(result.ndjson);
}
