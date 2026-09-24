import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const source = "C:/Users/32933/Downloads/savedrecs.xls";
const input = await FileBlob.load(source);
const workbook = await SpreadsheetFile.importXlsx(input);

const summary = await workbook.inspect({
  kind: "workbook,sheet,table",
  maxChars: 20000,
  tableMaxRows: 30,
  tableMaxCols: 40,
  tableMaxCellChars: 300,
});

process.stdout.write(summary.ndjson);
