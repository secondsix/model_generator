import fs from "node:fs/promises";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const files = process.argv.slice(2);
for (const path of files) {
  const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(path));
  const summary = await wb.inspect({
    kind: "workbook,sheet,table,match",
    searchTerm: "数据类型配置",
    maxChars: 12000,
    tableMaxRows: 12,
    tableMaxCols: 20,
    options: { maxResults: 50 },
  });
  console.log(`FILE:${path}`);
  console.log(summary.ndjson);
  const sheets = await wb.inspect({ kind: "sheet", include: "id,name", maxChars: 5000 });
  console.log(sheets.ndjson);
}
